"use client";

import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";
import {
  detectDetectionsFromDocx,
  detectFieldsFromDocx,
  type DetectedField,
  type Field
} from "./docxFields";

type LLMField = {
  field_id: string;
  canonical_id: string;
  raw_placeholder: string;
  label_for_user: string;
  who_should_fill: "company" | "counterparty" | "either" | "system";
  expected_type:
    | "string"
    | "date"
    | "money"
    | "state"
    | "email"
    | "name"
    | "title"
    | "address"
    | "jurisdiction"
    | "other";
  how_to_fill: string;
  confidence: number;
  evidence_quote: string;
  location_hint: string;
  ask_user: boolean;
};

type ChatField = {
  key: string; // canonicalized field id
  label: string; // label_for_user
  question: string;
  group: string;
  who?: LLMField["who_should_fill"];
  expected_type?: LLMField["expected_type"];
  ask_user?: boolean;
};

type Occurrence = {
  raw_placeholder: string;
  fieldKey: string; // key used to look up value (canonical_id)
  location_hint: string;
};

// test.py-equivalent patterns
const BRACKET_TOKEN_RE = /\[[^\[\]\n]{1,120}\]/g; // e.g. [Company Name]
// Match only the bracket portion so leading "$" isn't highlighted (e.g. "$[_____]" -> highlight "[_____]")
const BRACKET_BLANK_RE = /\[[\s_]{3,}\]/g; // e.g. [_____]
const UNDERLINE_RE = /_{2,}/g; // e.g. __ or ________

function shouldHighlightToken(token: string) {
  // Avoid highlighting overly-generic [name]/[title] tokens? test.py includes them, so keep them.
  return Boolean(token && token.length >= 2);
}

function isBlankLike(text: string) {
  // Treat whitespace + common invisible chars + line/underscore glyphs as "blank".
  const t = (text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b\u200c\u200d\u2060\ufeff]/g, "") // zero-width/invisible
    .trim();
  if (!t) return true;
  // underscore and common horizontal line glyphs (hyphens/dashes/box drawing)
  return /^[\s_‐‑‒–—―─━]+$/.test(t);
}

function ensureHighlightInRun(runXml: string) {
  if (runXml.includes("<w:highlight")) return runXml;
  if (runXml.includes("<w:rPr")) {
    return runXml.replace(
      /<w:rPr>/,
      `<w:rPr><w:highlight w:val="yellow"/>`
    );
  }
  return runXml.replace(
    /<w:r>/,
    `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>`
  );
}

function xmlEscape(text: string) {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function stripHighlightFromRun(runXml: string) {
  return runXml.replace(/<w:highlight\b[^>]*\/>/g, "").replace(/<w:highlight\b[^>]*>[\s\S]*?<\/w:highlight>/g, "");
}

function fillUnderscoreTextInRun(runXml: string, value: string) {
  const esc = xmlEscape(value);
  // Replace the first underscore text node
  return stripHighlightFromRun(runXml).replace(
    /<w:t[^>]*>_{2,}<\/w:t>/,
    `<w:t xml:space="preserve">${esc}</w:t>`
  );
}

function replaceSplitBracketToken(xml: string, token: string, value: string) {
  // token like "[COMPANY]" might be split across three runs: "[", "COMPANY", "]"
  const inner = token.slice(1, -1);
  const esc = xmlEscape(value);
  const re = new RegExp(
    `<w:r\\b[\\s\\S]*?<w:t[^>]*>\\[<\\/w:t>[\\s\\S]*?<\\/w:r>\\s*` +
      `(<w:r\\b[\\s\\S]*?<w:t[^>]*>${inner}<\\/w:t>[\\s\\S]*?<\\/w:r>)\\s*` +
      `<w:r\\b[\\s\\S]*?<w:t[^>]*>\\]<\\/w:t>[\\s\\S]*?<\\/w:r>`,
    "g"
  );
  return xml.replace(re, (_m, middleRun: string) => {
    const rPrMatch = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(middleRun);
    const rPr =
      rPrMatch?.[0]
        ?.replace(/<w:caps\/>\s*/g, "")
        ?.replace(/<w:smallCaps\/>\s*/g, "") ?? "";
    return `<w:r>${rPr}<w:t xml:space="preserve">${esc}</w:t></w:r>`;
  });
}

function stripCapsSmallCaps(runXml: string) {
  return runXml
    .replace(/<w:caps\/>\s*/g, "")
    .replace(/<w:smallCaps\/>\s*/g, "");
}

function replaceTokenInRuns(xml: string, raw: string, value: string) {
  const esc = xmlEscape(value);
  const tNeedle = raw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const re = new RegExp(
    `<w:r\\b[\\s\\S]*?<w:t[^>]*>${tNeedle}<\\/w:t>[\\s\\S]*?<\\/w:r>`,
    "g"
  );
  return xml.replace(re, (run) => {
    const cleaned = stripCapsSmallCaps(run);
    return cleaned.replace(
      new RegExp(`<w:t([^>]*)>${tNeedle}<\\/w:t>`),
      `<w:t$1>${esc}</w:t>`
    );
  });
}

async function patchDocxArrayBufferForPreview(
  arrayBuffer: ArrayBuffer,
  values: Record<string, string>,
  occurrences: Occurrence[]
) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  // IMPORTANT: only patch word/document.xml. Regex patching can produce malformed XML and cause docx-preview to render only the tail.
  const file = zip.file("word/document.xml");
  if (!file) return arrayBuffer;
  const xml = await file.async("string");

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  const parseErr = doc.getElementsByTagName("parsererror")[0];
  if (parseErr) return arrayBuffer;

  const nsW = doc.documentElement.lookupNamespaceURI("w") || doc.documentElement.namespaceURI || "";
  const nsXml = "http://www.w3.org/XML/1998/namespace";

  const getChildrenByLocal = (parent: Element, local: string) =>
    Array.from(parent.childNodes).filter(
      (n): n is Element => n.nodeType === 1 && (n as Element).localName === local
    );

  const ensureRPr = (run: Element) => {
    const rPr = getChildrenByLocal(run, "rPr")[0];
    if (rPr) return rPr;
    const created = nsW ? doc.createElementNS(nsW, "w:rPr") : doc.createElement("w:rPr");
    run.insertBefore(created, run.firstChild);
    return created;
  };

  const ensureHighlightDom = (run: Element) => {
    const rPr = ensureRPr(run);
    const has = getChildrenByLocal(rPr, "highlight")[0];
    if (has) return;
    const hl = nsW ? doc.createElementNS(nsW, "w:highlight") : doc.createElement("w:highlight");
    hl.setAttributeNS(nsW || null, "w:val", "yellow");
    rPr.appendChild(hl);
  };

  const stripCapsSmallCapsDom = (run: Element) => {
    const rPr = getChildrenByLocal(run, "rPr")[0];
    if (!rPr) return;
    for (const el of Array.from(rPr.childNodes)) {
      if (el.nodeType !== 1) continue;
      const e = el as Element;
      if (e.localName === "caps" || e.localName === "smallCaps") rPr.removeChild(e);
    }
  };

  // 1) Make signature blanks visible: underlined runs with only <w:tab/> become "________" with highlight.
  const runs = Array.from(doc.getElementsByTagNameNS("*", "r"));
  for (const run of runs) {
    const hasTab = run.getElementsByTagNameNS("*", "tab").length > 0;
    if (!hasTab) continue;
    const hasText = run.getElementsByTagNameNS("*", "t").length > 0;
    if (hasText) continue;
    const rPr = run.getElementsByTagNameNS("*", "rPr")[0] as Element | undefined;
    const underlined = Boolean(rPr && rPr.getElementsByTagNameNS("*", "u").length > 0);
    if (!underlined) continue;

    // remove tabs
    for (const tab of Array.from(run.getElementsByTagNameNS("*", "tab"))) {
      tab.parentNode?.removeChild(tab);
    }

    ensureHighlightDom(run);
    const t = nsW ? doc.createElementNS(nsW, "w:t") : doc.createElement("w:t");
    t.setAttributeNS(nsXml, "xml:space", "preserve");
    t.textContent = "________";
    run.appendChild(t);
  }

  const parseParaIndex = (hint: string) => {
    const m = /paragraph:(\d+)/.exec(hint || "");
    return m ? Number(m[1]) : null;
  };

  // Apply replacements BY LOCATION (paragraph) so identical placeholders like [_____________] don't overwrite other fields.
  const occByPara = new Map<number, Array<{ raw: string; key: string }>>();
  for (const occ of occurrences) {
    const pi = parseParaIndex(occ.location_hint);
    if (pi === null) continue;
    const arr = occByPara.get(pi) || [];
    arr.push({ raw: occ.raw_placeholder, key: occ.fieldKey });
    occByPara.set(pi, arr);
  }

  // Per paragraph, attempt to match placeholders even if split across multiple <w:t>.
  const paragraphs = Array.from(doc.getElementsByTagNameNS("*", "p"));
  const allRaw = Array.from(occurrences.map((o) => o.raw_placeholder)).filter(Boolean);
  const maxRawLen = allRaw.reduce((m, s) => Math.max(m, s.length), 0);

  const getTextNodesInPara = (p: Element) =>
    Array.from(p.getElementsByTagNameNS("*", "t")) as Element[];

  const norm = (s: string) =>
    (s || "")
      .replace(/\u00a0/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  for (let pIdx = 0; pIdx < paragraphs.length; pIdx++) {
    const p = paragraphs[pIdx]!;
    const tNodes = getTextNodesInPara(p);
    if (!tNodes.length) continue;
    const occList = (occByPara.get(pIdx) || []).slice();
    if (!occList.length) continue;

    // Sliding window matching across consecutive <w:t> nodes.
    for (let i = 0; i < tNodes.length; i++) {
      let acc = "";
      const idxs: number[] = [];
      for (let j = i; j < tNodes.length; j++) {
        const txt = tNodes[j]!.textContent || "";
        acc += txt;
        idxs.push(j);
        if (acc.length > maxRawLen + 2) break;

        // Find an occurrence in this paragraph whose raw_placeholder matches this acc (with $-flex).
        let matchIdx = -1;
        let matchKey: string | null = null;
        let includeDollar = false;
        for (let oi = 0; oi < occList.length; oi++) {
          const raw = occList[oi]!.raw || "";
          if (!raw) continue;
          const accN = norm(acc);
          const rawN = norm(raw);
          if (accN === rawN) {
            matchIdx = oi;
            matchKey = occList[oi]!.key;
            includeDollar = accN.startsWith("$");
            break;
          }
          // raw is "[___]" but acc is "$[___]"
          if (!rawN.startsWith("$") && accN.startsWith("$") && accN.slice(1) === rawN) {
            matchIdx = oi;
            matchKey = occList[oi]!.key;
            includeDollar = true;
            break;
          }
          // raw is "$[___]" but acc is "[___]"
          if (rawN.startsWith("$") && !accN.startsWith("$") && rawN.slice(1) === accN) {
            matchIdx = oi;
            matchKey = occList[oi]!.key;
            includeDollar = false;
            break;
          }
        }

        if (matchIdx !== -1 && matchKey) {
          const v = (values[matchKey] || "").trim();
          if (v) {
            // Preserve any leading/trailing whitespace that was part of the matched token span
            // (e.g. templates often have ", [Company Name]" where the leading space can be inside the same run).
            const leadWs = /^[\s\u00a0]+/.exec(acc)?.[0] ?? "";
            const trailWs = /[\s\u00a0]+$/.exec(acc)?.[0] ?? "";
            const replacement = `${leadWs}${includeDollar ? `$${v}` : v}${trailWs}`;
            tNodes[i]!.textContent = replacement;
            for (let k = 1; k < idxs.length; k++) tNodes[idxs[k]!]!.textContent = "";
            const run = tNodes[i]!.closest("w\\:r, r") as Element | null;
            if (run) stripCapsSmallCapsDom(run);
          }
          // consume this occurrence so we don't apply it again
          occList.splice(matchIdx, 1);
          break;
        }

        // Handle bracket blank where occurrences give "$[___]" but in XML it's split into "$" + "[___]"
        if (acc === "$") continue;
      }
    }

    // No global raw matching fallback; we intentionally avoid interconnectedness.
  }

  const serializer = new XMLSerializer();
  const outXml = serializer.serializeToString(doc);
  zip.file("word/document.xml", outXml);

  const out = await zip.generateAsync({ type: "arraybuffer" });
  return out;
}

function collectTextNodes(container: HTMLElement) {
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT, {
    acceptNode(n: Node) {
      const p = (n.parentNode as HTMLElement | null);
      if (!p) return NodeFilter.FILTER_REJECT;
      const tag = p.tagName?.toLowerCase?.() || "";
      if (tag === "script" || tag === "style") return NodeFilter.FILTER_REJECT;
      if (p.classList?.contains("phHl")) return NodeFilter.FILTER_REJECT;
      if (p.classList?.contains("phHlBlank")) return NodeFilter.FILTER_REJECT;
      if (!n.nodeValue) return NodeFilter.FILTER_REJECT;
      return NodeFilter.FILTER_ACCEPT;
    }
  } as any);

  const nodes: Text[] = [];
  let cur: Node | null;
  while ((cur = walker.nextNode())) nodes.push(cur as Text);
  return nodes;
}

function highlightPatternAcrossTextNodes(container: HTMLElement, re: RegExp) {
  const nodes = collectTextNodes(container);
  if (!nodes.length) return;

  const parts: Array<{ node: Text; start: number; end: number }> = [];
  let cursor = 0;
  let full = "";
  for (const n of nodes) {
    const t = n.nodeValue || "";
    parts.push({ node: n, start: cursor, end: cursor + t.length });
    cursor += t.length;
    full += t;
  }

  re.lastIndex = 0;
  const matches: Array<{ start: number; end: number; value: string }> = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(full)) !== null) {
    const value = m[0] || "";
    if (!shouldHighlightToken(value)) continue;
    matches.push({ start: m.index, end: m.index + value.length, value });
  }
  if (!matches.length) return;

  const findPartIndex = (pos: number) => {
    // linear is fine for our doc sizes
    for (let i = 0; i < parts.length; i++) {
      if (pos >= parts[i]!.start && pos < parts[i]!.end) return i;
    }
    return parts.length - 1;
  };

  // Process from end -> start so indices remain valid.
  for (let i = matches.length - 1; i >= 0; i--) {
    const mm = matches[i]!;
    const startIdx = findPartIndex(mm.start);
    const endIdx = findPartIndex(Math.max(mm.end - 1, mm.start));
    const startPart = parts[startIdx]!;
    const endPart = parts[endIdx]!;

    const startOffset = mm.start - startPart.start;
    const endOffset = mm.end - endPart.start;

    // Guard: if nodes were mutated by earlier operations, skip safely.
    if (!startPart.node.parentNode || !endPart.node.parentNode) continue;
    if (startOffset < 0 || endOffset < 0) continue;

    const range = document.createRange();
    try {
      range.setStart(startPart.node, startOffset);
      range.setEnd(endPart.node, endOffset);
    } catch {
      continue;
    }

    const span = document.createElement("span");
    span.className = "phHl";
    span.appendChild(range.extractContents());
    range.insertNode(span);
    range.detach();
  }
}

function highlightTextNode(node: Text, re: RegExp) {
  const text = node.nodeValue || "";
  re.lastIndex = 0;
  const matches: Array<{ start: number; end: number; value: string }> = [];
  let m: RegExpExecArray | null;
  while ((m = re.exec(text)) !== null) {
    const value = m[0] || "";
    if (!shouldHighlightToken(value)) continue;
    matches.push({ start: m.index, end: m.index + value.length, value });
  }
  if (!matches.length) return;

  const frag = document.createDocumentFragment();
  let last = 0;
  for (const mm of matches) {
    if (mm.start > last) frag.appendChild(document.createTextNode(text.slice(last, mm.start)));
    const span = document.createElement("span");
    span.className = "phHl";
    span.textContent = text.slice(mm.start, mm.end);
    frag.appendChild(span);
    last = mm.end;
  }
  if (last < text.length) frag.appendChild(document.createTextNode(text.slice(last)));
  node.parentNode?.replaceChild(frag, node);
}

function highlightUnderlineLikeBlanks(container: HTMLElement) {
  const els = Array.from(container.querySelectorAll<HTMLElement>("*"));
  for (const el of els) {
    if (el.classList.contains("phHl") || el.classList.contains("phHlBlank")) continue;
    const cs = window.getComputedStyle(el);
    const textDecor = (cs.textDecorationLine || "").toLowerCase();
    const borderBottomWidth = parseFloat(cs.borderBottomWidth || "0");
    const borderBottomStyle = (cs.borderBottomStyle || "").toLowerCase();
    const bgImg = (cs.backgroundImage || "").toLowerCase();
    const looksUnderlined =
      el.tagName.toLowerCase() === "u" ||
      textDecor.includes("underline") ||
      (borderBottomStyle !== "none" && borderBottomWidth > 0) ||
      // Some docx HTML renderers use gradient backgrounds as "lines"
      (bgImg !== "none" &&
        (bgImg.includes("linear-gradient") || bgImg.includes("repeating-linear-gradient")));
    if (!looksUnderlined) continue;

    const rawText = el.textContent || "";
    if (!isBlankLike(rawText)) continue;

    // Avoid false positives like tiny underlined whitespace between words (e.g., in underlined headings).
    // Only highlight if the visual blank is "line-like" (wide enough), OR it's an underscore run.
    const rect = el.getBoundingClientRect();
    const hasUnderscores = /_/.test(rawText);
    if (!hasUnderscores && rect.width < 30) continue;
    if (hasUnderscores && rect.width < 8) continue;

    // Highlight the underlined element itself.
    el.classList.add("phHlBlank");
    // Ensure the highlight has visible width even if empty.
    if (!rawText.trim()) el.innerHTML = "&nbsp;";
  }
}

function highlightUnderlinedWhitespaceSegments(container: HTMLElement) {
  const nodes = collectTextNodes(container);
  for (const n of nodes) {
    const parent = n.parentNode as HTMLElement | null;
    if (!parent) continue;

    const cs = window.getComputedStyle(parent);
    const textDecor = (cs.textDecorationLine || "").toLowerCase();
    const borderBottomWidth = parseFloat(cs.borderBottomWidth || "0");
    const borderBottomStyle = (cs.borderBottomStyle || "").toLowerCase();
    const isUnderlinedContext =
      textDecor.includes("underline") ||
      (borderBottomStyle !== "none" && borderBottomWidth > 0);
    if (!isUnderlinedContext) continue;

    const text = n.nodeValue || "";
    // Underlined "blanks" are typically sequences of spaces/NBSP in DOCX.
    // Keep threshold low enough to catch these lines, but high enough to avoid normal word spacing.
    const re = /[ \u00a0]{3,}/g;
    // If the whole node is whitespace and underlined, highlight it.
    if (!text.trim() && text.replace(/\u00a0/g, " ").length >= 3) {
      const span = document.createElement("span");
      span.className = "phHlBlank";
      span.textContent = text.replace(/\u00a0/g, " ");
      n.parentNode?.replaceChild(span, n);
      continue;
    }

    // Otherwise, highlight only the whitespace segments.
    highlightTextNode(n, re);
    // highlightTextNode wraps matches with .phHl; convert those highlights to blank highlights
    // when they are purely whitespace.
    const hlSpans = Array.from(parent.querySelectorAll<HTMLElement>("span.phHl"));
    for (const s of hlSpans) {
      if ((s.textContent || "").trim() === "") {
        s.classList.remove("phHl");
        s.classList.add("phHlBlank");
        s.innerHTML = "&nbsp;"; // preserve width
      }
    }
  }
}

function highlightPlaceholdersInContainer(container: HTMLElement) {
  // Robust highlighter: supports tokens split across multiple text nodes (common in DOCX renderers).
  highlightPatternAcrossTextNodes(container, BRACKET_BLANK_RE);
  highlightPatternAcrossTextNodes(container, BRACKET_TOKEN_RE);
  highlightPatternAcrossTextNodes(container, UNDERLINE_RE);

  // Note: underline-only blanks are now made visible/highlightable by patching DOCX tabs into underscores.
}

export default function HomePage() {
  const [file, setFile] = useState<File | null>(null);
  const [filename, setFilename] = useState<string | null>(null);
  const [previewType, setPreviewType] = useState<"docx" | "pdf" | null>(null);
  const [pdfUrl, setPdfUrl] = useState<string | null>(null);
  const docxContainerRef = useRef<HTMLDivElement | null>(null);
  const originalDocxRef = useRef<ArrayBuffer | null>(null);

  const [fields, setFields] = useState<ChatField[]>([]);
  const [classified, setClassified] = useState<LLMField[]>([]);
  const [occurrences, setOccurrences] = useState<Occurrence[]>([]);
  const [values, setValues] = useState<Record<string, string>>({});
  const [draftValues, setDraftValues] = useState<Record<string, string>>({});
  const [skipped, setSkipped] = useState<Set<string>>(new Set());
  const [chat, setChat] = useState<Array<{ role: "assistant" | "user"; content: string }>>([]);
  const [currentKey, setCurrentKey] = useState<string | null>(null);
  const [input, setInput] = useState("");
  const [loadingFields, setLoadingFields] = useState(false);
  const [fieldsError, setFieldsError] = useState<string | null>(null);
  const hasUnappliedDrafts = Object.entries(draftValues).some(
    ([k, v]) => (values[k] ?? "") !== v
  );
  const [jumpToTopOnUpdate, setJumpToTopOnUpdate] = useState(false);

  const nextAskableKey = (vals: Record<string, string>, skippedSet: Set<string>) => {
    for (const f of fields) {
      if (skippedSet.has(f.key)) continue;
      if (!vals[f.key] || !vals[f.key]!.trim()) return f.key;
    }
    return null;
  };

  const prettyLabel = (f: ChatField) => {
    // Prefer human label from detector; fallback to key.
    const base = (f.label || f.key || "").toString();
    // Collapse underline-only labels into "Blank"
    const cleaned = base.replace(/_+/g, "_").trim();
    if (/^_+$/.test(cleaned)) return "Blank";
    // Title-case-ish for readability
    return cleaned
      .replace(/_/g, " ")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/\b\w/g, (c) => c.toUpperCase());
  };

  const groupName = (f: ChatField) => f.group || "Core terms";

  const groupedFields = (() => {
    const groups = new Map<string, ChatField[]>();
    for (const f of fields) {
      const g = groupName(f);
      const arr = groups.get(g) || [];
      arr.push(f);
      groups.set(g, arr);
    }
    // Stable-ish ordering
    return Array.from(groups.entries()).sort(([a], [b]) => a.localeCompare(b));
  })();

  useEffect(() => {
    return () => {
      if (pdfUrl) URL.revokeObjectURL(pdfUrl);
    };
  }, [pdfUrl]);

  useEffect(() => {
    async function renderDocx() {
      if (!file || previewType !== "docx") return;
      const el = docxContainerRef.current;
      if (!el) return;
      const prevScrollTop = el.scrollTop;
      const prevScrollLeft = el.scrollLeft;
      el.innerHTML = "";

      const { renderAsync } = await import("docx-preview");
      const original = originalDocxRef.current ?? (await file.arrayBuffer());
      originalDocxRef.current = original;
      const arrayBuffer = await patchDocxArrayBufferForPreview(original, values, occurrences);

      await renderAsync(arrayBuffer, el, undefined, {
        // Keep fidelity: do not intentionally normalize or alter content.
        // docx-preview renders the DOCX structure/styles into HTML.
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true
      });

      // After render, highlight placeholders using test.py-like detection rules.
      // This does NOT change the underlying content; it only wraps matching text in spans.
      highlightPlaceholdersInContainer(el);

      // docx-preview can keep laying out after renderAsync resolves; restore scroll after layout settles.
      const targetTop = jumpToTopOnUpdate ? 0 : prevScrollTop;
      const targetLeft = jumpToTopOnUpdate ? 0 : prevScrollLeft;
      requestAnimationFrame(() => {
        requestAnimationFrame(() => {
          el.scrollTop = targetTop;
          el.scrollLeft = targetLeft;
          // One extra micro-delay to handle late layout shifts (fonts/images).
          window.setTimeout(() => {
            el.scrollTop = targetTop;
            el.scrollLeft = targetLeft;
          }, 0);
        });
      });
    }

    void renderDocx();
  }, [file, previewType, values, occurrences, jumpToTopOnUpdate]);

  // When a DOCX is uploaded, detect fields and start the chat.
  useEffect(() => {
    async function init() {
      if (!file || previewType !== "docx") return;
      const ab = await file.arrayBuffer();
      originalDocxRef.current = ab;
      setFields([]);
      setClassified([]);
      setOccurrences([]);
      setValues({});
      setDraftValues({});
      setSkipped(new Set());
      setChat([]);
      setCurrentKey(null);
      setLoadingFields(true);
      setFieldsError(null);

      let fallback: Field[] = [];
      try {
        fallback = await detectFieldsFromDocx(ab);
      } catch {
        fallback = [];
      }

      // fields.json-like classification (server-side OpenAI)
      try {
        const detections: DetectedField[] = await detectDetectionsFromDocx(ab, 1);
        const res = await fetch("/api/fields", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ detections })
        });
        if (res.ok) {
          const data: unknown = await res.json();
          const list = Array.isArray(data) ? (data as LLMField[]) : [];
          setClassified(list);

          const occ: Occurrence[] = list.map((f) => ({
            raw_placeholder: f.raw_placeholder,
            fieldKey: f.canonical_id || f.field_id,
            location_hint: f.location_hint
          }));
          setOccurrences(occ);

          // Ask once per canonical_id (semantic propagation).
          const byKey = new Map<string, ChatField>();
          const firstLoc = new Map<string, number>();
          for (const f of list) {
            const key = f.canonical_id || f.field_id;
            const existing = byKey.get(key);
            const label = f.label_for_user || key;
            const q =
              f.how_to_fill && f.how_to_fill.trim().length
                ? f.how_to_fill.trim().endsWith("?")
                  ? f.how_to_fill.trim()
                  : `Please provide: ${label}. ${f.how_to_fill.trim()}`
                : `Please provide: ${label}.`;

            const m = /paragraph:(\d+)/.exec(f.location_hint || "");
            const locNum = m ? Number(m[1]) : 1e9;
            if (!firstLoc.has(key) || locNum < (firstLoc.get(key) ?? 1e9)) firstLoc.set(key, locNum);

            if (!existing || f.confidence > (existing as any)._confidence) {
              byKey.set(key, {
                key,
                label,
                question: q,
                group:
                  f.who_should_fill === "company"
                    ? "Company"
                    : f.who_should_fill === "counterparty"
                      ? "Counterparty"
                      : f.who_should_fill === "either"
                        ? "Either party"
                        : "System",
                who: f.who_should_fill,
                expected_type: f.expected_type,
                ask_user: f.ask_user
              } as ChatField & { _confidence?: number });
              (byKey.get(key) as any)._confidence = f.confidence;
            } else if (existing) {
              if (f.ask_user) existing.ask_user = true;
            }
          }

          const deduped = Array.from(byKey.values()).sort((a, b) => {
            return (firstLoc.get(a.key) ?? 1e9) - (firstLoc.get(b.key) ?? 1e9);
          });
          setFields(deduped);
          const firstKey = deduped[0]?.key ?? null;
          setCurrentKey(firstKey);
          if (firstKey) {
            setChat([
              {
                role: "assistant",
                content: deduped.find((x) => x.key === firstKey)!.question
              }
            ]);
          }
        } else {
          let msg = `Field classification failed (${res.status}).`;
          try {
            const j: any = await res.json();
            if (j?.error) msg = String(j.error);
          } catch {
            // ignore
          }
          setFieldsError(msg);
          setClassified([]);
          setOccurrences([]);

          const fb: ChatField[] = fallback.map((f) => ({
            key: f.key,
            label: f.label,
            question: f.question,
            group: "Detected (no AI labeling)",
            ask_user: true
          }));
          setFields(fb);
          const firstKey = fb[0]?.key ?? null;
          setCurrentKey(firstKey);
          if (firstKey) setChat([{ role: "assistant", content: fb[0]!.question }]);
        }
      } catch {
        setFieldsError("Error calling /api/fields. (Check OPENAI_API_KEY)");
        setClassified([]);
        setOccurrences([]);

        const fb: ChatField[] = fallback.map((f) => ({
          key: f.key,
          label: f.label,
          question: f.question,
          group: "Detected (no AI labeling)",
          ask_user: true
        }));
        setFields(fb);
        const firstKey = fb[0]?.key ?? null;
        setCurrentKey(firstKey);
        if (firstKey) setChat([{ role: "assistant", content: fb[0]!.question }]);
      }
      setLoadingFields(false);
    }
    void init();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [file, previewType]);

  return (
    <div className="grid" style={{ gridTemplateColumns: "1fr" }}>
      <section className="card">
        <div className="cardHeader">
          <div>
            <div className="h">Upload & preview</div>
            <div className="small">
              Supports <code>.docx</code> and <code>.pdf</code>. Preview is rendered without
              placeholder insertion or text normalization.
            </div>
          </div>
          <span className="pill">{filename ? "Loaded" : "No document yet"}</span>
        </div>
        <div className="cardBody">
          <div className="row">
            <input
              className="input"
              type="file"
              accept=".docx,.pdf,application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
              onChange={(e) => {
                const f = e.target.files?.[0] || null;
                setFile(f);
                setFilename(null);
                setPreviewType(null);
                if (pdfUrl) URL.revokeObjectURL(pdfUrl);
                setPdfUrl(null);
                setFields([]);
                setClassified([]);
                setOccurrences([]);
                setValues({});
                setDraftValues({});
                setSkipped(new Set());
                setChat([]);
                setCurrentKey(null);
                setInput("");
                setLoadingFields(false);
                setFieldsError(null);
                originalDocxRef.current = null;

                if (!f) return;
                setFilename(f.name);
                const ext = (f.name || "").toLowerCase().split(".").pop();
                if (ext === "pdf") {
                  setPreviewType("pdf");
                  setPdfUrl(URL.createObjectURL(f));
                } else if (ext === "docx") {
                  setPreviewType("docx");
                } else {
                  setPreviewType(null);
                }
              }}
            />
          </div>

          {filename ? (
            <div style={{ marginTop: 12 }} className="small">
              Loaded: <b>{filename}</b>
            </div>
          ) : null}

          <div style={{ marginTop: 14 }}>
            {!previewType ? (
              <div className="small">Upload a document to preview it here.</div>
            ) : previewType === "pdf" && pdfUrl ? (
              <iframe
                title="PDF preview"
                src={pdfUrl}
                style={{
                  width: "100%",
                  height: 720,
                  border: "1px solid var(--border)",
                  borderRadius: 16,
                  background: "rgba(0,0,0,0.25)"
                }}
              />
            ) : (
              <div style={{ display: "grid", gridTemplateColumns: "1fr", gap: 16 }}>
                <div
                  ref={docxContainerRef}
                  style={{
                    width: "100%",
                    overflow: "auto",
                    border: "1px solid var(--border)",
                    borderRadius: 16,
                    background: "rgba(255,255,255,0.95)",
                    color: "#111",
                    padding: 16,
                    maxHeight: 720
                  }}
                />

                <div
                  style={{
                    border: "1px solid var(--border)",
                    borderRadius: 16,
                    background: "rgba(0,0,0,0.15)",
                    padding: 16,
                    overflow: "auto"
                  }}
                >
                  <div className="h">Fill fields (chat)</div>
                  <div className="small" style={{ marginTop: 6 }}>
                    Answer one-by-one. You can skip any field.
                  </div>

                  {loadingFields ? (
                    <div className="pill" style={{ marginTop: 12, display: "inline-block" }}>
                      Extracting fields…
                    </div>
                  ) : null}

                  {fieldsError ? (
                    <div style={{ marginTop: 12 }} className="small">
                      <span className="pill pillWarn">AI labeling unavailable</span>{" "}
                      <span style={{ opacity: 0.85 }}>{fieldsError}</span>
                    </div>
                  ) : null}

                  <div style={{ marginTop: 12 }}>
                    {chat.map((m, idx) => (
                      <div key={idx} style={{ marginBottom: 10 }}>
                        <div className="small">{m.role === "assistant" ? "Agent" : "You"}</div>
                        <div style={{ whiteSpace: "pre-wrap" }}>{m.content}</div>
                      </div>
                    ))}
                  </div>

                  <div className="row" style={{ marginTop: 12 }}>
                    <input
                      className="input"
                      value={input}
                      placeholder={
                        loadingFields
                          ? "Extracting fields…"
                          : currentKey
                            ? `Answer for ${currentKey}…`
                            : "Done"
                      }
                      onChange={(e) => setInput(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key !== "Enter") return;
                        if (loadingFields) return;
                        if (!currentKey) return;
                        const v = input.trim();
                        if (!v) return;
                        setInput("");
                        setValues((prev) => ({ ...prev, [currentKey]: v }));
                        setChat((prev) => [...prev, { role: "user", content: v }]);
                        const next = nextAskableKey({ ...values, [currentKey]: v }, skipped);
                        setCurrentKey(next);
                        if (next) {
                          const q = fields.find((x) => x.key === next)?.question ?? "Please provide the next value.";
                          setChat((prev) => [...prev, { role: "assistant", content: q }]);
                        } else {
                          setChat((prev) => [
                            ...prev,
                            { role: "assistant", content: "All done (or skipped). Preview has been updated." }
                          ]);
                        }
                      }}
                    />
                    <button
                      className="btn btnPrimary"
                      disabled={loadingFields || !currentKey || !input.trim()}
                      onClick={() => {
                        if (loadingFields) return;
                        if (!currentKey) return;
                        const v = input.trim();
                        if (!v) return;
                        setInput("");
                        setValues((prev) => ({ ...prev, [currentKey]: v }));
                        setChat((prev) => [...prev, { role: "user", content: v }]);
                        const next = nextAskableKey({ ...values, [currentKey]: v }, skipped);
                        setCurrentKey(next);
                        if (next) {
                          const q = fields.find((x) => x.key === next)?.question ?? "Please provide the next value.";
                          setChat((prev) => [...prev, { role: "assistant", content: q }]);
                        } else {
                          setChat((prev) => [
                            ...prev,
                            { role: "assistant", content: "All done (or skipped). Preview has been updated." }
                          ]);
                        }
                      }}
                    >
                      Send
                    </button>
                    <button
                      className="btn"
                      disabled={loadingFields || !currentKey}
                      onClick={() => {
                        if (loadingFields) return;
                        if (!currentKey) return;
                        const k = currentKey;
                        const newSkipped = new Set(skipped);
                        newSkipped.add(k);
                        setSkipped(newSkipped);
                        setChat((prev) => [...prev, { role: "user", content: "(skipped)" }]);
                        const next = nextAskableKey(values, newSkipped);
                        setCurrentKey(next);
                        if (next) {
                          const q = fields.find((x) => x.key === next)?.question ?? "Please provide the next value.";
                          setChat((prev) => [...prev, { role: "assistant", content: q }]);
                        } else {
                          setChat((prev) => [
                            ...prev,
                            { role: "assistant", content: "All done (or skipped). Preview has been updated." }
                          ]);
                        }
                      }}
                    >
                      Skip
                    </button>
                  </div>

                  <div style={{ marginTop: 14 }} className="small">
                    <b>Fields</b>
                    <div className="row" style={{ marginTop: 10 }}>
                      <button
                        className="btn btnPrimary"
                        disabled={!hasUnappliedDrafts}
                        onClick={() => {
                          setValues((prev) => ({ ...prev, ...draftValues }));
                          setDraftValues({});
                        }}
                      >
                        Update preview
                      </button>
                      <button
                        className="btn"
                        disabled={!file || previewType !== "docx"}
                        onClick={async () => {
                          if (!file || previewType !== "docx") return;
                          try {
                            const original = originalDocxRef.current ?? (await file.arrayBuffer());
                            originalDocxRef.current = original;
                            const patched = await patchDocxArrayBufferForPreview(original, values, occurrences);
                            const blob = new Blob([patched], {
                              type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            });
                            const url = URL.createObjectURL(blob);
                            const a = document.createElement("a");
                            const base = (filename || "document").replace(/\.docx$/i, "");
                            a.href = url;
                            a.download = `${base}-patched.docx`;
                            document.body.appendChild(a);
                            a.click();
                            a.remove();
                            URL.revokeObjectURL(url);
                          } catch {
                            // ignore
                          }
                        }}
                      >
                        Download patched DOCX (debug)
                      </button>
                      <button
                        className="btn"
                        disabled={!hasUnappliedDrafts}
                        onClick={() => setDraftValues({})}
                      >
                        Discard changes
                      </button>
                      <label className="small" style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <input
                          type="checkbox"
                          checked={jumpToTopOnUpdate}
                          onChange={(e) => setJumpToTopOnUpdate(e.target.checked)}
                        />
                        Jump to top after update
                      </label>
                      {hasUnappliedDrafts ? (
                        <span className="pill pillWarn">unapplied changes</span>
                      ) : (
                        <span className="pill">in sync</span>
                      )}
                    </div>
                    <div style={{ marginTop: 10 }}>
                      {!loadingFields && groupedFields.length === 0 ? (
                        <div className="small" style={{ opacity: 0.85 }}>
                          No fields detected yet.
                        </div>
                      ) : null}
                      {groupedFields.map(([group, items]) => (
                        <div key={group} style={{ marginBottom: 18 }}>
                          <div className="h" style={{ fontSize: 14 }}>
                            {group}
                          </div>
                          <div style={{ marginTop: 10 }}>
                            {items.map((f) => (
                              <div key={f.key} style={{ marginBottom: 12 }}>
                                <div className="row" style={{ justifyContent: "space-between" }}>
                                  <div className="small">
                                    <b>{prettyLabel(f)}</b>{" "}
                                    {skipped.has(f.key) ? (
                                      <span className="pill pillWarn">skipped</span>
                                    ) : null}
                                  </div>
                                  <div className="small" style={{ opacity: 0.7 }}>
                                    <code>{f.key}</code>
                                  </div>
                                </div>
                                <input
                                  className="input"
                                  value={draftValues[f.key] ?? values[f.key] ?? ""}
                                  placeholder="(empty)"
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    setDraftValues((prev) => ({ ...prev, [f.key]: v }));
                                  }}
                                />
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </section>
    </div>
  );
}


