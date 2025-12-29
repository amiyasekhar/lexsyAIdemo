"use client";

import { useEffect, useRef, useState } from "react";
import JSZip from "jszip";

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

async function patchDocxArrayBufferForHighlights(arrayBuffer: ArrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const xmlFiles = Object.keys(zip.files).filter((p) =>
    /^word\/(document|header\d+|footer\d+)\.xml$/.test(p)
  );

  for (const path of xmlFiles) {
    const file = zip.file(path);
    if (!file) continue;
    let xml = await file.async("string");

    // Convert underlined TAB runs (signature blanks) into highlightable underscore text.
    // DOCX signature lines in this template are underlined <w:tab/> sequences.
    xml = xml.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
      const hasTab = /<w:tab\/>/.test(run);
      if (!hasTab) return run;
      const isUnderlined = /<w:u\b/.test(run);
      if (!isUnderlined) return run;
      // Don't touch runs that already have visible text besides tabs.
      const hasText = /<w:t[^>]*>[\s\S]*?<\/w:t>/.test(run);
      if (hasText) return run;

      const highlighted = ensureHighlightInRun(run);
      // Replace each tab with a fixed underscore run so it's visible and highlightable.
      return highlighted.replace(
        /<w:tab\/>/g,
        `<w:t xml:space="preserve">________</w:t>`
      );
    });

    zip.file(path, xml);
  }

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
      el.innerHTML = "";

      const { renderAsync } = await import("docx-preview");
      const original = await file.arrayBuffer();
      const arrayBuffer = await patchDocxArrayBufferForHighlights(original);

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
    }

    void renderDocx();
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
            )}
          </div>
        </div>
      </section>
    </div>
  );
}


