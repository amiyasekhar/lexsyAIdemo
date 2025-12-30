import JSZip from "jszip";

export type Field = {
  key: string;
  label: string;
  question: string;
  kind: "bracket_token" | "bracket_blank" | "signature_blank";
  party?: "company" | "investor" | "unknown";
  // For targeting replacements
  paragraphIndex?: number;
};

export type DetectedField = {
  raw_placeholder: string;
  kind: "bracket_token" | "bracket_blank" | "underline" | "signature_label";
  context_before: string;
  context_line: string;
  context_after: string;
  location_hint: string;
};

const BRACKET_TOKEN_RE = /\[[^\[\]\n]{1,120}\]/g;
const BRACKET_BLANK_RE = /\$\[[\s_]{3,}\]|\[[\s_]{3,}\]/g;
const UNDERLINE_RE = /_{3,}/g;
const SIGNATURE_LABEL_RE = /(By:|Name:|Title:|Address:|Email:)/gi;

function normalizeKey(s: string) {
  return s
    .toLowerCase()
    .trim()
    .replace(/[^\w\s.-]/g, "")
    .replace(/\s+/g, "_");
}

function decodeXmlEntities(s: string) {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, hex) =>
      String.fromCharCode(parseInt(hex, 16))
    )
    .replace(/&#([0-9]+);/g, (_m, num) => String.fromCharCode(parseInt(num, 10)));
}

function extractParagraphText(pXml: string) {
  const ts = pXml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [];
  const parts: string[] = [];
  for (const t of ts) {
    const inner = /<w:t[^>]*>([\s\S]*?)<\/w:t>/.exec(t)?.[1] ?? "";
    parts.push(decodeXmlEntities(inner));
  }
  // Tabs are meaningful in signature blocks
  const tabCount = (pXml.match(/<w:tab\/>/g) || []).length;
  const text = parts.join("");
  return { text, tabCount };
}

function partyFromLine(line: string): "company" | "investor" | "unknown" {
  const s = line.toLowerCase();
  if (s.includes("investor")) return "investor";
  if (s.includes("company") || s.includes("issuer")) return "company";
  return "unknown";
}

function addUnique(map: Map<string, Field>, field: Field) {
  if (!map.has(field.key)) map.set(field.key, field);
}

export async function detectFieldsFromDocx(arrayBuffer: ArrayBuffer): Promise<Field[]> {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) return [];
  const xml = await docFile.async("string");
  const paragraphs = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];

  const fields = new Map<string, Field>();
  let currentParty: "company" | "investor" | "unknown" = "unknown";
  let pendingAddressContinuationParty: "company" | "investor" | "unknown" | null =
    null;
  let addressContinuationCount = 0;

  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i]!;
    const { text, tabCount } = extractParagraphText(pXml);
    const maybeParty = partyFromLine(text);
    if (maybeParty !== "unknown") currentParty = maybeParty;

    // Bracket tokens: [Company Name], [Date of Safe], etc.
    for (const m of text.matchAll(BRACKET_TOKEN_RE)) {
      const raw = m[0]!;
      const label = raw.slice(1, -1).trim();
      const norm = normalizeKey(label);

      const known: Record<string, { key: string; q: string }> = {
        company_name: { key: "company_name", q: "What is the Company legal name?" },
        investor_name: { key: "investor_name", q: "What is the Investor legal name?" },
        date_of_safe: { key: "date_of_safe", q: "What is the Date of Safe?" },
        state_of_incorporation: {
          key: "state_of_incorporation",
          q: "What is the State of Incorporation?"
        },
        governing_law_jurisdiction: {
          key: "governing_law_jurisdiction",
          q: "What is the Governing Law Jurisdiction?"
        },
        company: { key: "company_signature_name", q: "What is the Company name for the signature block?" },
        name: { key: "company_signer_print_name", q: "What is the printed name of the Company signer?" },
        title: { key: "company_signer_title", q: "What is the title of the Company signer?" }
      };

      const mapped = known[norm];
      const key = mapped?.key ?? `bracket_${norm}`;
      const question = mapped?.q ?? `Please provide: ${label}`;

      addUnique(fields, {
        key,
        label,
        question,
        kind: "bracket_token",
        paragraphIndex: i
      });
    }

    // Bracket blanks: [_____], used for purchase amount / valuation cap.
    for (const m of text.matchAll(BRACKET_BLANK_RE)) {
      const raw = m[0]!;
      const lower = text.toLowerCase();
      let key = `blank_${i}`;
      let label = "Blank";
      let question = "Please provide the missing value.";
      if (lower.includes("purchase amount")) {
        key = "purchase_amount";
        label = "Purchase Amount";
        question = "What is the Purchase Amount?";
      } else if (lower.includes("post-money valuation cap") || lower.includes("valuation cap")) {
        key = "post_money_valuation_cap";
        label = "Post-Money Valuation Cap";
        question = "What is the Post-Money Valuation Cap?";
      }

      addUnique(fields, {
        key,
        label,
        question,
        kind: "bracket_blank",
        paragraphIndex: i
      });

      // Keep raw around by making key-specific label if multiple unknown blanks exist.
      void raw;
    }

    // Signature blanks in this template are underlined TAB runs (no text, many <w:tab/>, <w:u/>).
    const isUnderlined = /<w:u\b/.test(pXml);
    const hasTabs = tabCount > 0;
    const hasLabel = /\b(By:|Name:|Title:|Address:|Email:)\b/.test(text);

    if (hasLabel) {
      // Address has a continuation blank line (next paragraph is mostly tabs, no label).
      if (text.includes("Address:")) {
        pendingAddressContinuationParty = currentParty;
        addressContinuationCount = 0;
      } else {
        pendingAddressContinuationParty = null;
      }
    } else if (
      pendingAddressContinuationParty &&
      pendingAddressContinuationParty !== "unknown" &&
      isUnderlined &&
      hasTabs &&
      !text.trim()
    ) {
      addressContinuationCount += 1;
      if (addressContinuationCount >= 1) {
        // treat as second line of address
        addUnique(fields, {
          key: `${pendingAddressContinuationParty}_address_line2`,
          label:
            pendingAddressContinuationParty === "company"
              ? "Company Address (line 2)"
              : "Investor Address (line 2)",
          question:
            pendingAddressContinuationParty === "company"
              ? "What is the Company address (line 2)?"
              : "What is the Investor address (line 2)?",
          kind: "signature_blank",
          party: pendingAddressContinuationParty,
          paragraphIndex: i
        });
      }
    }

    if (hasLabel && isUnderlined && hasTabs) {
      const party = currentParty;
      const prefix = party === "investor" ? "investor" : party === "company" ? "company" : "unknown";

      if (text.includes("By:")) {
        addUnique(fields, {
          key: `${prefix}_by`,
          label: prefix === "investor" ? "Investor By" : "Company By",
          question: prefix === "investor" ? "Who is signing on behalf of the Investor?" : "Who is signing on behalf of the Company?",
          kind: "signature_blank",
          party,
          paragraphIndex: i
        });
      }
      if (text.includes("Name:")) {
        addUnique(fields, {
          key: `${prefix}_name`,
          label: prefix === "investor" ? "Investor Name (signature)" : "Company Name (signature)",
          question: prefix === "investor" ? "What is the printed name of the Investor signer?" : "What is the printed name of the Company signer?",
          kind: "signature_blank",
          party,
          paragraphIndex: i
        });
      }
      if (text.includes("Title:")) {
        addUnique(fields, {
          key: `${prefix}_title`,
          label: prefix === "investor" ? "Investor Title" : "Company Title",
          question: prefix === "investor" ? "What is the title of the Investor signer?" : "What is the title of the Company signer?",
          kind: "signature_blank",
          party,
          paragraphIndex: i
        });
      }
      if (text.includes("Address:")) {
        addUnique(fields, {
          key: `${prefix}_address_line1`,
          label: prefix === "investor" ? "Investor Address" : "Company Address",
          question: prefix === "investor" ? "What is the Investor address?" : "What is the Company address?",
          kind: "signature_blank",
          party,
          paragraphIndex: i
        });
      }
      if (text.includes("Email:")) {
        addUnique(fields, {
          key: `${prefix}_email`,
          label: prefix === "investor" ? "Investor Email" : "Company Email",
          question: prefix === "investor" ? "What is the Investor email?" : "What is the Company email?",
          kind: "signature_blank",
          party,
          paragraphIndex: i
        });
      }
    }
  }

  return Array.from(fields.values());
}

export async function detectDetectionsFromDocx(
  arrayBuffer: ArrayBuffer,
  window = 1
): Promise<DetectedField[]> {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const docFile = zip.file("word/document.xml");
  if (!docFile) return [];
  const xml = await docFile.async("string");
  const paragraphs = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];

  type ParaInfo = {
    idx: number;
    text: string;
    hasTabs: boolean;
    underlined: boolean;
    location: string;
  };

  const infos: ParaInfo[] = [];
  for (let i = 0; i < paragraphs.length; i++) {
    const pXml = paragraphs[i]!;
    const { text, tabCount } = extractParagraphText(pXml);
    infos.push({
      idx: i,
      text: (text || "").replace(/\s+/g, " ").trim(),
      hasTabs: tabCount > 0,
      underlined: /<w:u\b/.test(pXml),
      location: `paragraph:${i}`
    });
  }

  const nearestPrevText = (i: number) => {
    for (let k = i - 1; k >= 0; k--) {
      if (infos[k]!.text) return infos[k]!.text;
    }
    return "";
  };
  const nearestNextText = (i: number) => {
    for (let k = i + 1; k < infos.length; k++) {
      if (infos[k]!.text) return infos[k]!.text;
    }
    return "";
  };

  const detected: DetectedField[] = [];
  let section: "company" | "investor" | "unknown" = "unknown";
  let lastLabel: string | null = null;

  for (let i = 0; i < infos.length; i++) {
    const info = infos[i]!;
    const line = info.text;
    const loc = info.location;

    if (/INVESTOR:/i.test(line)) section = "investor";
    if (/\[COMPANY\]/i.test(line) || /^\s*COMPANY\s*:?/i.test(line)) section = "company";

    const before = nearestPrevText(i);
    const after = nearestNextText(i);

    // Bracket tokens / blanks from actual visible text.
    for (const m of line.matchAll(BRACKET_TOKEN_RE)) {
      detected.push({
        raw_placeholder: m[0]!,
        kind: "bracket_token",
        context_before: before,
        context_line: line,
        context_after: after,
        location_hint: loc
      });
    }

    for (const m of line.matchAll(BRACKET_BLANK_RE)) {
      detected.push({
        raw_placeholder: m[0]!,
        kind: "bracket_blank",
        context_before: before,
        context_line: line,
        context_after: after,
        location_hint: loc
      });
    }

    // Signature/tab underline blanks: these often have NO placeholder text.
    // We emit an "underline" placeholder for the blank itself so /api/fields can label it.
    // Our preview patcher converts these to literal "________" in the same paragraph, so use that as raw_placeholder.
    if (info.underlined && info.hasTabs) {
      // Find any signature labels in this same paragraph text.
      const labels = Array.from(line.matchAll(SIGNATURE_LABEL_RE)).map((m) => m[1] || m[0] || "");
      if (labels.length) {
        lastLabel = labels[labels.length - 1] || lastLabel;
        // One underline field per label (By/Name/Title/Address/Email)
        for (const lab of labels) {
          detected.push({
            raw_placeholder: "________",
            kind: "underline",
            context_before: before,
            context_line: `${section.toUpperCase()}: ${lab} ________`,
            context_after: after,
            location_hint: loc
          });
        }
      } else if (!line) {
        // Common: Address continuation line is a blank underline line with tabs only.
        const synthetic = lastLabel?.toLowerCase().includes("address")
          ? `${section.toUpperCase()}: Address (line 2) ________`
          : `${section.toUpperCase()}: Blank ________`;
        detected.push({
          raw_placeholder: "________",
          kind: "underline",
          context_before: before,
          context_line: synthetic,
          context_after: after,
          location_hint: loc
        });
      } else {
        detected.push({
          raw_placeholder: "________",
          kind: "underline",
          context_before: before,
          context_line: `${section.toUpperCase()}: ${line} ________`,
          context_after: after,
          location_hint: loc
        });
      }
    } else if (UNDERLINE_RE.test(line) && !BRACKET_BLANK_RE.test(line)) {
      // literal underscore runs in text
      detected.push({
        raw_placeholder: "________",
        kind: "underline",
        context_before: before,
        context_line: line,
        context_after: after,
        location_hint: loc
      });
    }
  }

  // Dedupe by (raw_placeholder, location, kind)
  const uniq = new Map<string, DetectedField>();
  for (const d of detected) {
    const k = `${d.raw_placeholder}::${d.kind}::${d.location_hint}::${d.context_line}`;
    uniq.set(k, d);
  }
  return Array.from(uniq.values());
}


