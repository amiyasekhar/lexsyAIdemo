import { NextResponse } from "next/server";
import { z } from "zod";
import OpenAI from "openai";

export const runtime = "nodejs";

const DetectedFieldSchema = z.object({
  raw_placeholder: z.string(),
  kind: z.enum(["bracket_token", "bracket_blank", "underline", "signature_label"]),
  context_before: z.string(),
  context_line: z.string(),
  context_after: z.string(),
  location_hint: z.string()
});

const BodySchema = z.object({
  detections: z.array(DetectedFieldSchema)
});

const LLMFieldSchema = z.object({
  field_id: z.string(),
  raw_placeholder: z.string(),
  label_for_user: z.string(),
  who_should_fill: z.enum(["company", "counterparty", "either", "system"]),
  expected_type: z.enum([
    "string",
    "date",
    "money",
    "state",
    "email",
    "name",
    "title",
    "address",
    "jurisdiction",
    "other"
  ]),
  how_to_fill: z.string(),
  confidence: z.number().min(0).max(1),
  evidence_quote: z.string(),
  location_hint: z.string(),
  ask_user: z.boolean()
});

const OutputSchema = z.array(LLMFieldSchema);

const SYSTEM_INSTRUCTIONS = `You are a document field-extraction engine for legal documents.
You will receive detected placeholders with local context.
Your job: infer what the field should be, who should fill it, and how to ask the user.

Rules:
- Output ONLY valid JSON: an array of objects matching the schema described.
- Do NOT invent actual values. Only describe what should be entered.
- Prefer stable dot-notation field_id values, e.g. company.name, investor.name, safe.date, safe.purchase_amount.
- If the placeholder is clearly for the counterparty (e.g., Investor name, Purchase Amount), set who_should_fill="counterparty" and ask_user=false.
- Typed signatures only: for signature blocks, ask for printed name/title/email/address where appropriate, not a drawn signature.
- If uncertain, set lower confidence and explain briefly in how_to_fill.
`;

export async function POST(req: Request) {
  try {
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      return NextResponse.json(
        { error: "OPENAI_API_KEY is not set on the server." },
        { status: 500 }
      );
    }

    const json = await req.json();
    const parsed = BodySchema.safeParse(json);
    if (!parsed.success) {
      return NextResponse.json(
        { error: "Invalid request body.", issues: parsed.error.issues },
        { status: 400 }
      );
    }

    const model = process.env.OPENAI_MODEL || "gpt-4o-mini";
    const client = new OpenAI({ apiKey });

    const schemaHint = `[
  {
    "field_id": "string",
    "raw_placeholder": "string",
    "label_for_user": "string",
    "who_should_fill": "company|counterparty|either|system",
    "expected_type": "string|date|money|state|email|name|title|address|jurisdiction|other",
    "how_to_fill": "string (1 sentence)",
    "confidence": "number 0..1",
    "evidence_quote": "string",
    "location_hint": "string",
    "ask_user": "boolean"
  }
]`;

    const res = await client.chat.completions.create({
      model,
      temperature: 0.2,
      messages: [
        {
          role: "system",
          content: [
            "Return ONLY valid JSON (no markdown).",
            `JSON schema hint:\n${schemaHint}`,
            "",
            SYSTEM_INSTRUCTIONS
          ].join("\n")
        },
        {
          role: "user",
          content: JSON.stringify(
            {
              placeholders: parsed.data.detections
            },
            null,
            2
          )
        }
      ]
    });

    const raw = res.choices?.[0]?.message?.content?.trim() || "[]";
    let data: unknown;
    try {
      data = JSON.parse(raw);
    } catch {
      return NextResponse.json(
        { error: "Model did not return valid JSON.", raw },
        { status: 502 }
      );
    }

    const out = OutputSchema.safeParse(data);
    if (!out.success) {
      return NextResponse.json(
        { error: "Model JSON failed schema validation.", issues: out.error.issues, raw },
        { status: 502 }
      );
    }

    return NextResponse.json(out.data);
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return NextResponse.json({ error: message }, { status: 500 });
  }
}


