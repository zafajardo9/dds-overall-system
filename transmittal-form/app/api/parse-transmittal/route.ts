import { NextResponse } from "next/server";
import { parseTransmittalDocument } from "@/services/geminiService";

export const runtime = "nodejs";

const MAX_CONTENT_LENGTH = 25 * 1024 * 1024;

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const content = String(body?.content || "");
    const mimeType = String(body?.mimeType || "");
    const isText = Boolean(body?.isText);
    const fileName = body?.fileName ? String(body.fileName) : undefined;

    if (!content || !mimeType) {
      return NextResponse.json(
        { error: "Missing content or mimeType." },
        { status: 400 },
      );
    }

    if (content.length > MAX_CONTENT_LENGTH) {
      return NextResponse.json(
        { error: "File content is too large to analyze." },
        { status: 413 },
      );
    }

    const result = await parseTransmittalDocument(
      content,
      mimeType,
      isText,
      fileName,
    );

    return NextResponse.json(result);
  } catch (error: any) {
    console.error("Parse transmittal error:", error);
    return NextResponse.json(
      { error: String(error?.message || "Failed to analyze document.") },
      { status: 500 },
    );
  }
}
