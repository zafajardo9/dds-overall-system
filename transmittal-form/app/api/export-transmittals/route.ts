import { NextRequest, NextResponse } from "next/server";
import { db } from "@/server/auth";

export const runtime = "nodejs";

const stripTransmittalPrefix = (value: string) =>
  value.startsWith("TR-FP-") ? value.slice("TR-FP-".length) : value;

export async function GET(request: NextRequest) {
  try {
    const SEND_API_TOKEN = process.env.SEND_API_TOKEN;
    if (!SEND_API_TOKEN) {
      return NextResponse.json(
        { error: "SEND_API_TOKEN not configured" },
        { status: 500 },
      );
    }

    const tokenFromHeader = request.headers.get("x-api-token");
    const tokenFromQuery = request.nextUrl.searchParams.get("token");
    const token = tokenFromHeader || tokenFromQuery;

    if (!token || token !== SEND_API_TOKEN) {
      return NextResponse.json({ error: "Invalid token" }, { status: 401 });
    }

    const transmittals = await db.transmittal.findMany({
      include: {
        items: true,
        recipients: true,
        user: true,
      },
    });

    const rows = transmittals.flatMap((t) => {
      const project = (t.project || {}) as any;
      const transmittalNo = stripTransmittalPrefix(
        String(project.transmittalNumber || ""),
      );
      const date = String(project.date || "");

      const primaryRecipient = t.recipients?.[0];
      const client =
        primaryRecipient?.recipientOrganization ||
        primaryRecipient?.recipientName ||
        "";

      const mode = [
        t.handDelivery ? "Hand Delivery" : null,
        t.pickUp ? "Pick Up" : null,
        t.courier ? "Courier" : null,
        t.registeredMail ? "Registered Mail" : null,
      ]
        .filter(Boolean)
        .join(", ");

      const transmittedBy = t.preparedBy || t.user?.name || t.user?.email || "";

      const items = (t.items || []).length ? t.items : [null];
      return items.map((item) => {
        const docs = item
          ? [item.documentNumber, item.description].filter(Boolean).join(" - ")
          : "";

        return {
          client,
          documents: docs,
          transmittalDocumentNo: transmittalNo,
          dateOfTransmittal: date,
          transmittedBy,
          mode,
          gdLink: item && item.fileType === "gdrive" ? item.fileSource || "" : "",
          remarks: item?.remarks || "",
        };
      });
    });

    return NextResponse.json({
      rows,
      totalRows: rows.length,
      totalTransmittals: transmittals.length,
    });
  } catch (error: any) {
    console.error("Export transmittals error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
