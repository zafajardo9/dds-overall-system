import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

export async function GET(request: Request) {
  try {
    const session = await auth.api.getSession({
      headers: request.headers as any,
    });

    if (!session?.user) {
      return NextResponse.json({ error: "Not authenticated" }, { status: 401 });
    }

    const account = await db.account.findFirst({
      where: {
        userId: session.user.id,
        providerId: "google",
      },
    });

    if (!account) {
      return NextResponse.json(
        { error: "No Google account linked" },
        { status: 404 },
      );
    }

    const expiresAt = account.accessTokenExpiresAt
      ? new Date(account.accessTokenExpiresAt)
      : new Date(0);

    if (expiresAt < new Date() && account.refreshToken) {
      const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: new URLSearchParams({
          client_id: process.env.GOOGLE_CLIENT_ID!,
          client_secret: process.env.GOOGLE_CLIENT_SECRET!,
          refresh_token: account.refreshToken,
          grant_type: "refresh_token",
        }),
      });

      const tokens = await tokenResponse.json();
      if (tokens.access_token) {
        const newExpiresAt = new Date(Date.now() + tokens.expires_in * 1000);
        await db.account.update({
          where: { id: account.id },
          data: {
            accessToken: tokens.access_token,
            accessTokenExpiresAt: newExpiresAt,
          },
        });

        return NextResponse.json({ accessToken: tokens.access_token });
      }
    }

    return NextResponse.json({ accessToken: account.accessToken });
  } catch (error: any) {
    console.error("Token fetch error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
