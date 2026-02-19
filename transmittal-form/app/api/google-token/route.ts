import { NextResponse } from "next/server";
import { auth, db } from "@/server/auth";

export const runtime = "nodejs";

const GOOGLE_PROVIDER_IDS = ["google", "google-dds"] as const;
type GoogleProviderId = (typeof GOOGLE_PROVIDER_IDS)[number];

const getProviderCredentials = (providerId: string) => {
  if (providerId === "google-dds") {
    return {
      clientId: process.env.GOOGLE_DDS_CLIENT_ID,
      clientSecret: process.env.GOOGLE_DDS_CLIENT_SECRET,
    };
  }

  return {
    clientId: process.env.GOOGLE_CLIENT_ID,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET,
  };
};

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
        providerId: {
          in: [...GOOGLE_PROVIDER_IDS],
        },
      },
      orderBy: {
        updatedAt: "desc",
      },
    });

    if (!account) {
      return NextResponse.json(
        { error: "No Google account linked for this session" },
        { status: 404 },
      );
    }

    const providerId = account.providerId as GoogleProviderId;
    const credentials = getProviderCredentials(providerId);

    if (!credentials.clientId || !credentials.clientSecret) {
      return NextResponse.json(
        {
          error:
            providerId === "google-dds"
              ? "Missing GOOGLE_DDS_CLIENT_ID/GOOGLE_DDS_CLIENT_SECRET"
              : "Missing GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET",
        },
        { status: 500 },
      );
    }

    const expiresAt = account.accessTokenExpiresAt
      ? new Date(account.accessTokenExpiresAt)
      : new Date(0);

    const isExpired = expiresAt < new Date();

    if (isExpired && account.refreshToken) {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 15000);

      try {
        const tokenResponse = await fetch("https://oauth2.googleapis.com/token", {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: credentials.clientId,
            client_secret: credentials.clientSecret,
            refresh_token: account.refreshToken,
            grant_type: "refresh_token",
          }),
          signal: controller.signal,
        });
        clearTimeout(timeout);

        const tokens = await tokenResponse.json();

        if (tokens.error) {
          console.error(
            `[${providerId}] Token refresh error from Google:`,
            tokens.error,
            tokens.error_description,
          );
          return NextResponse.json(
            { error: `Google token refresh failed: ${tokens.error_description || tokens.error}` },
            { status: 401 },
          );
        }

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

        return NextResponse.json(
          { error: "Token refresh returned no access token" },
          { status: 500 },
        );
      } catch (refreshErr: any) {
        clearTimeout(timeout);
        console.error(
          `[${providerId}] Token refresh network error:`,
          refreshErr.message,
        );
        return NextResponse.json(
          { error: "Network error refreshing Google token. Please try again." },
          { status: 502 },
        );
      }
    }

    if (isExpired && !account.refreshToken) {
      return NextResponse.json(
        { error: "Google token expired and no refresh token available. Please sign in again." },
        { status: 401 },
      );
    }

    return NextResponse.json({ accessToken: account.accessToken });
  } catch (error: any) {
    console.error("Token fetch error:", error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
