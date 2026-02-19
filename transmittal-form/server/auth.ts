import "dotenv/config";
import { betterAuth } from "better-auth";
import { prismaAdapter } from "better-auth/adapters/prisma";
import { nextCookies } from "better-auth/next-js";
import { genericOAuth } from "better-auth/plugins";
import { PrismaClient } from "@prisma/client";
import { withAccelerate } from "@prisma/extension-accelerate";

const parseOrigins = (value: string | undefined, fallback: string[]) => {
  const items = String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean);
  return items.length ? items : fallback;
};

export const prisma = new PrismaClient({
  accelerateUrl: process.env.DATABASE_URL,
}).$extends(withAccelerate());

const globalForPrisma = globalThis as unknown as {
  prisma?: typeof prisma;
};

export const db = globalForPrisma.prisma || prisma;

if (!globalForPrisma.prisma) {
  globalForPrisma.prisma = db;
}

export const auth = betterAuth({
  baseURL: process.env.BETTER_AUTH_URL,
  trustedOrigins: parseOrigins(process.env.BETTER_AUTH_TRUSTED_ORIGINS, [
    "http://localhost:3000",
  ]),
  database: prismaAdapter(db, { provider: "postgresql" }),
  plugins: [
    nextCookies(),
    genericOAuth({
      config: [
        {
          providerId: "google",
          clientId: process.env.GOOGLE_CLIENT_ID as string,
          clientSecret: process.env.GOOGLE_CLIENT_SECRET as string,
          authorizationUrl: "https://accounts.google.com/o/oauth2/v2/auth",
          tokenUrl: "https://oauth2.googleapis.com/token",
          userInfoUrl: "https://www.googleapis.com/oauth2/v3/userinfo",
          scopes: [
            "openid",
            "email",
            "profile",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
          ],
          redirectURI: `${process.env.BETTER_AUTH_URL}/api/auth/callback/google`,
          accessType: "offline",
          prompt: "select_account",
          pkce: false,
        },
        {
          providerId: "google-dds",
          clientId: process.env.GOOGLE_DDS_CLIENT_ID as string,
          clientSecret: process.env.GOOGLE_DDS_CLIENT_SECRET as string,
          authorizationUrl: "https://accounts.google.com/o/oauth2/v2/auth",
          tokenUrl: "https://oauth2.googleapis.com/token",
          userInfoUrl: "https://www.googleapis.com/oauth2/v3/userinfo",
          scopes: [
            "openid",
            "email",
            "profile",
            "https://www.googleapis.com/auth/drive.file",
            "https://www.googleapis.com/auth/drive.readonly",
            "https://www.googleapis.com/auth/spreadsheets",
          ],
          redirectURI: `${process.env.BETTER_AUTH_URL}/api/auth/callback/google-dds`,
          accessType: "offline",
          prompt: "select_account",
          pkce: false,
        },
      ],
    }),
  ],
  account: {
    accountLinking: {
      enabled: true,
    },
  },
});
