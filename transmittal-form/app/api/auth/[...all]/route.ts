import { toNextJsHandler } from "better-auth/next-js";
import { auth } from "@/server/auth";

export const runtime = "nodejs";

const handler = toNextJsHandler(auth);

export const { GET, POST, PUT, PATCH, DELETE } = handler;
