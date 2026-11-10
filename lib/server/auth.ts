// Mã truy cập chung + phiên đăng nhập bằng cookie ký HMAC (không cần cơ sở dữ liệu).
import crypto from "node:crypto";
import { cookies } from "next/headers";
import { accessCode } from "./config";

const COOKIE = "vtvl_split_session";
const SESSION_HOURS = 12;

function secret(): Buffer {
  return crypto.createHash("sha256").update(`vtvl-split|${accessCode()}|${process.env.SESSION_SECRET ?? ""}`).digest();
}

function sign(payload: string): string {
  return crypto.createHmac("sha256", secret()).update(payload).digest("base64url");
}

export function checkAccessCode(input: string): boolean {
  const a = crypto.createHash("sha256").update(input).digest();
  const b = crypto.createHash("sha256").update(accessCode()).digest();
  return crypto.timingSafeEqual(a, b);
}

export async function createSession(): Promise<void> {
  const exp = Date.now() + SESSION_HOURS * 3600 * 1000;
  const payload = `${exp}.${crypto.randomBytes(9).toString("base64url")}`;
  (await cookies()).set(COOKIE, `${payload}.${sign(payload)}`, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: SESSION_HOURS * 3600,
  });
}

export async function destroySession(): Promise<void> {
  (await cookies()).delete(COOKIE);
}

export async function isAuthed(): Promise<boolean> {
  const raw = (await cookies()).get(COOKIE)?.value;
  if (!raw) return false;
  const i = raw.lastIndexOf(".");
  if (i < 0) return false;
  const payload = raw.slice(0, i);
  const sig = Buffer.from(raw.slice(i + 1));
  const expect = Buffer.from(sign(payload));
  if (sig.length !== expect.length || !crypto.timingSafeEqual(sig, expect)) return false;
  return Number(payload.split(".")[0]) > Date.now();
}

// Chống dò mã: giới hạn số lần thử theo IP (trong bộ nhớ của từng instance – mức bảo vệ cơ bản).
const attempts = new Map<string, { n: number; reset: number }>();
export function loginAllowed(ip: string): boolean {
  const now = Date.now();
  const a = attempts.get(ip);
  if (!a || a.reset < now) {
    attempts.set(ip, { n: 1, reset: now + 10 * 60 * 1000 });
    if (attempts.size > 5000) for (const [k, v] of attempts) if (v.reset < now) attempts.delete(k);
    return true;
  }
  a.n += 1;
  return a.n <= 10;
}
