import { checkAccessCode, createSession, destroySession, isAuthed, loginAllowed } from "@/lib/server/auth";
import { LIMITS, storageDriver } from "@/lib/server/config";
import { clientIp, fail, json, route } from "@/lib/server/http";

export const GET = route(
  async () => {
    const authed = await isAuthed();
    return json({
      authed,
      storage: authed ? storageDriver() : undefined,
      limits: { maxFiles: LIMITS.maxFilesPerBatch, maxFileMB: Math.floor(LIMITS.maxFileBytes / 1048576), retentionHours: LIMITS.retentionHours },
    });
  },
  { auth: false },
);

export const POST = route(
  async (req: Request) => {
    if (!loginAllowed(clientIp(req))) return fail(429, "TOO_MANY", "Nhập sai quá nhiều lần. Hãy thử lại sau 10 phút.");
    const body = (await req.json().catch(() => null)) as { code?: unknown } | null;
    const code = typeof body?.code === "string" ? body.code : "";
    if (!code || !checkAccessCode(code)) {
      await new Promise((r) => setTimeout(r, 400));
      return fail(401, "WRONG_CODE", "Mã truy cập không đúng.");
    }
    await createSession();
    return json({ authed: true });
  },
  { auth: false },
);

export const DELETE = route(
  async () => {
    await destroySession();
    return json({ authed: false });
  },
  { auth: false },
);
