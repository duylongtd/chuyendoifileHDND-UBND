import { isAuthed } from "./auth";
import { ConfigError } from "./config";

export function json(data: unknown, status = 200): Response {
  return Response.json(data, { status, headers: { "Cache-Control": "no-store" } });
}

export function fail(status: number, code: string, message: string): Response {
  return json({ error: { code, message } }, status);
}

/** Bọc route handler: yêu cầu đăng nhập, chuyển lỗi cấu hình/lỗi bất ngờ thành JSON rõ ràng. */
export function route<A extends unknown[]>(handler: (...args: A) => Promise<Response>, opts: { auth?: boolean } = {}) {
  return async (...args: A): Promise<Response> => {
    try {
      if (opts.auth !== false && !(await isAuthed())) return fail(401, "UNAUTHORIZED", "Phiên làm việc đã hết hạn. Hãy nhập lại mã truy cập.");
      return await handler(...args);
    } catch (e) {
      if (e instanceof ConfigError) return fail(503, "CONFIG", e.message);
      console.error("[api] lỗi không mong muốn:", e);
      return fail(500, "INTERNAL", "Hệ thống gặp lỗi. Hãy thử lại sau ít phút.");
    }
  };
}

export function clientIp(req: Request): string {
  return (req.headers.get("x-forwarded-for") ?? "").split(",")[0].trim() || "local";
}

const buckets = new Map<string, { n: number; reset: number }>();
/** Giới hạn tần suất đơn giản theo khóa (trong bộ nhớ instance). */
export function rateLimit(key: string, max: number, windowMs: number, cost = 1): boolean {
  const now = Date.now();
  const b = buckets.get(key);
  if (!b || b.reset < now) {
    buckets.set(key, { n: cost, reset: now + windowMs });
    if (buckets.size > 5000) for (const [k, v] of buckets) if (v.reset < now) buckets.delete(k);
    return cost <= max;
  }
  b.n += cost;
  return b.n <= max;
}
