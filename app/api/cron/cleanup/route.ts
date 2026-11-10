// Dọn file quá hạn (mặc định 24 giờ). Vercel Cron gọi hằng ngày kèm header Authorization: Bearer $CRON_SECRET.
import { isVercel, LIMITS } from "@/lib/server/config";
import { fail, json } from "@/lib/server/http";
import { deleteObjects, listObjects } from "@/lib/server/storage";

export const maxDuration = 120;

export async function GET(req: Request): Promise<Response> {
  const secret = process.env.CRON_SECRET;
  if (isVercel && !secret) return fail(503, "CONFIG", "Chưa cấu hình CRON_SECRET.");
  if (secret && req.headers.get("authorization") !== `Bearer ${secret}`) return fail(401, "UNAUTHORIZED", "Không có quyền.");
  try {
    const cutoff = Date.now() - LIMITS.retentionHours * 3600 * 1000;
    const all = await listObjects("jobs/");
    // Nhóm theo tác vụ: xóa cả tác vụ khi file mới nhất của nó đã quá hạn
    const newest = new Map<string, number>();
    for (const o of all) {
      const id = o.path.split("/")[1] ?? "";
      newest.set(id, Math.max(newest.get(id) ?? 0, o.uploadedAt));
    }
    const expired = all.filter((o) => (newest.get(o.path.split("/")[1] ?? "") ?? 0) < cutoff);
    await deleteObjects(expired.map((o) => o.path));
    return json({ scanned: all.length, deleted: expired.length, jobsDeleted: [...newest.values()].filter((t) => t < cutoff).length });
  } catch (e) {
    console.error("[cron] dọn file lỗi:", e);
    return fail(500, "INTERNAL", "Không dọn được file.");
  }
}
