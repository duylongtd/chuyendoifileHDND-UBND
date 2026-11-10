// Người dùng đã xác nhận bảng HĐND/UBND -> đưa tác vụ vào hàng đợi.
import { LIMITS } from "@/lib/server/config";
import { fail, json, route } from "@/lib/server/http";
import { loadJob, saveJob } from "@/lib/server/jobs";
import { enqueue } from "@/lib/server/queue";
import type { Org } from "@/lib/xlsx/classify";

export const maxDuration = 300;
type Ctx = { params: Promise<{ id: string }> };

export const POST = route(async (req: Request, ctx: Ctx) => {
  const job = await loadJob((await ctx.params).id);
  if (!job) return fail(404, "NOT_FOUND", "Không tìm thấy tác vụ (có thể đã hết hạn 24 giờ).");
  if (!job.positions?.length) return fail(409, "NOT_READY", "File chưa được đọc xong danh sách vị trí.");

  const age = Date.now() - job.updatedAt;
  const stale = (job.status === "queued" && age > LIMITS.staleQueuedMs) || (job.status === "processing" && age > LIMITS.staleProcessingMs);
  if ((job.status === "queued" || job.status === "processing") && !stale) return json({ job });

  const body = (await req.json().catch(() => null)) as { mapping?: Record<string, unknown> } | null;
  const mapping: Record<string, Org> = {};
  for (const p of job.positions) {
    const v = body?.mapping?.[p.code] ?? job.mapping?.[p.code] ?? p.suggested;
    if (v !== "HDND" && v !== "UBND") return fail(400, "BAD_MAPPING", `Vị trí ${p.code} chưa được chọn HĐND hay UBND.`);
    mapping[p.code] = v;
  }
  const nH = Object.values(mapping).filter((o) => o === "HDND").length;
  if (nH === 0) return fail(400, "ONE_SIDE_EMPTY", "Chưa có vị trí nào được chọn là HĐND nên không có gì để tách.");
  if (nH === job.positions.length) return fail(400, "ONE_SIDE_EMPTY", "Tất cả vị trí đang là HĐND nên không có gì để tách.");

  job.mapping = mapping;
  job.status = "queued";
  job.error = undefined;
  job.report = undefined;
  job.outputs = undefined;
  if (stale || job.attempts >= LIMITS.maxAttempts) job.attempts = 0;
  await saveJob(job);
  job.queuedVia = await enqueue(job.id, Date.now(), stale);
  return json({ job });
});
