import { XLSX_MIME } from "@/lib/server/config";
import { fail, route } from "@/lib/server/http";
import { loadJob, outputFile } from "@/lib/server/jobs";
import { getObjectStream } from "@/lib/server/storage";
import type { Org } from "@/lib/xlsx/classify";

export const maxDuration = 120;
type Ctx = { params: Promise<{ id: string }> };

export const GET = route(async (req: Request, ctx: Ctx) => {
  const org = new URL(req.url).searchParams.get("org") as Org;
  if (org !== "HDND" && org !== "UBND") return fail(400, "BAD_ORG", "Thiếu tham số org=HDND|UBND.");
  const job = await loadJob((await ctx.params).id);
  if (!job || job.status !== "done" || !job.outputs?.[org]) return fail(404, "NOT_FOUND", "File kết quả không còn (đã quá 24 giờ) hoặc chưa xử lý xong.");
  const obj = await getObjectStream(outputFile(job.id, org));
  if (!obj) return fail(404, "NOT_FOUND", "File kết quả đã bị xóa (quá 24 giờ). Hãy tải file gốc lên và tách lại.");
  const name = job.outputs[org].name;
  const ascii = name.normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/đ/g, "d").replace(/Đ/g, "D").replace(/[^\x20-\x7e]/g, "_").replace(/["\\]/g, "_");
  const headers: Record<string, string> = {
    "Content-Type": XLSX_MIME,
    "Content-Disposition": `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(name)}`,
    "Cache-Control": "private, no-store",
    "X-Content-Type-Options": "nosniff",
  };
  if (obj.size) headers["Content-Length"] = String(obj.size);
  return new Response(obj.stream, { headers });
});
