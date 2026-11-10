// Nhận file theo từng phần nhỏ (mặc định 4 MB) để nằm dưới giới hạn 4,5 MB/request của Vercel Function.
// Server tự ghi từng phần vào kho file nên chạy được với mọi kiểu xác thực của Vercel Blob
// (BLOB_READ_WRITE_TOKEN hoặc OIDC của store Private), không cần cấp token cho trình duyệt.
import { LIMITS, XLSX_MIME } from "@/lib/server/config";
import { fail, json, route } from "@/lib/server/http";
import { loadJob, partFile } from "@/lib/server/jobs";
import { putObject } from "@/lib/server/storage";

export const maxDuration = 60;
type Ctx = { params: Promise<{ id: string }> };

export const PUT = route(async (req: Request, ctx: Ctx) => {
  const url = new URL(req.url);
  const part = Number(url.searchParams.get("part"));
  const total = Number(url.searchParams.get("total"));
  const maxParts = Math.ceil(LIMITS.maxFileBytes / LIMITS.minPartBytes);
  if (!Number.isInteger(part) || !Number.isInteger(total) || total < 1 || total > maxParts || part < 0 || part >= total) {
    return fail(400, "BAD_PART", "Thông tin phần tải lên không hợp lệ.");
  }
  const job = await loadJob((await ctx.params).id);
  if (!job) return fail(404, "NOT_FOUND", "Không tìm thấy tác vụ (có thể đã hết hạn 24 giờ).");
  if (job.status !== "created") return fail(409, "BAD_STATE", "File này đã được tải lên rồi.");

  const declared = Number(req.headers.get("content-length") ?? 0);
  if (declared > LIMITS.maxPartBytes) return fail(413, "PART_TOO_LARGE", "Phần tải lên vượt kích thước cho phép.");
  const body = Buffer.from(await req.arrayBuffer());
  if (!body.length) return fail(400, "NO_BODY", "Không nhận được nội dung file.");
  if (body.length > LIMITS.maxPartBytes) return fail(413, "PART_TOO_LARGE", "Phần tải lên vượt kích thước cho phép.");

  await putObject(partFile(job.id, part), body, XLSX_MIME);
  return json({ part, size: body.length });
});
