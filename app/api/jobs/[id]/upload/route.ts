// Nhận file tải lên khi chạy trên máy (driver local). Trên Vercel, file đi thẳng từ trình duyệt lên Blob.
import { LIMITS, storageDriver } from "@/lib/server/config";
import { fail, json, route } from "@/lib/server/http";
import { inputFile, loadJob } from "@/lib/server/jobs";
import { putLocalStream } from "@/lib/server/storage";

export const maxDuration = 120;
type Ctx = { params: Promise<{ id: string }> };

export const PUT = route(async (req: Request, ctx: Ctx) => {
  if (storageDriver() !== "local") return fail(400, "WRONG_DRIVER", "Máy chủ này nhận file qua Vercel Blob.");
  const job = await loadJob((await ctx.params).id);
  if (!job) return fail(404, "NOT_FOUND", "Không tìm thấy tác vụ.");
  if (job.status !== "created") return fail(409, "BAD_STATE", "File này đã được tải lên rồi.");
  if (!req.body) return fail(400, "NO_BODY", "Không nhận được nội dung file.");
  try {
    const size = await putLocalStream(inputFile(job.id), req.body, LIMITS.maxFileBytes);
    return json({ size });
  } catch (e) {
    if ((e as Error).message === "FILE_TOO_LARGE") return fail(413, "TOO_LARGE", `File vượt ${Math.floor(LIMITS.maxFileBytes / 1048576)} MB.`);
    throw e;
  }
});
