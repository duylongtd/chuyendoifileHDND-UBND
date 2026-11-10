// Cấp quyền tải lên trực tiếp từ trình duyệt tới Vercel Blob (không đi qua function -> không vướng giới hạn 4,5 MB).
import { handleUpload, type HandleUploadBody } from "@vercel/blob/client";
import { isAuthed } from "@/lib/server/auth";
import { LIMITS, XLSX_MIME } from "@/lib/server/config";
import { fail, json } from "@/lib/server/http";
import { inputFile, isJobId, loadJob } from "@/lib/server/jobs";

export async function POST(req: Request): Promise<Response> {
  let body: HandleUploadBody;
  try {
    body = (await req.json()) as HandleUploadBody;
  } catch {
    return fail(400, "BAD_REQUEST", "Yêu cầu không hợp lệ.");
  }
  try {
    const result = await handleUpload({
      request: req,
      body,
      onBeforeGenerateToken: async (pathname, clientPayload) => {
        if (!(await isAuthed())) throw new Error("Phiên làm việc đã hết hạn. Hãy nhập lại mã truy cập.");
        const jobId = clientPayload ?? "";
        if (!isJobId(jobId) || pathname !== inputFile(jobId)) throw new Error("Đường dẫn tải lên không hợp lệ.");
        const job = await loadJob(jobId);
        if (!job || job.status !== "created") throw new Error("Tác vụ không tồn tại hoặc đã có file.");
        return {
          allowedContentTypes: [XLSX_MIME],
          maximumSizeInBytes: LIMITS.maxFileBytes,
          addRandomSuffix: false,
          allowOverwrite: true,
          validUntil: Date.now() + 15 * 60 * 1000,
        };
      },
    });
    return json(result);
  } catch (e) {
    return fail(400, "UPLOAD_DENIED", (e as Error).message || "Không cấp được quyền tải lên.");
  }
}
