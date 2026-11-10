// Consumer của Vercel Queues (topic "split-jobs"), đăng ký trong vercel.json.
// Thông điệp được xác nhận khi handler trả về; ném lỗi -> hàng đợi tự giao lại.
import { handleCallback } from "@vercel/queue";
import { LIMITS } from "@/lib/server/config";
import type { QueueMessage } from "@/lib/server/queue";
import { processJob } from "@/lib/server/worker";

export const maxDuration = 300;

const handler = handleCallback<QueueMessage>(
  async (message) => {
    if (!message?.jobId) return;
    await processJob(message.jobId);
  },
  {
    visibilityTimeoutSeconds: 300,
    retry: (_error, metadata) => {
      // worker tự đánh dấu lỗi sau maxAttempts; ở đây chặn thêm thông điệp "độc" lặp vô hạn
      if (metadata.deliveryCount > LIMITS.maxAttempts + 3) return { acknowledge: true };
      return { afterSeconds: Math.min(120, 15 * metadata.deliveryCount) };
    },
  },
);

export function POST(request: Request): Promise<Response> {
  return handler(request);
}
