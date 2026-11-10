// Đưa tác vụ vào hàng đợi. Mặc định trên Vercel dùng Vercel Queues (bền vững, tự giao lại khi lỗi);
// nếu Queues chưa bật hoặc gửi lỗi -> tự chuyển sang chạy nền ngay trong function (after()).
import { after } from "next/server";
import { LIMITS, QUEUE_TOPIC, queueDriver } from "./config";
import { processJob, RetryLater } from "./worker";

export interface QueueMessage {
  jobId: string;
}

async function runInline(jobId: string): Promise<void> {
  for (let i = 0; i < LIMITS.maxAttempts; i++) {
    try {
      await processJob(jobId);
      return;
    } catch (e) {
      if (e instanceof RetryLater) return;
      await new Promise((r) => setTimeout(r, 1500 * (i + 1)));
    }
  }
}

export async function enqueue(jobId: string, attempt: number, forceInline = false): Promise<"vercel" | "inline"> {
  if (!forceInline && queueDriver() === "vercel") {
    try {
      const { send } = await import("@vercel/queue");
      await send<QueueMessage>(QUEUE_TOPIC, { jobId }, { idempotencyKey: `${jobId}:${attempt}`, retentionSeconds: 3600 });
      return "vercel";
    } catch (e) {
      console.error("[queue] Không gửi được vào Vercel Queues, chuyển sang chạy nền:", e);
    }
  }
  after(() => runInline(jobId));
  return "inline";
}
