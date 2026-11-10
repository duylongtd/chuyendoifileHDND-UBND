// Cấu hình và giới hạn hệ thống. Thiếu cấu hình bắt buộc -> báo lỗi rõ ràng (không chạy âm thầm sai).

export const LIMITS = {
  /** Số file tối đa mỗi lượt tải lên */
  maxFilesPerBatch: 3,
  /** Dung lượng tối đa mỗi file (.xlsx) */
  maxFileBytes: 60 * 1024 * 1024,
  /** File và kết quả tự xóa sau số giờ này */
  retentionHours: 24,
  /** Số lần thử lại tối đa cho một tác vụ */
  maxAttempts: 3,
  /** Sau khoảng này mà tác vụ vẫn "đang chờ/đang xử lý" thì coi là treo, cho phép chạy lại */
  staleQueuedMs: 3 * 60 * 1000,
  staleProcessingMs: 8 * 60 * 1000,
} as const;

export const QUEUE_TOPIC = "split-jobs";
export const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

export class ConfigError extends Error {}

export const isVercel = Boolean(process.env.VERCEL);

export type StorageDriver = "blob" | "local";
export function storageDriver(): StorageDriver {
  const forced = process.env.STORAGE_DRIVER;
  if (forced === "blob" || forced === "local") return forced;
  if (process.env.BLOB_READ_WRITE_TOKEN || process.env.BLOB_STORE_ID) return "blob";
  if (isVercel) {
    throw new ConfigError("Chưa kết nối Vercel Blob cho project (thiếu BLOB_READ_WRITE_TOKEN / BLOB_STORE_ID). Vào Storage → tạo Blob store (private) và Connect vào project.");
  }
  return "local";
}

export type QueueDriver = "vercel" | "inline";
export function queueDriver(): QueueDriver {
  const forced = process.env.QUEUE_DRIVER;
  if (forced === "vercel" || forced === "inline") return forced;
  return isVercel ? "vercel" : "inline";
}

export function blobAccess(): "private" | "public" {
  return process.env.BLOB_ACCESS === "public" ? "public" : "private";
}

export function accessCode(): string {
  const code = process.env.ACCESS_CODE;
  if (!code || code.length < 8) {
    throw new ConfigError("Chưa cấu hình biến môi trường ACCESS_CODE (mã truy cập, tối thiểu 8 ký tự).");
  }
  return code;
}
