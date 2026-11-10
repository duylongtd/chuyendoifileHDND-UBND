// Gọi API và tải file lên theo từng phần (chạy trên trình duyệt). Tách riêng để kiểm thử được.

export class ApiError extends Error {
  constructor(
    message: string,
    public status: number,
    public code?: string,
    public data?: unknown,
  ) {
    super(message);
  }
}

export async function api<T>(url: string, init?: RequestInit): Promise<T> {
  let res: Response;
  try {
    res = await fetch(url, { ...init, headers: { ...(init?.body ? { "Content-Type": "application/json" } : {}), ...init?.headers } });
  } catch {
    throw new ApiError("Không kết nối được máy chủ. Hãy kiểm tra mạng rồi thử lại.", 0);
  }
  const text = await res.text();
  let data: { error?: { code?: string; message?: string } } | null = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = null;
  }
  if (!res.ok) {
    const fallback = res.status === 413 ? "File quá lớn so với giới hạn của máy chủ." : res.status === 504 ? "Máy chủ xử lý quá lâu. Hãy thử lại." : `Máy chủ trả lỗi ${res.status}. Hãy thử lại.`;
    throw new ApiError(data?.error?.message ?? fallback, res.status, data?.error?.code, data);
  }
  return data as T;
}

async function sha256Hex(file: File): Promise<string | undefined> {
  // crypto.subtle chỉ có trên https / localhost; thiếu thì bỏ qua (máy chủ vẫn đối chiếu số phần và kích thước)
  if (!globalThis.crypto?.subtle) return undefined;
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

/** Các mức tải lên, thử lần lượt khi đường truyền lỗi: phần nhỏ dần, ít luồng song song dần. */
const UPLOAD_PLANS = [
  { divide: 1, concurrency: 4 },
  { divide: 2, concurrency: 2 },
  { divide: 4, concurrency: 1 },
];
const PART_TIMEOUT_MS = 120_000;

function isNetworkTrouble(e: unknown): boolean {
  return e instanceof ApiError && (e.status === 0 || e.status >= 500);
}

async function uploadParts(jobId: string, file: File, partBytes: number, concurrency: number, onProgress: (pct: number) => void, retryDelayMs: number): Promise<number> {
  const total = Math.max(1, Math.ceil(file.size / partBytes));
  let next = 0;
  let sent = 0;
  let failed: unknown = null;
  const worker = async () => {
    while (next < total && !failed) {
      const i = next++;
      const chunk = file.slice(i * partBytes, (i + 1) * partBytes);
      for (let attempt = 1; ; attempt++) {
        try {
          await api(`/api/jobs/${jobId}/upload?part=${i}&total=${total}`, {
            method: "PUT",
            body: chunk,
            headers: { "Content-Type": "application/octet-stream" },
            signal: typeof AbortSignal !== "undefined" && "timeout" in AbortSignal ? AbortSignal.timeout(PART_TIMEOUT_MS) : undefined,
          });
          break;
        } catch (e) {
          if (!isNetworkTrouble(e) || attempt >= 4) {
            failed = e;
            return;
          }
          await new Promise((r) => setTimeout(r, retryDelayMs * attempt));
        }
      }
      sent += chunk.size;
      onProgress((sent / Math.max(1, file.size)) * 100);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, total) }, worker));
  if (failed) throw failed;
  return total;
}

/**
 * Tải file lên theo từng phần nhỏ qua API của hệ thống (mỗi phần dưới giới hạn 4,5 MB của máy chủ).
 * Mỗi phần tự thử lại 4 lần; nếu vẫn lỗi đường truyền thì tải lại cả file với phần nhỏ hơn và ít luồng hơn
 * (4 MB × 4 luồng → 2 MB × 2 luồng → 1 MB × 1 luồng) để qua được mạng chậm hoặc mạng cơ quan hạn chế gói lớn.
 * Trả về số phần và mã SHA-256 của file gốc để máy chủ đối chiếu toàn vẹn.
 */
export async function uploadFile(
  jobId: string,
  file: File,
  partBytes: number,
  onProgress: (pct: number) => void,
  onNote: (note: string) => void,
  retryDelayMs = 1500,
): Promise<{ parts: number; sha256?: string }> {
  const hashing = sha256Hex(file);
  hashing.catch(() => {});
  let lastError: unknown = null;
  for (const plan of UPLOAD_PLANS) {
    const size = Math.max(1024 * 1024, Math.floor(partBytes / plan.divide));
    try {
      const parts = await uploadParts(jobId, file, size, plan.concurrency, onProgress, retryDelayMs);
      onNote("");
      return { parts, sha256: await hashing };
    } catch (e) {
      lastError = e;
      if (!isNetworkTrouble(e)) throw e;
      onProgress(0);
      onNote("Đường truyền không ổn định, đang tự tải lại với các phần nhỏ hơn…");
    }
  }
  onNote("");
  if (lastError instanceof ApiError && lastError.status === 0) {
    throw new ApiError("Không tải được file lên máy chủ sau nhiều lần thử (mạng chậm, chập chờn hoặc mạng cơ quan chặn). Hãy kiểm tra mạng rồi bấm “Tải lại file này”; nếu vẫn lỗi, thử mạng khác (ví dụ 4G).", 0);
  }
  throw lastError;
}

