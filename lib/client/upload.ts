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

const NET_MSG = "Không kết nối được máy chủ. Hãy kiểm tra mạng rồi thử lại.";

export interface SendHooks {
  /** Số byte của phần này đã đẩy lên đường truyền */
  onBytes: (loaded: number) => void;
  signal: AbortSignal;
}
/** Gửi một phần của file. Lỗi đường truyền/bị hủy -> ApiError status 0; máy chủ từ chối -> ApiError kèm status. */
export type PartSender = (url: string, chunk: Blob, hooks: SendHooks) => Promise<void>;

/** Dùng XMLHttpRequest (không phải fetch) vì cần biết đã đẩy lên được bao nhiêu byte để phát hiện mạng chậm/đứng. */
const xhrSender: PartSender = (url, chunk, { onBytes, signal }) =>
  new Promise<void>((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("PUT", url);
    xhr.setRequestHeader("Content-Type", "application/octet-stream");
    xhr.upload.onprogress = (e) => onBytes(e.loaded);
    xhr.onload = () => {
      if (xhr.status >= 200 && xhr.status < 300) return resolve();
      let data: { error?: { code?: string; message?: string } } | null = null;
      try {
        data = JSON.parse(xhr.responseText);
      } catch {
        data = null;
      }
      const fallback = xhr.status === 413 ? "File quá lớn so với giới hạn của máy chủ." : xhr.status === 504 ? "Máy chủ xử lý quá lâu. Hãy thử lại." : `Máy chủ trả lỗi ${xhr.status}. Hãy thử lại.`;
      reject(new ApiError(data?.error?.message ?? fallback, xhr.status, data?.error?.code, data));
    };
    xhr.onerror = xhr.ontimeout = xhr.onabort = () => reject(new ApiError(NET_MSG, 0));
    if (signal.aborted) return reject(new ApiError(NET_MSG, 0));
    signal.addEventListener("abort", () => xhr.abort(), { once: true });
    xhr.send(chunk);
  });

/** Các mức tải lên: phần nhỏ dần, ít luồng song song dần. Mạng chậm/đứng thì chuyển xuống mức phù hợp. */
const UPLOAD_PLANS = [
  { divide: 1, concurrency: 4 },
  { divide: 2, concurrency: 2 },
  { divide: 4, concurrency: 1 },
  { divide: 8, concurrency: 1 },
];
const MIN_PART_BYTES = 512 * 1024;

export interface UploadOptions {
  send?: PartSender;
  /** Số lần thử mỗi phần khi lỗi đường truyền */
  attempts?: number;
  retryDelayMs?: number;
  /** Không đẩy thêm được byte nào trong ngần này thời gian -> coi là đường truyền đứng */
  stallMs?: number;
  /** Đã đẩy hết phần nhưng máy chủ chưa trả lời trong ngần này thời gian -> coi là đứng */
  responseWaitMs?: number;
  /** Sau ngần này thời gian thì bắt đầu đo tốc độ */
  slowCheckMs?: number;
  /** Ước tính một phần mất lâu hơn ngần này giây -> chuyển sang phần nhỏ hơn */
  slowPartSec?: number;
  /** Chọn mức mà mỗi phần ước tính xong trong ngần này giây */
  targetPartSec?: number;
  tickMs?: number;
}
const DEFAULTS: Required<UploadOptions> = {
  send: xhrSender,
  attempts: 4,
  retryDelayMs: 1500,
  stallMs: 30_000,
  responseWaitMs: 150_000,
  slowCheckMs: 12_000,
  slowPartSec: 90,
  targetPartSec: 45,
  tickMs: 1000,
};

/** Báo cho uploadFile biết cần đổi mức tải (không phải lỗi để hiện cho người dùng). */
class PlanSwitch extends Error {
  constructor(
    public reason: "slow" | "stalled",
    public bytesPerSec: number,
  ) {
    super(reason);
  }
}

function isNetworkTrouble(e: unknown): boolean {
  return e instanceof ApiError && (e.status === 0 || e.status >= 500);
}

async function uploadParts(jobId: string, file: File, partBytes: number, concurrency: number, onProgress: (pct: number) => void, o: Required<UploadOptions>, isLastPlan: boolean): Promise<number> {
  const total = Math.max(1, Math.ceil(file.size / partBytes));
  let next = 0;
  let done = 0;
  let failed: unknown = null;
  const inflight = new Map<number, { loaded: number; last: number; size: number; ctrl: AbortController }>();
  const startedAt = Date.now();
  const pushed = () => {
    let b = done;
    for (const f of inflight.values()) b += Math.min(f.loaded, f.size);
    return b;
  };
  const report = () => onProgress(Math.min(99, (pushed() / Math.max(1, file.size)) * 100));
  const abortAll = (e: unknown) => {
    if (!failed) failed = e;
    for (const f of inflight.values()) f.ctrl.abort();
  };
  const timer = setInterval(() => {
    const t = Date.now();
    for (const f of inflight.values()) {
      if (t - f.last > (f.loaded >= f.size ? o.responseWaitMs : o.stallMs)) {
        // mức cuối: hủy riêng phần này để thử lại; các mức trên: chuyển ngay xuống phần nhỏ hơn
        if (isLastPlan) f.ctrl.abort();
        else abortAll(new PlanSwitch("stalled", 0));
        return;
      }
    }
    if (!isLastPlan && inflight.size && t - startedAt >= o.slowCheckMs) {
      const speed = pushed() / ((t - startedAt) / 1000);
      if (speed > 0 && (partBytes * inflight.size) / speed > o.slowPartSec) abortAll(new PlanSwitch("slow", speed));
    }
  }, o.tickMs);
  const worker = async () => {
    while (next < total && !failed) {
      const i = next++;
      const chunk = file.slice(i * partBytes, (i + 1) * partBytes);
      for (let attempt = 1; ; attempt++) {
        const f = { loaded: 0, last: Date.now(), size: chunk.size, ctrl: new AbortController() };
        inflight.set(i, f);
        try {
          await o.send(`/api/jobs/${jobId}/upload?part=${i}&total=${total}`, chunk, {
            signal: f.ctrl.signal,
            onBytes: (n) => {
              if (n > f.loaded) {
                f.loaded = n;
                f.last = Date.now();
                report();
              }
            },
          });
          inflight.delete(i);
          break;
        } catch (e) {
          inflight.delete(i);
          if (failed) return;
          if (!isNetworkTrouble(e) || attempt >= o.attempts) return abortAll(e);
          await new Promise((r) => setTimeout(r, o.retryDelayMs * attempt));
          if (failed) return;
        }
      }
      done += chunk.size;
      report();
    }
  };
  try {
    await Promise.all(Array.from({ length: Math.min(concurrency, total) }, worker));
  } finally {
    clearInterval(timer);
  }
  if (failed) throw failed;
  return total;
}

/**
 * Tải file lên theo từng phần nhỏ qua API của hệ thống (mỗi phần dưới giới hạn 4,5 MB của máy chủ).
 * Bắt đầu với 4 MB × 4 luồng. Trong lúc tải, đo số byte thực sự đẩy lên được:
 *  - mạng chậm (ước tính một phần mất hơn 90 giây) -> chuyển thẳng xuống mức mà mỗi phần xong trong ~45 giây;
 *  - đường truyền đứng (30 giây không đẩy thêm được byte nào) hoặc lỗi lặp lại -> xuống mức kế tiếp.
 * Các mức: 4 MB × 4 → 2 MB × 2 → 1 MB × 1 → 512 KB × 1. Nhờ vậy không có yêu cầu nào treo hàng phút rồi mới báo lỗi.
 * Trả về số phần và mã SHA-256 của file gốc để máy chủ đối chiếu toàn vẹn.
 */
export async function uploadFile(
  jobId: string,
  file: File,
  partBytes: number,
  onProgress: (pct: number) => void,
  onNote: (note: string) => void,
  options: UploadOptions = {},
): Promise<{ parts: number; sha256?: string }> {
  const o = { ...DEFAULTS, ...options };
  const hashing = sha256Hex(file);
  hashing.catch(() => {});
  const sizeOf = (i: number) => Math.max(MIN_PART_BYTES, Math.floor(partBytes / UPLOAD_PLANS[i].divide));
  const last = UPLOAD_PLANS.length - 1;
  let lastError: unknown = null;
  for (let idx = 0; idx <= last; ) {
    try {
      const parts = await uploadParts(jobId, file, sizeOf(idx), UPLOAD_PLANS[idx].concurrency, onProgress, o, idx === last);
      onNote("");
      return { parts, sha256: await hashing };
    } catch (e) {
      if (e instanceof PlanSwitch) {
        idx++;
        if (e.reason === "slow") {
          while (idx < last && (sizeOf(idx) * UPLOAD_PLANS[idx].concurrency) / e.bytesPerSec > o.targetPartSec) idx++;
          onNote(`Mạng tải lên chậm (khoảng ${Math.max(1, Math.round(e.bytesPerSec / 1024))} KB/giây), đang chuyển sang tải từng phần nhỏ hơn…`);
        } else {
          onNote("Đường truyền bị đứng, đang tự tải lại với các phần nhỏ hơn…");
        }
      } else {
        lastError = e;
        if (!isNetworkTrouble(e)) throw e;
        idx++;
        onNote("Đường truyền không ổn định, đang tự tải lại với các phần nhỏ hơn…");
      }
      onProgress(0);
    }
  }
  onNote("");
  if (lastError instanceof ApiError && lastError.status === 0) {
    throw new ApiError("Không tải được file lên máy chủ sau nhiều lần thử (mạng chậm, chập chờn hoặc mạng cơ quan chặn). Hãy kiểm tra mạng rồi bấm “Tải lại file này”; nếu vẫn lỗi, thử mạng khác (ví dụ 4G).", 0);
  }
  throw lastError;
}
