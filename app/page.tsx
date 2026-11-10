"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { Job } from "@/lib/server/jobs";

type Org = "HDND" | "UBND";
type Phase = "uploading" | "inspecting" | "confirm" | "queued" | "processing" | "done" | "error";

interface Limits {
  partBytes: number;
  maxFiles: number;
  maxFileMB: number;
  retentionHours: number;
}

interface Item {
  key: string;
  name: string;
  size: number;
  jobId?: string;
  phase: Phase;
  progress: number;
  job?: Job;
  mapping: Record<string, Org>;
  error?: string;
  showAll: boolean;
}

const LS_KEY = "vtvl-split-jobs-v1";
const STALE_QUEUED_MS = 3 * 60 * 1000;
const STALE_PROCESSING_MS = 8 * 60 * 1000;
const SHEET_LABEL: Record<string, string> = { A2: "2. A2_CongViec", B1: "3. B1_KhoiLuong", C3: "8. C3_NhuCau", D1: "9. D1_NhanLuc", D2: "10. D2_PhuongAn" };

class ApiError extends Error {
  constructor(
    message: string,
    public status: number,
    public code?: string,
    public data?: unknown,
  ) {
    super(message);
  }
}

async function api<T>(url: string, init?: RequestInit): Promise<T> {
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

function mb(n: number): string {
  return (n / 1048576).toLocaleString("vi-VN", { maximumFractionDigits: 1 }) + " MB";
}

function phaseOf(job: Job): Phase | null {
  switch (job.status) {
    case "ready":
      return "confirm";
    case "queued":
      return "queued";
    case "processing":
      return "processing";
    case "done":
      return "done";
    case "error":
      return "error";
    default:
      return null;
  }
}

function mappingOf(job: Job): Record<string, Org> {
  const m: Record<string, Org> = {};
  for (const p of job.positions ?? []) m[p.code] = job.mapping?.[p.code] ?? p.suggested;
  return m;
}

async function sha256Hex(file: File): Promise<string | undefined> {
  // crypto.subtle chỉ có trên https / localhost; thiếu thì bỏ qua (máy chủ vẫn đối chiếu số phần và kích thước)
  if (!globalThis.crypto?.subtle) return undefined;
  const digest = await crypto.subtle.digest("SHA-256", await file.arrayBuffer());
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

const UPLOAD_CONCURRENCY = 4;

/**
 * Tải file lên theo từng phần nhỏ qua API của hệ thống (mỗi phần dưới giới hạn 4,5 MB của máy chủ),
 * gửi SONG SONG nhiều phần cùng lúc. Mỗi phần tự thử lại tối đa 3 lần khi rớt mạng.
 * Trả về số phần và mã SHA-256 của file gốc để máy chủ đối chiếu toàn vẹn.
 */
async function uploadFile(jobId: string, file: File, partBytes: number, onProgress: (pct: number) => void): Promise<{ parts: number; sha256?: string }> {
  const total = Math.max(1, Math.ceil(file.size / partBytes));
  const hashing = sha256Hex(file);
  let next = 0;
  let sent = 0;
  let failed: unknown = null;
  const worker = async () => {
    while (next < total && !failed) {
      const i = next++;
      const chunk = file.slice(i * partBytes, (i + 1) * partBytes);
      for (let attempt = 1; ; attempt++) {
        try {
          await api(`/api/jobs/${jobId}/upload?part=${i}&total=${total}`, { method: "PUT", body: chunk, headers: { "Content-Type": "application/octet-stream" } });
          break;
        } catch (e) {
          const retryable = e instanceof ApiError && (e.status === 0 || e.status >= 500);
          if (!retryable || attempt >= 3) {
            failed = e;
            return;
          }
          await new Promise((r) => setTimeout(r, 1200 * attempt));
        }
      }
      sent += chunk.size;
      onProgress((sent / Math.max(1, file.size)) * 100);
    }
  };
  await Promise.all(Array.from({ length: Math.min(UPLOAD_CONCURRENCY, total) }, worker));
  if (failed) throw failed;
  return { parts: total, sha256: await hashing };
}

export default function Page() {
  const [authed, setAuthed] = useState<boolean | null>(null);
  const [limits, setLimits] = useState<Limits>({ partBytes: 4 * 1024 * 1024, maxFiles: 3, maxFileMB: 60, retentionHours: 24 });
  const [fatal, setFatal] = useState("");

  const refreshSession = useCallback(async () => {
    try {
      const s = await api<{ authed: boolean; storage?: string; limits: Limits }>("/api/session");
      setLimits(s.limits);
      setAuthed(s.authed);
      setFatal("");
    } catch (e) {
      setFatal((e as Error).message);
      setAuthed(false);
    }
  }, []);

  useEffect(() => {
    void refreshSession();
  }, [refreshSession]);

  return (
    <>
      <header className="topbar">
        <div className="topbar-in">
          <div className="logo" aria-hidden>
            VTVL
          </div>
          <div>
            <h1>Tách bộ biểu VTVL: HĐND – UBND</h1>
            <p>Tải 1 file HĐND-UBND cấp xã, nhận về 2 file riêng: HĐND và UBND</p>
          </div>
          <div className="spacer" />
          {authed && (
            <button
              className="linkbtn"
              onClick={async () => {
                await api("/api/session", { method: "DELETE" }).catch(() => {});
                setAuthed(false);
              }}
            >
              Thoát
            </button>
          )}
        </div>
      </header>
      <main>
        {fatal && <div className="banner err">{fatal}</div>}
        {authed === null && !fatal && <p className="muted">Đang tải…</p>}
        {authed === false && <Login onDone={refreshSession} />}
        {authed && <Workspace limits={limits} onUnauthorized={() => setAuthed(false)} />}
      </main>
    </>
  );
}

function Login({ onDone }: { onDone: () => void }) {
  const [code, setCode] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);
  return (
    <form
      className="card login"
      onSubmit={async (e) => {
        e.preventDefault();
        setBusy(true);
        setErr("");
        try {
          await api("/api/session", { method: "POST", body: JSON.stringify({ code }) });
          onDone();
        } catch (e2) {
          setErr((e2 as Error).message);
        } finally {
          setBusy(false);
        }
      }}
    >
      <h2>Nhập mã truy cập</h2>
      <p className="muted small">Trang xử lý dữ liệu công vụ nên chỉ dành cho người có mã truy cập.</p>
      <label className="small" htmlFor="code">
        Mã truy cập
      </label>
      <input id="code" className="field" type="password" autoComplete="current-password" value={code} onChange={(e) => setCode(e.target.value)} autoFocus />
      {err && <div className="banner err" role="alert">{err}</div>}
      <div className="row" style={{ marginTop: 12 }}>
        <button className="btn primary" disabled={busy || !code}>
          {busy ? "Đang kiểm tra…" : "Vào trang"}
        </button>
      </div>
    </form>
  );
}

function Workspace({ limits, onUnauthorized }: { limits: Limits; onUnauthorized: () => void }) {
  const [items, setItems] = useState<Item[]>([]);
  const [notice, setNotice] = useState<{ kind: "err" | "info"; text: string } | null>(null);
  const [over, setOver] = useState(false);
  const [now, setNow] = useState(() => Date.now());
  const [restored, setRestored] = useState(false);
  const itemsRef = useRef(items);
  itemsRef.current = items;
  const inputRef = useRef<HTMLInputElement>(null);

  const patch = useCallback((key: string, p: Partial<Item>) => setItems((list) => list.map((it) => (it.key === key ? { ...it, ...p } : it))), []);

  const handleError = useCallback(
    (e: unknown): string => {
      if (e instanceof ApiError && e.status === 401) onUnauthorized();
      return (e as Error).message || "Có lỗi xảy ra.";
    },
    [onUnauthorized],
  );

  // Khôi phục danh sách sau khi tải lại trang (file còn giữ trong thời hạn lưu)
  useEffect(() => {
    (async () => {
      try {
        const saved = JSON.parse(localStorage.getItem(LS_KEY) ?? "[]") as { jobId: string }[];
        const ids = saved.map((s) => s.jobId).filter(Boolean);
        if (ids.length) {
          const r = await api<{ jobs: Job[] }>(`/api/jobs?ids=${ids.join(",")}`);
          const list: Item[] = [];
          for (const job of r.jobs) {
            const phase = phaseOf(job);
            if (!phase) continue;
            list.push({ key: job.id, name: job.fileName, size: job.size, jobId: job.id, phase, progress: 100, job, mapping: mappingOf(job), error: job.error?.message, showAll: false });
          }
          setItems(list);
        }
      } catch (e) {
        handleError(e);
      } finally {
        setRestored(true);
      }
    })();
  }, [handleError]);

  useEffect(() => {
    if (!restored) return;
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(items.filter((i) => i.jobId && i.phase !== "uploading").map((i) => ({ jobId: i.jobId }))));
    } catch {}
  }, [items, restored]);

  // Hỏi trạng thái định kỳ cho các file đang chờ / đang xử lý
  const waiting = items.some((i) => i.phase === "queued" || i.phase === "processing");
  useEffect(() => {
    if (!waiting) return;
    let stop = false;
    const tick = async () => {
      const ids = itemsRef.current.filter((i) => i.jobId && (i.phase === "queued" || i.phase === "processing")).map((i) => i.jobId!);
      if (!ids.length) return;
      try {
        const r = await api<{ jobs: Job[]; missing: string[] }>(`/api/jobs?ids=${ids.join(",")}`);
        if (stop) return;
        setNow(Date.now());
        setItems((list) =>
          list.map((it) => {
            if (!it.jobId) return it;
            if (r.missing.includes(it.jobId)) return { ...it, phase: "error", error: "Tác vụ không còn trên máy chủ (đã quá hạn lưu). Hãy tải lại file." };
            const job = r.jobs.find((j) => j.id === it.jobId);
            if (!job) return it;
            const phase = phaseOf(job) ?? it.phase;
            return { ...it, job, phase, error: job.error?.message };
          }),
        );
      } catch (e) {
        if (e instanceof ApiError && e.status === 401) onUnauthorized();
      }
    };
    const t = setInterval(tick, 2000);
    void tick();
    return () => {
      stop = true;
      clearInterval(t);
    };
  }, [waiting, onUnauthorized]);

  const addFiles = useCallback(
    async (files: File[]) => {
      setNotice(null);
      if (!files.length) return;
      const current = itemsRef.current;
      if (current.length + files.length > limits.maxFiles) {
        setNotice({ kind: "err", text: `Mỗi lượt chỉ xử lý tối đa ${limits.maxFiles} file. Đang có ${current.length} file trong danh sách, hãy xóa bớt hoặc bấm "Làm lượt mới".` });
        return;
      }
      for (const f of files) {
        const n = f.name.normalize("NFC");
        if (!/\.xlsx$/i.test(n)) return setNotice({ kind: "err", text: `"${n}" không phải file .xlsx. Nếu là file .xls, hãy mở bằng Excel và lưu lại thành .xlsx.` });
        if (n.startsWith("~$")) return setNotice({ kind: "err", text: `"${n}" là file tạm của Excel. Hãy đóng Excel và chọn file gốc.` });
        if (f.size === 0) return setNotice({ kind: "err", text: `"${n}" là file rỗng.` });
        if (f.size > limits.maxFileMB * 1048576) return setNotice({ kind: "err", text: `"${n}" nặng ${mb(f.size)}, vượt giới hạn ${limits.maxFileMB} MB mỗi file.` });
        if (current.some((i) => i.name === n) || files.filter((x) => x.name.normalize("NFC") === n).length > 1) return setNotice({ kind: "err", text: `File "${n}" đã có trong danh sách.` });
      }
      let jobs: Job[];
      try {
        jobs = (await api<{ jobs: Job[] }>("/api/jobs", { method: "POST", body: JSON.stringify({ files: files.map((f) => ({ name: f.name.normalize("NFC"), size: f.size })) }) })).jobs;
      } catch (e) {
        setNotice({ kind: "err", text: handleError(e) });
        return;
      }
      const fresh: Item[] = jobs.map((job) => ({ key: job.id, name: job.fileName, size: job.size, jobId: job.id, phase: "uploading", progress: 0, mapping: {}, showAll: false }));
      setItems((list) => [...list, ...fresh]);
      // Tải lần lượt từng file để đường truyền và máy chủ ổn định
      for (let i = 0; i < jobs.length; i++) {
        const job = jobs[i];
        try {
          const up = await uploadFile(job.id, files[i], limits.partBytes, (pct) => patch(job.id, { progress: Math.round(pct) }));
          patch(job.id, { phase: "inspecting", progress: 100 });
          const r = await api<{ job: Job }>(`/api/jobs/${job.id}/inspect`, { method: "POST", body: JSON.stringify(up) });
          patch(job.id, { phase: "confirm", job: r.job, mapping: mappingOf(r.job) });
        } catch (e) {
          patch(job.id, { phase: "error", error: handleError(e) });
        }
      }
    },
    [limits, patch, handleError],
  );

  const startSplit = useCallback(
    async (it: Item) => {
      if (!it.jobId) return;
      patch(it.key, { phase: "queued", error: undefined });
      try {
        const r = await api<{ job: Job }>(`/api/jobs/${it.jobId}/split`, { method: "POST", body: JSON.stringify({ mapping: it.mapping }) });
        patch(it.key, { job: r.job, phase: phaseOf(r.job) ?? "queued" });
        setNow(Date.now());
      } catch (e) {
        const msg = handleError(e);
        const backToConfirm = e instanceof ApiError && (e.code === "ONE_SIDE_EMPTY" || e.code === "BAD_MAPPING");
        patch(it.key, backToConfirm ? { phase: "confirm", error: msg } : { phase: "error", error: msg });
      }
    },
    [patch, handleError],
  );

  const remove = useCallback((it: Item) => {
    setItems((list) => list.filter((x) => x.key !== it.key));
    if (it.jobId) void api(`/api/jobs/${it.jobId}`, { method: "DELETE" }).catch(() => {});
  }, []);

  const done = items.filter((i) => i.phase === "done" && i.jobId);
  const ready = items.filter((i) => i.phase === "confirm");
  const busy = items.some((i) => i.phase === "uploading" || i.phase === "inspecting");
  const full = items.length >= limits.maxFiles;

  const downloadAll = () => {
    const links = done.flatMap((i) => (["HDND", "UBND"] as Org[]).map((org) => `/api/jobs/${i.jobId}/download?org=${org}`));
    links.forEach((href, idx) =>
      setTimeout(() => {
        const a = document.createElement("a");
        a.href = href;
        a.download = "";
        document.body.appendChild(a);
        a.click();
        a.remove();
      }, idx * 900),
    );
    setNotice({ kind: "info", text: `Đang tải ${links.length} file. Nếu trình duyệt hỏi "cho phép tải nhiều file", hãy chọn Cho phép.` });
  };

  return (
    <>
      <ol className="steps">
        <li>
          <b>1. Tải file lên</b>Tối đa {limits.maxFiles} file .xlsx mỗi lượt, mỗi file tối đa {limits.maxFileMB} MB.
        </li>
        <li>
          <b>2. Xác nhận vị trí HĐND</b>Kiểm tra bảng gợi ý, sửa nếu cần.
        </li>
        <li>
          <b>3. Tải 2 file kết quả</b>File tự xóa khỏi máy chủ sau {limits.retentionHours} giờ.
        </li>
      </ol>

      <div
        className={`drop${over ? " over" : ""}${full || busy ? " disabled" : ""}`}
        onDragOver={(e) => {
          e.preventDefault();
          if (!full && !busy) setOver(true);
        }}
        onDragLeave={() => setOver(false)}
        onDrop={(e) => {
          e.preventDefault();
          setOver(false);
          if (!full && !busy) void addFiles([...e.dataTransfer.files]);
        }}
      >
        <strong>Kéo thả file bộ biểu HĐND-UBND vào đây</strong>
        <span className="muted small">hoặc</span>
        <div className="row" style={{ justifyContent: "center", marginTop: 10 }}>
          <button className="btn primary" disabled={full || busy} onClick={() => inputRef.current?.click()}>
            Chọn file (.xlsx)
          </button>
          {items.length > 0 && (
            <button
              className="btn"
              disabled={busy || waiting}
              onClick={() => {
                if (!confirm("Xóa toàn bộ danh sách hiện tại (kể cả file kết quả trên máy chủ) để làm lượt mới?")) return;
                itemsRef.current.forEach((i) => i.jobId && void api(`/api/jobs/${i.jobId}`, { method: "DELETE" }).catch(() => {}));
                setItems([]);
                setNotice(null);
              }}
            >
              Làm lượt mới
            </button>
          )}
        </div>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          multiple
          hidden
          onChange={(e) => {
            const files = [...(e.target.files ?? [])];
            e.target.value = "";
            void addFiles(files);
          }}
        />
        <p className="muted small" style={{ margin: "10px 0 0" }}>
          {full ? `Đã đủ ${limits.maxFiles} file cho lượt này.` : `Còn nhận thêm ${limits.maxFiles - items.length} file trong lượt này.`}
        </p>
      </div>

      {notice && (
        <div className={`banner ${notice.kind}`} role={notice.kind === "err" ? "alert" : "status"}>
          {notice.text}
        </div>
      )}

      {(ready.length > 1 || done.length > 1) && (
        <div className="row" style={{ marginTop: 14 }}>
          {ready.length > 1 && (
            <button className="btn primary" onClick={() => ready.forEach((i) => void startSplit(i))}>
              Tách tất cả {ready.length} file
            </button>
          )}
          {done.length > 1 && (
            <button className="btn" onClick={downloadAll}>
              ⬇ Tải tất cả {done.length * 2} file kết quả
            </button>
          )}
        </div>
      )}

      <div style={{ marginTop: 14 }}>
        {items.map((it) => (
          <FileCard key={it.key} it={it} now={now} patch={patch} onSplit={() => void startSplit(it)} onRemove={() => remove(it)} />
        ))}
      </div>

      <p className="footer-note">
        Nội dung các ô được giữ nguyên, chỉ chia dòng theo vị trí HĐND/UBND và dồn lên liền nhau. Khi mở file kết quả, Excel sẽ tự tính lại công thức.
      </p>
    </>
  );
}

function StatusBadge({ it, stale }: { it: Item; stale: boolean }) {
  switch (it.phase) {
    case "uploading":
      return <span className="badge blue"><i className="dot" />Đang tải lên {it.progress}%</span>;
    case "inspecting":
      return <span className="badge blue"><i className="dot" />Đang đọc danh sách vị trí</span>;
    case "confirm":
      return <span className="badge warn">Chờ bạn xác nhận</span>;
    case "queued":
      return <span className={`badge ${stale ? "warn" : "gray"}`}>{!stale && <i className="dot" />}{stale ? "Chờ quá lâu" : "Trong hàng đợi"}</span>;
    case "processing":
      return <span className={`badge ${stale ? "warn" : "blue"}`}>{!stale && <i className="dot" />}{stale ? "Xử lý quá lâu" : "Đang tách file"}</span>;
    case "done":
      return <span className="badge ok">✓ Hoàn tất</span>;
    default:
      return <span className="badge err">Lỗi</span>;
  }
}

function FileCard({ it, now, patch, onSplit, onRemove }: { it: Item; now: number; patch: (key: string, p: Partial<Item>) => void; onSplit: () => void; onRemove: () => void }) {
  const job = it.job;
  const positions = useMemo(() => job?.positions ?? [], [job]);
  const nH = positions.filter((p) => it.mapping[p.code] === "HDND").length;
  const nFlag = positions.filter((p) => !p.confident).length;
  const age = job ? now - job.updatedAt : 0;
  const stale = (it.phase === "queued" && age > STALE_QUEUED_MS) || (it.phase === "processing" && age > STALE_PROCESSING_MS);
  const visible = it.showAll ? positions : positions.filter((p) => it.mapping[p.code] === "HDND" || !p.confident || p.related);
  const canRemove = it.phase !== "uploading" && it.phase !== "inspecting";

  return (
    <section className="card" aria-label={it.name}>
      <div className="file-head">
        <div>
          <div className="file-name">{it.name}</div>
          <div className="muted small">{mb(it.size)}</div>
        </div>
        <div className="row">
          <StatusBadge it={it} stale={stale} />
          {canRemove && (
            <button className="btn ghost sm" onClick={onRemove} aria-label={`Xóa ${it.name} khỏi danh sách`}>
              Xóa
            </button>
          )}
        </div>
      </div>

      {(it.phase === "uploading" || it.phase === "inspecting") && (
        <div className="bar" aria-hidden>
          <i style={{ width: `${it.phase === "inspecting" ? 100 : it.progress}%` }} />
        </div>
      )}

      {it.phase === "confirm" && job && (
        <>
          <div className="row" style={{ marginTop: 12 }}>
            <span className="badge hdnd">HĐND: {nH} vị trí</span>
            <span className="badge ubnd">UBND: {positions.length - nH} vị trí</span>
            {nFlag > 0 && <span className="badge warn">{nFlag} vị trí cần bạn xem lại</span>}
          </div>
          {nFlag > 0 && (
            <div className="banner warn">
              Có {nFlag} vị trí để trống tên ở biểu 6. C1_VTVL nên hệ thống chỉ gợi ý được (dòng tô vàng). Hãy kiểm tra và chọn đúng HĐND hoặc UBND trước khi tách.
            </div>
          )}
          {!!job.missingSheets?.length && (
            <div className="banner warn">File thiếu biểu: {job.missingSheets.map((r) => SHEET_LABEL[r] ?? r).join(", ")}. Các biểu còn lại vẫn được tách bình thường.</div>
          )}
          {it.error && <div className="banner err" role="alert">{it.error}</div>}
          <div className="tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Mã VTVL</th>
                  <th>Tên vị trí (biểu 6. C1_VTVL)</th>
                  <th>Thuộc khối</th>
                </tr>
              </thead>
              <tbody>
                {visible.map((p) => (
                  <tr key={p.code} className={p.confident ? "" : "flag"}>
                    <td className="code">{p.code}</td>
                    <td>
                      {p.name || <em className="muted">(Tên vị trí để trống)</em>}
                      {p.confident && p.related && <div className="hint">{p.reason}.</div>}
                      {!p.confident && (
                        <div className="hint">
                          {p.reason}.{p.hint ? ` Mô tả ở biểu C2: “${p.hint.slice(0, 160)}${p.hint.length > 160 ? "…" : ""}”` : ""}
                        </div>
                      )}
                    </td>
                    <td>
                      <div className="seg" role="group" aria-label={`Khối của ${p.code}`}>
                        {(["HDND", "UBND"] as Org[]).map((org) => (
                          <button
                            key={org}
                            type="button"
                            aria-pressed={it.mapping[p.code] === org}
                            className={`${org.toLowerCase()}${it.mapping[p.code] === org ? " on" : ""}`}
                            onClick={() => patch(it.key, { mapping: { ...it.mapping, [p.code]: org }, error: undefined })}
                          >
                            {org === "HDND" ? "HĐND" : "UBND"}
                          </button>
                        ))}
                      </div>
                    </td>
                  </tr>
                ))}
                {!visible.length && (
                  <tr>
                    <td colSpan={3} className="muted">
                      Chưa có vị trí nào được chọn là HĐND. Bấm “Hiện tất cả” để chọn.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
          <div className="row between" style={{ marginTop: 12 }}>
            <button className="btn sm" onClick={() => patch(it.key, { showAll: !it.showAll })}>
              {it.showAll ? "Chỉ hiện vị trí HĐND và vị trí liên quan" : `Hiện tất cả ${positions.length} vị trí`}
            </button>
            <button className="btn primary" onClick={onSplit} disabled={nH === 0 || nH === positions.length}>
              Tách thành 2 file
            </button>
          </div>
          {(nH === 0 || nH === positions.length) && <p className="small muted">Cần có ít nhất 1 vị trí HĐND và 1 vị trí UBND thì mới tách được.</p>}
        </>
      )}

      {(it.phase === "queued" || it.phase === "processing") && (
        <>
          <p className="muted small" style={{ margin: "10px 0 0" }}>
            {it.phase === "queued" ? "File đang chờ tới lượt. Bạn có thể để trang mở, kết quả sẽ tự hiện." : "Đang tách file, thường mất dưới 1 phút."}
          </p>
          {stale && (
            <div className="banner warn">
              Tác vụ chạy lâu hơn bình thường.{" "}
              <button className="btn sm" onClick={onSplit}>
                Chạy lại
              </button>
            </div>
          )}
        </>
      )}

      {it.phase === "done" && job?.outputs && (
        <>
          <div className="row" style={{ marginTop: 12 }}>
            <a className="btn hdnd" href={`/api/jobs/${job.id}/download?org=HDND`} download>
              ⬇ {job.outputs.HDND.name} ({mb(job.outputs.HDND.size)})
            </a>
            <a className="btn ubnd" href={`/api/jobs/${job.id}/download?org=UBND`} download>
              ⬇ {job.outputs.UBND.name} ({mb(job.outputs.UBND.size)})
            </a>
          </div>
          {job.report && (
            <>
              <p className="small muted" style={{ margin: "10px 0 0" }}>
                HĐND {job.report.positions.hdnd} vị trí · UBND {job.report.positions.ubnd} vị trí
                {job.durationMs ? ` · xử lý trong ${(job.durationMs / 1000).toLocaleString("vi-VN", { maximumFractionDigits: 1 })} giây` : ""}
              </p>
              {!!job.report.warnings.length && (
                <div className="banner warn">
                  Có {job.report.warnings.length} điểm cần lưu ý:
                  <ul className="warns">
                    {job.report.warnings.map((w) => (
                      <li key={w}>{w}</li>
                    ))}
                  </ul>
                </div>
              )}
              <details>
                <summary>Số dòng đã chia theo từng biểu</summary>
                <div className="tablewrap">
                  <table>
                    <thead>
                      <tr>
                        <th>Biểu</th>
                        <th className="num">Dòng gốc</th>
                        <th className="num">Vào file HĐND</th>
                        <th className="num">Vào file UBND</th>
                        <th className="num">Có ở cả hai</th>
                      </tr>
                    </thead>
                    <tbody>
                      {job.report.sheets.map((s) => (
                        <tr key={s.sheet}>
                          <td>{s.sheet}</td>
                          <td className="num">{s.total}</td>
                          <td className="num">{s.hdnd + s.both}</td>
                          <td className="num">{s.ubnd + s.both}</td>
                          <td className="num">{s.both}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
            </>
          )}
        </>
      )}

      {it.phase === "error" && (
        <>
          <div className="banner err" role="alert">{it.error ?? "Có lỗi xảy ra với file này."}</div>
          {!!positions.length && (
            <div className="row" style={{ marginTop: 10 }}>
              <button className="btn sm" onClick={() => patch(it.key, { phase: "confirm", error: undefined })}>
                Thử tách lại
              </button>
            </div>
          )}
        </>
      )}
    </section>
  );
}
