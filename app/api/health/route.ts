// Thông tin vận hành (cần đăng nhập): cấu hình đang dùng và bộ nhớ của instance.
// rss = bộ nhớ hệ điều hành đang cấp; heap + buffers = phần ứng dụng thực sự đang giữ.
import { queueDriver, storageDriver } from "@/lib/server/config";
import { json, route } from "@/lib/server/http";

export const GET = route(async () => {
  const m = process.memoryUsage();
  const mb = (n: number) => Math.round(n / 1048576);
  return json({
    ok: true,
    storage: storageDriver(),
    queue: queueDriver(),
    node: process.version,
    uptimeSeconds: Math.round(process.uptime()),
    memoryMB: { rss: mb(m.rss), heapUsed: mb(m.heapUsed), heapTotal: mb(m.heapTotal), buffers: mb(m.arrayBuffers), external: mb(m.external) },
  });
});
