import { fail, json, route } from "@/lib/server/http";
import { deleteJob, loadJob } from "@/lib/server/jobs";

type Ctx = { params: Promise<{ id: string }> };

export const GET = route(async (_req: Request, ctx: Ctx) => {
  const job = await loadJob((await ctx.params).id);
  return job ? json({ job }) : fail(404, "NOT_FOUND", "Không tìm thấy tác vụ (có thể đã hết hạn 24 giờ).");
});

export const DELETE = route(async (_req: Request, ctx: Ctx) => {
  await deleteJob((await ctx.params).id);
  return json({ ok: true });
});
