export type RouteLeaveGuard = (nextHash: string) => boolean | Promise<boolean>;

const routeLeaveGuards = new Map<string, RouteLeaveGuard>();

export function setRouteLeaveGuard(hash: string, guard: RouteLeaveGuard | null): void {
  if (guard) routeLeaveGuards.set(hash, guard);
  else routeLeaveGuards.delete(hash);
}

export async function confirmRouteLeave(currentHash: string, nextHash: string): Promise<boolean> {
  const guard = routeLeaveGuards.get(currentHash);
  if (!guard || currentHash === nextHash) return true;
  return guard(nextHash);
}
