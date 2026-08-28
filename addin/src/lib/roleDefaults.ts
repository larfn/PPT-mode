export interface RoleDefaultShape {
  type?: string;
  hasText?: boolean;
}

export function defaultRoleForShape(s: RoleDefaultShape | null | undefined): string {
  if (!s) return 'fixed';
  if (s.type === 'table') return 'table';
  if (s.type === 'picture') return 'ai_image';
  return s.hasText ? 'ai_text' : 'fixed';
}
