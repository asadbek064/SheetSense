import { z } from 'zod';

export const IssueType = z.enum(['formula', 'style', 'data']);
export type IssueType = z.infer<typeof IssueType>;

export const IssueSeverity = z.enum(['error', 'warning', 'info']);
export type IssueSeverity = z.infer<typeof IssueSeverity>;

export const IssueSchema = z.object({
  type: IssueType,
  severity: IssueSeverity,
  cell: z.string(),
  sheet: z.string(),
  message: z.string(),
  suggestion: z.string().optional(),
});

export type Issue = z.infer<typeof IssueSchema>;