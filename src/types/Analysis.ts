import { z } from 'zod';
import { IssueSchema } from './Issue';

export const AnalysisSchema = z.object({
  issues: z.array(IssueSchema),
  metadata: z.object({
    formulaCount: z.number(),
    sheetCount: z.number(),
    namedRanges: z.array(z.string()),
    volatileFunctions: z.number(),
    externalReferences: z.number(),
  }),
});

export type Analysis = z.infer<typeof AnalysisSchema>;