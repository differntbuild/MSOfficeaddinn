import { z } from 'zod';

export const AuditActionSchema = z.enum([
  "FIX_FORMULA",
  "DELETE_ROWS",
  "FORMAT_CELLS",
  "INVESTIGATE",
  "CREATE_DASHBOARD",
  "EXTERNAL_ENRICH"
]);

export const AuditFindingSchema = z.object({
  id: z.string(),
  title: z.string().describe("Clear, concise title without boilerplate."),
  desc: z.string().describe("Specific issue with exact cell addresses."),
  severity: z.enum(["CRITICAL", "HIGH", "MEDIUM", "LOW"]),
  impact: z.string().describe("Business or calculation impact in one sentence."),
  recommendation: z.string().describe("Actionable fix."),
  action_type: AuditActionSchema,
  affected_cells: z.array(z.string()).describe("Array of cell addresses like ['A1', 'B2']"),
  suggested_formula: z.string().optional().describe("For table creation, data investigation, or formatting actions, leave formula empty as it requires a UI action. ONLY populate if action_type is FIX_FORMULA with an exact, executable Excel formula.")
});

export const AuditReportSchema = z.object({
  health_score: z.number().min(0).max(100),
  summary: z.string().describe("Executive summary of the dataset health."),
  findings: z.array(AuditFindingSchema)
});

export type AuditAction = z.infer<typeof AuditActionSchema>;
export type AuditFinding = z.infer<typeof AuditFindingSchema>;
export type AuditReport = z.infer<typeof AuditReportSchema>;
