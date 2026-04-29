import * as z from 'zod';

export const AgentIntentSchema = z.enum([
  'FORMULA',
  'EXPLANATION',
  'AUDIT',
  'PIVOT',
  'DATA_CLEANING',
  'GENERAL_ASSISTANT',
]);

export const FormulaDraftSchema = z.object({
  formula: z.string().min(2),
  explanation: z.string().min(3),
  confidence: z.number().min(0).max(1).optional(),
  assumptions: z.array(z.string()).optional(),
});

export const ExplainDraftSchema = z.object({
  summary: z.string().min(20),
  keyPoints: z.array(z.string()).min(1),
  followUps: z.array(z.string()).optional(),
});

export const ExecutionPlanStepSchema = z.object({
  id: z.string(),
  name: z.string(),
  retries: z.number().int().min(0).default(0),
  timeoutMs: z.number().int().min(1000).default(30000),
});

export const ExecutionPlanSchema = z.object({
  intent: AgentIntentSchema,
  steps: z.array(ExecutionPlanStepSchema),
  expectedSchema: z.string(),
});

export const ExecutionLogSchema = z.object({
  timestamp: z.string(),
  level: z.enum(['info', 'warn', 'error']),
  stage: z.string(),
  message: z.string(),
  meta: z.object({}).catchall(z.any()).optional(),
});

export type FormulaDraft = z.infer<typeof FormulaDraftSchema>;
export type ExplainDraft = z.infer<typeof ExplainDraftSchema>;
export type ExecutionPlan = z.infer<typeof ExecutionPlanSchema>;
export type ExecutionLog = z.infer<typeof ExecutionLogSchema>;
