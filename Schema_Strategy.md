# LLM Schema Strategy: Why and How

## Will Writing Schemas Make a Big Impact?

**Yes, absolutely.** Implementing strict schemas is arguably the single most impactful architectural change you can make when building reliable AI agents. 

When you move from raw text generation to schema-driven generation, you shift from **Parsing Text** (error-prone, fragile, unpredictable) to **Contract Enforcement** (reliable, typed, deterministic).

### The Massive Impact of Schemas:
1. **Zero Parsing Errors:** No more regex hacks or `try/catch` blocks trying to find JSON hidden inside Markdown fences. 
2. **Guaranteed Data Types:** If you expect a `number` for a cell coordinate, you get a `number`. The LLM cannot accidentally return `"five"`.
3. **Self-Correction (The Repair Loop):** If the LLM generates output that violates the schema, the schema throws a specific error (e.g., `Expected array for "rows", received string`). You can automatically feed this exact error back to the LLM and say, "You made this mistake, fix it."
4. **UI Confidence:** Your React frontend components can safely assume the data shape is exact, removing the need for defensive coding in your views.

---

## How We Will Write Schemas (The Zod + TypeScript Approach)

We will use **Zod** (which is already installed in your project: `zod: "^4.1.5"`) because it provides runtime validation and statically infers TypeScript types automatically.

### 1. The Formula Intent Schema
Instead of asking the LLM to "return a formula", we force it to explain its reasoning *before* writing the formula to improve accuracy.

```typescript
import { z } from 'zod';

export const FormulaDraftSchema = z.object({
  thought_process: z.string().describe("Briefly explain the logic before writing the formula."),
  formula: z.string().startsWith("=").describe("The actual Excel/Word formula, starting with =."),
  confidence_score: z.number().min(0).max(100),
  assumptions: z.array(z.string()).describe("Any assumptions made about the data.")
});

// Automatically generate the TypeScript type!
export type FormulaDraft = z.infer<typeof FormulaDraftSchema>;
```

### 2. The Audit Finding Schema
This schema forces the LLM to categorize errors and provide quantified business impact.

```typescript
export const AuditFindingSchema = z.object({
  id: z.string(),
  title: z.string().describe("A specific, data-aware title."),
  severity: z.enum(["CRITICAL", "HIGH", "MEDIUM", "LOW"]),
  cell_references: z.array(z.string()).describe("Array of affected cells, e.g., ['A1', 'B2']"),
  business_impact: z.string().describe("Quantified consequence of this error."),
  suggested_fix: z.object({
    action: z.enum(["REWRITE_FORMULA", "DELETE_ROW", "FORMAT_CELL", "NONE"]),
    new_formula: z.string().optional()
  })
});

export const AuditReportSchema = z.object({
  findings: z.array(AuditFindingSchema),
  overall_health_score: z.number().min(0).max(100)
});
```

### 3. The Pivot / Dashboard Schema
This schema dictates exactly how a chart or pivot table should be built via Office.js.

```typescript
export const DashboardConfigSchema = z.object({
  title: z.string(),
  sheet_name_proposal: z.string(),
  pivots: z.array(z.object({
    name: z.string(),
    rows: z.array(z.string()).describe("Column headers to group by."),
    columns: z.array(z.string()).optional(),
    values: z.array(z.object({
      header: z.string(),
      aggregation: z.enum(["SUM", "COUNT", "AVERAGE", "MAX", "MIN"])
    }))
  })),
  charts: z.array(z.object({
    type: z.enum(["BAR", "LINE", "PIE", "SCATTER"]),
    source_pivot_name: z.string()
  }))
});
```

---

## How to Enforce Schemas in the Agent Core

When calling the LLM, you inject the schema description into the prompt, and then run the output through the Zod parser.

```typescript
import { zodToJsonSchema } from "zod-to-json-schema";

async function executeWithSchema(prompt: string, schema: z.ZodTypeAny) {
  // 1. Convert Zod schema to JSON Schema for the LLM prompt
  const jsonSchema = zodToJsonSchema(schema);
  const systemPrompt = `You must output valid JSON matching this schema: ${JSON.stringify(jsonSchema)}`;
  
  // 2. Call the LLM (using JSON mode if available)
  const rawResponse = await callLLM(prompt, systemPrompt);
  
  // 3. Validate
  const parsed = schema.safeParse(JSON.parse(rawResponse));
  
  if (!parsed.success) {
    // 4. If it fails, feed the error back to the LLM for self-correction!
    console.error("Schema Violation:", parsed.error.issues);
    return triggerRepairLoop(prompt, parsed.error.issues);
  }
  
  return parsed.data; // Type-safe data ready for UI
}
```

## Best Practices for LLM Schemas
1. **Use `.describe()` heavily:** Zod's `.describe()` adds metadata to the JSON schema. The LLM reads these descriptions as instructions.
2. **Keep it flat:** LLMs struggle with deeply nested JSON (more than 3 levels deep). Keep schemas as flat as possible.
3. **Use Enums:** If an action can only be `PREVIEW` or `COMMIT`, use `z.enum()`. This prevents the LLM from hallucinating an action like `APPLY_NOW`.
4. **Ask for thoughts first:** Always put a `thought_process` or `reasoning` string field at the *top* of your schema. Because LLMs generate tokens sequentially, forcing it to "think" before outputting the final formula dramatically improves the quality of the formula.
