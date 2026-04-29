import { runLocalAudit, computeHealthScore } from '../../utils/auditEngine';
import { AuditReportSchema } from '../schemas/auditSchemas';
import { ExecutionLogSchema } from '../schemas/formulaExplainSchemas';

type EngineParams = {
  streamChat: Function;
  host: string;
  domain: string;
  getSystemPromptForIntent: Function;
};

export class AuditEngine {
  private streamChat: Function;
  private host: string;
  private domain: string;
  private getSystemPromptForIntent: Function;

  constructor(params: EngineParams) {
    this.streamChat = params.streamChat;
    this.host = params.host;
    this.domain = params.domain;
    this.getSystemPromptForIntent = params.getSystemPromptForIntent;
  }

  private mkLog(stage: string, message: string, level: 'info' | 'warn' | 'error' = 'info', meta: any = {}) {
    return ExecutionLogSchema.parse({
      timestamp: new Date().toISOString(),
      stage,
      message,
      level,
      meta,
    });
  }

  private parseAuditReport(raw: string) {
    try {
      // Attempt to extract JSON if wrapped in markdown
      const match = raw.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/) || raw.match(/(\{[\s\S]*\})/);
      const jsonStr = match ? match[1] : raw;
      const parsed = JSON.parse(jsonStr);
      return AuditReportSchema.safeParse(parsed);
    } catch (e) {
      return { success: false, error: e };
    }
  }

  async executeAudit(contextData: any, hooks: any = {}) {
    const logs: any[] = [];
    logs.push(this.mkLog('audit-start', 'Starting Hybrid Audit execution.'));

    // Step 1: Deterministic Engine
    let localFindings = [];
    try {
      localFindings = runLocalAudit(contextData);
      logs.push(this.mkLog('local-engine', `Deterministic audit complete. Found ${localFindings.length} raw issues.`));
    } catch (e: any) {
      logs.push(this.mkLog('local-engine-error', `Failed to run local audit: ${e.message}`, 'error'));
      return { ok: false, error: 'Local audit engine failed.', logs };
    }

    const baselineScore = computeHealthScore(localFindings);

    // Step 2: Context Gathering for LLM
    const systemPrompt = `You are an Enterprise Data Auditor for ${this.host}.
Domain: ${this.domain}

Your task is to enrich the findings from the local deterministic engine, assess business impact, and return a strict JSON object matching this schema:
{
  "health_score": number (0-100),
  "summary": "Executive summary paragraph",
  "findings": [
    {
      "id": "string",
      "title": "Specific title with numbers",
      "desc": "Detailed issue description",
      "severity": "CRITICAL" | "HIGH" | "MEDIUM" | "LOW",
      "impact": "Business consequence",
      "recommendation": "Actionable fix",
      "action_type": "FIX_FORMULA" | "DELETE_ROWS" | "FORMAT_CELLS" | "INVESTIGATE" | "CREATE_DASHBOARD" | "EXTERNAL_ENRICH",
      "affected_cells": ["A1"],
      "suggested_formula": "Optional Excel formula"
    }
  ]
}

Input Context:
- Raw Findings: ${JSON.stringify(localFindings)}
- Data Stats: ${JSON.stringify(contextData.stats || {})}
- Dimensions: ${contextData.rowCount} rows, ${contextData.columnCount} columns

Rules:
1. ONLY return valid JSON. Do not use markdown code blocks.
2. Maintain the finding IDs from the raw findings.
3. If the dataset is clean, return an empty findings array and a high health score.
4. Provide concrete, actionable recommendations.
5. CRITICAL: ONLY provide a "suggested_formula" if the action is explicitly "FIX_FORMULA". DO NOT invent pseudocode formulas like =TABLE() or =PROPER(ColumnName). If the action is to create a table, delete rows, or investigate, leave "suggested_formula" completely out of the object. All formulas MUST be valid, executable Excel formulas.`;

    const attempts = [
      { label: 'base (fast)', prompt: 'Perform the audit enrichment.', model: 'meta-llama/llama-4-scout-17b-16e-instruct' },
      { label: 'strict-retry (deep)', prompt: 'Your previous response was invalid JSON. You MUST return ONLY a JSON object matching the exact schema.', model: 'llama-3.3-70b-versatile' }
    ];

    // Step 3 & 4: Model Execution and Validation
    for (let i = 0; i < attempts.length; i++) {
      const attempt = attempts[i];
      logs.push(this.mkLog('llm-enrichment', `Attempt ${i + 1}: ${attempt.label}`));
      
      if (hooks.onStatus) hooks.onStatus(`Analyzing business impact... (Attempt ${i + 1})`);

      let rawResponse = '';
      try {
        await this.streamChat(attempt.prompt, null, systemPrompt, (chunk: string) => {
          rawResponse = chunk;
          if (hooks.onChunk) hooks.onChunk(chunk); // Stream raw if needed, but mostly we wait for final
        }, { isJson: true, model: attempt.model });
      } catch (e: any) {
        logs.push(this.mkLog('llm-error', `API error: ${e.message}`, 'error'));
        continue;
      }

      const parsed = this.parseAuditReport(rawResponse);

      if (!parsed.success || !('data' in parsed)) {
        logs.push(this.mkLog('validator-schema', 'Schema validation failed.', 'warn', { issues: parsed.error }));
        continue;
      }

      logs.push(this.mkLog('audit-complete', 'Audit report validated successfully.'));
      
      // Step 5: Action Return
      return {
        ok: true,
        action: { type: 'AUDIT_REPORT_READY' },
        report: parsed.data,
        logs
      };
    }

    logs.push(this.mkLog('fail-safe', 'All LLM attempts failed; falling back to local only.', 'warn'));
    
    // Fallback to local findings formatted manually
    const fallbackReport = {
      health_score: baselineScore,
      summary: `Found ${localFindings.length} issues in the dataset. Detailed AI enrichment failed, showing raw findings.`,
      findings: localFindings.map(f => ({
        id: f.id,
        title: f.title,
        desc: f.desc,
        severity: f.priority.toUpperCase() as "CRITICAL" | "HIGH" | "MEDIUM" | "LOW",
        impact: "Unknown impact (enrichment failed)",
        recommendation: "Review manually",
        action_type: "INVESTIGATE" as const,
        affected_cells: [f.loc]
      }))
    };

    return {
      ok: true,
      action: { type: 'AUDIT_REPORT_READY' },
      report: fallbackReport,
      logs
    };
  }
}
