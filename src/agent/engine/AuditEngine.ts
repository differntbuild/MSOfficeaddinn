import { runLocalAudit, computeHealthScore, generateDataProfile } from '../../utils/auditEngine';
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

  private normalizePythonCode(code: string): string {
    const codeMatch = code.match(/```python\s*([\s\S]*?)```/) || code.match(/```\s*([\s\S]*?)```/);
    let clean = codeMatch ? codeMatch[1] : code;
    const lines = clean.split('\n');
    return lines.map(line => line.replace(/\t/g, '    ')).join('\n').trim();
  }

  private parseAuditReport(raw: string, localFindings: any[]) {
    try {
      const match = raw.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/) || raw.match(/(\{[\s\S]*\})/);
      const jsonStr = match ? match[1] : raw;
      let parsed = JSON.parse(jsonStr);

      if (parsed.findings && Array.isArray(parsed.findings)) {
        parsed.findings = parsed.findings.map((f: any) => {
          const localMatch = localFindings.find(lf => lf.id === f.id);

          if (localMatch) {
            f.id = localMatch.id;
            f.loc = f.loc || localMatch.loc;
            f.category = f.category || localMatch.category;
            f.type = f.type || localMatch.type;
            f.affectedCount = localMatch.affectedCount;
          }

          f.priority = (f.severity || f.priority || 'medium').toLowerCase();
          f.loc = f.loc || (Array.isArray(f.affected_cells) && f.affected_cells.length > 0 ? f.affected_cells[0] : '—');
          f.affectedCount = f.affectedCount ?? (Array.isArray(f.affected_cells) ? f.affected_cells.length : 0);
          f.category = f.category || 'insight';
          f.type = f.type || 'improvement';
          f.id = f.id || `ai-${Math.random().toString(36).substr(2, 9)}`;

          return f;
        });
      }

      return AuditReportSchema.safeParse(parsed);
    } catch (e) {
      return { success: false, error: e };
    }
  }

  async executeAudit(contextData: any, hooks: any = {}) {
    const logs: any[] = [];
    logs.push(this.mkLog('audit-start', 'Starting Resilient Audit execution.'));

    // Step 1: Local audit — always runs, never skipped
    const profile = generateDataProfile(contextData);
    const localFindings = runLocalAudit(contextData);
    const baselineScore = computeHealthScore(localFindings);
    logs.push(this.mkLog('profiler', 'Local findings ready.'));

    // Step 2: Coder pass — optional, failure is non-fatal
    let pythonResults: any = { status: 'Skipped' };
    if (hooks.runPython) {
      if (hooks.onStatus) hooks.onStatus('Generating analytical strategy...');
      try {
        const coderPrompt = `You are an expert Data Scientist. Analyze this Data Profile and write a Python (Pandas) function to identify business anomalies.

Data Profile:
${JSON.stringify(profile, null, 2)}

Requirements:
1. Define: \`def execute_analysis(df):\`
2. Return a dictionary of findings.
3. Use 4-space indentation consistently.
4. ONLY return the code.`;

        let pythonCode = '';
        await this.streamChat(coderPrompt, null, "You are a Python Data Analyst.", (chunk: string) => {
          pythonCode = chunk;
        }, { model: 'llama-3.3-70b-versatile' });
        
        pythonCode = this.normalizePythonCode(pythonCode);

        if (hooks.onStatus) hooks.onStatus('Executing local Python analysis...');
        const raw = await hooks.runPython(pythonCode, contextData);
        pythonResults = raw?.error ? { status: 'Python analysis unavailable' } : raw;
        logs.push(this.mkLog('python-execution', 'Python analysis complete.'));
      } catch (e: any) {
        console.warn('AuditEngine Coder Error:', e);
        logs.push(this.mkLog('coder-warn', `Python enrichment skipped: ${e.message}`, 'warn', { error: e }));
      }
    }

    // Step 3: Narrator pass — optional, failure falls back to local findings
    if (hooks.onStatus) hooks.onStatus('Synthesizing executive report...');
    const truncatedLocal = localFindings.slice(0, 20);
    
    // Suppress Python errors from becoming findings
    const safeResults = pythonResults?.error || pythonResults?.status === 'Python analysis unavailable'
      ? { status: 'Python analysis unavailable' } 
      : pythonResults;

    const narratorPrompt = `You are a Senior Auditor. Synthesize findings into a professional report.

Local Findings (PRESERVE IDs/LOCATIONS):
${JSON.stringify(truncatedLocal, null, 2)}

Python Analytical Results:
${JSON.stringify(safeResults, null, 2)}

Health Score: ${baselineScore}

VALID ENUMS:
Severity: CRITICAL, HIGH, MEDIUM, LOW
Action Type: FIX_FORMULA, DELETE_ROWS, FORMAT_CELLS, INVESTIGATE, CREATE_DASHBOARD

INSTRUCTIONS:
1. Include ALL findings from 'Local Findings'. Use their 'id', 'loc', 'category', and 'type'.
2. Enrich 'title' and 'desc' using 'Python Results' if they correlate.
3. If 'Python Results' found NEW issues, add them with unique IDs.
4. If 'Python Results' contains an 'error' key, DO NOT create a finding for it. Ignore it.
5. Return a strict JSON object:
{
  "health_score": number (0-100),
  "summary": "Executive summary paragraph",
  "findings": [
    {
      "id": "string",
      "title": "Finding title",
      "desc": "Detailed description",
      "severity": "CRITICAL|HIGH|MEDIUM|LOW",
      "impact": "Business impact",
      "recommendation": "Specific fix",
      "action_type": "ENUM_VALUE_FROM_LIST",
      "affected_cells": ["A1"]
    }
  ]
}`;

    let rawResponse = '';
    try {
      await this.streamChat(narratorPrompt, null, "You are an Enterprise Auditor. Return ONLY JSON.", (chunk: string) => {
        rawResponse = chunk;
      }, { model: 'llama-3.3-70b-versatile' }); 
    } catch (e: any) {
      console.warn('AuditEngine Narrator Error:', e);
      logs.push(this.mkLog('narrator-warn', `Narrator synthesis skipped: ${e.message}`, 'warn', { error: e }));
    }

    const parsed = rawResponse ? this.parseAuditReport(rawResponse, localFindings) : { success: false };

    // Always return ok: true — local findings are always valid
    const findings = (parsed.success && 'data' in parsed)
      ? (parsed.data as any).findings
      : localFindings.map((f: any) => ({
          ...f,
          severity: (f.priority || 'medium').toUpperCase(),
          impact: 'Potential data integrity issue.',
          recommendation: 'Review the affected cells for accuracy.',
          action_type: 'INVESTIGATE',
          affected_cells: f.loc ? [f.loc] : [],
          aiEnriched: false,
        }));

    const health_score = (parsed.success && 'data' in parsed && (parsed.data as any).health_score)
      ? (parsed.data as any).health_score
      : baselineScore;

    const summary = (parsed.success && 'data' in parsed)
      ? (parsed.data as any).summary
      : `Local audit complete. ${localFindings.length} findings detected. AI enrichment unavailable.`;

    return {
      ok: true,
      action: { type: 'AUDIT_REPORT_READY' },
      report: { health_score, summary, findings },
      logs,
    };
  }
}
