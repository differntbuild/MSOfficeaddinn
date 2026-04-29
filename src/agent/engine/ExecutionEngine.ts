import { detectIntent } from '../../utils/intelligence';
import { FormulaDraftSchema, ExecutionPlanSchema, ExecutionLogSchema } from '../schemas/formulaExplainSchemas';
import { sanitizeFormulaText, validateFormulaSafety, validateXlookupShape } from '../validators/formulaValidators';
import { validateFormulaSemantics } from '../validators/semanticValidators';

type EngineParams = {
  streamChat: Function;
  host: string;
  domain: string;
  getSystemPromptForIntent: Function;
};

export class ExecutionEngine {
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

  routeIntent(input: string) {
    return detectIntent(input);
  }

  plan(intent: string) {
    return ExecutionPlanSchema.parse({
      intent,
      expectedSchema: intent === 'FORMULA' ? 'FormulaDraftSchema' : 'Unknown',
      steps: [
        { id: 'generate', name: 'GenerateDraft', retries: 2, timeoutMs: 30000 },
        { id: 'validate', name: 'ValidateDraft', retries: 0, timeoutMs: 5000 },
      ],
    });
  }

  private async callFormulaModel(input: string, contextData: any, opts: any = {}) {
    const system = `${this.getSystemPromptForIntent('FORMULA', { host: this.host, domain: this.domain, data: contextData })}

Return ONLY valid JSON with this shape:
{
  "formula": "=XLOOKUP(...)",
  "explanation": "one sentence"
}
No markdown fences.`;

    let full = '';
    await this.streamChat(input, contextData, system, (chunk: string) => {
      full = chunk;
      if (opts.onChunk) opts.onChunk(chunk);
    }, { intent: 'FORMULA', isJson: true });
    return full;
  }

  private parseFormulaDraft(raw: string) {
    let parsed: any = null;
    try {
      parsed = JSON.parse(raw);
    } catch {
      const formulaMatch = raw.match(/`(=[^`]+)`/) || raw.match(/(=[A-Z0-9_:\+\-\*\/\^\(\),."'%!<>=& ]+)/i);
      parsed = {
        formula: formulaMatch ? formulaMatch[1] : '',
        explanation: 'Generated formula draft.',
      };
    }
    parsed.formula = sanitizeFormulaText(parsed.formula || '');
    return FormulaDraftSchema.safeParse(parsed);
  }

  async executeFormula(input: string, contextData: any, hooks: any = {}) {
    const logs: any[] = [];
    logs.push(this.mkLog('router', 'Intent routed.', 'info', { intent: 'FORMULA' }));
    const plan = this.plan('FORMULA');
    logs.push(this.mkLog('planner', 'Plan created.', 'info', { plan }));

    const attempts = [
      { label: 'base', prompt: input, context: contextData },
      { label: 'strict-retry', prompt: `${input}\n\nUse exact selected headers. Return valid XLOOKUP range vectors only.`, context: contextData },
      { label: 'reduced-context', prompt: input, context: { ...contextData, selection: contextData?.selection || contextData, sample: undefined, stats: undefined } },
    ];

    for (let i = 0; i < attempts.length; i++) {
      const attempt = attempts[i];
      logs.push(this.mkLog('runner', `Attempt ${i + 1}: ${attempt.label}`));

      const raw = await this.callFormulaModel(attempt.prompt, attempt.context, { onChunk: hooks.onChunk });
      const parsed = this.parseFormulaDraft(raw);

      if (!parsed.success) {
        logs.push(this.mkLog('validator-schema', 'Schema validation failed.', 'warn', { issues: parsed.error.issues }));
        continue;
      }

      const draft = parsed.data;
      const safety = validateFormulaSafety(draft.formula);
      if (!safety.ok) {
        logs.push(this.mkLog('validator-formula', safety.reason || 'Formula safety failed.', 'warn'));
        continue;
      }

      const shape = validateXlookupShape(draft.formula);
      if (!shape.ok) {
        logs.push(this.mkLog('validator-shape', shape.reason || 'XLOOKUP validation failed.', 'warn'));
        continue;
      }

      const semantic = validateFormulaSemantics(draft.formula, input, contextData?.selection || contextData);
      if (!semantic.ok) {
        logs.push(this.mkLog('validator-semantic', semantic.reason || 'Semantic validation failed.', 'warn'));
        continue;
      }

      logs.push(this.mkLog('commit-gate', 'Draft validated and ready for preview.'));
      return {
        ok: true,
        intent: 'FORMULA',
        plan,
        proposal: draft,
        action: { type: 'FORMULA_PREVIEW', mode: 'preview_first' },
        logs,
      };
    }

    logs.push(this.mkLog('fail-safe', 'All attempts failed; blocked write.', 'error'));
    return {
      ok: false,
      intent: 'FORMULA',
      plan,
      error: 'Could not generate a safe, schema-valid formula after retries.',
      action: { type: 'BLOCK_WRITE' },
      logs,
    };
  }
}
