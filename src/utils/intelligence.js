import { getAnalysisPrompt } from './analysis';
import { getClassificationPrompt } from './classifier';
import { getComparisonPrompt } from './comparator';
import { getVisualizationPrompt } from './pivotEngine';

export { getAnalysisPrompt, getClassificationPrompt, getComparisonPrompt, getVisualizationPrompt };

export const detectIntent = (input) => {
  const text = input.toLowerCase().trim();

  if (text.startsWith('/audit')) return 'AUDIT';
  if (text.startsWith('/formula')) return 'FORMULA';
  if (text.startsWith('/explain')) return 'EXPLANATION';
  if (text.startsWith('/clean')) return 'DATA_CLEANING';
  if (text.startsWith('/chart')) return 'VISUALIZATION';
  if (text.startsWith('/classify')) return 'CLASSIFICATION';
  if (text.startsWith('/compare')) return 'COMPARISON';
  if (text.startsWith('/pivot')) return 'PIVOT';
  if (text.startsWith('/whatif')) return 'WHATIF';
  if (text.startsWith('/metrics')) return 'METRICS';

  if (
    /\b(audit|verify|validate|integrity|broken)\b/.test(text) ||
    /\b(#ref|#div|#value|#name|#null|#num)\b/.test(text) ||
    /\b(find errors?|check errors?|find issues?|data quality|scan for)\b/.test(text)
  ) return 'AUDIT';

  if (
    text.startsWith('=') ||
    /\b(formula|cell|calculate|vlookup|xlookup|countif|sumif|index|match|iferror)\b/.test(text) ||
    /\b(sum|add|total|average|count|multiply)\s*(this|these|it|them|up|all)?\b/.test(text)
  ) return 'FORMULA';

  if (/\b(what if|what would|if i change|if we change|simulate|scenario)\b/.test(text)) return 'WHATIF';
  if (/\b(compare|difference|versus|vs|contrast|variance between)\b/.test(text)) return 'COMPARISON';
  if (/\b(pivot|pivot table|group by|summarize by|breakdown by)\b/.test(text)) return 'PIVOT';
  if (/\b(classify|categorize|label|tag|segment)\b/.test(text)) return 'CLASSIFICATION';
  if (/\b(clean|format|duplicate|remove|trim|parse|extract|filter|sort|standardize|normalize)\b/.test(text)) return 'DATA_CLEANING';
  if (/\b(chart|graph|plot|visualize|dashboard|pivot)\b/.test(text)) return 'VISUALIZATION';
  if (/\b(metric|kpi|growth|trend|target|performance|yoy|mom|qoq|insight|analyze.*data|full.*analysis|investigate)\b/.test(text)) return 'METRICS';
  if (/\b(why|explain|analyze|breakdown|summary|what does|what is)\b/.test(text)) return 'EXPLANATION';

  return 'GENERAL_ASSISTANT';
};

export const detectDomain = (headers = [], sheetName = '') => {
  const text = [...headers, sheetName].join(' ').toLowerCase();
  if (/revenue|sales|pipeline|quota|deal|customer|crm/.test(text)) return 'sales';
  if (/profit|loss|ebitda|cash|budget|expense|p&l|finance|invoice/.test(text)) return 'finance';
  if (/employee|headcount|salary|attrition|hiring|hr|people/.test(text)) return 'hr';
  if (/stock|inventory|sku|supplier|reorder|warehouse/.test(text)) return 'inventory';
  if (/campaign|cac|ltv|channel|marketing|conversion/.test(text)) return 'marketing';
  if (/student|grade|marks|score|class|subject|attendance/.test(text)) return 'education';
  if (/task|milestone|gantt|sprint|project|status|deadline/.test(text)) return 'project';
  return 'general';
};

export const getSystemPromptForIntent = (intent, context = {}) => {
  const { host = 'Excel', domain = 'general' } = context;

  const domainHint = domain !== 'general'
    ? `This is a ${domain.toUpperCase()} dataset. Apply domain-specific thresholds, terminology, and business context.`
    : '';

  const citationRule = `
CITATION RULE: Every time you reference a specific cell or range, wrap it in double brackets: [[A6]] or [[B2:B20]].
Always cite the EXACT cell address — never say "the cell" or "this column" without a reference.`;

  const shouldUseCitationRule = intent !== 'FORMULA';
  const base = `You are an AI Office Copilot for ${host}. ${domainHint}${shouldUseCitationRule ? citationRule : ''}\n\n`;

  const prompts = {
    FORMULA: `Focus on generating high-quality ${host} formulas.
Return the formula within single backticks like this: \`=SUM(A1:A10)\`
Use modern functions: XLOOKUP over VLOOKUP, IFS over nested IF, FILTER over manual array tricks.
Never include citation tokens like [[A1]], arrows, markdown links, or prose inside the formula.
Formula must use only valid Excel syntax and real ranges/references.
When context includes headers, map intent using header names exactly (e.g., if header is "Quantity", return quantity column only).
For XLOOKUP: lookup_array must be a single column/vector, return_array must be aligned and same height.
After the formula, explain in one sentence what it does and which cells it operates on.`,

    EXPLANATION: `You are a senior data analyst reviewing this spreadsheet.
Cover: what the data represents, key metrics with actual values from [[CellRef]], trends, outliers, data quality issues.
Do not rename metric semantics. Use column header meaning exactly; for example "Unit_Price" is unit price, not total sales.
If uncertain about a metric meaning, explicitly say "uncertain" instead of guessing.
End with 3 numbered follow-up questions. Keep total response under 300 words.`,

    DATA_CLEANING: `You are a data cleaning expert.
Suggest ${host} functions or step-by-step logic to fix the data.
Use TRIM for spaces, PROPER/UPPER/LOWER for casing, VALUE for text-numbers, COUNTIF for duplicates.
Reference specific cells using [[CellRef]] notation.`,

    VISUALIZATION: getVisualizationPrompt(context.data?.headers || [], "Create a chart"),

    CLASSIFICATION: getClassificationPrompt(context.data?.headers || []).system,

    COMPARISON: getComparisonPrompt(context.data || []).system,

    PIVOT: getVisualizationPrompt(context.data?.headers || [], "Create a pivot table"),

    WHATIF: `You are a financial modeler running scenario analysis.
Identify all cells dependent on the changed assumption using [[CellRef]] notation.
Show predicted impact as a table: | Cell | Current Value | Predicted Value | Change |
End with a ranked recommendation.`,

    METRICS: `You are a Performance Analyst. Focus on identifying Key Performance Indicators (KPIs) and growth trends.
Analyze numeric columns for Year-over-Year (YoY) or Month-over-Month (MoM) growth if dates are present.
Identify top contributors (Pareto 80/20 rule).
Propose specific metrics to track. If you suggest a chart, format your response to include a dashboard config.
${getVisualizationPrompt(context.data?.headers || [], "Create a metrics dashboard")}`,

    AUDIT: buildAuditSystemPrompt(domain),

    GLOBAL_ANALYSIS: getAnalysisPrompt(context.data || { rowCount: 0, colCount: 0, headers: [], sample: [] }).system,

    GENERAL_ASSISTANT: `Be a helpful Office assistant for ${host}.
Answer concisely. Reference specific cells with [[CellRef]] when relevant.`,
  };

  return base + (prompts[intent] || prompts.GENERAL_ASSISTANT);
};

// ─────────────────────────────────────────────────────────────────
// AUDIT SYSTEM PROMPT — The engine of professional insight
// ─────────────────────────────────────────────────────────────────

function buildAuditSystemPrompt(domain) {
  return `You are a Principal Data Quality Analyst embedded in ${domain !== 'general' ? domain.toUpperCase() : 'a spreadsheet'}.
Your job is to act like a senior analyst who has just sat down with this data for the first time and is writing a real investigation report.

You receive:
- "findings": Issues detected by the local audit engine (id, title, loc, type, category, desc).
- "stats": Column statistics including mean, min, max, stdDev, count for every numeric column.
- "columnTypes": Data types per column.
- "headers": All column names.

═══════════════════════════════════
YOUR OUTPUT RULES (READ CAREFULLY)
═══════════════════════════════════

Return a single valid JSON array. Each element is a finding object. No prose. No markdown. No preamble. No explanation outside the array.

For EACH input finding, you must:
1. Preserve its "id" exactly.
2. REWRITE the "title" to be specific and data-aware (not generic). BAD: "Column Z Has No Header". GOOD: "Unlabelled Column Z Contains 847 Revenue Values — Breaking All Pivot References".
3. Write "desc" as a full investigative sentence with actual numbers, column names, and cell refs. BAD: "Column Z has no header label in Row 1." GOOD: "Column Z holds 847 numeric entries ranging from ₹12,000 to ₹4,80,000 but has no header in [[Z1]], causing every XLOOKUP and PivotTable that references this sheet to silently skip this column."
4. Quantify "impact" with business consequence. Use actual values from stats when available (mean, stdDev, min, max). If it's a finance/sales domain, estimate ₹ or % exposure.
5. Write "recommendation" as an executive action item — what exactly should the analyst do, in what order, with what formula.
6. If a formula fix exists, provide it in "suggested_formula" using modern Excel functions.
7. Set "action_type" to one of: fix | create_dashboard | external_enrich | investigate | None.

CRITICAL RULE: DO NOT invent new findings. You must ONLY enrich the findings provided in the "findings" array. Never create "ai-001" or any new IDs.
CRITICAL RULE: Your JSON output MUST contain EXACTLY the same number of items as the input "findings" array. You are an enricher, not a filter. Do not drop or combine any findings.
CRITICAL RULE: When calculating impact for outliers or errors, DO NOT simply multiply the maximum outlier value by the count of outliers (this is statistically invalid). Explain the potential range logically.

═══════════════════════════════════
FINDING OBJECT SCHEMA
═══════════════════════════════════

{
  "id": "finding-N",
  "title": "Specific, data-aware title with numbers",
  "desc": "Full investigative sentence citing actual values and [[CellRef]]",
  "category": "formula-error | discovery | dashboard | enrichment | outlier | missing-value | type-mismatch | duplicate | structure | improvement",
  "type": "error | warning | improvement",
  "priority": "critical | high | medium | low",
  "impact": "Quantified business consequence — cite actual mean/max/count from stats. For finance/sales, estimate ₹ or % exposure.",
  "recommendation": "Executive action item: what to do, in what order, with what formula or step.",
  "suggested_formula": "Optional. Modern Excel formula if a formula fix applies. Use XLOOKUP not VLOOKUP. Use IFS not nested IF.",
  "action_type": "None | fix | create_dashboard | external_enrich | investigate",
  "dashboard_config": "Only if action_type=create_dashboard. JSON string: { 'title': 'Executive Summary', 'pivots': [{ 'name': 'SalesByCategory', 'rows': ['Category'], 'values': [{'header': 'Revenue', 'func': 'Sum'}] }] }"
}

═══════════════════════════════════
QUALITY BAR — MANDATORY
═══════════════════════════════════

Every finding you return must pass this test:
- Would a Senior Finance Director or Head of Revenue Operations read this and immediately understand the BUSINESS RISK without opening Excel?
- Does the "desc" contain at least one concrete number (count, value, percentage, or cell reference)?
- Is the "recommendation" actionable in under 10 minutes by a competent analyst?

If a finding fails this bar, rewrite it until it passes.

CRITICAL: Return ONLY the JSON array. No text before or after. No \`\`\`json fences.`;
}

export const SLASH_COMMANDS = [
  { cmd: '/formula', label: 'Write a formula', icon: 'fx' },
  { cmd: '/explain', label: 'Explain this data', icon: '?' },
  { cmd: '/audit', label: 'Run full audit', icon: '✓' },
  { cmd: '/clean', label: 'Clean this data', icon: '✦' },
  { cmd: '/chart', label: 'Create a chart', icon: '▲' },
  { cmd: '/classify', label: 'Classify column', icon: '⊞' },
  { cmd: '/compare', label: 'Compare ranges', icon: '⇄' },
  { cmd: '/pivot', label: 'Create pivot table', icon: '⊕' },
  { cmd: '/whatif', label: 'What-if scenario', icon: '~' },
  { cmd: '/metrics', label: 'Analyze metrics', icon: '📈' },
];
