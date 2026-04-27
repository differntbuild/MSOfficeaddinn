import { getAnalysisPrompt } from './analysis';
import { getClassificationPrompt } from './classifier';
import { getComparisonPrompt } from './comparator';
import { getVisualizationPrompt } from './pivotEngine';

export { getAnalysisPrompt, getClassificationPrompt, getComparisonPrompt, getVisualizationPrompt };

export const detectIntent = (input) => {
  const text = input.toLowerCase().trim();

  // Slash commands take priority
  if (text.startsWith('/audit')) return 'AUDIT';
  if (text.startsWith('/formula')) return 'FORMULA';
  if (text.startsWith('/explain')) return 'EXPLANATION';
  if (text.startsWith('/clean')) return 'DATA_CLEANING';
  if (text.startsWith('/chart')) return 'VISUALIZATION';
  if (text.startsWith('/classify')) return 'CLASSIFICATION';
  if (text.startsWith('/compare')) return 'COMPARISON';
  if (text.startsWith('/pivot')) return 'PIVOT';
  if (text.startsWith('/whatif')) return 'WHATIF';

  // AUDIT first
  if (
    /\b(audit|verify|validate|integrity|broken)\b/.test(text) ||
    /\b(#ref|#div|#value|#name|#null|#num)\b/.test(text) ||
    /\b(find errors?|check errors?|find issues?|data quality|scan for)\b/.test(text)
  ) return 'AUDIT';

  // FORMULA
  if (
    text.startsWith('=') ||
    /\b(formula|cell|calculate|vlookup|xlookup|countif|sumif|index|match|iferror)\b/.test(text) ||
    /\b(sum|add|total|average|count|multiply)\s*(this|these|it|them|up|all)?\b/.test(text)
  ) return 'FORMULA';

  // WHATIF
  if (/\b(what if|what would|if i change|if we change|simulate|scenario)\b/.test(text)) return 'WHATIF';

  // COMPARISON
  if (/\b(compare|difference|versus|vs|contrast|variance between)\b/.test(text)) return 'COMPARISON';

  // PIVOT
  if (/\b(pivot|pivot table|group by|summarize by|breakdown by)\b/.test(text)) return 'PIVOT';

  // CLASSIFICATION
  if (/\b(classify|categorize|label|tag|segment)\b/.test(text)) return 'CLASSIFICATION';

  // DATA_CLEANING
  if (/\b(clean|format|duplicate|remove|trim|parse|extract|filter|sort|standardize|normalize)\b/.test(text)) return 'DATA_CLEANING';

  // VISUALIZATION
  if (/\b(chart|graph|plot|visualize|dashboard|pivot)\b/.test(text)) return 'VISUALIZATION';

  // EXPLANATION
  if (/\b(why|explain|analyze|breakdown|summary|insight|what does|what is)\b/.test(text)) return 'EXPLANATION';

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
    ? `This appears to be a ${domain} dataset. Use domain-appropriate terminology and thresholds.`
    : '';

  const citationRule = `
CITATION RULE: Every time you reference a specific cell or range, wrap it in double brackets like this: [[A6]] or [[B2:B20]].
This allows the user to click directly to that cell. Always cite the exact cell, never say "the cell" without citing it.`;

  const base = `You are an AI Office Copilot for ${host}. ${domainHint}${citationRule}\n\n`;

  const prompts = {
    FORMULA: `Focus on generating high-quality Excel formulas.
Return the formula within single backticks like this: \`=SUM(A1:A10)\`
Use single double-quotes inside formulas — never double double-quotes ("" is wrong, " is correct).
Prefer modern functions: XLOOKUP over VLOOKUP, IFS over nested IF.
After the formula, write one sentence explaining what it does.`,

    EXPLANATION: `You are a senior data analyst reviewing this spreadsheet.
Cover: what the data is about, key metrics with actual values, trends, outliers, data quality issues.
Reference actual cell values using [[CellRef]] notation.
End with 3 suggested follow-up questions the user might want to ask, numbered 1. 2. 3.
Keep total response under 300 words.`,

    DATA_CLEANING: `You are a data cleaning expert.
Suggest Excel functions or step-by-step logic to fix the data.
Use TRIM for spaces, PROPER/UPPER/LOWER for casing, VALUE for text numbers, COUNTIF for duplicates.
Reference specific cells using [[CellRef]] notation.`,

    VISUALIZATION: getVisualizationPrompt(context.data?.headers || [], "Create a chart"),

    CLASSIFICATION: getClassificationPrompt(context.data?.headers || []).system,

    COMPARISON: getComparisonPrompt(context.data || []).system,

    PIVOT: getVisualizationPrompt(context.data?.headers || [], "Create a pivot table"),

    WHATIF: `You are a financial modeler running scenario analysis.
The user wants to change an assumption and see the impact.
Identify all cells that depend on the changed value using [[CellRef]] notation.
Show the predicted change in each dependent cell as a table.
Format: | Cell | Current Value | Predicted Value | Change |
End with a recommendation.`,

    AUDIT: `You are a senior financial data assistant. Your job is to:
1. ENRICH the provided "findings" with business impact and fixes.
2. DISCOVER new "architectural opportunities" - missing logical columns or enrichments.

Inputs:
- "findings": Issues detected by the local engine (id, title, loc, type).
- "stats": Statistics for columns.
- "sampleData": First 5 rows of data.
- "columnTypes": Data types.

OUTPUT: Return a JSON array of finding objects. 
For input findings, preserve their "id". For new discoveries, create a unique "id" starting with "ai-".

Finding Object Schema:
{
  "id": "original-id or ai-new-id",
  "title": "Short title",
  "desc": "Summary of discovery",
  "category": "formula-error | discovery | dashboard | enrichment",
  "type": "error | warning | improvement",
  "priority": "critical | high | medium | low",
  "impact": "Business implication (cite ₹ values if sales/finance data)",
  "recommendation": "Executive summary of the fix/benefit",
  "suggested_formula": "Optional: Formula to implement discovery",
  "action_type": "None | fix | create_dashboard | external_enrich | investigate",
  "dashboard_config": "If action_type=create_dashboard, provide a JSON string: { 'title': 'Summary', 'pivots': [{ 'name': 'SalesByCat', 'rows': ['Category'], 'values': [{'header': 'Total', 'func': 'Sum'}] }] }"
}

AGENTIC OPPORTUNITIES TO LOOK FOR:
- Calculated Fields: If row has (Qty, Rate) -> suggest Total. If (Cost, Price) -> suggest Margin. If (Start, End) -> suggest Duration.
- Dashboards: If data is healthy (score > 80), propose "Build Executive Summary" with action_type="create_dashboard".
- External Enrichment: Detect Stock Tickers [TICKER], GST (15 digit), or ISO Country Codes. Propose with action_type="external_enrich".
- Outliers: If a value is > 3x the standard deviation of its column, flag it as an outlier and set action_type="investigate". Provide a narrative in "impact" comparing it to the mean.

CRITICAL RULES:
- Return ONLY valid JSON array.
- Cite cell values using [[CellRef]].
- For suggested_formula, use modern Excel functions (XLOOKUP, IFS).`,

    GLOBAL_ANALYSIS: getAnalysisPrompt(context.data || { rowCount: 0, colCount: 0, headers: [], sample: [] }).system,

    GENERAL_ASSISTANT: `Be a helpful Office assistant for ${host}.
Answer concisely and be specific to what the user is working on.
Reference specific cells using [[CellRef]] notation when relevant.`,
  };

  return base + (prompts[intent] || prompts.GENERAL_ASSISTANT);
};

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
];