/**
 * analysis.js — Global analysis prompt & column statistics engine
 *
 * The getAnalysisPrompt is used for the GLOBAL_ANALYSIS intent (full-sheet
 * deep-dive) and is also re-used as a context block inside the AUDIT enrichment.
 * It must produce a professional, quantified narrative — not a generic checklist.
 */

export const getAnalysisPrompt = (sheetData) => {
  // Build a richer context block so the AI has enough to work with
  const statsBlock = sheetData.stats
    ? Object.entries(sheetData.stats)
        .map(([col, s]) =>
          `  "${col}": mean=${s.mean?.toFixed(2)}, min=${s.min}, max=${s.max}, stdDev=${s.stdDev?.toFixed(2)}, n=${s.count}`
        )
        .join('\n')
    : '  (no numeric columns detected)';

  return {
    system: `You are a Principal Data Analyst conducting a first-look investigation of a business dataset.

DATASET CONTEXT:
  Rows: ${sheetData.rowCount} (${sheetData.rowCount - 1} data rows + 1 header)
  Columns: ${sheetData.colCount}
  Headers: ${JSON.stringify(sheetData.headers)}

COLUMN STATISTICS (numeric columns only):
${statsBlock}

SAMPLE DATA (first 5 rows):
${JSON.stringify(sheetData.sample, null, 2)}

════════════════════════════════════
YOUR INVESTIGATION FRAMEWORK
════════════════════════════════════

1. DATASET SUMMARY
   - What does this dataset represent? Who likely owns it and what decisions does it support?
   - State the row count, column count, and the date range (if dates exist).

2. KEY METRICS WITH ACTUAL VALUES
   - Pick the 3-5 most important numeric columns.
   - For each: state the mean, the range (min–max), and whether the spread (stdDev) is wide or tight.
   - Highlight any column where max/mean > 3 — that is a red flag for outliers.

3. DATA QUALITY FINDINGS
   - Anomalies: values that break the column's expected pattern (e.g., text in a number column).
   - Outliers: values more than 2.5 standard deviations from the mean. Cite the actual value and the mean.
   - Missing data: identify columns with blank cells. State the count and percentage.
   - Structural issues: duplicate rows, missing headers, mixed types.

4. CLEAN-UP PLAN
   Provide a prioritized action list:
   | Priority | Issue | Fix | Formula |
   Each fix must reference specific columns or cell ranges.

5. OPPORTUNITY INSIGHTS
   - What derived columns would make this data more useful? (e.g., if Qty and Rate exist → suggest Total)
   - What aggregation would answer the most obvious business question this data implies?

════════════════════════════════════
WRITING STANDARDS
════════════════════════════════════
- Use a professional, direct analyst voice — not conversational.
- Every claim must cite a concrete number. "Revenue varies widely" is not acceptable. "Revenue ranges from ₹4,200 to ₹4,80,000 with a mean of ₹52,300 and stdDev of ₹38,700" is.
- Reference columns by their actual header name, not column letters.
- Keep the report under 500 words total.
- Format with clean Markdown. Use tables where appropriate.`,

    user: `Conduct a full dataset investigation. Write the report in Markdown.
Be specific. Cite actual values. Identify the most critical data quality risk first.`,
  };
};

/**
 * buildColumnStats — computes mean, min, max, stdDev, median, and
 * a simple outlier list for every numeric column in the sheet.
 *
 * @param {Array<Array>} vals  — 2D array where vals[0] is the header row
 * @returns {Object}           — keyed by header name
 */
export const buildColumnStats = (vals) => {
  const stats = {};
  if (!vals || vals.length < 2) return stats;

  const headers = vals[0];

  headers.forEach((header, colIdx) => {
    if (!header) return; // skip unlabelled columns

    const colVals = vals
      .slice(1)
      .map((row) => row[colIdx])
      .filter((v) => typeof v === 'number' && isFinite(v));

    if (colVals.length === 0) return;

    const n = colVals.length;
    const sum = colVals.reduce((a, b) => a + b, 0);
    const mean = sum / n;
    const min = Math.min(...colVals);
    const max = Math.max(...colVals);

    // Population standard deviation
    const variance = colVals.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / n;
    const stdDev = Math.sqrt(variance);

    // Median
    const sorted = [...colVals].sort((a, b) => a - b);
    const mid = Math.floor(n / 2);
    const median = n % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];

    // Outlier list (Z-score > 2.8)
    const Z_THRESHOLD = 2.8;
    const outliers = stdDev > 0
      ? colVals
          .map((v, i) => ({ val: v, rowIdx: i + 1, z: Math.abs((v - mean) / stdDev) }))
          .filter((o) => o.z > Z_THRESHOLD)
          .slice(0, 5) // cap at 5 for context window efficiency
      : [];

    stats[header] = {
      mean,
      median,
      min,
      max,
      stdDev,
      count: n,
      totalRows: vals.length - 1,
      blankCount: (vals.length - 1) - n,
      blankPct: Math.round(((vals.length - 1 - n) / (vals.length - 1)) * 100),
      outliers,
    };
  });

  return stats;
};

/**
 * inferColumnTypes — classifies each column as 'numeric', 'date', 'boolean', or 'text'.
 * Used by the audit engine and passed to the AI as context.
 *
 * @param {Array<Array>} vals — 2D array with header row at index 0
 * @returns {Object}          — keyed by header name
 */
export const inferColumnTypes = (vals) => {
  const types = {};
  if (!vals || vals.length < 2) return types;

  const headers = vals[0];
  const DATE_RE = /^\d{1,4}[-/]\d{1,2}[-/]\d{1,4}$/;
  const BOOL_SET = new Set(['true', 'false', 'yes', 'no', '1', '0', 'y', 'n']);

  headers.forEach((header, colIdx) => {
    if (!header) return;

    const sample = vals
      .slice(1, 21) // check up to 20 rows
      .map((row) => row[colIdx])
      .filter((v) => v !== null && v !== undefined && v !== '');

    if (sample.length === 0) {
      types[header] = 'empty';
      return;
    }

    const numericCount = sample.filter((v) => typeof v === 'number').length;
    const dateCount = sample.filter(
      (v) => typeof v === 'string' && DATE_RE.test(v.trim())
    ).length;
    const boolCount = sample.filter(
      (v) => typeof v === 'string' && BOOL_SET.has(v.toLowerCase().trim())
    ).length;

    const ratio = (count) => count / sample.length;

    if (ratio(numericCount) >= 0.8) types[header] = 'numeric';
    else if (ratio(dateCount) >= 0.6) types[header] = 'date';
    else if (ratio(boolCount) >= 0.8) types[header] = 'boolean';
    else types[header] = 'text';
  });

  return types;
};