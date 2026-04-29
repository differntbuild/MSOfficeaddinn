/**
 * auditEngine.js — Local-first analysis engine (Phase 1)
 *
 * Runs all detection checks in pure JS — no API, instant results.
 * Every finding produced here must be specific enough to be useful on its own,
 * BEFORE the AI enrichment layer (Phase 2) adds business context.
 *
 * Quality bar for every desc field:
 *   - Must contain at least one concrete number (count, value, cell address).
 *   - Must explain the CONSEQUENCE, not just the symptom.
 *   - Must be readable by a non-technical stakeholder.
 */

// ─── Helpers ────────────────────────────────────────────────────

const colLetter = (index) => {
  let letter = '', i = index;
  while (i >= 0) {
    letter = String.fromCharCode((i % 26) + 65) + letter;
    i = Math.floor(i / 26) - 1;
  }
  return letter;
};

const cellAddr = (row, col) => `${colLetter(col)}${row + 1}`;

const fmt = (v) => {
  if (typeof v !== 'number') return String(v);
  return Math.abs(v) >= 10000
    ? v.toLocaleString('en-IN')
    : v % 1 === 0
    ? String(v)
    : v.toFixed(2);
};

const pct = (part, total) => total > 0 ? `${Math.round((part / total) * 100)}%` : '0%';

/**
 * Prunes phantom columns at the end of the used range.
 * A column is "active" if it has a header OR at least one non-empty data cell.
 */
function pruneNoisyColumns(data) {
  let lastActiveCol = -1;
  for (let c = 0; c < data.columnCount; c++) {
    const header = (data.headers[c] || '').toString().trim();
    const hasHeader = header !== '';
    let hasData = false;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val !== '' && val !== null && val !== undefined && String(val).trim() !== '') { 
        hasData = true; 
        break; 
      }
    }
    if (hasHeader || hasData) lastActiveCol = c;
  }
  if (lastActiveCol === -1) return data;
  const pruned = lastActiveCol + 1;
  if (pruned === data.columnCount) return data;
  return { ...data, columnCount: pruned, headers: data.headers.slice(0, pruned) };
}

// ─── Constants ──────────────────────────────────────────────────

const Z_THRESHOLD = 2.8;

const NON_NEGATIVE_PATTERNS = /\b(qty|quantity|price|amount|revenue|cost|cogs|age|count|units|sales|profit|stock|inventory|rate|score|hours|days|months|years)\b/i;
const KEY_COL_PATTERNS = /\b(id|code|key|no|num|number|ref|reference|sku|order|invoice|employee|emp|customer|cust|product|serial)\b/i;
const VOLATILE_FUNCS = /\b(NOW|TODAY|RAND|RANDBETWEEN|OFFSET|INDIRECT)\s*\(/i;
const MAGIC_NUMBER_PATTERN = /[+\-*/,]\s*(0\.\d+|\d{1,3}(?!\d))\s*(?:[+\-*/,)]|$)/;
const OLD_LOOKUP = /\bVLOOKUP\s*\(/i;
const NO_IFERROR = /^(?!.*IFERROR).*\b(VLOOKUP|XLOOKUP|INDEX|MATCH)\s*\(/i;

// ─── Check 1: Formula Errors ────────────────────────────────────

function checkFormulaErrors(data) {
  const findings = [];
  const ERRORS = ['#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NULL!', '#NUM!', '#N/A'];

  // Group errors by type and column for concise, high-signal reporting
  const byTypeCol = {};
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const val = data.values[r][c];
      if (typeof val === 'string' && ERRORS.includes(val.trim())) {
        const key = `${val.trim()}|${c}`;
        if (!byTypeCol[key]) byTypeCol[key] = { errType: val.trim(), col: c, cells: [] };
        byTypeCol[key].cells.push(cellAddr(r, c));
      }
    }
  }

  Object.values(byTypeCol).forEach(({ errType, col, cells }) => {
    const header = data.headers[col] || colLetter(col);
    const count = cells.length;
    const example = cells[0];
    const downstream = errType === '#REF!' ? 'any formula that references this cell will also break' :
                       errType === '#DIV/0!' ? 'division-by-zero — the denominator cell is blank or zero' :
                       errType === '#VALUE!' ? 'a text value is being used where a number is expected' :
                       errType === '#N/A' ? 'a lookup returned no match — check the lookup key exists in the source table' :
                       'a formula error is propagating through dependent cells';

    findings.push({
      title: `${errType} in "${header}" — ${count} Cell${count > 1 ? 's' : ''} Broken`,
      desc: `${count} cell${count > 1 ? 's' : ''} in the "${header}" column contain ${errType} (e.g. ${example}). This means ${downstream}. ${count > 1 ? `All ${count} instances must be resolved before this column can be used in aggregations or reports.` : ''}`,
      loc: count === 1 ? example : `${colLetter(col)}2:${colLetter(col)}${data.rowCount}`,
      type: 'error',
      category: 'formula-error',
      priority: 'critical',
      effort: 'easy',
      affectedCount: count,
    });
  });

  return findings;
}

// ─── Check 2: Type Mismatch ─────────────────────────────────────

function checkTypeMismatch(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    let nums = 0, texts = 0, textExamples = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) continue;
      if (typeof val === 'number') nums++;
      else { texts++; if (textExamples.length < 3) textExamples.push(`"${val}" at ${cellAddr(r, c)}`); }
    }
    const total = nums + texts;
    if (total < 10) continue;
    if (nums > 0 && texts > 0 && Math.min(nums, texts) / total > 0.05) {
      const header = data.headers[c] || colLetter(c);
      const minorType = nums >= texts ? 'text' : 'numeric';
      const minorCount = nums >= texts ? texts : nums;
      const majorCount = nums >= texts ? nums : texts;
      findings.push({
        title: `Mixed Types in "${header}" — ${minorCount} ${minorType === 'text' ? 'Text' : 'Numeric'} Values Among ${majorCount} ${minorType === 'text' ? 'Numbers' : 'Strings'}`,
        desc: `The "${header}" column is predominantly ${minorType === 'text' ? 'numeric' : 'text'} (${majorCount} cells) but contains ${minorCount} ${minorType} values — for example: ${textExamples.slice(0, 2).join(', ')}. This silently breaks =SUM(), AVERAGE(), and any conditional aggregation on this column, which will under-count by up to ${pct(minorCount, total)} of the total.`,
        loc: `${colLetter(c)}2:${colLetter(c)}${data.rowCount}`,
        type: 'warning',
        category: 'type-mismatch',
        priority: 'high',
        effort: 'medium',
        affectedCount: Math.min(nums, texts),
      });
    }
  }
  return findings;
}

// ─── Check 3: Inconsistent Formulas (hardcoded in formula column) ─

function checkInconsistentFormulas(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    let formulaCells = 0, hardcodedCells = 0, hardcodedExamples = [];
    for (let r = 1; r < data.rowCount; r++) {
      const formula = data.formulas[r][c];
      const val = data.values[r][c];
      if (val === '' || val === null) continue;
      if (formula?.toString().startsWith('=')) formulaCells++;
      else {
        hardcodedCells++;
        if (hardcodedExamples.length < 2) hardcodedExamples.push(`${cellAddr(r, c)}=${fmt(val)}`);
      }
    }
    const total = formulaCells + hardcodedCells;
    if (total < 5) continue;
    if (formulaCells / total > 0.7 && hardcodedCells > 0) {
      const header = data.headers[c] || colLetter(c);
      findings.push({
        title: `${hardcodedCells} Hardcoded Value${hardcodedCells > 1 ? 's' : ''} Breaking Formula Consistency in "${header}"`,
        desc: `"${header}" is a formula column (${formulaCells} calculated cells) but ${hardcodedCells} cell${hardcodedCells > 1 ? 's' : ''} contain hardcoded values instead (e.g. ${hardcodedExamples.join(', ')}). These are likely stale — manually entered during data cleanup and never updated since. When source data changes, these cells will not recalculate, silently corrupting aggregations that rely on this column.`,
        loc: `${colLetter(c)}2:${colLetter(c)}${data.rowCount}`,
        type: 'warning',
        category: 'formula-inconsistency',
        priority: 'high',
        effort: 'medium',
        affectedCount: hardcodedCells,
      });
    }
  }
  return findings;
}

// ─── Check 4: Volatile Functions ────────────────────────────────

function checkVolatileFunctions(data) {
  const findings = [];
  const seen = new Set();
  const colCounts = {};

  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (!formula.startsWith('=') || !VOLATILE_FUNCS.test(formula)) continue;
      const match = formula.match(VOLATILE_FUNCS);
      const funcName = match[1];
      const colKey = `${c}-${funcName}`;
      if (!colCounts[colKey]) colCounts[colKey] = { col: c, func: funcName, count: 0, example: cellAddr(r, c) };
      colCounts[colKey].count++;
    }
  }

  Object.values(colCounts).forEach(({ col, func, count, example }) => {
    const header = data.headers[col] || colLetter(col);
    const risk = func === 'RAND' || func === 'RANDBETWEEN'
      ? 'values change every time ANY cell in the workbook is edited, making audit trails impossible'
      : func === 'NOW' || func === 'TODAY'
      ? 'timestamps recalculate on every open, destroying historical accuracy'
      : 'the range it references can shift unpredictably as rows are inserted or deleted';
    findings.push({
      title: `${func}() in "${header}" Recalculates on Every Workbook Change — ${count} Instance${count > 1 ? 's' : ''}`,
      desc: `${count} cell${count > 1 ? 's' : ''} in the "${header}" column (e.g. ${example}) use ${func}(), a volatile function that ${risk}. In a workbook with many such functions, this also causes significant performance degradation during recalculation.`,
      loc: example,
      type: 'warning',
      category: 'volatile-function',
      priority: 'medium',
      effort: 'medium',
      affectedCount: count,
    });
  });

  return findings;
}

// ─── Check 5: Statistical Outliers ──────────────────────────────

function checkOutliers(data) {
  const findings = [];
  if (!data.stats) return findings;

  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c];
    const colStats = data.stats[header];
    if (!colStats || colStats.stdDev === 0 || colStats.count < 20) continue;
    const { mean, stdDev, max, min } = colStats;

    const flagged = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val !== 'number') continue;
      const z = Math.abs((val - mean) / stdDev);
      if (z > Z_THRESHOLD) flagged.push({ addr: cellAddr(r, c), val, z: z.toFixed(1) });
    }

    if (flagged.length > 0 && flagged.length <= 15) {
      const topOutlier = flagged.sort((a, b) => b.z - a.z)[0];
      const sumOutliers = flagged.reduce((acc, curr) => acc + curr.val, 0);
      const direction = topOutlier.val > mean ? 'above' : 'below';
      const deviationPct = Math.round(Math.abs(topOutlier.val - mean) / mean * 100);
      findings.push({
        title: `${flagged.length} Statistical Outlier${flagged.length > 1 ? 's' : ''} in "${header}" — Values Up to ${topOutlier.z}σ from Mean`,
        desc: `"${header}" has a mean of ${fmt(mean)} (stdDev ${fmt(stdDev)}). ${flagged.length} value${flagged.length > 1 ? 's' : ''} exceed the ${Z_THRESHOLD}σ threshold (the sum of these specific outliers is ${fmt(sumOutliers)}) — the most extreme is ${fmt(topOutlier.val)} at ${topOutlier.addr}, which is ${deviationPct}% ${direction} average and ${topOutlier.z} standard deviations out. This is either a genuine exceptional event or a data entry error (e.g., an extra zero).`,
        loc: topOutlier.addr,
        type: 'warning',
        category: 'outlier',
        priority: 'high',
        effort: 'easy',
        affectedCount: flagged.length,
      });
    }
  }
  return findings;
}

// ─── Check 6: Negative Values in Non-Negative Columns ───────────

function checkNegativeValues(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = (data.headers[c] || '').toString();
    if (!NON_NEGATIVE_PATTERNS.test(header)) continue;

    const negatives = [];
    let mostNegative = null;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val === 'number' && val < 0) {
        negatives.push({ addr: cellAddr(r, c), val });
        if (!mostNegative || val < mostNegative.val) mostNegative = { addr: cellAddr(r, c), val };
      }
    }
    if (negatives.length === 0) continue;

    const totalNeg = negatives.reduce((sum, n) => sum + n.val, 0);
    findings.push({
      title: `${negatives.length} Negative Value${negatives.length > 1 ? 's' : ''} in "${header}" — Logically Should Be Non-Negative`,
      desc: `"${header}" contains ${negatives.length} negative value${negatives.length > 1 ? 's' : ''} totalling ${fmt(totalNeg)} (most extreme: ${fmt(mostNegative.val)} at ${mostNegative.addr}). Columns representing ${header.toLowerCase()} are expected to be ≥ 0. Negative values will distort SUM totals and any MIN/MAX-based KPIs unless these are intentional adjustments or returns.`,
      loc: negatives.length === 1 ? negatives[0].addr : `${colLetter(c)}:${colLetter(c)}`,
      type: 'error',
      category: 'negative-illogical',
      priority: 'high',
      effort: 'easy',
      affectedCount: negatives.length,
    });
  }
  return findings;
}

// ─── Check 7: Missing Values in Key/ID Columns ──────────────────

function checkMissingValues(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = (data.headers[c] || '').toString();
    if (!KEY_COL_PATTERNS.test(header)) continue;

    const blanks = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) blanks.push(cellAddr(r, c));
    }
    if (blanks.length === 0) continue;

    const blankPct = Math.round((blanks.length / (data.rowCount - 1)) * 100);
    findings.push({
      title: `${blanks.length} Missing ${header} Values — ${blankPct}% of Rows Have No Primary Key`,
      desc: `${blanks.length} row${blanks.length > 1 ? 's' : ''} (${blankPct}% of the dataset) have a blank "${header}" field (e.g. ${blanks.slice(0, 3).join(', ')}). Primary key columns must never be empty — these rows cannot be joined to other tables, will be excluded from XLOOKUP-based reports, and may cause duplicate-detection logic to incorrectly merge unrelated records.`,
      loc: blanks.length === 1 ? blanks[0] : `${colLetter(c)}:${colLetter(c)}`,
      type: 'error',
      category: 'missing-value',
      priority: 'critical',
      effort: 'medium',
      affectedCount: blanks.length,
    });
  }
  return findings;
}

// ─── Check 8: Duplicate Rows ─────────────────────────────────────

function checkDuplicates(data) {
  const findings = [];
  const seen = new Map();
  for (let r = 1; r < data.rowCount; r++) {
    const key = JSON.stringify(data.values[r]);
    if (seen.has(key)) seen.get(key).push(r + 1);
    else seen.set(key, [r + 1]);
  }
  const groups = [...seen.values()].filter(rows => rows.length > 1);
  if (groups.length === 0) return findings;

  const totalDupeRows = groups.reduce((acc, rows) => acc + rows.length - 1, 0);
  const exampleRows = groups[0].join(', ');

  findings.push({
    title: `${totalDupeRows} Duplicate Row${totalDupeRows > 1 ? 's' : ''} Detected Across ${groups.length} Group${groups.length > 1 ? 's' : ''}`,
    desc: `${groups.length} group${groups.length > 1 ? 's' : ''} of completely identical rows found. The first duplicate group appears at rows ${exampleRows}. Duplicate rows inflate every SUM, COUNT, and SUMIF in this sheet. If this is a transactions or orders dataset, they likely represent double-processing or a failed data import that ran twice.`,
    loc: `A${groups[0][1]}`,
    type: 'error',
    category: 'duplicate',
    priority: 'critical',
    effort: 'easy',
    affectedCount: totalDupeRows,
  });
  return findings;
}

// ─── Check 9: Trailing/Leading Spaces ───────────────────────────

function checkTrailingSpaces(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    if (data.columnTypes && data.columnTypes[data.headers[c]] === 'numeric') continue;
    const spacyCells = [];
    const examples = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val === 'string' && val !== val.trim()) {
        spacyCells.push(cellAddr(r, c));
        if (examples.length < 2) examples.push(`${cellAddr(r, c)}="${val.trim()}"`);
      }
    }
    if (spacyCells.length === 0) continue;

    const header = data.headers[c] || colLetter(c);
    findings.push({
      title: `${spacyCells.length} Cells in "${header}" Have Invisible Whitespace — XLOOKUP Will Miss These Matches`,
      desc: `${spacyCells.length} cell${spacyCells.length > 1 ? 's' : ''} in "${header}" contain leading or trailing spaces (e.g. ${examples.join(', ')}). Whitespace is invisible in the cell but makes exact-match functions like XLOOKUP, COUNTIF, and MATCH return no results — causing lookup failures that look like missing data rather than a data quality issue.`,
      loc: spacyCells.length === 1 ? spacyCells[0] : `${colLetter(c)}:${colLetter(c)}`,
      type: 'warning',
      category: 'trailing-space',
      priority: 'medium',
      effort: 'easy',
      affectedCount: spacyCells.length,
    });
  }
  return findings;
}

// ─── Check 10: Magic Numbers in Formulas ────────────────────────

function checkMagicNumbers(data) {
  const findings = [];
  const seen = new Set();
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (!formula.startsWith('=') || !MAGIC_NUMBER_PATTERN.test(formula)) continue;
      const key = colLetter(c);
      if (seen.has(key)) continue;
      seen.add(key);
      const match = formula.match(MAGIC_NUMBER_PATTERN);
      const constant = match ? match[0].replace(/[+\-*/,\s()]/g, '') : '?';
      const header = data.headers[c] || colLetter(c);
      findings.push({
        title: `Hardcoded Constant (${constant}) in "${header}" Formula — Change-Resistant Technical Debt`,
        desc: `Formula at ${cellAddr(r, c)} in "${header}" contains a hardcoded constant "${constant}": "${formula.slice(0, 80)}...". If this is a rate (GST, tax, commission, discount), it is invisible to reviewers and cannot be updated centrally. When the rate changes, every formula using it must be hunted down and updated manually — a high-risk, error-prone process.`,
        loc: cellAddr(r, c),
        type: 'warning',
        category: 'magic-number',
        priority: 'medium',
        effort: 'medium',
        affectedCount: 1,
      });
    }
  }
  return findings;
}

// ─── Check 11: Unprotected Lookups ──────────────────────────────

function checkUnprotectedLookups(data) {
  const findings = [];
  let count = 0;
  let firstExample = null;

  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (formula.startsWith('=') && NO_IFERROR.test(formula)) {
        count++;
        if (!firstExample) firstExample = { addr: cellAddr(r, c), formula: formula.slice(0, 80) };
      }
    }
  }

  if (count > 0) {
    findings.push({
      title: `${count} Lookup Formula${count > 1 ? 's' : ''} Without IFERROR — Will Show #N/A When Match Fails`,
      desc: `${count} lookup formula${count > 1 ? 's are' : ' is'} not wrapped in IFERROR (first found at ${firstExample.addr}: "${firstExample.formula}..."). When a lookup key is not found in the source table, these cells display #N/A, which then propagates to every SUM, AVERAGE, or report that references them — turning a single missing record into a cascading report failure.`,
      loc: firstExample.addr,
      type: 'warning',
      category: 'formula-inconsistency',
      priority: 'high',
      effort: 'easy',
      affectedCount: count,
    });
  }
  return findings;
}

// ─── Check 12: VLOOKUP → XLOOKUP Upgrade ────────────────────────

function checkOldFunctions(data) {
  const findings = [];
  let vlookupCount = 0, example = null;
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (formula.startsWith('=') && OLD_LOOKUP.test(formula)) {
        vlookupCount++;
        if (!example) example = cellAddr(r, c);
      }
    }
  }
  if (vlookupCount === 0) return findings;

  findings.push({
    title: `${vlookupCount} VLOOKUP Formula${vlookupCount > 1 ? 's' : ''} Should Be Upgraded to XLOOKUP`,
    desc: `Found ${vlookupCount} VLOOKUP formula${vlookupCount > 1 ? 's' : ''} (first at ${example}). VLOOKUP can only search left-to-right, breaks when columns are inserted, and requires a manual column index number that becomes wrong whenever the source table changes. XLOOKUP eliminates all of these risks and also handles missing matches natively without requiring IFERROR.`,
    loc: example,
    type: 'improvement',
    category: 'improvement',
    priority: 'low',
    effort: 'medium',
    affectedCount: vlookupCount,
  });
  return findings;
}

// ─── Check 13: Convert to Excel Table ────────────────────────────

function checkTableFormat(data) {
  const findings = [];
  if (!data.isTable && data.rowCount > 20 && data.columnCount > 2) {
    findings.push({
      title: `${data.rowCount}-Row Dataset Is Not an Excel Table — Missing Auto-Expand, Structured References`,
      desc: `The data at ${data.address} has ${data.rowCount} rows and ${data.columnCount} columns but is formatted as a plain range, not an Excel Table (Ctrl+T). Without a Table: formulas do not auto-expand when new rows are added, column references in SUMIF/XLOOKUP use fragile A:A notation instead of [@ColumnName], and PivotTables do not auto-refresh to include new data.`,
      loc: data.address,
      type: 'improvement',
      category: 'improvement',
      priority: 'low',
      effort: 'easy',
      affectedCount: data.rowCount,
    });
  }
  return findings;
}

// ─── Check 14: Missing Column Headers ───────────────────────────

function checkMissingHeaders(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c];
    if (header !== '' && header !== null && header !== undefined) continue;

    // Only flag if the column actually has significant data
    let nonEmpty = 0;
    let numericCount = 0;
    let sampleVals = [];
    for (let r = 1; r < data.rowCount; r++) {
      const v = data.values[r][c];
      if (v !== '' && v !== null && v !== undefined && String(v).trim() !== '') {
        nonEmpty++;
        if (typeof v === 'number') numericCount++;
        if (sampleVals.length < 3) sampleVals.push(fmt(v));
      }
    }
    // Skip if column is empty, contains only a few rogue entries, or has very low density
    const fillRate = nonEmpty / (data.rowCount > 1 ? data.rowCount - 1 : 1);
    if (nonEmpty < 5 || (fillRate < 0.3 && nonEmpty < 20)) continue;

    const typeHint = numericCount / nonEmpty > 0.8
      ? `numeric (e.g. ${sampleVals.join(', ')})`
      : `text/mixed (e.g. ${sampleVals.join(', ')})`;

    findings.push({
      title: `Column ${colLetter(c)} Has ${nonEmpty} Data Values But No Header — All References to This Column Are Broken`,
      desc: `Column ${colLetter(c)} contains ${nonEmpty} ${typeHint} values across rows 2–${data.rowCount} but has no label in row 1 (cell ${colLetter(c)}1 is blank). Every XLOOKUP, SUMIF, and PivotTable that references this sheet by header name will silently skip this column. Without a header, the data is effectively invisible to structured queries.`,
      loc: `${colLetter(c)}1`,
      type: 'error',
      category: 'structure',
      priority: 'high',
      effort: 'easy',
      affectedCount: nonEmpty,
    });
  }
  return findings;
}

// ─── Check 15: Data Completeness (non-key columns) ───────────────

function checkDataCompleteness(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c] || colLetter(c);
    if (KEY_COL_PATTERNS.test(header)) continue; // already caught by checkMissingValues

    let blanks = 0;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) blanks++;
    }
    const totalRows = data.rowCount - 1;
    const blankPct = blanks / totalRows;
    
    // Ignore highly sparse columns with no real header (likely trash at the sheet edge)
    const hasMeaningfulHeader = data.headers[c] && data.headers[c] !== '' && !/^[A-Z]{1,3}$/.test(data.headers[c]);
    if (!hasMeaningfulHeader && blankPct > 0.95) continue;

    if (blankPct <= 0.2 || blankPct >= 1.0 || blanks <= 3) continue;

    // Compute impact on column aggregation
    const colStats = data.stats?.[header];
    const aggImpact = colStats
      ? ` AVERAGE() for "${header}" is currently ${fmt(colStats.mean)}, computed over only ${colStats.count} of ${totalRows} rows — the ${blanks} blanks are excluded from all aggregations.`
      : '';

    findings.push({
      title: `${Math.round(blankPct * 100)}% of "${header}" Is Blank — ${blanks} Rows Missing Data`,
      desc: `${blanks} of ${totalRows} data rows (${Math.round(blankPct * 100)}%) have no value in the "${header}" column.${aggImpact} High incompleteness means any report based on this column represents only a partial picture — aggregations will be biased toward the rows that do have data.`,
      loc: `${colLetter(c)}2:${colLetter(c)}${data.rowCount}`,
      type: 'warning',
      category: 'missing-value',
      priority: blankPct > 0.5 ? 'high' : 'medium',
      effort: 'hard',
      affectedCount: blanks,
    });
  }
  return findings;
}

// ─── Health Score ────────────────────────────────────────────────

export const computeHealthScore = (findings) => {
  const penalties = { critical: 15, high: 8, medium: 3, low: 1 };
  const deduction = findings.reduce((acc, f) => acc + (penalties[f.priority] || 0), 0);
  return Math.max(0, 100 - deduction);
};

/**
 * Generates a one-paragraph narrative summary of audit results
 * for the AISummary component — replacing the generic boilerplate.
 */
export const generateAuditNarrative = (findings, domain, score) => {
  if (findings.length === 0) return 'No issues detected. This dataset is clean and ready for analysis.';

  const criticals = findings.filter(f => f.priority === 'critical');
  const errors = findings.filter(f => f.type === 'error');
  const warnings = findings.filter(f => f.type === 'warning');
  const improvements = findings.filter(f => f.type === 'improvement');

  const parts = [];

  if (criticals.length > 0) {
    parts.push(`${criticals.length} critical issue${criticals.length > 1 ? 's' : ''} require immediate attention before this data can be trusted in reports`);
  }
  if (errors.length > criticals.length) {
    parts.push(`${errors.length - criticals.length} additional error${errors.length - criticals.length > 1 ? 's' : ''} affecting data integrity`);
  }
  if (warnings.length > 0) {
    parts.push(`${warnings.length} warning${warnings.length > 1 ? 's' : ''} that may cause silent calculation errors`);
  }
  if (improvements.length > 0) {
    parts.push(`${improvements.length} structural improvement${improvements.length > 1 ? 's' : ''} that would increase reliability`);
  }

  const topIssue = criticals[0] || errors[0] || warnings[0];
  const topIssueStr = topIssue ? ` Top priority: ${topIssue.title}.` : '';

  return `Found ${parts.join(', ')}.${topIssueStr}${domain !== 'general' ? ` Dataset classified as ${domain.toUpperCase()}.` : ''}`;
};

// ─── Main Export ─────────────────────────────────────────────────

/**
 * runLocalAudit — runs all checks and returns a deduplicated, ID-stamped
 * findings array. No API call — instant, runs entirely in the browser.
 *
 * @param {Object} data  — full sheet context from useOffice()
 * @returns {Array}      — finding objects
 */
export const runLocalAudit = (data) => {
  const prunedData = pruneNoisyColumns(data);

  const checks = [
    checkMissingHeaders,
    checkFormulaErrors,
    checkMissingValues,
    checkDuplicates,
    checkTypeMismatch,
    checkNegativeValues,
    checkInconsistentFormulas,
    checkUnprotectedLookups,
    checkVolatileFunctions,
    checkMagicNumbers,
    checkTrailingSpaces,
    checkOutliers,
    checkDataCompleteness,
    checkOldFunctions,
    checkTableFormat,
  ];

  let rawFindings = [];
  for (const check of checks) {
    try {
      rawFindings = [...rawFindings, ...check(prunedData)];
    } catch (e) {
      console.warn(`Audit check ${check.name} failed:`, e);
    }
  }

  // Deduplicate by loc + title
  const seen = new Set();
  const deduped = rawFindings.filter(f => {
    const key = `${f.loc}|${f.title}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  // Sort: critical first, then high, medium, low
  const priorityOrder = { critical: 0, high: 1, medium: 2, low: 3 };
  deduped.sort((a, b) => (priorityOrder[a.priority] ?? 4) - (priorityOrder[b.priority] ?? 4));

  return deduped.map((f, i) => ({ ...f, id: `finding-${i + 1}` }));
};