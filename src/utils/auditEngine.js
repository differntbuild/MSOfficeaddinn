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
      const row = data.values[r];
      if (!row) continue;
      const val = row[c];
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
  
  let newAddress = data.address;
  if (data.address && data.address.includes(':')) {
    const oldRangeParts = data.address.split(':');
    const startCell = oldRangeParts[0]; // e.g., 'Sheet1!A1'
    const newEndCell = `${colLetter(pruned - 1)}${data.rowCount}`; // e.g., 'T501'
    newAddress = `${startCell}:${newEndCell}`;
  }

  return { 
    ...data, 
    columnCount: pruned, 
    headers: data.headers.slice(0, pruned),
    address: newAddress
  };
}

function pruneNoisyRows(data) {
  let lastActiveRow = 0; // Header is row 0
  for (let r = data.rowCount - 1; r >= 1; r--) {
    const row = data.values[r];
    if (!row) continue;
    
    let hasData = false;
    for (let c = 0; c < data.columnCount; c++) {
      const val = row[c];
      if (val !== '' && val !== null && val !== undefined && String(val).trim() !== '') {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastActiveRow = r;
      break;
    }
  }

  if (lastActiveRow === data.rowCount - 1) return data;
  
  const pruned = lastActiveRow + 1;
  const prunedValues = data.values.slice(0, pruned);
  const prunedFormulas = data.formulas ? data.formulas.slice(0, pruned) : null;

  let newAddress = data.address;
  if (data.address && data.address.includes(':')) {
    const parts = data.address.split(':');
    const startCell = parts[0];
    const sheetPrefix = startCell.includes('!') ? startCell.split('!')[0] + '!' : '';
    const startCoords = startCell.split('!').pop();
    const startCol = startCoords.replace(/[0-9]/g, '');
    const newEndCell = `${colLetter(data.columnCount - 1)}${pruned}`;
    newAddress = `${startCell}:${sheetPrefix}${startCol}${startCoords.replace(/[A-Z]/g, '') === '1' ? '' : startCoords.replace(/[A-Z]/g, '')}${newEndCell}`; // This is getting complex, let's simplify
    // Simplest way: if it's A1:AE1048576, just make it A1:AE[pruned]
    const colPart = colLetter(data.columnCount - 1);
    newAddress = `${startCell}:${colPart}${pruned}`;
  }

  return {
    ...data,
    values: prunedValues,
    formulas: prunedFormulas,
    rowCount: pruned,
    address: newAddress
  };
}

function pruneGhostRows(data, findings) {
  const prunedValues = [data.values[0]]; // keep header
  const prunedFormulas = data.formulas ? [data.formulas[0]] : null;
  const ghostRows = [];
  
  // We only call it a "Ghost Row" if there is data AFTER it.
  // Otherwise, it was just trailing space (which pruneNoisyRows handles).
  for (let r = 1; r < data.rowCount; r++) {
    const row = data.values[r];
    if (!row) {
      ghostRows.push(r + 1);
      continue;
    }
    let emptyCount = 0;
    for (let c = 0; c < data.columnCount; c++) {
      const val = row[c];
      if (val === '' || val === null || val === undefined || String(val).trim() === '') {
        emptyCount++;
      }
    }
    if (emptyCount / data.columnCount > 0.95) {
      ghostRows.push(r + 1); // 1-indexed row number
    } else {
      prunedValues.push(data.values[r]);
      if (prunedFormulas) prunedFormulas.push(data.formulas[r]);
    }
  }

  if (ghostRows.length > 0) {
    findings.push({
      title: `Pruned ${ghostRows.length} Sparse Ghost Row${ghostRows.length > 1 ? 's' : ''}`,
      desc: `Detected ${ghostRows.length} nearly-empty row${ghostRows.length > 1 ? 's' : ''} (e.g. Row ${ghostRows[0]}) interspersed within your data range. These "ghost rows" are often invisible artifacts of deleted data that can corrupt pivot table counts and trigger false-positive missing-data warnings. They have been safely excluded from this analysis.`,
      loc: `Row ${ghostRows[0]}`,
      type: 'improvement',
      category: 'structure',
      priority: 'low',
      effort: 'easy',
      affectedCount: ghostRows.length,
    });
  }

  return { ...data, values: prunedValues, formulas: prunedFormulas || data.formulas, rowCount: prunedValues.length };
}

function coalesceColumns(data, findings) {
  const groups = {};
  
  const normalize = (h) => h.toLowerCase().replace(/[^a-z]/g, '');
  const getSemanticRoot = (h) => {
    const clean = normalize(h);
    
    // EXCLUSION: Prevent over-merging of hierarchical or meta data
    if (clean.includes('sub') || clean.includes('parent') || clean.includes('meta') || clean.includes('child')) {
      return null; 
    }

    if (clean.includes('email')) return 'email';
    if (clean.includes('phone') || clean === 'contact') return 'phone';
    if (clean === 'address' || clean === 'addr' || clean === 'location') return 'address';
    if (clean === 'id' || clean === 'id1' || clean === 'empid' || clean === 'employeeid') return 'empid';
    if (clean === 'product' || clean === 'productname') return 'product';
    if (clean === 'category' || clean === 'cat') return 'category';
    if (clean === 'rating') return 'rating';
    if (clean === 'qty' || clean === 'quantity') return 'quantity';
    if (clean.includes('revenue') || clean === 'rev') return 'revenue';
    if (clean === 'dob' || clean === 'birthdate') return 'dob';
    if (clean === 'name' || clean === 'fullname') return 'name';
    return null; // Don't guess for unknown roots
  };

  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c];
    if (!header) continue;
    const root = getSemanticRoot(header);
    if (root) {
      if (!groups[root]) groups[root] = [];
      groups[root].push(c);
    }
  }

  const newValues = data.values.map(row => row ? [...row] : []);
  const newHeaders = [...data.headers];
  const newFormulas = data.formulas ? data.formulas.map(row => row ? [...row] : []) : null;

  for (const [root, cols] of Object.entries(groups)) {
    if (cols.length > 1) {
      const primaryCol = cols[0];
      const mergedNames = cols.map(c => data.headers[c]);
      
      for (let r = 1; r < data.rowCount; r++) {
        let bestVal = '';
        let bestForm = '';
        for (const c of cols) {
          const rowVal = newValues[r];
          const rowForm = newFormulas ? newFormulas[r] : null;
          if (!rowVal) continue;
          
          const val = rowVal[c];
          if (val !== '' && val !== null && val !== undefined) {
            bestVal = val;
            if (rowForm) bestForm = rowForm[c];
            break;
          }
        }
        
        if (newValues[r]) newValues[r][primaryCol] = bestVal;
        if (newFormulas && newFormulas[r]) newFormulas[r][primaryCol] = bestForm;
        
        for (let i = 1; i < cols.length; i++) {
          if (newValues[r]) newValues[r][cols[i]] = '';
          if (newFormulas && newFormulas[r]) newFormulas[r][cols[i]] = '';
        }
      }
      
      for (let i = 1; i < cols.length; i++) {
        newHeaders[cols[i]] = '';
      }

      findings.push({
        title: `Semantically Merged ${cols.length} "${mergedNames[0]}" Columns`,
        desc: `Consolidated data from ${cols.length} fragmented columns (${mergedNames.join(', ')}) into a single column. This prevents the audit engine from falsely penalizing the dataset for missing values in redundant/phantom columns.`,
        loc: `${colLetter(cols[0])}1`,
        type: 'improvement',
        category: 'structure',
        priority: 'low',
        effort: 'easy',
        affectedCount: cols.length,
      });
    }
  }

  return { ...data, values: newValues, headers: newHeaders, formulas: newFormulas };
}

// ─── Constants ──────────────────────────────────────────────────



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
    const row = data.values[r];
    if (!row) continue;
    for (let c = 0; c < data.columnCount; c++) {
      const val = row[c];
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

    const cv = mean === 0 ? 0 : Math.abs(stdDev / mean);
    let dynamicThreshold = 2.8;
    if (cv > 1.5) dynamicThreshold = 3.5;
    else if (cv < 0.5) dynamicThreshold = 2.0;

    const flagged = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val !== 'number') continue;
      const z = Math.abs((val - mean) / stdDev);
      if (z > dynamicThreshold) flagged.push({ addr: cellAddr(r, c), val, z: z.toFixed(1) });
    }

    if (flagged.length > 0 && flagged.length <= 15) {
      const topOutlier = flagged.sort((a, b) => b.z - a.z)[0];
      const sumOutliers = flagged.reduce((acc, curr) => acc + curr.val, 0);
      const direction = topOutlier.val > mean ? 'above' : 'below';
      const deviationPct = Math.round(Math.abs(topOutlier.val - mean) / mean * 100);
      findings.push({
        title: `${flagged.length} Statistical Outlier${flagged.length > 1 ? 's' : ''} in "${header}" — Values Up to ${topOutlier.z}σ from Mean`,
        desc: `"${header}" has a mean of ${fmt(mean)} (stdDev ${fmt(stdDev)}). ${flagged.length} value${flagged.length > 1 ? 's' : ''} exceed the dynamic ${dynamicThreshold}σ threshold (the sum of these specific outliers is ${fmt(sumOutliers)}) — the most extreme is ${fmt(topOutlier.val)} at ${topOutlier.addr}, which is ${deviationPct}% ${direction} average and ${topOutlier.z} standard deviations out. This is either a genuine exceptional event or a data entry error (e.g., an extra zero).`,
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
  const keyCols = [];

  for (let c = 0; c < data.columnCount; c++) {
    const header = (data.headers[c] || '').toString();
    if (!KEY_COL_PATTERNS.test(header)) continue;

    const blanks = [];
    const uniqueVals = new Set();
    let nonBlanks = 0;

    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) {
        blanks.push(cellAddr(r, c));
      } else {
        nonBlanks++;
        uniqueVals.add(val);
      }
    }

    const fillRate = data.rowCount > 1 ? nonBlanks / (data.rowCount - 1) : 0;
    const uniqueness = nonBlanks > 0 ? uniqueVals.size / nonBlanks : 0;
    const score = fillRate * uniqueness;

    keyCols.push({ c, header, blanks, score });
  }

  if (keyCols.length === 0) return findings;

  keyCols.sort((a, b) => b.score - a.score);
  const primaryKeyCol = keyCols[0];

  for (const kc of keyCols) {
    if (kc.blanks.length === 0) continue;

    const blankPct = Math.round((kc.blanks.length / (data.rowCount - 1)) * 100);
    const isPrimary = kc.c === primaryKeyCol.c;

    if (isPrimary) {
      findings.push({
        title: `${kc.blanks.length} Missing ${kc.header} Values — ${blankPct}% of Rows Have No Primary Key`,
        desc: `${kc.blanks.length} row${kc.blanks.length > 1 ? 's' : ''} (${blankPct}% of the dataset) have a blank "${kc.header}" field (e.g. ${kc.blanks.slice(0, 3).join(', ')}). Primary key columns must never be empty — these rows cannot be joined to other tables, will be excluded from XLOOKUP-based reports, and may cause duplicate-detection logic to incorrectly merge unrelated records.`,
        loc: kc.blanks.length === 1 ? kc.blanks[0] : `${colLetter(kc.c)}:${colLetter(kc.c)}`,
        type: 'error',
        category: 'missing-value',
        priority: 'critical',
        effort: 'medium',
        affectedCount: kc.blanks.length,
      });
    } else {
      findings.push({
        title: `${kc.blanks.length} Missing ${kc.header} Values — ${blankPct}% of Rows Are Unassigned`,
        desc: `${kc.blanks.length} row${kc.blanks.length > 1 ? 's' : ''} (${blankPct}% of the dataset) have a blank "${kc.header}" field (e.g. ${kc.blanks.slice(0, 3).join(', ')}). Since this appears to be a secondary identifier or foreign key, missing values may be valid (e.g., an unassigned product), but they will drop out of any PivotTables or groupings based on this column.`,
        loc: kc.blanks.length === 1 ? kc.blanks[0] : `${colLetter(kc.c)}:${colLetter(kc.c)}`,
        type: 'warning',
        category: 'missing-value',
        priority: 'medium',
        effort: 'medium',
        affectedCount: kc.blanks.length,
      });
    }
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
      priority: 'medium',
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
function checkBusinessLogic(data) {
  const findings = [];
  
  let discountCol = -1, marginCol = -1, gstCol = -1, categoryCol = -1, returnCol = -1;
  for (let c = 0; c < data.columnCount; c++) {
    const h = (data.headers[c] || '').toLowerCase();
    if (h.includes('discount') && h.includes('%')) discountCol = c;
    if (h.includes('margin') && h.includes('%')) marginCol = c;
    if (h.includes('gst') || h.includes('tax')) gstCol = c;
    if (h.includes('category') || h.includes('segment')) categoryCol = c;
    if (h.includes('return') || h.includes('refund')) returnCol = c;
  }

  // Margin Violations
  if (discountCol !== -1 && marginCol !== -1) {
    let violationCount = 0, example = null;
    for (let r = 1; r < data.rowCount; r++) {
      const disc = parseFloat(data.values[r][discountCol]);
      const marg = parseFloat(data.values[r][marginCol]);
      if (!isNaN(disc) && !isNaN(marg) && disc > marg) {
        violationCount++;
        if (!example) example = cellAddr(r, discountCol);
      }
    }
    if (violationCount > 0) {
      findings.push({
        title: `Margin Violation: ${violationCount} Order${violationCount > 1 ? 's' : ''} Where Discount Exceeds Gross Margin`,
        desc: `Found ${violationCount} row${violationCount > 1 ? 's' : ''} (e.g. at ${example}) where the discount percentage is higher than the gross margin percentage, resulting in a net loss on the sale.`,
        loc: example,
        type: 'warning',
        category: 'business-logic',
        priority: 'high',
        effort: 'medium',
        affectedCount: violationCount,
      });
    }
  }

  // Tax/GST Anomalies
  if (gstCol !== -1) {
    let anomalyCount = 0, example = null;
    const validSlabs = [0, 5, 12, 18, 28];
    for (let r = 1; r < data.rowCount; r++) {
      let gst = parseFloat(data.values[r][gstCol]);
      if (isNaN(gst)) continue;
      if (gst > 0 && gst < 1) gst = gst * 100;
      gst = Math.round(gst);
      if (!validSlabs.includes(gst)) {
        anomalyCount++;
        if (!example) example = cellAddr(r, gstCol);
      }
    }
    if (anomalyCount > 0) {
      findings.push({
        title: `GST Compliance Risk: ${anomalyCount} Invalid Tax Rate${anomalyCount > 1 ? 's' : ''}`,
        desc: `Found ${anomalyCount} row${anomalyCount > 1 ? 's' : ''} (e.g. at ${example}) where the GST rate does not match standard Indian tax slabs (0%, 5%, 12%, 18%, 28%). This could lead to compliance issues during tax filing.`,
        loc: example,
        type: 'error',
        category: 'business-logic',
        priority: 'critical',
        effort: 'easy',
        affectedCount: anomalyCount,
      });
    }
  }

  return findings;
}

export const runLocalAudit = (data) => {
  let rawFindings = [];
  
  // Phase 0: Pruning & Normalization
  // 1. Truncate trailing empty columns
  let prunedData = pruneNoisyColumns(data);
  // 2. Truncate trailing empty rows (silently)
  prunedData = pruneNoisyRows(prunedData);
  // 3. Remove interspersed "ghost" rows (with finding)
  prunedData = pruneGhostRows(prunedData, rawFindings);
  // 4. Merge semantic duplicates
  prunedData = coalesceColumns(prunedData, rawFindings);

  const hasFormulas = prunedData.formulas && prunedData.formulas.some(row => row && row.some(f => typeof f === 'string' && f.startsWith('=')));
  
  const formulaChecks = hasFormulas ? [
    checkFormulaErrors,
    checkInconsistentFormulas,
    checkUnprotectedLookups,
    checkVolatileFunctions,
    checkMagicNumbers,
    checkOldFunctions,
  ] : [];

  const checks = [
    checkMissingHeaders,
    ...formulaChecks,
    checkMissingValues,
    checkDuplicates,
    checkTypeMismatch,
    checkNegativeValues,
    checkTrailingSpaces,
    checkOutliers,
    checkDataCompleteness,
    checkTableFormat,
    checkBusinessLogic,
  ];

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

/**
 * Generates a Data Profile for the LLM "Coder" pass.
 * Extracts schema, samples, and summary stats without sending full data.
 */
export const generateDataProfile = (data) => {
  const profile = {
    rowCount: data.rowCount,
    colCount: data.columnCount,
    headers: data.headers,
    types: {},
    stats: {},
    sample: data.values.slice(1, 6), // First 5 rows of data
    address: data.address,
    sheetName: data.sheetName
  };

  data.headers.forEach((h, i) => {
    let type = 'string';
    let nullCount = 0;
    let min = Infinity, max = -Infinity;
    let numericCount = 0;

    for (let r = 1; r < data.rowCount; r++) {
      const row = data.values[r];
      if (!row) { nullCount++; continue; }
      const val = row[i];
      if (val === '' || val === null || val === undefined) {
        nullCount++;
        continue;
      }
      
      if (typeof val === 'number') {
        type = 'number';
        numericCount++;
        if (val < min) min = val;
        if (val > max) max = val;
      }
    }

    profile.types[h] = type;
    profile.stats[h] = {
      nullCount,
      nullPct: Math.round((nullCount / Math.max(1, data.rowCount)) * 100) + '%'
    };

    if (numericCount > 0 && min !== Infinity) {
      profile.stats[h].min = min;
      profile.stats[h].max = max;
    }
  });

  // Sanitize sample to avoid [object Object] or null issues
  profile.sample = profile.sample.map(row => 
    row ? row.map(v => v === null || v === undefined ? "" : v) : []
  );

  return profile;
};