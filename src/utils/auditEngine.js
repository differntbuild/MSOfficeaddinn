/**
 * auditEngine.js — Local-first analysis engine (Phase 1)
 * Mimics Microsoft Copilot's Analyse Data approach:
 * runs all detection checks in pure JS (no API, instant results).
 * The AI (Phase 2) only enriches findings with recommendations.
 */

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

const colLetter = (index) => {
  let letter = '', i = index;
  while (i >= 0) {
    letter = String.fromCharCode((i % 26) + 65) + letter;
    i = Math.floor(i / 26) - 1;
  }
  return letter;
};

const cellAddr = (row, col) => `${colLetter(col)}${row + 1}`;

/** 
 * Identifies the logical boundary of the data.
 * Excel's usedRange often includes columns that are physically empty.
 * A column is considered "active" if it has a header OR any non-empty data row.
 */
function pruneNoisyColumns(data) {
  let lastActiveCol = -1;
  
  for (let c = 0; c < data.columnCount; c++) {
    const hasHeader = !!data.headers[c];
    let hasData = false;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val !== '' && val !== null && val !== undefined) {
        hasData = true;
        break;
      }
    }
    if (hasHeader || hasData) {
      lastActiveCol = c;
    }
  }

  if (lastActiveCol === -1) return data; // Keep as is if completely empty

  const prunedColumnCount = lastActiveCol + 1;
  if (prunedColumnCount === data.columnCount) return data;

  return {
    ...data,
    columnCount: prunedColumnCount,
    headers: data.headers.slice(0, prunedColumnCount),
    // Note: values, formulas, etc. can stay as is, we'll just loop till prunedColumnCount
  };
}

// Z-score outlier threshold
const Z_THRESHOLD = 2.8;

// Column-name patterns that imply non-negative values
const NON_NEGATIVE_PATTERNS = /\b(qty|quantity|price|amount|revenue|cost|cogs|age|count|units|sales|profit|stock|inventory|rate|score|hours|days|months|years)\b/i;

// Column-name patterns that suggest an ID or key field (should have no blanks)
const KEY_COL_PATTERNS = /\b(id|code|key|no|num|number|ref|reference|sku|order|invoice|employee|emp|customer|cust|product|serial)\b/i;

// Volatile Excel functions
const VOLATILE_FUNCS = /\b(NOW|TODAY|RAND|RANDBETWEEN|OFFSET|INDIRECT)\s*\(/i;

// Magic number pattern: formula contains a literal decimal or integer used as a rate/multiplier
const MAGIC_NUMBER_PATTERN = /[+\-*/,]\s*(0\.\d+|\d{1,3}(?!\d))\s*(?:[+\-*/,)]|$)/;

// Old-style lookup functions
const OLD_LOOKUP = /\bVLOOKUP\s*\(/i;
const NO_IFERROR = /^(?!.*IFERROR).*\b(VLOOKUP|XLOOKUP|INDEX|MATCH)\s*\(/i;

// ---------------------------------------------------------------------------
// Individual checks (each returns an array of finding objects)
// ---------------------------------------------------------------------------

/** 1. Formula errors: cells with Excel error strings */
function checkFormulaErrors(data) {
  const findings = [];
  const errors = ['#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#NULL!', '#NUM!', '#N/A'];
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const val = data.values[r][c];
      if (typeof val === 'string' && errors.includes(val.trim())) {
        findings.push({
          title: `${val.trim()} Error in ${data.headers[c] || colLetter(c)} column`,
          desc: `Cell ${cellAddr(r, c)} contains a ${val.trim()} error.`,
          loc: cellAddr(r, c),
          type: 'error',
          category: 'formula-error',
          priority: 'critical',
          effort: 'easy',
          affectedCount: 1,
        });
      }
    }
  }
  return findings;
}

/** 2. Type mismatch per column (skip header row r=0) */
function checkTypeMismatch(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    let nums = 0, texts = 0, blanks = 0;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) blanks++;
      else if (typeof val === 'number') nums++;
      else texts++;
    }
    const total = nums + texts;
    if (total === 0) continue;
    // Flag only if mixed AND neither type is trivially rare (>5% of the other)
    if (nums > 0 && texts > 0 && Math.min(nums, texts) / total > 0.05) {
      const header = data.headers[c] || colLetter(c);
      findings.push({
        title: `Mixed Data Types in "${header}" column`,
        desc: `Column ${colLetter(c)} has ${nums} numeric and ${texts} text values. This suggests data entry errors or numbers stored as text.`,
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

/** 3. Hardcoded values mixed into formula columns */
function checkInconsistentFormulas(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    let formulaCells = 0, hardcodedCells = 0, hardcodedExample = null;
    for (let r = 1; r < data.rowCount; r++) {
      const formula = data.formulas[r][c];
      const val = data.values[r][c];
      if (val === '' || val === null) continue;
      if (formula?.toString().startsWith('=')) formulaCells++;
      else { hardcodedCells++; if (!hardcodedExample) hardcodedExample = cellAddr(r, c); }
    }
    const total = formulaCells + hardcodedCells;
    if (total < 5) continue;
    // Flag if majority is formulas but some cells are hardcoded
    if (formulaCells / total > 0.7 && hardcodedCells > 0) {
      const header = data.headers[c] || colLetter(c);
      findings.push({
        title: `Hardcoded Values in "${header}" Formula Column`,
        desc: `${formulaCells} cells use formulas but ${hardcodedCells} are hardcoded (e.g. ${hardcodedExample}). These may be stale or accidentally overwritten formulas.`,
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

/** 4. Volatile functions (NOW, TODAY, RAND, OFFSET, INDIRECT) */
function checkVolatileFunctions(data) {
  const findings = [];
  const seen = new Set();
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (formula.startsWith('=') && VOLATILE_FUNCS.test(formula)) {
        const match = formula.match(VOLATILE_FUNCS);
        const key = `${colLetter(c)}-${match[1]}`;
        if (!seen.has(key)) {
          seen.add(key);
          findings.push({
            title: `Volatile Function ${match[1]}() in Column ${colLetter(c)}`,
            desc: `${match[1]}() recalculates on every workbook change in column ${colLetter(c)} (e.g. ${cellAddr(r, c)}). This can slow recalculation and cause audit trail issues.`,
            loc: cellAddr(r, c),
            type: 'warning',
            category: 'volatile-function',
            priority: 'medium',
            effort: 'medium',
            affectedCount: 1,
          });
        }
      }
    }
  }
  return findings;
}

/** 5. Statistical outliers using Z-score (using pre-computed stats from useOffice) */
function checkOutliers(data) {
  const findings = [];
  const headers = data.headers;
  if (!data.stats) return findings;

  for (let c = 0; c < data.columnCount; c++) {
    const header = headers[c];
    const colStats = data.stats[header];
    if (!colStats || colStats.stdDev === 0 || colStats.count < 10) continue;
    const { mean, stdDev } = colStats;
    const flagged = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val !== 'number') continue;
      const z = Math.abs((val - mean) / stdDev);
      if (z > Z_THRESHOLD) {
        flagged.push({ addr: cellAddr(r, c), val, z: z.toFixed(1) });
      }
    }
    if (flagged.length > 0 && flagged.length <= 10) {
      flagged.forEach(({ addr, val, z }) => {
        findings.push({
          title: `Statistical Outlier in "${header}"`,
          desc: `${addr} = ${val.toLocaleString('en-IN')} is ${z}σ from the column mean (${mean.toLocaleString('en-IN')}). Likely a typo or exceptional event.`,
          loc: addr,
          type: 'warning',
          category: 'outlier',
          priority: 'medium',
          effort: 'easy',
          affectedCount: 1,
        });
      });
    }
  }
  return findings;
}

/** 6. Negative values in logically non-negative columns */
function checkNegativeValues(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = (data.headers[c] || '').toString();
    if (!NON_NEGATIVE_PATTERNS.test(header)) continue;
    const negatives = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val === 'number' && val < 0) negatives.push(cellAddr(r, c));
    }
    if (negatives.length > 0) {
      findings.push({
        title: `Negative Values in "${header}"`,
        desc: `${negatives.length} cell(s) in "${header}" are negative (e.g. ${negatives[0]}). This column should logically be non-negative.`,
        loc: negatives.length === 1 ? negatives[0] : `${negatives[0]} (+${negatives.length - 1} more)`,
        type: 'error',
        category: 'negative-illogical',
        priority: 'high',
        effort: 'easy',
        affectedCount: negatives.length,
      });
    }
  }
  return findings;
}

/** 7. Missing values in key/ID columns */
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
    if (blanks.length > 0) {
      findings.push({
        title: `Missing Values in Key Column "${header}"`,
        desc: `${blanks.length} blank cell(s) in "${header}" (e.g. ${blanks[0]}). ID and reference columns should never be empty.`,
        loc: blanks.length === 1 ? blanks[0] : `${colLetter(c)}:${colLetter(c)}`,
        type: 'error',
        category: 'missing-value',
        priority: 'critical',
        effort: 'medium',
        affectedCount: blanks.length,
      });
    }
  }
  return findings;
}

/** 8. Duplicate rows */
function checkDuplicates(data) {
  const findings = [];
  const seen = new Map();
  for (let r = 1; r < data.rowCount; r++) {
    const key = JSON.stringify(data.values[r]);
    if (seen.has(key)) {
      seen.get(key).push(r + 1);
    } else {
      seen.set(key, [r + 1]);
    }
  }
  const dupes = [...seen.values()].filter(rows => rows.length > 1);
  if (dupes.length > 0) {
    const totalDupeRows = dupes.reduce((acc, rows) => acc + rows.length - 1, 0);
    findings.push({
      title: `${totalDupeRows} Duplicate Row${totalDupeRows > 1 ? 's' : ''} Detected`,
      desc: `${dupes.length} group(s) of identical rows found (e.g. rows ${dupes[0].join(', ')}). Duplicates inflate totals and distort analysis.`,
      loc: `A${dupes[0][1]}`,
      type: 'error',
      category: 'duplicate',
      priority: 'critical',
      effort: 'easy',
      affectedCount: totalDupeRows,
    });
  }
  return findings;
}

/** 9. Trailing/leading spaces in text columns */
function checkTrailingSpaces(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    if (data.columnTypes && data.columnTypes[data.headers[c]] === 'numeric') continue;
    const spacyCells = [];
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (typeof val === 'string' && val !== val.trim()) {
        spacyCells.push(cellAddr(r, c));
      }
    }
    if (spacyCells.length > 0) {
      const header = data.headers[c] || colLetter(c);
      findings.push({
        title: `Trailing/Leading Spaces in "${header}"`,
        desc: `${spacyCells.length} cell(s) have invisible whitespace (e.g. ${spacyCells[0]}). This breaks VLOOKUP/XLOOKUP matches silently.`,
        loc: spacyCells.length === 1 ? spacyCells[0] : `${colLetter(c)}:${colLetter(c)}`,
        type: 'warning',
        category: 'trailing-space',
        priority: 'medium',
        effort: 'easy',
        affectedCount: spacyCells.length,
      });
    }
  }
  return findings;
}

/** 10. Magic numbers in formulas (hardcoded tax/rate constants) */
function checkMagicNumbers(data) {
  const findings = [];
  const seen = new Set();
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (!formula.startsWith('=')) continue;
      if (MAGIC_NUMBER_PATTERN.test(formula)) {
        const key = colLetter(c);
        if (!seen.has(key)) {
          seen.add(key);
          findings.push({
            title: `Magic Number in Formula at ${cellAddr(r, c)}`,
            desc: `Formula "${formula.slice(0, 60)}..." contains a hardcoded constant. If this is a rate (tax, GST, discount), it should be a named range or cell reference.`,
            loc: cellAddr(r, c),
            type: 'warning',
            category: 'magic-number',
            priority: 'medium',
            effort: 'medium',
            affectedCount: 1,
          });
        }
      }
    }
  }
  return findings;
}

/** 11. VLOOKUP without IFERROR protection */
function checkUnprotectedLookups(data) {
  const findings = [];
  for (let r = 1; r < data.rowCount; r++) {
    for (let c = 0; c < data.columnCount; c++) {
      const formula = data.formulas[r][c]?.toString() || '';
      if (formula.startsWith('=') && NO_IFERROR.test(formula)) {
        findings.push({
          title: `Unprotected Lookup at ${cellAddr(r, c)}`,
          desc: `Formula "${formula.slice(0, 60)}..." uses a lookup function without IFERROR. A missing match will display #N/A and break dependent formulas.`,
          loc: cellAddr(r, c),
          type: 'warning',
          category: 'formula-inconsistency',
          priority: 'high',
          effort: 'easy',
          affectedCount: 1,
        });
        break; // one per sheet is enough to surface the pattern
      }
    }
  }
  return findings;
}

/** 12. Improvement: VLOOKUP → XLOOKUP upgrade suggestion */
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
  if (vlookupCount > 0) {
    findings.push({
      title: `Upgrade ${vlookupCount} VLOOKUP${vlookupCount > 1 ? 's' : ''} to XLOOKUP`,
      desc: `Found ${vlookupCount} VLOOKUP formula(s) (e.g. ${example}). XLOOKUP is faster, handles missing values natively, and works in both directions.`,
      loc: example,
      type: 'improvement',
      category: 'improvement',
      priority: 'low',
      effort: 'medium',
      affectedCount: vlookupCount,
    });
  }
  return findings;
}

/** 13. Improvement: Suggest converting range to Excel Table */
function checkTableFormat(data) {
  const findings = [];
  // Use the real isTable flag from useOffice instead of a regex heuristic
  if (!data.isTable && data.rowCount > 20 && data.columnCount > 2) {
    findings.push({
      title: 'Convert Range to Excel Table',
      desc: `Data has ${data.rowCount} rows but is not an Excel Table. Tables auto-expand formulas, support structured references, and make XLOOKUP/SUMIF more readable.`,
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

/** 14. Missing column headers */
function checkMissingHeaders(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c];
    if (header === '' || header === null || header === undefined) {
      // Check if column actually has data. If it's literally empty, don't flag "Missing Header"
      let hasAnyData = false;
      for (let r = 1; r < data.rowCount; r++) {
        if (data.values[r][c] !== '' && data.values[r][c] !== null) {
          hasAnyData = true;
          break;
        }
      }
      
      if (hasAnyData) {
        findings.push({
          title: `Column ${colLetter(c)} Has No Header`,
          desc: `Column ${colLetter(c)} has no header label in Row 1. Missing headers break XLOOKUP, pivot tables, and structured references.`,
          loc: `${colLetter(c)}1`,
          type: 'error',
          category: 'structure',
          priority: 'high',
          effort: 'easy',
          affectedCount: 1,
        });
      }
    }
  }
  return findings;
}

/** 15. Improvement: Columns with >20% blanks (data completeness) */
function checkDataCompleteness(data) {
  const findings = [];
  for (let c = 0; c < data.columnCount; c++) {
    const header = data.headers[c] || colLetter(c);
    // Skip key columns (already covered by checkMissingValues)
    if (KEY_COL_PATTERNS.test(header)) continue;
    let blanks = 0;
    for (let r = 1; r < data.rowCount; r++) {
      const val = data.values[r][c];
      if (val === '' || val === null || val === undefined) blanks++;
    }
    const pct = blanks / (data.rowCount - 1);
    // Only flag if it's missing data but NOT 100% empty (which suggests an intentionally unused column)
    if (pct > 0.2 && pct < 1.0 && blanks > 3) {
      findings.push({
        title: `${Math.round(pct * 100)}% Missing in "${header}"`,
        desc: `${blanks} of ${data.rowCount - 1} rows are blank in "${header}". High incompleteness affects aggregation accuracy.`,
        loc: `${colLetter(c)}2:${colLetter(c)}${data.rowCount}`,
        type: 'warning',
        category: 'missing-value',
        priority: 'medium',
        effort: 'hard',
        affectedCount: blanks,
      });
    }
  }
  return findings;
}

// ---------------------------------------------------------------------------
// Health Score
// ---------------------------------------------------------------------------

export const computeHealthScore = (findings) => {
  const penalties = { critical: 15, high: 8, medium: 3, low: 1 };
  const deduction = findings.reduce((acc, f) => acc + (penalties[f.priority] || 0), 0);
  return Math.max(0, 100 - deduction);
};

// ---------------------------------------------------------------------------
// Main export
// ---------------------------------------------------------------------------

/**
 * Runs all local checks on the full sheet context from useOffice.
 * Returns findings immediately — no API call.
 */
export const runLocalAudit = (data) => {
  // 1. Prune unused columns at the end of the range
  const prunedData = pruneNoisyColumns(data);

  // 2. Map through checks with the pruned data
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

  // 3. Deduplicate by loc+title
  const seen = new Set();
  const deduped = rawFindings.filter(f => {
    const key = `${f.loc}|${f.title}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });

  // 4. Assign unique IDs for Phase 2 tracking
  return deduped.map((f, i) => ({
    ...f,
    id: `finding-${i + 1}`
  }));
};
