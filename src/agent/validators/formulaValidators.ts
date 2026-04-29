export const sanitizeFormulaText = (formula: string) => {
  if (!formula) return formula;
  return formula
    .replace(/↗/g, '')
    .replace(/\[\[[^\]]+\]\]/g, '')
    .replace(/\s+/g, '')
    .trim();
};

export const validateFormulaSafety = (formula: string) => {
  if (!formula) return { ok: false, reason: 'No formula generated.' };
  if (!formula.startsWith('=')) return { ok: false, reason: 'Formula must start with "=".' };
  if (/\[\[|\]\]|↗/.test(formula)) return { ok: false, reason: 'Formula contains citation markers.' };
  if (/[^A-Za-z0-9_:\+\-\*\/\^\(\),."'%!<>=& ]/.test(formula)) {
    return { ok: false, reason: 'Formula contains invalid characters.' };
  }
  return { ok: true };
};

const getRangeHeight = (rangeText: string) => {
  const parts = rangeText.split(':');
  const rowOf = (token: string) => {
    const m = token.match(/\d+/);
    return m ? parseInt(m[0], 10) : null;
  };
  const r1 = rowOf(parts[0]);
  const r2 = rowOf(parts[1] || parts[0]);
  if (!r1 || !r2) return null;
  return Math.abs(r2 - r1) + 1;
};

export const validateXlookupShape = (formula: string) => {
  const xlookupMatch = formula.match(/XLOOKUP\s*\((.*)\)/i);
  if (!xlookupMatch) return { ok: true };

  const argsText = xlookupMatch[1];
  const args = argsText.split(',').map(s => s.trim());
  if (args.length < 3) return { ok: false, reason: 'XLOOKUP requires lookup value, lookup array, and return array.' };

  const lookupArray = args[1];
  const returnArray = args[2];
  const lookupHeight = getRangeHeight(lookupArray);
  const returnHeight = getRangeHeight(returnArray);

  if (lookupHeight && returnHeight && lookupHeight !== returnHeight) {
    return { ok: false, reason: 'XLOOKUP lookup and return ranges are not aligned.' };
  }
  return { ok: true };
};
