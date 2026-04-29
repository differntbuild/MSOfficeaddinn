const collectColumnLetters = (formula: string) => {
  const refs = formula.match(/[A-Z]{1,3}\d*(?::[A-Z]{1,3}\d*)?/g) || [];
  const cols = new Set<string>();
  refs.forEach((ref) => {
    const [left, right] = ref.split(':');
    const leftCol = (left.match(/[A-Z]{1,3}/) || [])[0];
    const rightCol = (right?.match(/[A-Z]{1,3}/) || [])[0];
    if (leftCol) cols.add(leftCol);
    if (rightCol) cols.add(rightCol);
  });
  return cols;
};

const findHeaderColumn = (headersByColumn: Record<string, string> | undefined, keyword: string) => {
  if (!headersByColumn) return null;
  const lower = keyword.toLowerCase();
  const hit = Object.entries(headersByColumn).find(([, header]) => String(header).toLowerCase().includes(lower));
  return hit ? hit[0] : null;
};

export const validateFormulaSemantics = (formula: string, prompt: string, selectionContext: any) => {
  const headersByColumn = selectionContext?.headersByColumn || {};
  const usedCols = collectColumnLetters(formula);
  const promptLower = (prompt || '').toLowerCase();

  if (promptLower.includes('quantity')) {
    const quantityCol = findHeaderColumn(headersByColumn, 'quantity');
    if (quantityCol && !usedCols.has(quantityCol)) {
      return { ok: false, reason: `Formula does not reference selected Quantity column (${quantityCol}).` };
    }
  }

  if (promptLower.includes('order') || promptLower.includes('id')) {
    const orderCol = findHeaderColumn(headersByColumn, 'order');
    if (orderCol && !usedCols.has(orderCol)) {
      return { ok: false, reason: `Formula does not reference selected Order column (${orderCol}).` };
    }
  }

  return { ok: true };
};
