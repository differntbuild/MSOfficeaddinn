export const getComparisonPrompt = (rangeA, rangeB) => {
  return {
    system: "You are a financial controller. Compare these two datasets and explain the variances, trends, and anomalies. Be precise.",
    user: `Dataset A: ${JSON.stringify(rangeA)}\n\nDataset B: ${JSON.stringify(rangeB)}`
  };
};