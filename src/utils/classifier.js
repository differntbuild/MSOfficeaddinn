export const getClassificationPrompt = (data, type) => {
  return {
    system: `You are a data labeling expert. Categorize the following data into ${type}. 
             Return ONLY a JSON array of strings. No extra text.
             Example: ["High", "Low", "Medium"]`,
    user: `Data: ${JSON.stringify(data)}`
  };
};