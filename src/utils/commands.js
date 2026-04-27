export const getSystemPrompt = (command, host) => {
  const base = `You are an expert Office ${host} assistant. Current Year: 2026.`;
  
  const strategies = {
    '/formula': `${base} You are a formula engine. Return ONLY the Excel formula starting with =. No prose, no backticks, no explanation.`,
    '/rewrite': `${base} Rewrite the selected text to be more professional and concise. Maintain the original meaning.`,
    '/summarize': `${base} Provide a 3-bullet point executive summary of the provided text/data.`,
    '/clean': `${base} Analyze the provided data for duplicates, formatting errors, or nulls. List the issues clearly.`,
    '/explain': `${base} Explain the logic behind this ${host === 'Excel' ? 'formula/data' : 'paragraph'} simply for a beginner.`
  };

  const activeCommand = Object.keys(strategies).find(cmd => command.startsWith(cmd));
  return activeCommand ? strategies[activeCommand] : `${base} Use Markdown for structure.`;
};