import React, {
  useState, useRef, useEffect,
  forwardRef, useImperativeHandle, useCallback
} from 'react';
import { useOffice } from '../hooks/useOffice';
import { streamChat } from '../hooks/useGroq';
import { getSettings, saveSettings } from '../../utils/storage';
import {
  detectIntent,
  detectDomain,
  getSystemPromptForIntent,
  SLASH_COMMANDS,
} from '../../utils/intelligence';
import { getVisualizationPrompt } from '../../utils/pivotEngine';

// ─── Helpers ────────────────────────────────────────────────────────────────

const extractFormula = (text) => {
  if (!text) return null;
  // Looks for anything starting with '=' inside triple backticks OR single backticks
  const match = text.match(/```(?:[\w]*\n)?(=[\s\S]*?)```/) || text.match(/`(=[^`]+)`/);
  if (!match) return null;
  
  const f = match[1]
    .replace(/&quot;/g, '"')
    .replace(/""/g, '"')
    .replace(/\\"/g, '"')
    .replace(/\s+/g, '')
    .trim();
    
  return f.startsWith('=') ? f : null;
};

const extractJSON = (text) => {
  if (!text) return null;
  try {
    // Look for JSON block or just the first { to last }
    const match = text.match(/```json\s*(\{[\s\S]*?\})/) || text.match(/(\{[\s\S]*\})/);
    return match ? JSON.parse(match[1]) : null;
  } catch (e) {
    return null;
  }
};

const extractSuggestions = (text) => {
  const lines = text.split('\n').filter(l => /^\d[\.\)]/.test(l.trim()));
  return lines.slice(0, 3).map(l => l.replace(/^\d[\.\)]\s*/, '').trim());
};

// ─── Message Renderer — parses [[CellRef]] citations ────────────────────────

const MessageContent = ({ text, onCitationClick }) => {
  if (!text) return null;

  const sanitized = text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  const parts = sanitized.split(/(\[\[[A-Z]+\d+(?::[A-Z]+\d+)?\]\])/g);

  return (
    <div style={{ fontSize: 12.5, lineHeight: 1.6, color: 'inherit' }}>
      {parts.map((part, i) => {
        const citMatch = part.match(/^\[\[([A-Z]+\d+(?::[A-Z]+\d+)?)\]\]$/);
        if (citMatch) {
          return (
            <button
              key={i}
              onClick={() => onCitationClick(citMatch[1])}
              style={{
                display: 'inline-flex', alignItems: 'center', gap: 2,
                background: '#E8F5EE', border: '0.5px solid #217346',
                color: '#1a7340', borderRadius: 3, padding: '0px 5px',
                fontSize: 11, fontFamily: 'Consolas, monospace',
                cursor: 'pointer', fontWeight: 600, margin: '0 1px',
                verticalAlign: 'middle',
              }}
            >
              ↗ {citMatch[1]}
            </button>
          );
        }
        // Bold
        const boldParts = part.split(/(\*\*.*?\*\*)/g);
        return boldParts.map((bp, j) => {
          if (bp.startsWith('**') && bp.endsWith('**')) {
            return <strong key={`${i}-${j}`}>{bp.slice(2, -2)}</strong>;
          }
          // Inline code
          const codeParts = bp.split(/(`[^`]+`)/g);
          return codeParts.map((cp, k) => {
            if (cp.startsWith('`') && cp.endsWith('`')) {
              return (
                <code key={`${i}-${j}-${k}`} style={{ background: '#f3f3f3', border: '0.5px solid #e0e0e0', borderRadius: 3, padding: '1px 5px', fontFamily: 'Consolas, monospace', fontSize: 11.5 }}>
                  {cp.slice(1, -1)}
                </code>
              );
            }
            return cp.split('\n').map((line, l) => (
              <React.Fragment key={`${i}-${j}-${k}-${l}`}>
                {l > 0 && <br />}
                {line}
              </React.Fragment>
            ));
          });
        });
      })}
    </div>
  );
};

// ─── Slash Command Palette ───────────────────────────────────────────────────

const SlashPalette = ({ onSelect }) => (
  <div style={{
    position: 'absolute', bottom: '100%', left: 0, right: 0,
    background: 'white', border: '0.5px solid #d0d0d0',
    borderRadius: '6px 6px 0 0', overflow: 'hidden',
    boxShadow: '0 -4px 12px rgba(0,0,0,0.08)', zIndex: 100,
  }}>
    <div style={{ padding: '6px 10px', fontSize: 10, fontWeight: 600, color: '#a09e9c', letterSpacing: '0.05em', textTransform: 'uppercase', borderBottom: '0.5px solid #f0f0f0' }}>
      Commands
    </div>
    {SLASH_COMMANDS.map(cmd => (
      <button key={cmd.cmd} onClick={() => onSelect(cmd.cmd + ' ')} style={{
        width: '100%', display: 'flex', alignItems: 'center', gap: 10,
        padding: '7px 12px', background: 'none', border: 'none',
        cursor: 'pointer', textAlign: 'left', fontFamily: 'Segoe UI, system-ui, sans-serif',
      }}
        onMouseEnter={e => e.currentTarget.style.background = '#f8f8f8'}
        onMouseLeave={e => e.currentTarget.style.background = 'none'}
      >
        <span style={{ width: 20, height: 20, background: '#e8f5ee', borderRadius: 4, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 11, color: '#217346', fontWeight: 700, flexShrink: 0 }}>
          {cmd.icon}
        </span>
        <div>
          <div style={{ fontSize: 12, fontWeight: 600, color: '#201f1e' }}>{cmd.cmd}</div>
          <div style={{ fontSize: 10.5, color: '#a09e9c' }}>{cmd.label}</div>
        </div>
      </button>
    ))}
  </div>
);

// ─── What-if Panel ───────────────────────────────────────────────────────────

const WhatIfPanel = ({ context, onClose, onApply, navigateToCell }) => {
  const [cellAddress, setCellAddress] = useState('');
  const [newValue, setNewValue] = useState('');
  const [analysis, setAnalysis] = useState('');
  const [loading, setLoading] = useState(false);

  const runWhatIf = async () => {
    if (!cellAddress || !newValue) return;
    setLoading(true);

    const prompt = `/whatif If I change [[${cellAddress}]] from its current value to ${newValue}, what happens to dependent cells?`;
    const systemMsg = getSystemPromptForIntent('WHATIF', { host: 'Excel' });

    let result = '';
    await streamChat(prompt, context, systemMsg, (chunk) => {
      result = chunk;
      setAnalysis(chunk);
    });

    setLoading(false);
  };

  return (
    <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, background: 'white', zIndex: 200, display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '12px 14px', borderBottom: '0.5px solid #e0e0e0', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#201f1e' }}>What-if Simulator</div>
          <div style={{ fontSize: 11, color: '#a09e9c' }}>Change an assumption, see the impact</div>
        </div>
        <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 16, color: '#a09e9c' }}>✕</button>
      </div>

      <div style={{ padding: '12px 14px', display: 'flex', flexDirection: 'column', gap: 10, flexShrink: 0 }}>
        <div>
          <label style={{ fontSize: 11, fontWeight: 600, color: '#605e5c', display: 'block', marginBottom: 4 }}>Cell to change</label>
          <input
            value={cellAddress}
            onChange={e => setCellAddress(e.target.value.toUpperCase())}
            placeholder="e.g. B2"
            style={{ width: '100%', border: '0.5px solid #d0d0d0', borderRadius: 4, padding: '6px 9px', fontSize: 12, fontFamily: 'Segoe UI, system-ui, sans-serif', outline: 'none', boxSizing: 'border-box' }}
          />
        </div>
        <div>
          <label style={{ fontSize: 11, fontWeight: 600, color: '#605e5c', display: 'block', marginBottom: 4 }}>New value</label>
          <input
            value={newValue}
            onChange={e => setNewValue(e.target.value)}
            placeholder="e.g. 0.25"
            style={{ width: '100%', border: '0.5px solid #d0d0d0', borderRadius: 4, padding: '6px 9px', fontSize: 12, fontFamily: 'Segoe UI, system-ui, sans-serif', outline: 'none', boxSizing: 'border-box' }}
          />
        </div>
        <button onClick={runWhatIf} disabled={loading || !cellAddress || !newValue} style={{ background: loading ? '#a0cfb5' : '#217346', color: 'white', border: 'none', borderRadius: 6, padding: '8px', fontSize: 12.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
          {loading ? 'Simulating…' : 'Run Simulation'}
        </button>
      </div>

      <div style={{ flex: 1, overflowY: 'auto', padding: '0 14px 14px' }}>
        {analysis && (
          <div style={{ background: '#f8f8f8', border: '0.5px solid #e0e0e0', borderRadius: 6, padding: 12 }}>
            <MessageContent text={analysis} onCitationClick={navigateToCell} />
          </div>
        )}
      </div>
    </div>
  );
};

// ─── Formula Explainer Panel ─────────────────────────────────────────────────

const FormulaExplainer = ({ formula, cellAddress, onClose, navigateToCell }) => {
  const [explanation, setExplanation] = useState('');
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const explain = async () => {
      const prompt = `Explain this Excel formula step by step: \`${formula}\` in cell [[${cellAddress}]]`;
      const systemMsg = getSystemPromptForIntent('EXPLANATION', { host: 'Excel' });
      await streamChat(prompt, null, systemMsg, (chunk) => setExplanation(chunk));
      setLoading(false);
    };
    explain();
  }, [formula, cellAddress]);

  return (
    <div style={{ position: 'absolute', top: 0, left: 0, right: 0, bottom: 0, background: 'white', zIndex: 200, display: 'flex', flexDirection: 'column' }}>
      <div style={{ padding: '12px 14px', borderBottom: '0.5px solid #e0e0e0', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0 }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 600, color: '#201f1e' }}>Formula Explainer</div>
          <code style={{ fontSize: 11, color: '#217346', fontFamily: 'Consolas, monospace' }}>{cellAddress}</code>
        </div>
        <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', fontSize: 16, color: '#a09e9c' }}>✕</button>
      </div>
      <div style={{ padding: '10px 14px', background: '#f8f8f8', borderBottom: '0.5px solid #e0e0e0', flexShrink: 0 }}>
        <code style={{ fontSize: 12, fontFamily: 'Consolas, monospace', color: '#201f1e', wordBreak: 'break-all' }}>{formula}</code>
      </div>
      <div style={{ flex: 1, overflowY: 'auto', padding: 14 }}>
        {loading ? (
          <div style={{ display: 'flex', gap: 4, padding: 10 }}>
            {[0, 1, 2].map(i => (
              <div key={i} style={{ width: 5, height: 5, borderRadius: '50%', background: '#a09e9c', animation: `bounce 1.2s ${i * 0.2}s infinite ease-in-out` }} />
            ))}
          </div>
        ) : (
          <MessageContent text={explanation} onCitationClick={() => {}} />
        )}
      </div>
    </div>
  );
};

// ─── Main ChatPane ───────────────────────────────────────────────────────────

const ChatPane = forwardRef(({ host }, ref) => {
  const {
    getFullSheetContext,
    getSelectionContext,
    insertFormulaBelow,
    navigateToCell,
    getFormulaAtCell,
    saveToWorkbookMemory,
    loadFromWorkbookMemory,
    createPivotAndChart,
  } = useOffice();

  const [messages, setMessages] = useState([{
    role: 'assistant',
    content: `Hi! I'm your AI assistant for ${host}. Select cells and ask anything — or type **/** .`,
    suggestions: ['Sum the selected cells', 'Explain this data', '/audit'],
  }]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [contextInfo, setContextInfo] = useState(null);
  const [overwritePrompt, setOverwritePrompt] = useState(null);
  const [showSlash, setShowSlash] = useState(false);
  const [showWhatIf, setShowWhatIf] = useState(false);
  const [formulaExplainer, setFormulaExplainer] = useState(null);
  const [sheetContext, setSheetContext] = useState(null);
  const [domain, setDomain] = useState('general');
  const [selectedModel, setSelectedModel] = useState(() => getSettings().model);
  const messagesEndRef = useRef(null);
  const inputRef = useRef(null);

  useImperativeHandle(ref, () => ({
    sendMessage: (text) => handleSend(text),
  }));

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  // Load selection + sheet context on mount, load workbook memory
  useEffect(() => {
    if (!window.Office) return;

    const init = async () => {
      try {
        const sel = await getSelectionContext();
        if (sel.cells.length > 0) setContextInfo(sel);

        const full = await getFullSheetContext();
        setSheetContext(full);

        const detectedDomain = detectDomain(full.headers, full.sheetName);
        setDomain(detectedDomain);

        // Load workbook memory
        const memory = await loadFromWorkbookMemory('lastAnalysis');
        if (memory) {
          setMessages(prev => [...prev, {
            role: 'assistant',
            content: `Welcome back! Last time I analyzed this workbook: **${memory.summary}**\n\nWant me to run a fresh analysis?`,
            suggestions: ['Yes, analyze again', 'Show last audit results', 'What changed?'],
          }]);
        }

        // Proactive insight on open
        if (full.cells.length > 0 && !memory) {
          const proactivePrompt = `Give me the top 3 most important things to know about this spreadsheet in 2 sentences each.`;
          const systemMsg = getSystemPromptForIntent('EXPLANATION', { host, domain: detectedDomain });
          let response = '';
          await streamChat(proactivePrompt, {
            dimensions: full.dimensions,
            stats: full.stats,
            columnTypes: full.columnTypes,
            sample: full.values?.slice(0, 10),
          }, systemMsg, (chunk) => {
            response = chunk;
            setMessages(prev => {
              const updated = [...prev];
              if (updated[updated.length - 1]?.role === 'assistant' && updated[updated.length - 1]?.isProactive) {
                updated[updated.length - 1] = { role: 'assistant', content: chunk, isProactive: true };
              } else {
                updated.push({ role: 'assistant', content: chunk, isProactive: true });
              }
              return updated;
            });
          });
        }
      } catch (e) {
        console.warn('Init failed:', e);
      }
    };

    init();
  }, []);

  const handleSend = useCallback(async (overrideInput) => {
    const text = (overrideInput || input).trim();
    if (!text || isLoading) return;

    setInput('');
    setShowSlash(false);
    setIsLoading(true);
    setMessages(prev => [...prev, { role: 'user', content: text }]);

    try {
      const intent = detectIntent(text);

      // /whatif opens the panel instead of chatting
      if (intent === 'WHATIF') {
        setShowWhatIf(true);
        setIsLoading(false);
        return;
      }

      if (intent === 'VISUALIZATION' && window.Office && host === 'Excel') {
        try {
          const full = sheetContext || await getFullSheetContext();
          const headers = full.values ? full.values[0] : []; // Get top row headers
          
          const vizPrompt = getVisualizationPrompt(headers, text);
          
          setMessages(prev => [...prev, { role: 'assistant', content: 'Analyzing data layout and building your dashboard...' }]);
          
          // Call Groq (forcing it to think strictly in JSON for this task)
          let jsonResponse = '';
          await streamChat(vizPrompt, { headers }, "Output ONLY valid JSON.", (chunk) => {
            jsonResponse = chunk;
          });

          // Parse AI response and build the chart
          const config = JSON.parse(jsonResponse.replace(/```json/g, '').replace(/```/g, '').trim());
          const newSheetName = await createPivotAndChart(config);

          setMessages(prev => [...prev, { 
            role: 'assistant', 
            content: `Dashboard successfully generated! I created a Pivot Table and Chart in a new sheet called **${newSheetName}**.` 
          }]);

        } catch (err) {
          setMessages(prev => [...prev, { role: 'assistant', content: `Failed to build visualization: ${err.message}` }]);
        }
        setIsLoading(false);
        return;
      }

      // Build context based on intent
      let contextData = null;
      if (window.Office) {
        if (intent === 'AUDIT' || intent === 'EXPLANATION' || intent === 'PIVOT') {
          const full = sheetContext || await getFullSheetContext();
          contextData = {
            sheetName: full.sheetName,
            dimensions: full.dimensions,
            headers: full.headers,
            columnTypes: full.columnTypes,
            stats: full.stats,
            sample: full.values?.slice(0, 30),
            cells: full.cells?.slice(0, 200),
          };
        } else {
          const sel = await getSelectionContext();
          contextData = sel.cells.length > 0 ? sel : (sheetContext ? {
            dimensions: sheetContext.dimensions,
            sample: sheetContext.values?.slice(0, 10),
          } : null);
        }
      }

      const systemMsg = getSystemPromptForIntent(intent, { host, domain, data: contextInfo });
      let fullResponse = '';
      let assistantAdded = false;

      await streamChat(text, contextData, systemMsg, (chunk) => {
        fullResponse = chunk;
        setMessages(prev => {
          const updated = [...prev];
          if (!assistantAdded) {
            updated.push({ role: 'assistant', content: chunk });
            assistantAdded = true;
          } else {
            updated[updated.length - 1] = { role: 'assistant', content: chunk };
          }
          return updated;
        });
      });

      // Formula auto-apply
      const formula = extractFormula(fullResponse);
      if (formula && intent === 'FORMULA' && window.Office) {
        const result = await insertFormulaBelow(formula);
        if (result?.blocked) {
          setOverwritePrompt({ formula, address: result.targetAddress, existingValue: result.existingValue });
        } else if (result?.success) {
          setMessages(prev => {
            const updated = [...prev];
            updated[updated.length - 1] = { ...updated[updated.length - 1], appliedFormula: formula, appliedAddress: result.address };
            return updated;
          });
        }
      }

      // Save to workbook memory after explanation/audit
      if ((intent === 'EXPLANATION' || intent === 'AUDIT') && window.Office) {
        const summary = fullResponse.slice(0, 200).replace(/\[\[.*?\]\]/g, '');
        await saveToWorkbookMemory('lastAnalysis', { summary, timestamp: new Date().toISOString(), intent });
      }

      // Visualization / Pivot auto-apply
      if (intent === 'PIVOT' && window.Office) {
        const config = extractJSON(fullResponse);
        if (config && (config.rows || config.chartType)) {
          try {
            const dashName = await createPivotAndChart(config);
            setMessages(prev => [...prev, {
              role: 'assistant',
              content: `I've created a new dashboard sheet **${dashName}** with the requested chart and pivot table.`,
            }]);
          } catch (e) {
            console.error("Dashboard creation failed:", e);
          }
        }
      }

      // Suggestions
      const suggestions = extractSuggestions(fullResponse);
      if (suggestions.length > 0) {
        setMessages(prev => {
          const updated = [...prev];
          updated[updated.length - 1] = { ...updated[updated.length - 1], suggestions };
          return updated;
        });
      }

    } catch (err) {
      setMessages(prev => [...prev, { role: 'assistant', content: `Error: ${err.message}` }]);
    }

    setIsLoading(false);
  }, [input, isLoading, sheetContext, host, domain]);

  const handleInputChange = (e) => {
    const val = e.target.value;
    setInput(val);
    setShowSlash(val === '/' || val.startsWith('/') && val.length < 3);
  };

  const handleOverwriteConfirm = async () => {
    if (!overwritePrompt) return;
    const result = await insertFormulaBelow(overwritePrompt.formula, { forceOverwrite: true });
    setOverwritePrompt(null);
    if (result?.success) {
      setMessages(prev => [...prev, { role: 'assistant', content: `Formula applied to [[${result.address}]].` }]);
    }
  };

  // Formula explainer — triggered when user clicks a formula cell
  const handleExplainFormula = async () => {
    if (!contextInfo?.address || !window.Office) return;
    try {
      const cellData = await getFormulaAtCell(contextInfo.address.split('!').pop().split(':')[0]);
      if (cellData.formula && cellData.formula.toString().startsWith('=')) {
        setFormulaExplainer({ formula: cellData.formula, cellAddress: cellData.address });
      } else {
        setMessages(prev => [...prev, { role: 'assistant', content: 'The selected cell does not contain a formula.' }]);
      }
    } catch (e) {
      console.warn('Formula explainer failed:', e);
    }
  };

  const handleModelChange = (model) => {
    setSelectedModel(model);
    const current = getSettings();
    saveSettings({ ...current, model });
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', position: 'relative' }}>

      {/* Formula Explainer overlay */}
      {formulaExplainer && (
        <FormulaExplainer
          formula={formulaExplainer.formula}
          cellAddress={formulaExplainer.cellAddress}
          onClose={() => setFormulaExplainer(null)}
          navigateToCell={navigateToCell}
        />
      )}

      {/* What-if overlay */}
      {showWhatIf && (
        <WhatIfPanel
          context={sheetContext}
          onClose={() => setShowWhatIf(false)}
          navigateToCell={navigateToCell}
        />
      )}

      {/* Selection chip */}
      {contextInfo && (
        <div style={{ margin: '6px 10px 0', padding: '5px 10px', background: 'white', border: '0.5px solid #e0e0e0', borderRadius: 4, display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: '#605e5c', flexShrink: 0 }}>
          <div style={{ width: 6, height: 6, borderRadius: '50%', background: '#217346', flexShrink: 0 }} />
          <span style={{ fontWeight: 500, color: '#201f1e' }}>{contextInfo.address}</span>
          <span>·</span>
          <span>{contextInfo.cells.length} cells</span>
          {domain !== 'general' && (
            <span style={{ marginLeft: 'auto', fontSize: 10, background: '#e8f5ee', color: '#217346', padding: '1px 6px', borderRadius: 10, fontWeight: 600 }}>
              {domain}
            </span>
          )}
          <button
            onClick={handleExplainFormula}
            style={{ marginLeft: domain !== 'general' ? 4 : 'auto', fontSize: 10, background: '#f3f3f3', border: '0.5px solid #e0e0e0', borderRadius: 3, padding: '2px 7px', cursor: 'pointer', color: '#605e5c', fontFamily: 'Segoe UI, system-ui, sans-serif' }}
          >
            fx?
          </button>
        </div>
      )}

      {/* Messages */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '10px 10px 6px', display: 'flex', flexDirection: 'column', gap: 10, background: '#f8f8f8' }}>
        {messages.map((msg, idx) => (
          <div key={idx} style={{ display: 'flex', flexDirection: 'column', alignSelf: msg.role === 'user' ? 'flex-end' : 'flex-start', maxWidth: '92%' }}>
            <div style={{
              padding: '9px 12px',
              borderRadius: msg.role === 'user' ? '12px 12px 3px 12px' : '3px 12px 12px 12px',
              background: msg.role === 'user' ? '#217346' : 'white',
              color: msg.role === 'user' ? 'white' : '#201f1e',
              border: msg.role === 'assistant' ? '0.5px solid #e0e0e0' : 'none',
            }}>
              {msg.role === 'user'
                ? <div style={{ fontSize: 12.5, lineHeight: 1.5 }}>{msg.content}</div>
                : <MessageContent text={msg.content} onCitationClick={navigateToCell} />
              }
            </div>

            {msg.appliedFormula && (
              <div style={{ marginTop: 4, fontSize: 11, background: '#e8f5ee', border: '0.5px solid #c3e6d0', borderRadius: 20, padding: '3px 10px', color: '#1a7340', alignSelf: 'flex-start' }}>
                ✓ Applied to {msg.appliedAddress}
              </div>
            )}

            {msg.suggestions && msg.suggestions.length > 0 && (
              <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginTop: 6 }}>
                {msg.suggestions.map((s, i) => (
                  <button key={i} onClick={() => handleSend(s)} style={{ fontSize: 11, padding: '3px 10px', borderRadius: 20, cursor: 'pointer', background: '#e8f5ee', border: '0.5px solid #c3e6d0', color: '#217346', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>
                    {s}
                  </button>
                ))}
              </div>
            )}
          </div>
        ))}

        {isLoading && (
          <div style={{ alignSelf: 'flex-start', background: 'white', border: '0.5px solid #e0e0e0', borderRadius: '3px 12px 12px 12px', padding: '10px 14px' }}>
            <div style={{ display: 'flex', gap: 4 }}>
              {[0, 1, 2].map(i => (
                <div key={i} style={{ width: 5, height: 5, borderRadius: '50%', background: '#a09e9c', animation: `bounce 1.2s ${i * 0.2}s infinite ease-in-out` }} />
              ))}
            </div>
          </div>
        )}
        <div ref={messagesEndRef} />
      </div>

      {/* Overwrite dialog */}
      {overwritePrompt && (
        <div style={{ margin: '0 10px', padding: 10, background: '#FFF4E5', border: '0.5px solid #EF9F27', borderRadius: 6, fontSize: 12, flexShrink: 0 }}>
          <div style={{ fontWeight: 600, color: '#633806', marginBottom: 4 }}>
            ⚠ Cell {overwritePrompt.address} has: "{overwritePrompt.existingValue}"
          </div>
          <div style={{ color: '#605e5c', marginBottom: 8 }}>
            Overwrite with <code style={{ fontFamily: 'Consolas, monospace' }}>{overwritePrompt.formula}</code>?
          </div>
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={handleOverwriteConfirm} style={{ background: '#217346', color: 'white', border: 'none', borderRadius: 4, padding: '4px 12px', fontSize: 11.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Overwrite</button>
            <button onClick={() => setOverwritePrompt(null)} style={{ background: 'white', color: '#605e5c', border: '0.5px solid #d0d0d0', borderRadius: 4, padding: '4px 12px', fontSize: 11.5, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Cancel</button>
          </div>
        </div>
      )}

      {/* Input area */}
      <div style={{ position: 'relative', padding: '8px 10px 10px', background: 'white', borderTop: '0.5px solid #e0e0e0', flexShrink: 0 }}>
        {showSlash && <SlashPalette onSelect={(cmd) => { setInput(cmd); setShowSlash(false); inputRef.current?.focus(); }} />}
        
        <div style={{ display: 'flex', alignItems: 'flex-end', gap: 8 }}>
          {/* Model Switcher Pill */}
          <div style={{ position: 'relative', flexShrink: 0, marginBottom: 4 }}>
            <select
              value={selectedModel}
              onChange={(e) => handleModelChange(e.target.value)}
              className="settings-select"
              style={{
                background: '#f3f3f3',
                border: '0.5px solid #d0d0d0',
                borderRadius: 6,
                padding: '2px 8px',
                fontSize: 10,
                fontWeight: 600,
                color: '#605e5c',
                height: 24,
                width: 100,
                appearance: 'none',
                textAlign: 'center'
              }}
            >
              <option value="groq/compound">Groq Comp</option>
              <option value="llama-3.3-70b-versatile">Llama 70B</option>
              <option value="llama-3.1-8b-instant">Llama 8B</option>
            </select>
            <div style={{ position: 'absolute', left: 6, top: '50%', transform: 'translateY(-50%)', width: 4, height: 4, borderRadius: '50%', background: '#217346' }} />
          </div>

          <textarea
            ref={inputRef}
            value={input}
            onChange={handleInputChange}
            onKeyDown={e => {
              if (e.key === 'Escape') setShowSlash(false);
              if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSend(); }
            }}
            onInput={e => { e.target.style.height = 'auto'; e.target.style.height = Math.min(e.target.scrollHeight, 80) + 'px'; }}
            placeholder="Ask anything or type / "
            rows={1}
            style={{ flex: 1, border: '0.5px solid #d0d0d0', borderRadius: 6, padding: '7px 10px', fontFamily: 'Segoe UI, system-ui, sans-serif', fontSize: 12.5, resize: 'none', lineHeight: 1.45, minHeight: 36, maxHeight: 80, background: '#f8f8f8', outline: 'none' }}
          />
          <button
            onClick={() => handleSend()}
            disabled={isLoading || !input.trim()}
            style={{ width: 34, height: 34, background: isLoading || !input.trim() ? '#a0cfb5' : '#217346', border: 'none', borderRadius: 6, cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}
          >
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none"><path d="M2 8L14 2L10 8L14 14L2 8Z" fill="white" /></svg>
          </button>
        </div>
      </div>
    </div>
  );
});

ChatPane.displayName = 'ChatPane';
export default ChatPane;