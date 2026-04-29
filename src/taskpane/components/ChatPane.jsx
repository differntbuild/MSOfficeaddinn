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
import { ExecutionEngine } from '../../agent/engine/ExecutionEngine';
import { AuditEngine } from '../../agent/engine/AuditEngine';
import { AuditDashboard } from './AuditDashboard';

const CHAT_HISTORY_PREFIX = 'groqflow_chat_history_v1';
const MAX_HISTORY_MESSAGES = 60;
const MAX_PROMPT_HISTORY = 50;
const baseWelcomeMessage = (host) => ({
  role: 'assistant',
  content: `Hi! I'm your AI assistant for ${host}. Select cells and ask anything — or type **/** .`,
  suggestions: ['Sum the selected cells', 'Explain this data', '/audit'],
});

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

const sanitizeFormula = (formula) => {
  if (!formula) return formula;
  return formula
    .replace(/↗/g, '')
    .replace(/\[\[[^\]]+\]\]/g, '')
    .replace(/\s+/g, '')
    .trim();
};

const validateFormula = (formula) => {
  if (!formula) return { ok: false, reason: 'No formula was generated.' };
  if (!formula.startsWith('=')) return { ok: false, reason: 'Formula must start with "=".' };
  if (/\[\[|\]\]/.test(formula)) return { ok: false, reason: 'Formula still contains citation markers.' };
  if (/[^A-Za-z0-9_:\+\-\*\/\^\(\),."'%!<>=& ]/.test(formula)) {
    return { ok: false, reason: 'Formula contains invalid characters.' };
  }

  const xlookupMatch = formula.match(/XLOOKUP\s*\((.+)\)/i);
  if (xlookupMatch) {
    const args = xlookupMatch[1].split(',').map(a => a.trim());
    if (args.length >= 2) {
      const lookupArray = args[1];
      const rangeMatch = lookupArray.match(/^([A-Z]+)\d*:[A-Z]+\d*$/i);
      if (rangeMatch) {
        const leftCol = lookupArray.split(':')[0].match(/[A-Z]+/i)?.[0]?.toUpperCase();
        const rightCol = lookupArray.split(':')[1].match(/[A-Z]+/i)?.[0]?.toUpperCase();
        if (leftCol && rightCol && leftCol !== rightCol) {
          return { ok: false, reason: 'Invalid XLOOKUP lookup_array: must be a single row or single column range.' };
        }
      }
    }
  }
  return { ok: true };
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

const getChatHistoryKey = (host) => `${CHAT_HISTORY_PREFIX}_${host || 'Excel'}`;

const isPersistableMessage = (msg) =>
  msg && (msg.role === 'assistant' || msg.role === 'user') && typeof msg.content === 'string';

const sanitizeHistory = (msgs = []) =>
  msgs
    .filter(isPersistableMessage)
    .map(m => ({
      role: m.role,
      content: m.content,
      suggestions: Array.isArray(m.suggestions) ? m.suggestions.slice(0, 3) : undefined,
      appliedFormula: m.appliedFormula || undefined,
      appliedAddress: m.appliedAddress || undefined,
    }))
    .slice(-MAX_HISTORY_MESSAGES);

const SCAN_INTENTS = new Set(['AUDIT', 'EXPLANATION', 'PIVOT', 'METRICS', 'VISUALIZATION']);

const normalizeColumnTypes = (columnTypes) => {
  if (!columnTypes) return undefined;
  if (Array.isArray(columnTypes)) {
    const mapped = {};
    columnTypes.forEach(item => {
      if (item?.header) mapped[item.header] = item.type || 'unknown';
    });
    return mapped;
  }
  return columnTypes;
};

const summarizeSheet = (full, limits) => {
  const sampleRows = Math.max(5, limits?.sampleRows || 30);
  const maxCells = Math.max(100, limits?.fullCells || 400);
  const rows = full.values || [];
  const header = rows[0] || [];
  const bodyRows = rows.slice(1, sampleRows + 1);
  const maxCols = header.length ? Math.max(1, Math.min(header.length, Math.floor(maxCells / Math.max(1, sampleRows)))) : 0;

  return {
    scope: 'sheet-summary',
    sheetName: full.sheetName,
    dimensions: full.dimensions,
    headers: maxCols > 0 ? header.slice(0, maxCols) : full.headers,
    columnTypes: normalizeColumnTypes(full.columnTypes),
    stats: full.stats,
    sample: bodyRows.map(r => (maxCols > 0 ? r.slice(0, maxCols) : r)),
    rowCount: full.rowCount,
    columnCount: full.columnCount,
    truncated: true,
  };
};

const packSelectionContext = (selection, limits) => {
  if (!selection) return null;
  const maxCells = Math.max(50, limits?.selectionCells || 200);
  const cells = Array.isArray(selection.cells) ? selection.cells.slice(0, maxCells) : [];
  return {
    scope: 'selection',
    address: selection.address,
    cells,
    cellCount: Array.isArray(selection.cells) ? selection.cells.length : cells.length,
    truncated: (selection.cells?.length || 0) > cells.length,
    headersByColumn: selection.headersByColumn || {},
    selectedHeaders: selection.selectedHeaders || [],
  };
};

const buildContextPayload = ({ intent, selection, full, settings, host }) => {
  const toggles = settings?.toggles || {};
  const contextMode = settings?.contextMode || 'selection';
  const limits = settings?.contextLimits || {};

  if (!toggles.context) {
    return {
      scope: 'disabled',
      host,
      note: 'Context sharing disabled by user.',
    };
  }

  const selectionPacked = packSelectionContext(selection, limits);
  const fullSummary = summarizeSheet(full, limits);

  if (!SCAN_INTENTS.has(intent)) {
    if (contextMode === 'sheet') return fullSummary;
    if (contextMode === 'table') {
      return {
        ...fullSummary,
        scope: 'table-region',
        note: 'Using active used-range as current table/region context.',
      };
    }
    return selectionPacked || fullSummary;
  }

  const payload = {
    ...(contextMode === 'selection' ? (selectionPacked || fullSummary) : fullSummary),
    scope: contextMode === 'selection' ? 'selection' : contextMode === 'table' ? 'table-region' : 'sheet-summary',
    selection: selectionPacked || undefined,
  };

  if (!toggles.sheetNames) delete payload.sheetNames;
  else payload.sheetNames = full?.sheetNames;

  if (!toggles.namedRanges) delete payload.namedRanges;
  else payload.namedRanges = full?.namedRanges;

  if (!toggles.dataTypes) delete payload.columnTypes;

  return payload;
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

const ToolbarButton = ({ onClick, title, label, disabled = false, icon }) => (
  <button
    onClick={onClick}
    disabled={disabled}
    title={title}
    style={{
      display: 'inline-flex',
      alignItems: 'center',
      gap: 5,
      height: 28,
      borderRadius: 6,
      border: disabled ? '0.5px dashed #d0d0d0' : '0.5px solid #d0d0d0',
      background: disabled ? '#fafafa' : '#fff',
      color: disabled ? '#a09e9c' : '#605e5c',
      padding: '0 8px',
      fontSize: 10.5,
      cursor: disabled ? 'not-allowed' : 'pointer',
      fontFamily: 'Segoe UI, system-ui, sans-serif',
      whiteSpace: 'nowrap',
    }}
  >
    {icon}
    <span>{label}</span>
  </button>
);

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

  const [messages, setMessages] = useState([baseWelcomeMessage(host)]);
  const [input, setInput] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [contextInfo, setContextInfo] = useState(null);
  const [overwritePrompt, setOverwritePrompt] = useState(null);
  const [showSlash, setShowSlash] = useState(false);
  const [showWhatIf, setShowWhatIf] = useState(false);
  const [auditReport, setAuditReport] = useState(null);
  const [formulaExplainer, setFormulaExplainer] = useState(null);
  const [sheetContext, setSheetContext] = useState(null);
  const [domain, setDomain] = useState('general');
  const [selectedModel, setSelectedModel] = useState(() => getSettings().model);
  const [contextMode, setContextMode] = useState(() => getSettings().contextMode || 'selection');
  const [promptHistory, setPromptHistory] = useState([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [historyPreview, setHistoryPreview] = useState('');
  const [pendingFormulaProposal, setPendingFormulaProposal] = useState(null);
  const messagesEndRef = useRef(null);
  const inputRef = useRef(null);
  const hasHydratedHistoryRef = useRef(false);

  const refreshSelectionContext = useCallback(async () => {
    if (!window.Office) return;
    try {
      const sel = await getSelectionContext();
      if (sel?.address) setContextInfo(sel);
    } catch (e) {
      console.warn('Selection refresh failed:', e);
    }
  }, [getSelectionContext]);

  useImperativeHandle(ref, () => ({
    sendMessage: (text) => handleSend(text),
  }));

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages]);

  useEffect(() => {
    refreshSelectionContext();
  }, [refreshSelectionContext]);

  useEffect(() => {
    if (!window.Office || !Office?.context?.document?.addHandlerAsync) return;

    const onSelectionChanged = () => {
      refreshSelectionContext();
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      onSelectionChanged,
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn('Failed to attach selection listener:', result.error?.message);
        }
      }
    );

    return () => {
      if (!Office?.context?.document?.removeHandlerAsync) return;
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        { handler: onSelectionChanged },
        () => {}
      );
    };
  }, [refreshSelectionContext]);

  useEffect(() => {
    if (hasHydratedHistoryRef.current) return;
    hasHydratedHistoryRef.current = true;
    try {
      const raw = localStorage.getItem(getChatHistoryKey(host));
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        setMessages(sanitizeHistory(parsed));
      }
    } catch (e) {
      console.warn('Failed to hydrate chat history:', e);
    }
  }, [host]);

  useEffect(() => {
    if (!hasHydratedHistoryRef.current) return;
    try {
      localStorage.setItem(getChatHistoryKey(host), JSON.stringify(sanitizeHistory(messages)));
    } catch (e) {
      console.warn('Failed to persist chat history:', e);
    }
  }, [messages, host]);

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

        const hasChatHistory = (() => {
          try {
            const raw = localStorage.getItem(getChatHistoryKey(host));
            const parsed = raw ? JSON.parse(raw) : [];
            return Array.isArray(parsed) && parsed.length > 0;
          } catch {
            return false;
          }
        })();

        // Load workbook memory
        const memory = await loadFromWorkbookMemory('lastAnalysis');
        if (memory && !hasChatHistory) {
          setMessages(prev => [...prev, {
            role: 'assistant',
            content: `Welcome back! Last time I analyzed this workbook: **${memory.summary}**\n\nWant me to run a fresh analysis?`,
            suggestions: ['Yes, analyze again', 'Show last audit results', 'What changed?'],
          }]);
        }

        // Proactive insight on open
        if (full.cells.length > 0 && !memory && !hasChatHistory) {
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
    setHistoryPreview('');
    setHistoryIndex(-1);
    setPromptHistory(prev => [...prev, text].slice(-MAX_PROMPT_HISTORY));
    setIsLoading(true);
    setMessages(prev => [...prev, { role: 'user', content: text }]);

    try {
      const intent = detectIntent(text);
      const engine = new ExecutionEngine({
        streamChat,
        host,
        domain,
        getSystemPromptForIntent,
      });

      // /whatif opens the panel instead of chatting
      if (intent === 'WHATIF') {
        setShowWhatIf(true);
        setIsLoading(false);
        return;
      }

      if (intent === 'FORMULA' && window.Office) {
        const runtimeSettings = getSettings();
        const full = sheetContext || await getFullSheetContext();
        const sel = await getSelectionContext();
        if (sel?.cells?.length >= 1) setContextInfo(sel);
        const contextData = buildContextPayload({
          intent,
          selection: sel,
          full,
          settings: runtimeSettings,
          host,
        });

        setMessages(prev => [...prev, { role: 'assistant', content: 'Generating formula draft...' }]);

        const result = await engine.executeFormula(text, contextData, {
          onChunk: (chunk) => {
            setMessages(prev => {
              const updated = [...prev];
              updated[updated.length - 1] = { role: 'assistant', content: chunk };
              return updated;
            });
          },
        });

        if (!result.ok) {
          setMessages(prev => [...prev, {
            role: 'assistant',
            content: `I could not produce a safe formula after validation retries.\n\nReason: ${result.error}`,
          }]);
          setIsLoading(false);
          return;
        }

        const draft = result.proposal;
        setPendingFormulaProposal({
          formula: draft.formula,
          explanation: draft.explanation,
          logs: result.logs,
        });
        setMessages(prev => {
          const updated = [...prev];
          updated[updated.length - 1] = {
            role: 'assistant',
            content: `\`${draft.formula}\`\n${draft.explanation}\n\nPreview ready. Click **Apply Formula** to commit.`,
          };
          return updated;
        });
        setIsLoading(false);
      }

      if (intent === 'AUDIT' && window.Office) {
        const runtimeSettings = getSettings();
        const full = sheetContext || await getFullSheetContext();
        const sel = await getSelectionContext();
        if (sel?.cells?.length >= 1) setContextInfo(sel);
        const contextData = buildContextPayload({
          intent,
          selection: sel,
          full,
          settings: runtimeSettings,
          host,
        });

        setMessages(prev => [...prev, { role: 'assistant', content: 'Running hybrid intelligence audit...' }]);

        const engine = new AuditEngine({
          streamChat,
          host,
          domain,
          getSystemPromptForIntent,
        });

        const result = await engine.executeAudit(contextData, {
          onStatus: (msg) => {
            setMessages(prev => {
              const updated = [...prev];
              updated[updated.length - 1] = { role: 'assistant', content: msg };
              return updated;
            });
          }
        });

        if (result.ok && result.report) {
          setAuditReport(result.report);
          setMessages(prev => {
             const updated = [...prev];
             updated[updated.length - 1] = { role: 'assistant', content: `Audit complete. Health Score: ${result.report.health_score}` };
             return updated;
          });
          
          // Save to workbook memory
          await saveToWorkbookMemory('lastAnalysis', { summary: result.report.summary, timestamp: new Date().toISOString(), intent });
        } else {
          setMessages(prev => [...prev, { role: 'assistant', content: `Audit failed: ${result.error}` }]);
        }
        setIsLoading(false);
        return;
      }

      if ((intent === 'VISUALIZATION' || intent === 'METRICS') && window.Office && host === 'Excel') {
        try {
          const full = sheetContext || await getFullSheetContext();
          const headers = full.headers || [];
          
          const vizPrompt = getVisualizationPrompt(headers, text);
          
          setMessages(prev => [...prev, { 
            role: 'assistant', 
            thought: `Analyzing ${full.dimensions}... Mapping headers: ${headers.join(', ')}`,
            content: intent === 'METRICS' ? 'Designing your Executive Performance Dashboard...' : 'Analyzing data layout and building your chart...' 
          }]);
          
          let jsonResponse = '';
          await streamChat(vizPrompt, { headers }, "Output ONLY valid JSON.", (chunk) => {
            jsonResponse = chunk;
            // Update the "thought" block so user sees progress
            setMessages(prev => {
              const updated = [...prev];
              updated[updated.length - 1].thought = `Building specification...\n${chunk.slice(-100)}`;
              return updated;
            });
          }, { isJson: true, intent });

          // Process the JSON response
          let config;
          try {
            config = JSON.parse(jsonResponse.replace(/```json/g, '').replace(/```/g, '').trim());
          } catch (parseErr) {
            // FALLBACK: If AI returned text instead of JSON, just show it as a normal message
            setMessages(prev => {
              const updated = [...prev];
              updated[updated.length - 1].content = jsonResponse;
              updated[updated.length - 1].thought = "AI provided a textual response instead of a dashboard specification.";
              return updated;
            });
            setIsLoading(false);
            return;
          }
          
          let result;
          if (intent === 'METRICS' || config.pivots || config.metrics) {
            result = await createAutonomousDashboard(config);
          } else {
            result = await createPivotAndChart(config);
          }

          setMessages(prev => {
            const updated = [...prev];
            updated[updated.length - 1].content = `✨ Done! I've generated the **${config.title || 'Insights'}** on a new sheet called **${result.sheetName}**.`;
            updated[updated.length - 1].thought = `Dashboard layout: ${JSON.stringify(config, null, 2)}`;
            return updated;
          });

        } catch (err) {
          setMessages(prev => [...prev, { role: 'assistant', content: `Failed to build dashboard: ${err.message}` }]);
        }
        setIsLoading(false);
        return;
      }

      // Build context based on intent + user context policy settings.
      let contextData = null;
      if (window.Office) {
        const runtimeSettings = getSettings();
        const full = sheetContext || await getFullSheetContext();
        const sel = await getSelectionContext();
        if (sel?.cells?.length >= 1) setContextInfo(sel);
        contextData = buildContextPayload({
          intent,
          selection: sel,
          full,
          settings: runtimeSettings,
          host,
        });
      }

      const systemMsg = getSystemPromptForIntent(intent, { host, domain, data: contextData });
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
    if (val.length > 0 && historyPreview) {
      setHistoryPreview('');
      setHistoryIndex(-1);
    }
  };

  const previewHistory = (direction) => {
    if (!promptHistory.length) return;
    const nextIndex = historyIndex < 0
      ? promptHistory.length - 1
      : Math.max(0, Math.min(promptHistory.length - 1, historyIndex + direction));
    setHistoryIndex(nextIndex);
    setHistoryPreview(promptHistory[nextIndex]);
  };

  const clearChat = () => {
    setMessages([baseWelcomeMessage(host)]);
    setInput('');
    setHistoryPreview('');
    setHistoryIndex(-1);
    setPendingFormulaProposal(null);
    localStorage.removeItem(getChatHistoryKey(host));
  };

  const copyResponse = async (text) => {
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
    } catch (e) {
      console.warn('Copy failed:', e);
    }
  };

  const handleOverwriteConfirm = async () => {
    if (!overwritePrompt) return;
    const result = await insertFormulaBelow(overwritePrompt.formula, { forceOverwrite: true });
    setOverwritePrompt(null);
    if (result?.success) {
      setMessages(prev => [...prev, { role: 'assistant', content: `Formula applied to [[${result.address}]].` }]);
    }
  };

  const handleApplyFormulaProposal = async () => {
    if (!pendingFormulaProposal) return;
    const result = await insertFormulaBelow(pendingFormulaProposal.formula);
    if (result?.blocked) {
      setOverwritePrompt({ formula: pendingFormulaProposal.formula, address: result.targetAddress, existingValue: result.existingValue });
      return;
    }
    if (result?.success) {
      setMessages(prev => [...prev, { role: 'assistant', content: `Formula applied to [[${result.address}]].` }]);
      setPendingFormulaProposal(null);
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

  const handleContextModeChange = (mode) => {
    setContextMode(mode);
    const current = getSettings();
    saveSettings({ ...current, contextMode: mode });
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

      <div style={{ margin: '6px 10px 0', padding: '8px 10px', background: 'white', border: '0.5px solid #e0e0e0', borderRadius: 8, display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
        <code style={{ fontSize: 10.5, fontFamily: 'Consolas, monospace', background: '#f8f8f8', border: '0.5px solid #e0e0e0', borderRadius: 6, padding: '2px 8px', color: '#201f1e' }} title="Active selection">
          {contextInfo?.address || 'No selection'}
        </code>
        <span style={{ fontSize: 10.5, color: '#605e5c' }}>{contextInfo?.cells?.length || 0} cells</span>
        <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 6 }}>
          <div title="Choose how much spreadsheet context is sent to AI" style={{ display: 'inline-flex', alignItems: 'center', gap: 5, border: '0.5px solid #d0d0d0', borderRadius: 6, padding: '0 6px', height: 28, background: '#fff' }}>
            <span style={{ width: 12, height: 12, color: '#605e5c' }}>
              <svg width="12" height="12" viewBox="0 0 16 16" fill="none"><path d="M2 3H14M2 8H14M2 13H14" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>
            </span>
            <select
              value={contextMode}
              onChange={(e) => handleContextModeChange(e.target.value)}
              className="settings-select"
              style={{ border: 'none', outline: 'none', fontSize: 11, background: 'transparent', color: '#201f1e', fontFamily: 'Segoe UI, system-ui, sans-serif' }}
            >
              <option value="selection">Scope: Selection</option>
              <option value="table">Scope: Table/Region</option>
              <option value="sheet">Scope: Whole Sheet</option>
            </select>
          </div>
          <ToolbarButton
            onClick={clearChat}
            title="Clear chat history for this host"
            label="Clear"
            icon={<svg width="12" height="12" viewBox="0 0 16 16" fill="none"><path d="M3 4H13M6 4V3A1 1 0 017 2H9A1 1 0 0110 3V4M5 4V13A1 1 0 006 14H10A1 1 0 0011 13V4" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round"/></svg>}
          />
        </div>
      </div>

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
              boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
            }}>
              {msg.role === 'assistant' && (
                <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 4 }}>
                  <button
                    onClick={() => copyResponse(msg.content)}
                    style={{ fontSize: 10, border: '0.5px solid #d0d0d0', background: '#fafafa', color: '#605e5c', borderRadius: 4, padding: '1px 6px', cursor: 'pointer', display: 'inline-flex', alignItems: 'center', gap: 4 }}
                    title="Copy exact response text"
                  >
                    <svg width="10" height="10" viewBox="0 0 16 16" fill="none"><rect x="6" y="2" width="8" height="10" rx="1" stroke="currentColor" strokeWidth="1.2"/><rect x="2" y="6" width="8" height="8" rx="1" stroke="currentColor" strokeWidth="1.2"/></svg>
                    Copy
                  </button>
                </div>
              )}
              {msg.thought && (
                <details style={{ marginBottom: 6, opacity: 0.8, fontSize: 11 }}>
                  <summary style={{ fontStyle: 'italic', color: '#666', cursor: 'pointer', outline: 'none' }}>Analysis process...</summary>
                  <div style={{ marginTop: 4, padding: '4px 6px', borderLeft: '1.5px solid #217346', background: '#f9f9f9', fontFamily: 'monospace', whiteSpace: 'pre-wrap', fontSize: 10.5 }}>
                    {msg.thought}
                  </div>
                </details>
              )}
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
      {pendingFormulaProposal && (
        <div style={{ margin: '0 10px 8px', padding: 10, background: '#EFF6FF', border: '0.5px solid #93C5FD', borderRadius: 6, fontSize: 12, flexShrink: 0 }}>
          <div style={{ fontWeight: 600, color: '#1E3A8A', marginBottom: 4 }}>
            Formula Preview (Validated)
          </div>
          <code style={{ display: 'block', marginBottom: 6, fontFamily: 'Consolas, monospace', fontSize: 11, color: '#1E3A8A', background: '#fff', border: '0.5px solid #bfdbfe', borderRadius: 4, padding: '4px 6px' }}>
            {pendingFormulaProposal.formula}
          </code>
          <div style={{ color: '#1f2937', marginBottom: 8 }}>{pendingFormulaProposal.explanation}</div>
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={handleApplyFormulaProposal} style={{ background: '#217346', color: 'white', border: 'none', borderRadius: 4, padding: '4px 12px', fontSize: 11.5, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Apply Formula</button>
            <button onClick={() => setPendingFormulaProposal(null)} style={{ background: 'white', color: '#605e5c', border: '0.5px solid #d0d0d0', borderRadius: 4, padding: '4px 12px', fontSize: 11.5, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif' }}>Discard</button>
          </div>
        </div>
      )}

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
              <option value="auto">Auto</option>
              <option value="groq/compound">Compound</option>
              <option value="llama-3.3-70b-versatile">Llama 70B</option>
              <option value="meta-llama/llama-4-scout-17b-16e-instruct">Llama 17B</option>
            </select>
            <div style={{ position: 'absolute', left: 6, top: '50%', transform: 'translateY(-50%)', width: 4, height: 4, borderRadius: '50%', background: '#217346' }} />
          </div>

          <textarea
            ref={inputRef}
            value={input}
            onChange={handleInputChange}
            onKeyDown={e => {
              if (e.key === 'Escape') setShowSlash(false);
              if (e.key === 'ArrowUp' && !input.trim()) {
                e.preventDefault();
                previewHistory(-1);
                return;
              }
              if (e.key === 'ArrowDown' && !input.trim()) {
                e.preventDefault();
                previewHistory(1);
                return;
              }
              if (e.key === 'Tab' && historyPreview) {
                e.preventDefault();
                setInput(historyPreview);
                setHistoryPreview('');
                setHistoryIndex(-1);
                return;
              }
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
        {historyPreview && (
          <div style={{ marginTop: 6, fontSize: 10.5, color: '#605e5c', background: '#f3f3f3', border: '0.5px solid #e0e0e0', borderRadius: 5, padding: '4px 8px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
            Prompt history: {historyPreview} · Press Tab to apply
          </div>
        )}
      </div>

      {auditReport && (
        <AuditDashboard 
          report={auditReport} 
          onClose={() => setAuditReport(null)} 
          onApplyFix={(finding) => {
             setAuditReport(null);
             if (finding.suggested_formula && finding.affected_cells) {
               setMessages(prev => [...prev, { 
                 role: 'assistant', 
                 content: `To apply the fix for ${finding.affected_cells.join(', ')}, copy this formula:\n\`${finding.suggested_formula}\`` 
               }]);
             }
          }} 
        />
      )}
    </div>
  );
});

ChatPane.displayName = 'ChatPane';
export default ChatPane;
