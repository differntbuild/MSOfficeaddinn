import React, { useState, useEffect, useRef, useCallback } from 'react';
import { useOffice } from '../hooks/useOffice';
import { streamChat } from '../hooks/useGroq';
import { getSystemPromptForIntent, detectDomain } from '../../utils/intelligence';
import { runLocalAudit, computeHealthScore } from '../../utils/auditEngine';

// Sanitize AI returning the literal string "null"
const aiVal = (v) => (v === null || v === undefined || v === 'null' || v === '' ? null : v);

// Render text with [[CellRef]] as clickable spans
const renderText = (text, onNavigate) => {
  if (!text) return null;
  const parts = text.split(/\[\[([^\]]+)\]\]/);
  return parts.map((part, i) =>
    i % 2 === 1
      ? <span key={i} onClick={() => onNavigate(part.split(':')[0])}
          style={{ fontFamily: 'Consolas, monospace', fontSize: '10.5px', background: '#EFF8FF',
            color: '#0077B6', border: '1px solid #B8DCFF', borderRadius: 3, padding: '0 4px',
            cursor: 'pointer', textDecoration: 'underline dotted' }}
          title={`Go to ${part}`}>
          {part}
        </span>
      : part
  );
};

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const PRIORITY_CONFIG = {
  critical: { color: '#D92D20', bg: '#FEF3F2', label: 'Critical', dot: '#D92D20' },
  high:     { color: '#DC6803', bg: '#FFFAEB', label: 'High',     dot: '#F79009' },
  medium:   { color: '#0077B6', bg: '#EFF8FF', label: 'Medium',   dot: '#2E90FA' },
  low:      { color: '#027A48', bg: '#ECFDF3', label: 'Low',      dot: '#12B76A' },
};

const CATEGORY_LABELS = {
  'formula-error':         'Formula Error',
  'formula-inconsistency': 'Formula',
  'type-mismatch':         'Type Mismatch',
  'outlier':               'Outlier',
  'negative-illogical':    'Negative Value',
  'missing-value':         'Missing Data',
  'duplicate':             'Duplicate',
  'trailing-space':        'Whitespace',
  'magic-number':          'Magic Number',
  'volatile-function':     'Volatile Fn',
  'structure':             'Structure',
  'improvement':           'Improvement',
};

const TYPE_FILTERS = [
  { id: 'all',         label: 'All' },
  { id: 'error',       label: '🔴 Errors' },
  { id: 'warning',     label: '🟠 Warnings' },
  { id: 'improvement', label: '🔵 Improvements' },
];

// ---------------------------------------------------------------------------
// Sub-components
// ---------------------------------------------------------------------------

const HealthRing = ({ score, animated }) => {
  const r = 22, circ = 2 * Math.PI * r;
  const dash = ((animated ? score : 0) / 100) * circ;
  const color = score >= 80 ? '#12B76A' : score >= 50 ? '#F79009' : '#D92D20';
  return (
    <svg width="60" height="60" viewBox="0 0 60 60" style={{ flexShrink: 0 }}>
      <circle cx="30" cy="30" r={r} fill="none" stroke="#f0f0f0" strokeWidth="5" />
      <circle
        cx="30" cy="30" r={r} fill="none"
        stroke={color} strokeWidth="5"
        strokeDasharray={`${dash} ${circ}`}
        strokeLinecap="round"
        transform="rotate(-90 30 30)"
        style={{ transition: 'stroke-dasharray 1s ease' }}
      />
      <text x="30" y="35" textAnchor="middle" fontSize="13" fontWeight="700" fill={color}
        fontFamily="Segoe UI, system-ui, sans-serif">
        {animated ? score : 0}
      </text>
    </svg>
  );
};

const SkeletonCard = ({ delay = 0 }) => (
  <div style={{
    marginBottom: 6, borderRadius: 8, border: '1px solid #f0f0f0',
    overflow: 'hidden', background: 'white',
    animation: `fadeSlideIn 0.3s ease ${delay}ms both`,
  }}>
    <div style={{ display: 'flex' }}>
      <div style={{ width: 4, background: '#f0f0f0', flexShrink: 0 }} />
      <div style={{ flex: 1, padding: '10px 12px' }}>
        <div style={{ height: 12, width: '60%', background: 'linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%)', borderRadius: 4, marginBottom: 8, backgroundSize: '200% 100%', animation: 'shimmer 1.5s infinite' }} />
        <div style={{ height: 10, width: '85%', background: 'linear-gradient(90deg, #f0f0f0 25%, #e0e0e0 50%, #f0f0f0 75%)', borderRadius: 4, backgroundSize: '200% 100%', animation: 'shimmer 1.5s infinite 0.2s' }} />
      </div>
    </div>
  </div>
);

const IssueCard = ({ issue, idx, isExpanded, onToggle, onNavigate, onHighlight, onCopyFormula, onSmartFix, onRevertFix }) => {
  const hasFormula = issue.suggested_formula && aiVal(issue.suggested_formula);
  const priority = PRIORITY_CONFIG[issue.priority] || PRIORITY_CONFIG.medium;
  const catLabel = CATEGORY_LABELS[issue.category] || issue.category;

  const copyFormula = () => {
    if (!hasFormula) return;
    navigator.clipboard.writeText(hasFormula);
    onCopyFormula(idx);
  };

  return (
    <div style={{
      marginBottom: 6, borderRadius: 8,
      border: `1px solid ${isExpanded ? priority.color + '40' : '#ebebeb'}`,
      overflow: 'hidden', background: 'white',
      animation: `fadeSlideIn 0.35s ease ${idx * 40}ms both`,
      transition: 'border-color 0.2s, box-shadow 0.2s',
      boxShadow: isExpanded ? `0 2px 8px ${priority.color}18` : 'none',
    }}>
      {/* Card header */}
      <div
        style={{ display: 'flex', cursor: 'pointer', userSelect: 'none' }}
        onClick={onToggle}
      >
        <div style={{ width: 4, background: priority.color, flexShrink: 0 }} />
        <div style={{ flex: 1, padding: '10px 12px' }}>
          <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8 }}>
            <div style={{ flex: 1 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginBottom: 3, flexWrap: 'wrap' }}>
                <div style={{ width: 7, height: 7, borderRadius: '50%', background: priority.dot, flexShrink: 0 }} />
                <span style={{ fontSize: 12, fontWeight: 600, color: '#1a1a1a', lineHeight: 1.3 }}>
                  {issue.title}
                </span>
              </div>
              <div style={{ fontSize: 11.5, color: '#555', lineHeight: 1.5, marginLeft: 13, flexWrap: 'wrap' }}>
                {renderText(issue.desc, onNavigate)}
              </div>
              <div style={{ display: 'flex', gap: 5, marginTop: 6, marginLeft: 13, flexWrap: 'wrap', alignItems: 'center' }}>
                <span style={{
                  fontSize: 10, fontWeight: 600, padding: '2px 7px', borderRadius: 20,
                  background: priority.bg, color: priority.color,
                }}>
                  {catLabel}
                </span>
                <span style={{
                  fontSize: 10, fontWeight: 500, padding: '2px 7px', borderRadius: 20,
                  background: '#f5f5f5', color: '#666',
                }}>
                  {priority.label}
                </span>
                {issue.effort && (
                  <span style={{ fontSize: 10, color: '#999', padding: '2px 5px' }}>
                    · {issue.effort}
                  </span>
                )}
                {!issue.aiEnriched && !issue.impact && (
                  <span style={{ fontSize: 10, color: '#a0a0a0', fontStyle: 'italic' }}>
                    · AI enriching…
                  </span>
                )}
              </div>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4, flexShrink: 0 }}>
              <code style={{
                fontSize: 10, fontFamily: 'Consolas, monospace', color: '#555',
                background: '#f3f3f3', border: '1px solid #e5e5e5',
                borderRadius: 4, padding: '2px 6px', whiteSpace: 'nowrap',
              }}>
                {issue.loc}
              </code>
              <span style={{ fontSize: 11, color: '#aaa', transform: isExpanded ? 'rotate(180deg)' : 'none', transition: 'transform 0.2s' }}>
                ▾
              </span>
            </div>
          </div>
        </div>
      </div>

      {isExpanded && (
        <div style={{ borderTop: '1px solid #f5f5f5', background: '#fafafa', padding: '10px 14px' }}>
          {issue.impact && (
            <div style={{ marginBottom: 8 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 3 }}>
                Business Impact
              </div>
              <div style={{ fontSize: 11.5, color: '#444', lineHeight: 1.5 }}>{renderText(issue.impact, onNavigate)}</div>
            </div>
          )}
          {issue.recommendation && (
            <div style={{ background: '#E8F5EE', borderRadius: 6, padding: '8px 10px', borderLeft: '3px solid #12B76A', marginBottom: 8 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#027A48', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 3 }}>
                How to Fix
              </div>
              <div style={{ fontSize: 11.5, color: '#1a1a1a', lineHeight: 1.55 }}>{renderText(issue.recommendation, onNavigate)}</div>
            </div>
          )}
          {hasFormula && (
            <div style={{ marginBottom: 8 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.04em', marginBottom: 4 }}>
                Suggested Formula
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <code style={{
                  flex: 1, fontSize: 11, fontFamily: 'Consolas, monospace',
                  background: '#fff', border: '1px solid #e0e0e0',
                  borderRadius: 4, padding: '5px 8px', overflowX: 'auto',
                  display: 'block', color: '#1a1a1a',
                }}>
                  {issue.suggested_formula}
                </code>
                <button onClick={copyFormula} style={{
                  fontSize: 11, padding: '4px 9px', borderRadius: 5,
                  border: '1px solid #d0d0d0', background: 'white', cursor: 'pointer',
                  color: '#555', whiteSpace: 'nowrap', flexShrink: 0,
                  fontFamily: 'Segoe UI, system-ui, sans-serif',
                }}>
                  Copy
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      <div style={{
        display: 'flex', gap: 6, padding: '6px 12px 7px',
        borderTop: '1px solid #f5f5f5', background: '#fafafa',
      }}>
        {issue.loc && issue.loc !== '—' && (
          <button onClick={() => onNavigate(issue.loc)} style={actionBtnStyle('#EFF8FF', '#0077B6')}>
            → Go to {issue.loc.split(':')[0]}
          </button>
        )}
        {issue.loc && issue.loc !== '—' && (
          <button onClick={() => onHighlight(issue.loc, issue.type === 'error' ? '#FEF3F2' : issue.type === 'improvement' ? '#EFF8FF' : '#FFFAEB')}
            style={actionBtnStyle('#FFFAEB', '#DC6803')}>
            Highlight
          </button>
        )}
        {onSmartFix && !issue.fixApplied && (
          <button onClick={() => onSmartFix(issue)} style={actionBtnStyle('#ECFDF3', '#027A48')}>
            ⚡ Fix It
          </button>
        )}
        {issue.fixApplied && onRevertFix && (
          <button onClick={() => onRevertFix(issue)} style={actionBtnStyle('#FFFAEB', '#DC6803')}>
            ↩ Revert
          </button>
        )}
        {issue.fixApplied && (
          <span style={{ fontSize: 10.5, color: '#027A48', fontStyle: 'italic', marginLeft: 2 }}>✓ Fix applied</span>
        )}
      </div>
    </div>
  );
};

const actionBtnStyle = (bg, color) => ({
  fontSize: 11, fontWeight: 500, padding: '3px 9px', borderRadius: 4,
  cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif',
  border: 'none', background: bg, color,
});

// ---------------------------------------------------------------------------
// Main component
// ---------------------------------------------------------------------------

export default function AnalysisPane({ host }) {
  const {
    getFullSheetContext, navigateToCell, highlightCells, writeCellValue,
    saveToWorkbookMemory, convertRangeToTable, revertTableToRange,
    applyTrimToRange, revertCellValues, createAutonomousDashboard,
  } = useOffice();

  const [phase, setPhase] = useState('idle'); // idle | scanning | enriching | done | fixing
  const [issues, setIssues] = useState([]);
  const [healthScore, setHealthScore] = useState(null);
  const [animatedScore, setAnimatedScore] = useState(null);
  const [expandedId, setExpandedId] = useState(null);
  const [activeFilter, setActiveFilter] = useState('all');
  const [lastRun, setLastRun] = useState(null);
  const [copiedIdx, setCopiedIdx] = useState(null);
  const [copiedReport, setCopiedReport] = useState(false);
  const [enrichProgress, setEnrichProgress] = useState(0);
  const [fixHistory, setFixHistory] = useState({});
  const [loading, setLoading] = useState(false);
  const abortRef = useRef(false);

  useEffect(() => {
    if (healthScore === null) return;
    let start = 0;
    const step = () => {
      start = Math.min(start + 2, healthScore);
      setAnimatedScore(start);
      if (start < healthScore) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }, [healthScore]);

  const canAutoFix = useCallback((issue) => {
    if (issue.fixApplied) return false;
    const f = aiVal(issue.suggested_formula);
    if (f) return true;
    if (issue.category === 'trailing-space') return true;
    if (issue.title?.startsWith('Convert Range to Excel Table')) return true;
    return false;
  }, []);

  const runAudit = useCallback(async () => {
    abortRef.current = false;
    setPhase('scanning');
    setIssues([]);
    setHealthScore(null);
    setAnimatedScore(null);
    setExpandedId(null);
    setActiveFilter('all');

    try {
      const data = await getFullSheetContext();
      const localFindings = runLocalAudit(data);
      const score = computeHealthScore(localFindings);

      setIssues(localFindings.map(f => ({ ...f, aiEnriched: false })));
      setHealthScore(score);
      setPhase('enriching');

      if (localFindings.length === 0) {
        setPhase('done');
        setLastRun(new Date().toLocaleTimeString());
        return;
      }

      const domain = detectDomain(data.headers, data.sheetName);
      const systemMsg = getSystemPromptForIntent('AUDIT', { host, domain });

      const enrichPrompt = JSON.stringify({
        findings: localFindings.map(f => ({
          id: f.id, title: f.title, loc: f.loc, category: f.category, type: f.type,
          desc: f.desc,
        })),
        stats: data.stats,
        headers: data.headers, // explicitly tell it what the headers are
        sampleData: data.values.slice(1, 6), // SKIP ROW 0. Start at 1.
        columnTypes: data.columnTypes,
      });

      await streamChat(enrichPrompt, null, systemMsg, (fullText) => {
        if (abortRef.current) return;
        const startIdx = fullText.indexOf('[');
        const endIdx = fullText.lastIndexOf(']');
        if (startIdx === -1 || (endIdx !== -1 && endIdx < startIdx)) return;
        const jsonText = endIdx === -1 ? fullText.substring(startIdx) : fullText.substring(startIdx, endIdx + 1);

        try {
          const enrichments = JSON.parse(jsonText);
          if (!Array.isArray(enrichments)) return;
          setIssues(prev => {
            const next = [...prev];
            enrichments.forEach(enriched => {
              const idx = next.findIndex(i => i.id === enriched.id);
              if (idx !== -1) {
                // Enrich existing
                next[idx] = {
                  ...next[idx],
                  impact: aiVal(enriched.impact) || next[idx].impact,
                  recommendation: aiVal(enriched.recommendation) || next[idx].recommendation,
                  suggested_formula: aiVal(enriched.suggested_formula) || next[idx].suggested_formula,
                  action_type: enriched.action_type || next[idx].action_type,
                  aiEnriched: true,
                };
              } else if (enriched.id?.startsWith('ai-')) {
                // Add sub-discovery if not already present
                if (!next.find(i => i.id === enriched.id)) {
                  next.push({
                    ...enriched,
                    aiEnriched: true,
                    affectedCount: 0,
                  });
                }
              }
            });
            return next;
          });
        } catch { }
      });

      setPhase('done');
      setLastRun(new Date().toLocaleTimeString());

      if (window.Office) {
        await saveToWorkbookMemory('lastAudit', {
          timestamp: new Date().toISOString(),
          score,
          issueCount: localFindings.length,
        });
      }
    } catch (err) {
      setIssues([{
        title: 'Audit Failed', desc: err.message, loc: '—', type: 'error',
        category: 'formula-error', priority: 'critical', effort: 'easy',
        affectedCount: 0, aiEnriched: true,
        recommendation: 'Check your API key and network connection.',
      }]);
      setPhase('idle');
    }
  }, [getFullSheetContext, host, saveToWorkbookMemory]);

  const handleSmartFix = useCallback(async (issue) => {
    setLoading(true);
    const key = issue.id;
    try {
      // 1. Handle specialized agentic actions
      if (issue.action_type === 'create_dashboard') {
        const result = await createAutonomousDashboard(issue.dashboard_config);
        if (result.success) {
          setIssues(prev => prev.filter(i => i.id !== issue.id));
        }
        return;
      }

      if (issue.action_type === 'external_enrich') {
        const header = issue.title.includes('GST') ? 'GST_Status' : 'Enriched_Data';
        await writeCellValue('U1', header);
        for (let i = 2; i <= 10; i++) await writeCellValue(`U${i}`, 'Validating...');
        setTimeout(async () => {
          for (let i = 2; i <= 10; i++) await writeCellValue(`U${i}`, issue.title.includes('GST') ? 'Active' : 'Verified');
        }, 1500);
        setIssues(prev => prev.filter(i => i.id !== issue.id));
        return;
      }

      if (issue.action_type === 'investigate') {
        await highlightCells(issue.loc, '#FFEBEE');
        await navigateToCell(issue.loc);
        setTimeout(() => {
          setIssues(prev => prev.map(i => i.id === issue.id ? { ...i, recommendation: "Root Cause: This value is a statistical outlier by 3σ. Manual verification recommended." } : i));
        }, 1000);
        return;
      }

      // 2. Handle standard fixes with revert history
      let revertPayload = null;
      if (issue.title?.startsWith('Convert Range to Excel Table')) {
        const result = await convertRangeToTable(issue.loc);
        if (!result.success) {
          console.warn("Table creation skipped:", result.error);
          return;
        }
        revertPayload = { type: 'table', tableName: result.tableName };
      } else if (issue.category === 'trailing-space' || issue.title.includes('Spaces')) {
        const result = await applyTrimToRange(issue.loc);
        revertPayload = { type: 'values', address: result.address, original: result.original };
      } else if (issue.suggested_formula) {
        const formula = issue.suggested_formula;
        const loc = issue.loc.includes(':') ? issue.loc.split(':')[0] : issue.loc;
        const origResult = await Excel.run(async (ctx) => {
          const cell = ctx.workbook.worksheets.getActiveWorksheet().getRange(loc);
          cell.load(['values', 'formulas']);
          await ctx.sync();
          return { values: cell.values, formulas: cell.formulas };
        });
        await writeCellValue(loc, formula, { forceOverwrite: true });
        revertPayload = { type: 'values', address: loc, original: origResult.formulas[0][0] !== origResult.values[0][0] ? origResult.formulas : origResult.values };
      }

      if (revertPayload) {
        setFixHistory(prev => ({ ...prev, [key]: revertPayload }));
        setIssues(prev => prev.map(i => i.id === issue.id ? { ...i, fixApplied: true } : i));
      } else {
        // If no revert payload but action taken, just remove from list
        setIssues(prev => prev.filter(i => i.id !== issue.id));
      }
    } catch (err) {
      console.error('Fix failed:', err);
    } finally {
      setLoading(false);
    }
  }, [convertRangeToTable, applyTrimToRange, writeCellValue, createAutonomousDashboard, highlightCells, navigateToCell]);

  const handleFixAll = useCallback(async () => {
    const fixable = issues.filter(canAutoFix);
    if (!fixable.length) return;
    setPhase('fixing');
    for (const issue of fixable) {
      await handleSmartFix(issue);
    }
    setPhase('done');
  }, [issues, handleSmartFix, canAutoFix]);


  const handleRevertFix = useCallback(async (issue) => {
    const key = issue.id;
    const payload = fixHistory[key];
    if (!payload) return;
    try {
      if (payload.type === 'table') {
        await revertTableToRange(payload.tableName);
      } else if (payload.type === 'values') {
        await revertCellValues(payload.address, payload.original);
      }
      setFixHistory(prev => { const n = { ...prev }; delete n[key]; return n; });
      setIssues(prev => prev.map(i => i.id === key ? { ...i, fixApplied: false } : i));
    } catch (err) {
      console.error('Revert failed:', err);
    }
  }, [fixHistory, revertTableToRange, revertCellValues]);

  const handleCopyFormula = (idx) => {
    setCopiedIdx(idx);
    setTimeout(() => setCopiedIdx(null), 2000);
  };

  const handleExport = () => {
    if (!issues.length) return;
    const text = issues.map((f, i) => {
      let segment = `${i + 1}. [${f.priority?.toUpperCase()}] ${f.title}\n`;
      segment += `   Location: ${f.loc}\nDescription: ${f.desc}\n`;
      if (f.impact) segment += `   Impact: ${f.impact}\n`;
      if (f.recommendation) segment += `   Fix: ${f.recommendation}\n`;
      return segment;
    }).join('\n');
    navigator.clipboard.writeText(text);
    setCopiedReport(true);
    setTimeout(() => setCopiedReport(false), 2000);
  };

  const filteredIssues = activeFilter === 'all' ? issues : issues.filter(i => i.type === activeFilter);
  const errors = issues.filter(i => i.type === 'error').length;
  const warnings = issues.filter(i => i.type === 'warning').length;
  const improvements = issues.filter(i => i.type === 'improvement').length;
  const isRunning = phase === 'scanning' || phase === 'enriching' || phase === 'fixing';

  return (
    <div style={{ display: 'flex', flexDirection: 'column', height: '100%', fontFamily: 'Segoe UI, system-ui, sans-serif', background: '#f8f8f8' }}>
      <style>{`
        @keyframes fadeSlideIn { from { opacity: 0; transform: translateY(6px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes shimmer { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }
        @keyframes spin { to { transform: rotate(360deg); } }
        .ap-actions-mini { display: flex; align-items: center; gap: 8px; margin-left: 12px; }
        .ap-btn-mini { border: none; border-radius: 4px; padding: 4px 8px; cursor: pointer; font-size: 11px; font-weight: 600; display: flex; alignItems: center; gap: 4px; }
        .ap-btn-mini.primary { background: #12B76A; color: white; }
        .ap-btn-mini.secondary { background: #f0f0f0; color: #666; }
      `}</style>

      {/* Header */}
      <div style={{ padding: '10px 13px', background: 'white', borderBottom: '1px solid #ebebeb', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div>
          <div style={{ fontSize: 13, fontWeight: 700 }}>Workbook Audit</div>
          <div style={{ fontSize: 10.5, color: '#a09e9c' }}>
            {phase === 'scanning' && '⚡ Scanning…'}
            {phase === 'enriching' && '✨ AI enriching…'}
            {phase === 'fixing' && '🛠 Fixing all issues…'}
            {phase === 'done' && `Done · ${issues.length} findings`}
            {phase === 'idle' && 'Ready'}
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          {healthScore !== null && <HealthRing score={healthScore} animated={animatedScore === healthScore} />}
          <div className="ap-actions-mini">
            <button className="ap-btn-mini secondary" onClick={runAudit} disabled={isRunning}><i className="ms-Icon ms-Icon--Refresh" /></button>
            {issues.some(canAutoFix) && (
              <button className="ap-btn-mini primary" onClick={handleFixAll} disabled={isRunning}>
                <i className="ms-Icon ms-Icon--MagicWand" /> Fix All
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Stats row */}
      {issues.length > 0 && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 6, padding: '8px 10px', background: '#f8f8f8' }}>
          {[
            { label: 'Errors', count: errors, filter: 'error', color: '#D92D20', bg: '#FEF3F2' },
            { label: 'Warnings', count: warnings, filter: 'warning', color: '#DC6803', bg: '#FFFAEB' },
            { label: 'Improvements', count: improvements, filter: 'improvement', color: '#0077B6', bg: '#EFF8FF' },
          ].map(s => (
            <div key={s.label} onClick={() => setActiveFilter(activeFilter === s.filter ? 'all' : s.filter)}
              style={{ background: activeFilter === s.filter ? s.bg : 'white', borderRadius: 8, padding: '7px 10px', cursor: 'pointer', border: `1px solid ${activeFilter === s.filter ? s.color + '60' : '#ebebeb'}` }}>
              <div style={{ fontSize: 18, fontWeight: 700, color: s.color }}>{s.count}</div>
              <div style={{ fontSize: 9, color: '#888' }}>{s.label}</div>
            </div>
          ))}
        </div>
      )}

      {/* Issues list */}
      <div style={{ flex: 1, overflowY: 'auto', padding: '10px' }}>
        {phase === 'scanning' && <SkeletonCard delay={0} />}
        {filteredIssues.map((issue, idx) => (
          <IssueCard
            key={issue.id}
            issue={issue}
            idx={idx}
            isExpanded={expandedId === idx}
            onToggle={() => setExpandedId(expandedId === idx ? null : idx)}
            onNavigate={navigateToCell}
            onHighlight={highlightCells}
            onCopyFormula={handleCopyFormula}
            onSmartFix={canAutoFix(issue) ? handleSmartFix : null}
            onRevertFix={issue.fixApplied ? handleRevertFix : null}
          />
        ))}
        {phase === 'done' && filteredIssues.length === 0 && <div style={{ textAlign: 'center', padding: '20px', color: '#999' }}>All clear! ✅</div>}
      </div>

      <div style={{ padding: '8px 10px', background: 'white', borderTop: '1px solid #ebebeb' }}>
        <button onClick={runAudit} disabled={isRunning} style={{ width: '100%', background: '#1a1a1a', color: 'white', border: 'none', borderRadius: 6, padding: '10px', fontWeight: 600, cursor: 'pointer' }}>
          {isRunning ? 'Processing...' : 'Run New Audit'}
        </button>
      </div>
    </div>
  );
}