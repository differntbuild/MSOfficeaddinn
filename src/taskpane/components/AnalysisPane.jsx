import React, { useState, useEffect, useRef, useCallback, useMemo } from 'react';
import { useOffice } from '../hooks/useOffice';
import { usePython } from '../hooks/usePython';
import { streamChat } from '../hooks/useGroq';
import { getSystemPromptForIntent, detectDomain } from '../../utils/intelligence';
import { runLocalAudit, computeHealthScore, generateAuditNarrative } from '../../utils/auditEngine';
import { ActionExecutor } from '../../agent/engine/ActionExecutor';
import { AuditEngine } from '../../agent/engine/AuditEngine';

/* ─── Sanitize AI nulls ─── */
const aiVal = (v) => (v === null || v === undefined || v === 'null' || v === '' ? null : v);

/* ─── Render [[CellRef]] citations as clickable chips ─── */
const renderText = (text, onNavigate) => {
  if (!text) return null;
  const parts = text.split(/\[\[([^\]]+)\]\]/);
  return parts.map((part, i) =>
    i % 2 === 1 ? (
      <button
        key={i}
        onClick={() => onNavigate(part.split(':')[0])}
        style={{
          display: 'inline-flex', alignItems: 'center', gap: 2,
          fontFamily: "'Cascadia Code', 'Fira Code', monospace", fontSize: 10.5,
          background: 'rgba(59, 139, 212, 0.1)', color: '#2563EB',
          border: '0.5px solid rgba(59, 139, 212, 0.3)',
          borderRadius: 4, padding: '0 5px', height: 18,
          cursor: 'pointer', verticalAlign: 'middle', margin: '0 1px',
          fontWeight: 500,
        }}
        title={`Navigate to ${part}`}
      >
        ↗ {part}
      </button>
    ) : (
      <span key={i}>{part}</span>
    )
  );
};

/* ─── Priority/type config ─── */
const PRIORITY = {
  critical: { label: 'Critical', color: '#EF4444', bg: 'rgba(239,68,68,0.08)', accent: '#FCA5A5', dot: '#EF4444', score: 15 },
  high:     { label: 'High',     color: '#F59E0B', bg: 'rgba(245,158,11,0.08)', accent: '#FCD34D', dot: '#F59E0B', score: 8 },
  medium:   { label: 'Medium',   color: '#3B82F6', bg: 'rgba(59,130,246,0.08)', accent: '#93C5FD', dot: '#3B82F6', score: 3 },
  low:      { label: 'Low',      color: '#10B981', bg: 'rgba(16,185,129,0.08)', accent: '#6EE7B7', dot: '#10B981', score: 1 },
};

const CAT_LABELS = {
  'formula-error':         { label: 'Formula Error',    icon: '⊘' },
  'formula-inconsistency': { label: 'Inconsistent',     icon: '≈' },
  'type-mismatch':         { label: 'Type Mismatch',    icon: '⇄' },
  'outlier':               { label: 'Statistical',      icon: '◈' },
  'negative-illogical':    { label: 'Negative Value',   icon: '⊖' },
  'missing-value':         { label: 'Missing Data',     icon: '○' },
  'duplicate':             { label: 'Duplicate',        icon: '⊜' },
  'trailing-space':        { label: 'Whitespace',       icon: '∵' },
  'magic-number':          { label: 'Magic Number',     icon: '⊛' },
  'volatile-function':     { label: 'Volatile',         icon: '⚡' },
  'structure':             { label: 'Structure',        icon: '⊞' },
  'improvement':           { label: 'Improvement',      icon: '◆' },
  'discovery':             { label: 'AI Discovery',     icon: '✦' },
  'enrichment':            { label: 'Enrichment',       icon: '◉' },
  'dashboard':             { label: 'Dashboard',        icon: '▦' },
};

const TYPE_FILTERS = [
  { id: 'all',         label: 'All' },
  { id: 'error',       label: 'Errors' },
  { id: 'warning',     label: 'Warnings' },
  { id: 'improvement', label: 'Insights' },
];

/* ─── Animated Score Ring ─── */
const ScoreRing = ({ score, animated }) => {
  const r = 28, circ = 2 * Math.PI * r;
  const pct = (animated ? score : 0) / 100;
  const dash = pct * circ;
  const color = score >= 80 ? '#10B981' : score >= 50 ? '#F59E0B' : '#EF4444';
  const grade = score >= 90 ? 'A' : score >= 75 ? 'B' : score >= 60 ? 'C' : score >= 40 ? 'D' : 'F';
  return (
    <div style={{ position: 'relative', width: 72, height: 72 }}>
      <svg width="72" height="72" viewBox="0 0 72 72">
        <circle cx="36" cy="36" r={r} fill="none" stroke="rgba(0,0,0,0.06)" strokeWidth="5"/>
        <circle
          cx="36" cy="36" r={r} fill="none"
          stroke={color} strokeWidth="5"
          strokeDasharray={`${dash} ${circ}`}
          strokeLinecap="round"
          transform="rotate(-90 36 36)"
          style={{ transition: 'stroke-dasharray 1.2s cubic-bezier(0.4,0,0.2,1)' }}
        />
      </svg>
      <div style={{
        position: 'absolute', inset: 0,
        display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
      }}>
        <span style={{ fontSize: 16, fontWeight: 700, color, lineHeight: 1 }}>{animated ? score : 0}</span>
        <span style={{ fontSize: 9, color, fontWeight: 600, opacity: 0.7 }}>{grade}</span>
      </div>
    </div>
  );
};

/* ─── Skeleton Loader ─── */
const SkeletonCard = ({ delay = 0 }) => (
  <div style={{
    borderRadius: 10, border: '0.5px solid rgba(0,0,0,0.07)',
    background: '#fff', padding: '14px 16px', marginBottom: 8,
    animation: `fadeUp 0.4s ease ${delay}ms both`,
  }}>
    <div style={{ display: 'flex', gap: 12, alignItems: 'flex-start' }}>
      <div style={{ width: 8, height: 8, borderRadius: '50%', background: 'rgba(0,0,0,0.1)', marginTop: 4, flexShrink: 0 }}/>
      <div style={{ flex: 1 }}>
        <div style={{ height: 13, width: '65%', background: 'linear-gradient(90deg,#f0f0f0 25%,#e0e0e0 50%,#f0f0f0 75%)', borderRadius: 4, marginBottom: 8, backgroundSize: '200% 100%', animation: 'shimmer 1.5s infinite' }}/>
        <div style={{ height: 11, width: '90%', background: 'linear-gradient(90deg,#f0f0f0 25%,#e0e0e0 50%,#f0f0f0 75%)', borderRadius: 4, backgroundSize: '200% 100%', animation: 'shimmer 1.5s infinite 0.2s' }}/>
        <div style={{ height: 11, width: '70%', background: 'linear-gradient(90deg,#f0f0f0 25%,#e0e0e0 50%,#f0f0f0 75%)', borderRadius: 4, marginTop: 6, backgroundSize: '200% 100%', animation: 'shimmer 1.5s infinite 0.35s' }}/>
      </div>
    </div>
  </div>
);

/* ─── AI Typing Indicator ─── */
const TypingDots = () => (
  <span style={{ display: 'inline-flex', gap: 3, alignItems: 'center', marginLeft: 4 }}>
    {[0,1,2].map(i => (
      <span key={i} style={{
        width: 4, height: 4, borderRadius: '50%', background: 'currentColor', opacity: 0.4,
        animation: `dotBounce 1.2s ${i*0.2}s infinite ease-in-out`,
      }}/>
    ))}
  </span>
);

/* ─── Stat Mini Card ─── */
const StatCard = ({ count, label, color, bg, active, onClick }) => (
  <button onClick={onClick} style={{
    background: active ? bg : 'rgba(255,255,255,0.7)',
    border: `0.5px solid ${active ? color + '40' : 'rgba(0,0,0,0.07)'}`,
    borderRadius: 10, padding: '10px 12px',
    cursor: 'pointer', textAlign: 'left', width: '100%',
    transition: 'all 0.15s ease',
    transform: active ? 'translateY(-1px)' : 'none',
    boxShadow: active ? `0 2px 8px ${color}20` : 'none',
  }}>
    <div style={{ fontSize: 22, fontWeight: 700, color, lineHeight: 1 }}>{count}</div>
    <div style={{ fontSize: 10, color: active ? color : '#888', marginTop: 2, fontWeight: 500 }}>{label}</div>
  </button>
);

/* ─── Action Tag ─── */
const ActionTag = ({ label, type }) => {
  const colors = {
    fix: { c: '#10B981', bg: 'rgba(16,185,129,0.1)' },
    create_dashboard: { c: '#8B5CF6', bg: 'rgba(139,92,246,0.1)' },
    external_enrich: { c: '#F59E0B', bg: 'rgba(245,158,11,0.1)' },
    investigate: { c: '#EF4444', bg: 'rgba(239,68,68,0.1)' },
    None: { c: '#94A3B8', bg: 'rgba(148,163,184,0.1)' },
  };
  const c = colors[type] || colors.None;
  return (
    <span style={{
      fontSize: 10, fontWeight: 600, padding: '2px 7px', borderRadius: 20,
      color: c.c, background: c.bg, letterSpacing: '0.02em',
    }}>{label}</span>
  );
};

/* ─── Issue Card ─── */
const IssueCard = ({
  issue, idx, expanded, onToggle, onNavigate, onHighlight,
  onSmartFix, onRevertFix, canFix,
}) => {
  const p = PRIORITY[issue.priority] || PRIORITY.medium;
  const cat = CAT_LABELS[issue.category] || { label: issue.category, icon: '·' };
  const hasFormula = aiVal(issue.suggested_formula);
  const [justCopied, setJustCopied] = useState(false);

  const copyFormula = () => {
    if (!hasFormula) return;
    navigator.clipboard.writeText(hasFormula);
    setJustCopied(true);
    setTimeout(() => setJustCopied(false), 2000);
  };

  const actionLabel = {
    fix: '⚡ Auto Fix',
    create_dashboard: '▦ Build Dashboard',
    external_enrich: '◉ Enrich Data',
    investigate: '◈ Investigate',
  }[issue.action_type] || null;

  return (
    <div style={{
      borderRadius: 10,
      border: `0.5px solid ${expanded ? p.color + '30' : 'rgba(0,0,0,0.06)'}`,
      background: '#fff',
      marginBottom: 7,
      overflow: 'hidden',
      animation: `fadeUp 0.35s ease ${idx * 35}ms both`,
      boxShadow: expanded ? `0 4px 16px ${p.color}12` : '0 1px 3px rgba(0,0,0,0.04)',
      transition: 'border-color 0.2s, box-shadow 0.2s',
    }}>
      {/* Priority accent bar */}
      <div style={{ height: 2, background: `linear-gradient(90deg, ${p.color}, transparent)` }}/>

      {/* Card header */}
      <div
        style={{ padding: '12px 14px 10px', cursor: 'pointer', userSelect: 'none' }}
        onClick={onToggle}
      >
        <div style={{ display: 'flex', gap: 10, alignItems: 'flex-start' }}>
          {/* Priority dot + icon */}
          <div style={{ marginTop: 2, flexShrink: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 4 }}>
            <div style={{ width: 8, height: 8, borderRadius: '50%', background: p.dot }}/>
            <span style={{ fontSize: 11, opacity: 0.5 }}>{cat.icon}</span>
          </div>

          {/* Content */}
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: 8 }}>
              <p style={{ margin: 0, fontSize: 12.5, fontWeight: 600, color: '#1a1a1a', lineHeight: 1.4, flex: 1 }}>
                {issue.title}
              </p>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0 }}>
                {issue.loc && issue.loc !== '—' && (
                  <code style={{
                    fontSize: 10, fontFamily: "'Cascadia Code', monospace",
                    background: 'rgba(0,0,0,0.05)', borderRadius: 4,
                    padding: '2px 6px', color: '#555',
                  }}>{issue.loc}</code>
                )}
                <span style={{
                  fontSize: 11, color: '#aaa',
                  transform: expanded ? 'rotate(180deg)' : 'none',
                  transition: 'transform 0.2s', lineHeight: 1,
                }}>▾</span>
              </div>
            </div>

            {/* Description — always visible, gives immediate context */}
            <p style={{ margin: '5px 0 0', fontSize: 11.5, color: '#555', lineHeight: 1.6 }}>
              {renderText(issue.desc, onNavigate)}
            </p>

            {/* Tags row */}
            <div style={{ display: 'flex', gap: 5, marginTop: 8, flexWrap: 'wrap', alignItems: 'center' }}>
              <span style={{
                fontSize: 10, fontWeight: 600, padding: '2px 7px', borderRadius: 20,
                background: p.bg, color: p.color, letterSpacing: '0.02em',
              }}>{p.label}</span>
              <span style={{
                fontSize: 10, fontWeight: 500, padding: '2px 7px', borderRadius: 20,
                background: 'rgba(0,0,0,0.04)', color: '#777',
              }}>{cat.label}</span>
              {issue.affectedCount > 0 && (
                <span style={{ fontSize: 10, color: '#bbb' }}>· {issue.affectedCount.toLocaleString()} affected</span>
              )}
              {issue.effort && (
                <span style={{ fontSize: 10, color: '#bbb' }}>· {issue.effort} fix</span>
              )}
              {!issue.aiEnriched && (
                <span style={{ fontSize: 10, color: '#8B5CF6', fontStyle: 'italic', display: 'flex', alignItems: 'center' }}>
                  · AI analyzing<TypingDots/>
                </span>
              )}
              {issue.fixApplied && (
                <span style={{ fontSize: 10, color: '#10B981', fontWeight: 600 }}>✓ Applied</span>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Expanded detail */}
      {expanded && (
        <div style={{
          borderTop: '0.5px solid rgba(0,0,0,0.06)',
          background: '#f8f8f8',
          padding: '12px 14px 14px',
          animation: 'expandIn 0.2s ease',
        }}>
          {issue.impact && (
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 9.5, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 5 }}>
                Business Impact
              </div>
              <p style={{ margin: 0, fontSize: 12, color: '#444', lineHeight: 1.6 }}>
                {renderText(issue.impact, onNavigate)}
              </p>
            </div>
          )}

          {issue.recommendation && (
            <div style={{
              background: 'rgba(16,185,129,0.06)', borderRadius: 8,
              padding: '10px 12px', borderLeft: '2px solid #10B981', marginBottom: 12,
            }}>
              <div style={{ fontSize: 9.5, fontWeight: 700, color: '#059669', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 4 }}>
                Recommendation
              </div>
              <p style={{ margin: 0, fontSize: 12, color: '#1a1a1a', lineHeight: 1.6 }}>
                {renderText(issue.recommendation, onNavigate)}
              </p>
            </div>
          )}

          {hasFormula && (
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 9.5, fontWeight: 700, color: '#888', textTransform: 'uppercase', letterSpacing: '0.08em', marginBottom: 6 }}>
                Suggested Formula
              </div>
              <div style={{ display: 'flex', gap: 6, alignItems: 'stretch' }}>
                <code style={{
                  flex: 1, fontSize: 11, fontFamily: "'Cascadia Code', 'Fira Code', monospace",
                  background: '#e8f5ee', color: '#217346', borderRadius: 4,
                  padding: '8px 10px', display: 'block', overflowX: 'auto',
                  whiteSpace: 'nowrap', lineHeight: 1.5,
                  border: '0.5px solid #c3e6d0'
                }}>
                  {hasFormula}
                </code>
                <button onClick={copyFormula} style={{
                  padding: '0 12px', borderRadius: 7, border: '0.5px solid rgba(0,0,0,0.1)',
                  background: justCopied ? '#10B981' : '#fff', color: justCopied ? '#fff' : '#555',
                  cursor: 'pointer', fontSize: 11, fontWeight: 600, whiteSpace: 'nowrap',
                  transition: 'all 0.2s', flexShrink: 0,
                }}>
                  {justCopied ? '✓' : 'Copy'}
                </button>
              </div>
            </div>
          )}

          {/* Action buttons */}
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {issue.loc && issue.loc !== '—' && (
              <button onClick={() => onNavigate(issue.loc)} style={btnStyle('#EFF6FF', '#2563EB')}>
                ↗ Go to {issue.loc.split(':')[0]}
              </button>
            )}
            {issue.loc && issue.loc !== '—' && (
              <button onClick={() => onHighlight(issue.loc, issue.type === 'error' ? '#FEF2F2' : issue.type === 'improvement' ? '#EFF6FF' : '#FFFBEB')} style={btnStyle('#FFFBEB', '#D97706')}>
                ◈ Highlight
              </button>
            )}
            {canFix && !issue.fixApplied && (
              <button onClick={() => onSmartFix(issue)} style={btnStyle('#ECFDF5', '#059669', true)}>
                {actionLabel || '⚡ Fix'}
              </button>
            )}
            {issue.fixApplied && onRevertFix && (
              <button onClick={() => onRevertFix(issue)} style={btnStyle('#FFFBEB', '#D97706')}>
                ↩ Undo
              </button>
            )}
          </div>
        </div>
      )}
    </div>
  );
};

const btnStyle = (bg, color, primary = false) => ({
  fontSize: 11, fontWeight: 600, padding: '5px 10px', borderRadius: 6,
  cursor: 'pointer', background: bg, color,
  border: primary ? `1px solid ${color}30` : '0.5px solid rgba(0,0,0,0.08)',
  letterSpacing: '0.01em', transition: 'all 0.15s',
});

/* ─── Smart Search / Filter Bar ─── */
const FilterBar = ({ activeFilter, setActiveFilter, search, setSearch, issues }) => (
  <div style={{ padding: '6px 12px', borderBottom: '0.5px solid #e0e0e0', background: '#f8f8f8' }}>
    <div style={{ display: 'flex', gap: 4, marginBottom: 0 }}>
      {TYPE_FILTERS.map(f => {
        const count = f.id === 'all' ? issues.length : issues.filter(i => i.type === f.id).length;
        return (
          <button key={f.id} onClick={() => setActiveFilter(f.id)} style={{
            flex: 1, padding: '4px 2px', borderRadius: 4, fontSize: 11, fontWeight: 600,
            cursor: 'pointer', transition: 'all 0.1s',
            background: activeFilter === f.id ? '#e8f5ee' : 'transparent',
            color: activeFilter === f.id ? '#217346' : '#605e5c',
            border: activeFilter === f.id ? '0.5px solid #c3e6d0' : '0.5px solid transparent',
          }}>
            {f.label} {count > 0 && `(${count})`}
          </button>
        );
      })}
    </div>
  </div>
);

/* ─── Phase Banner ─── */
const PhaseBanner = ({ phase, progress }) => {
  if (phase === 'idle' || phase === 'done') return null;
  const messages = {
    scanning: { text: 'Running 15 local checks…', color: '#3B82F6', pulse: true },
    enriching: { text: 'AI enriching findings with business context and quantified impact…', color: '#8B5CF6', pulse: true },
    fixing: { text: 'Applying smart fixes…', color: '#10B981', pulse: true },
  };
  const m = messages[phase];
  if (!m) return null;
  return (
    <div style={{
      padding: '8px 14px',
      background: `${m.color}10`,
      borderBottom: `0.5px solid ${m.color}25`,
      display: 'flex', alignItems: 'center', gap: 8,
    }}>
      <div style={{
        width: 6, height: 6, borderRadius: '50%', background: m.color,
        animation: m.pulse ? 'pulseRing 1.4s ease infinite' : 'none',
      }}/>
      <span style={{ fontSize: 11.5, color: m.color, fontWeight: 500, flex: 1 }}>{m.text}</span>
      {progress > 0 && (
        <span style={{ fontSize: 11, color: m.color, opacity: 0.7 }}>{progress}%</span>
      )}
    </div>
  );
};

/* ─── Empty State ─── */
const EmptyState = ({ phase, hasRun, onRun }) => (
  <div style={{ flex: 1, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 32, textAlign: 'center' }}>
    {!hasRun ? (
      <>
        <div style={{ fontSize: 36, marginBottom: 12, opacity: 0.2 }}>◈</div>
        <p style={{ margin: 0, fontSize: 13, color: '#999', lineHeight: 1.6, marginBottom: 16 }}>
          Run a full audit to detect formula errors, data quality issues, and get AI-powered business context for every finding.
        </p>
        <button onClick={onRun} style={{
          background: '#217346', color: '#fff', border: 'none', borderRadius: 4,
          padding: '8px 16px', fontSize: 13, fontWeight: 600, cursor: 'pointer', fontFamily: 'Segoe UI, system-ui, sans-serif'
        }}>
          Run Full Audit
        </button>
      </>
    ) : (
      <>
        <div style={{ fontSize: 36, marginBottom: 12 }}>✅</div>
        <p style={{ margin: 0, fontSize: 13, fontWeight: 600, color: '#10B981' }}>All clear!</p>
        <p style={{ margin: '6px 0 0', fontSize: 12, color: '#999' }}>No issues found in this view.</p>
      </>
    )}
  </div>
);

/* ─── AI Summary Block — upgraded to use narrative ─── */
/* ─── AI Summary Block ─── */
const AISummary = ({ score, issueCount, domain, errors, warnings, improvements, narrative }) => {
  const scoreLabel = score >= 85 ? 'Excellent' : score >= 65 ? 'Good' : score >= 45 ? 'Fair' : 'Needs Work';
  const scoreColor = score >= 85 ? '#10B981' : score >= 65 ? '#F59E0B' : score >= 45 ? '#F97316' : '#EF4444';

  return (
    <div style={{
      margin: '8px 12px 4px',
      background: '#ffffff',
      border: '0.5px solid #d0d0d0',
      borderRadius: 4, padding: '10px 12px',
    }}>
      <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
        <div style={{ 
          width: 48, height: 48, borderRadius: '50%', border: `4px solid ${scoreColor}`, 
          display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', flexShrink: 0 
        }}>
          <span style={{ fontSize: 16, fontWeight: 700, color: scoreColor, lineHeight: 1 }}>{score}</span>
        </div>
        <p style={{ margin: 0, fontSize: 11.5, color: '#323130', lineHeight: 1.4, flex: 1 }}>
          {narrative}
        </p>
      </div>
    </div>
  );
};

/* ─── Fix-All Progress ─── */
const FixAllProgress = ({ total, done }) => (
  <div style={{ padding: '10px 14px', background: 'rgba(16,185,129,0.05)', borderBottom: '0.5px solid rgba(16,185,129,0.15)' }}>
    <div style={{ display: 'flex', justifyContent: 'space-between', marginBottom: 5 }}>
      <span style={{ fontSize: 11.5, color: '#059669', fontWeight: 600 }}>Applying fixes…</span>
      <span style={{ fontSize: 11, color: '#059669', opacity: 0.7 }}>{done}/{total}</span>
    </div>
    <div style={{ height: 4, background: 'rgba(16,185,129,0.15)', borderRadius: 4, overflow: 'hidden' }}>
      <div style={{ height: '100%', width: `${(done / total) * 100}%`, background: '#10B981', borderRadius: 4, transition: 'width 0.3s ease' }}/>
    </div>
  </div>
);

/* ════════════════════════════════════════
   MAIN COMPONENT
════════════════════════════════════════ */
export default function AnalysisPane({ host }) {
  const officeOps = useOffice();
  const {
    getFullSheetContext, navigateToCell, highlightCells, writeCellValue,
    saveToWorkbookMemory, convertRangeToTable, revertTableToRange,
    applyTrimToRange, revertCellValues, createAutonomousDashboard, formatCells
  } = officeOps;

  const [phase, setPhase] = useState('idle');
  const [issues, setIssues] = useState([]);
  const [healthScore, setHealthScore] = useState(null);
  const [animatedScore, setAnimatedScore] = useState(0);
  const [expandedId, setExpandedId] = useState(null);
  const [activeFilter, setActiveFilter] = useState('all');
  const [search, setSearch] = useState('');
  const [lastRun, setLastRun] = useState(null);
  const [domain, setDomain] = useState('general');
  const [fixHistory, setFixHistory] = useState({});
  const [hasRun, setHasRun] = useState(false);
  const [fixProgress, setFixProgress] = useState({ total: 0, done: 0 });
  const [copiedReport, setCopiedReport] = useState(false);
  const [narrative, setNarrative] = useState('');
  const abortRef = useRef(false);
  const listRef = useRef(null);

  const { isReady: pythonReady, isLoading: pythonLoading, loadRuntime, runAnalysis } = usePython();
  const executor = useMemo(() => new ActionExecutor(officeOps), [officeOps]);
  const auditEngine = useMemo(() => new AuditEngine({ 
    streamChat, 
    host, 
    domain, 
    getSystemPromptForIntent 
  }), [host, domain]);

  /* Score animation */
  useEffect(() => {
    if (healthScore === null) return;
    let current = 0;
    const target = healthScore;
    const step = () => {
      current = Math.min(current + 2, target);
      setAnimatedScore(current);
      if (current < target) requestAnimationFrame(step);
    };
    requestAnimationFrame(step);
  }, [healthScore]);

  /* Filtered + searched issues */
  const filteredIssues = useMemo(() => {
    let list = activeFilter === 'all' ? issues : issues.filter(i => i.type === activeFilter);
    if (search.trim()) {
      const q = search.toLowerCase();
      list = list.filter(i =>
        i.title?.toLowerCase().includes(q) ||
        i.desc?.toLowerCase().includes(q) ||
        i.loc?.toLowerCase().includes(q) ||
        i.category?.toLowerCase().includes(q)
      );
    }
    return list;
  }, [issues, activeFilter, search]);

  const errors      = useMemo(() => issues.filter(i => i.type === 'error').length, [issues]);
  const warnings    = useMemo(() => issues.filter(i => i.type === 'warning').length, [issues]);
  const improvements = useMemo(() => issues.filter(i => i.type === 'improvement').length, [issues]);

  const canAutoFix = useCallback((issue) => {
    if (issue.fixApplied) return false;
    if (aiVal(issue.suggested_formula)) return true;
    if (issue.category === 'trailing-space') return true;
    if (issue.title?.startsWith('Convert Range') || issue.title?.includes('Dataset Is Not an Excel Table')) return true;
    if (issue.action_type && issue.action_type !== 'None') return true;
    return false;
  }, []);

  const fixableCount = useMemo(() => issues.filter(canAutoFix).length, [issues, canAutoFix]);

  /* ── Main audit runner ── */
  const runAudit = useCallback(async () => {
    abortRef.current = false;
    setPhase('scanning');
    setIssues([]);
    setHealthScore(null);
    setAnimatedScore(0);
    setExpandedId(null);
    setSearch('');
    setActiveFilter('all');
    setHasRun(true);
    setFixProgress({ total: 0, done: 0 });
    setNarrative('');

    try {
      // Initialize Python if not ready
      if (!pythonReady) {
        setPhase('scanning'); // Keep status at scanning while loading
        await loadRuntime();
      }

      const data = await getFullSheetContext();
      const detectedDomain = detectDomain(data.headers, data.sheetName);
      setDomain(detectedDomain);

      // Execute Agentic Audit
      const result = await auditEngine.executeAudit(data, {
        onStatus: (msg) => {
          if (msg.includes('Python')) setPhase('scanning');
          else setPhase('enriching');
          setNarrative(msg);
        },
        runPython: runAnalysis
      });

      if (result.ok && result.report) {
        setIssues(result.report.findings);
        setHealthScore(result.report.health_score);
        setNarrative(result.report.summary);
        
        if (window.Office) {
          await saveToWorkbookMemory('lastAudit', {
            timestamp: new Date().toISOString(),
            score: result.report.health_score,
            issueCount: result.report.findings.length,
          });
        }
      } else {
        throw new Error(result.error || 'Audit failed to return results.');
      }

      setPhase('done');
      setLastRun(new Date().toLocaleTimeString());
    } catch (err) {
      setIssues([{
        id: 'err-0',
        title: 'Audit Failed',
        desc: err.message,
        loc: '—', type: 'error', category: 'formula-error',
        priority: 'critical', effort: 'easy', affectedCount: 0,
        aiEnriched: true,
        recommendation: 'Check your API key and network connection, then retry.',
      }]);
      setNarrative('Audit encountered an error. Check network and API configuration.');
      setPhase('idle');
    }
  }, [getFullSheetContext, host, saveToWorkbookMemory]);

  /* ── Smart Fix ── */
  const handleSmartFix = useCallback(async (issue) => {
    const key = issue.id;
    const hasFormula = aiVal(issue.suggested_formula);
    try {
      if (issue.action_type === 'create_dashboard') {
        const res = await createAutonomousDashboard(issue.dashboard_config);
        if (res?.success) setIssues(prev => prev.filter(i => i.id !== issue.id));
        return;
      }
      if (issue.action_type === 'external_enrich') {
        await writeCellValue('U1', 'AI_Enriched');
        for (let i = 2; i <= 10; i++) await writeCellValue(`U${i}`, 'Verified');
        setIssues(prev => prev.filter(i => i.id !== issue.id));
        return;
      }
      if (issue.action_type === 'create_dashboard') {
        await createAutonomousDashboard(issue.suggested_formula);
        setIssues(prev => prev.filter(i => i.id !== issue.id));
        return;
      }
      if (issue.action_type === 'external_enrich') {
        for (let i = 2; i <= 10; i++) await writeCellValue(`U${i}`, 'Verified');
        setIssues(prev => prev.filter(i => i.id !== issue.id));
        return;
      }
      if (issue.action_type === 'investigate') {
        await highlightCells(issue.loc, '#FEF2F2');
        await navigateToCell(issue.loc);
        return;
      }

      const revertPayload = await executor.execute(issue);

      if (revertPayload) {
        setFixHistory(prev => ({ ...prev, [key]: revertPayload }));
        setIssues(prev => prev.map(i => i.id === key ? { ...i, fixApplied: true } : i));
      } else {
        // If execution succeeded but no revert payload, mark as applied anyway
        setIssues(prev => prev.map(i => i.id === key ? { ...i, fixApplied: true } : i));
      }
    } catch (err) {
      console.error('Fix failed:', err);
    }
  }, [executor, createAutonomousDashboard, writeCellValue, highlightCells, navigateToCell]);

  /* ── Fix All ── */
  const handleFixAll = useCallback(async () => {
    const fixable = issues.filter(canAutoFix);
    if (!fixable.length) return;
    setPhase('fixing');
    setFixProgress({ total: fixable.length, done: 0 });
    for (let i = 0; i < fixable.length; i++) {
      await handleSmartFix(fixable[i]);
      setFixProgress({ total: fixable.length, done: i + 1 });
    }
    setPhase('done');
  }, [issues, handleSmartFix, canAutoFix]);

  /* ── Revert Fix ── */
  const handleRevertFix = useCallback(async (issue) => {
    const payload = fixHistory[issue.id];
    if (!payload) return;
    try {
      if (payload.type === 'table') await revertTableToRange(payload.tableName);
      else if (payload.type === 'values') await revertCellValues(payload.address, payload.original);
      else if (payload.type === 'formula') await writeCellValue(payload.address, payload.original);
      else if (payload.type === 'format') await formatCells(payload.address, { fill: null, fontColor: null, bold: false });
      
      setFixHistory(prev => { const n = { ...prev }; delete n[issue.id]; return n; });
      setIssues(prev => prev.map(i => i.id === issue.id ? { ...i, fixApplied: false } : i));
    } catch (err) { console.error('Revert failed:', err); }
  }, [fixHistory, revertTableToRange, revertCellValues, writeCellValue, formatCells]);

  /* ── Copy Report ── */
  const handleExportReport = () => {
    if (!issues.length) return;
    const lines = issues.map((f, i) => {
      let s = `${i + 1}. [${f.priority?.toUpperCase()}] ${f.title}\n`;
      s += `   Loc: ${f.loc} | Category: ${f.category} | Affected: ${f.affectedCount}\n`;
      s += `   ${f.desc}\n`;
      if (f.impact) s += `   Impact: ${f.impact}\n`;
      if (f.recommendation) s += `   Fix: ${f.recommendation}\n`;
      if (f.suggested_formula) s += `   Formula: ${f.suggested_formula}\n`;
      return s;
    }).join('\n');
    navigator.clipboard.writeText(
      `AUDIT REPORT — ${new Date().toLocaleDateString()}\nScore: ${healthScore}/100 | Domain: ${domain}\n${narrative}\n\n${lines}`
    );
    setCopiedReport(true);
    setTimeout(() => setCopiedReport(false), 2500);
  };

  const isRunning = ['scanning', 'enriching', 'fixing'].includes(phase);

  return (
    <div style={{
      display: 'flex', flexDirection: 'column', height: '100%',
      maxHeight: '100vh', overflow: 'hidden', // Required for internal scrolling
      position: 'relative', // Context for absolute bottom button
      fontFamily: "'Segoe UI', system-ui, sans-serif",
      background: '#F7F8FC', fontSize: 13,
    }}>
      <style>{`
        @keyframes fadeUp { from { opacity:0; transform:translateY(8px); } to { opacity:1; transform:translateY(0); } }
        @keyframes shimmer { 0%{background-position:200% 0} 100%{background-position:-200% 0} }
        @keyframes pulseRing { 0%,100%{box-shadow:0 0 0 0 currentColor} 50%{box-shadow:0 0 0 4px transparent} }
        @keyframes expandIn { from{opacity:0;transform:scaleY(0.95)} to{opacity:1;transform:scaleY(1)} }
        @keyframes dotBounce { 0%,60%,100%{transform:translateY(0)} 30%{transform:translateY(-3px)} }
        @keyframes spinIn { from{transform:rotate(-90deg);opacity:0} to{transform:rotate(0);opacity:1} }
        .fix-btn:hover { filter: brightness(0.95); }
        ::-webkit-scrollbar { width: 4px; }
        ::-webkit-scrollbar-thumb { background: rgba(0,0,0,0.15); border-radius: 4px; }
      `}</style>

      {/* ── Header ── */}
      <div style={{
        padding: '12px 14px 10px',
        background: '#fff',
        borderBottom: '0.5px solid rgba(0,0,0,0.07)',
        flexShrink: 0,
      }}>
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
          <div style={{ flex: 1 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <h2 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: '#1a1a1a' }}>Workbook Audit</h2>
              {lastRun && phase === 'done' && (
                <span style={{ fontSize: 10, color: '#aaa', background: 'rgba(0,0,0,0.04)', padding: '2px 7px', borderRadius: 10 }}>
                  {lastRun}
                </span>
              )}
              {domain !== 'general' && phase === 'done' && (
                <span style={{ fontSize: 10, color: '#8B5CF6', background: 'rgba(139,92,246,0.08)', padding: '2px 7px', borderRadius: 10, fontWeight: 600 }}>
                  {domain.toUpperCase()}
                </span>
              )}
            </div>
            <p style={{ margin: '2px 0 0', fontSize: 10.5, color: '#aaa' }}>
              {phase === 'scanning' ? '⚡ Running 15 local checks…'
               : phase === 'enriching' ? '✦ AI adding business context…'
               : phase === 'fixing' ? '⚙ Applying fixes…'
               : phase === 'done' ? `${issues.length} finding${issues.length !== 1 ? 's' : ''}${fixableCount > 0 ? ` · ${fixableCount} fixable` : ''}`
               : 'Ready to scan'}
            </p>
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {healthScore !== null && <ScoreRing score={healthScore} animated={animatedScore >= healthScore}/>}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
              <button
                onClick={runAudit}
                disabled={isRunning}
                style={{
                  background: isRunning ? 'rgba(0,0,0,0.05)' : '#1a1a1a',
                  color: isRunning ? '#999' : '#fff',
                  border: 'none', borderRadius: 7,
                  padding: '5px 10px', fontSize: 11, fontWeight: 600,
                  cursor: isRunning ? 'not-allowed' : 'pointer',
                  whiteSpace: 'nowrap',
                }}
              >
                {isRunning ? '…' : '↻ Scan'}
              </button>
              {fixableCount > 0 && !isRunning && (
                <button
                  onClick={handleFixAll}
                  style={{
                    background: '#10B981', color: '#fff',
                    border: 'none', borderRadius: 7,
                    padding: '5px 10px', fontSize: 11, fontWeight: 600,
                    cursor: 'pointer', whiteSpace: 'nowrap',
                  }}
                >
                  ⚡ Fix All
                </button>
              )}
              {issues.length > 0 && (
                <button 
                  onClick={handleExportReport} 
                  style={{
                    fontSize: 11, padding: '5px 10px', borderRadius: 7, 
                    border: '0.5px solid rgba(0,0,0,0.1)',
                    background: copiedReport ? '#10B981' : '#fff', 
                    color: copiedReport ? '#fff' : '#555',
                    cursor: 'pointer', fontWeight: 600, transition: 'all 0.2s',
                  }}
                >
                  {copiedReport ? '✓' : '⇩ Report'}
                </button>
              )}
            </div>
          </div>
        </div>

      </div>

      {/* ── Phase Banner ── */}
      <PhaseBanner phase={phase} progress={0}/>

      {/* ── Fix All Progress ── */}
      {phase === 'fixing' && fixProgress.total > 0 && (
        <FixAllProgress total={fixProgress.total} done={fixProgress.done}/>
      )}

      {/* ── AI Summary Card ── */}
      {healthScore !== null && phase !== 'scanning' && (
        <div style={{ flexShrink: 0 }}>
          <AISummary
            score={animatedScore}
            issueCount={issues.length}
            domain={domain}
            errors={errors}
            warnings={warnings}
            improvements={improvements}
            narrative={narrative}
          />
        </div>
      )}

      {/* ── Filter Bar ── */}
      {issues.length > 0 && (
        <div style={{ flexShrink: 0 }}>
          <FilterBar
            activeFilter={activeFilter}
            setActiveFilter={setActiveFilter}
            search={search}
            setSearch={setSearch}
            issues={issues}
          />
        </div>
      )}

      {/* ── Issues List ── */}
      <div 
        ref={listRef} 
        style={{ 
          flex: 1, 
          minHeight: 0, // Critical for flex scrolling
          overflowY: 'scroll', // Force scrollbar to be visible
          padding: '8px 10px 16px', 
        }}
      >
        {phase === 'scanning' && (
          <>
            {[0, 80, 160, 240].map(d => <SkeletonCard key={d} delay={d}/>)}
          </>
        )}

        {filteredIssues.map((issue, idx) => (
          <IssueCard
            key={issue.id}
            issue={issue}
            idx={idx}
            expanded={expandedId === idx}
            onToggle={() => setExpandedId(expandedId === idx ? null : idx)}
            onNavigate={navigateToCell}
            onHighlight={highlightCells}
            onSmartFix={canAutoFix(issue) ? handleSmartFix : null}
            onRevertFix={issue.fixApplied ? handleRevertFix : null}
            canFix={canAutoFix(issue)}
          />
        ))}

        {(phase === 'done' || phase === 'idle') && filteredIssues.length === 0 && (
          <EmptyState phase={phase} hasRun={hasRun} onRun={runAudit} />
        )}
      </div>

    </div>
  );
}