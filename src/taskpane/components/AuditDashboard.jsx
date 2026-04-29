import React from 'react';
import '../styles/audit-dashboard.css';

const SeverityIcons = {
  CRITICAL: '🔴',
  HIGH: '🟠',
  MEDIUM: '🔵',
  LOW: '⚪'
};

export const AuditDashboard = ({ report, onClose, onApplyFix }) => {
  if (!report) return null;

  const scoreClass = report.health_score > 80 ? 'excellent' : report.health_score > 50 ? 'warning' : 'poor';

  return (
    <div className="audit-dashboard-overlay">
      
      {/* Header */}
      <div className="audit-dashboard-header">
        <div>
          <h2 className="audit-dashboard-title">Data Health Audit</h2>
          <div className="audit-dashboard-subtitle">Hybrid Intelligence Engine</div>
        </div>
        <button className="audit-dashboard-close" onClick={onClose}>✕</button>
      </div>

      {/* Scrollable Content */}
      <div className="audit-dashboard-content scroll-area">
        
        {/* Score Ring & Summary */}
        <div className="audit-score-card">
          <div className={`audit-score-ring ${scoreClass}`}>
            <span className={`audit-score-value ${scoreClass}`}>{report.health_score}</span>
          </div>
          <div>
            <div className="audit-summary-title">Executive Summary</div>
            <div className="audit-summary-text">{report.summary}</div>
          </div>
        </div>

        {/* Findings List */}
        <div className="audit-findings-label">
          Detailed Findings ({report.findings?.length || 0})
        </div>
        
        <div className="audit-findings-list">
          {report.findings?.map((finding) => {
            const icon = SeverityIcons[finding.severity] || SeverityIcons.LOW;
            return (
              <div key={finding.id} className={`audit-finding-card severity-${finding.severity}`}>
                <div className="audit-finding-body">
                  <div className="audit-finding-head">
                    <span className="audit-finding-icon">{icon}</span>
                    <div>
                      <div className="audit-finding-title">{finding.title}</div>
                      <div className={`audit-finding-badge severity-${finding.severity}`}>
                        {finding.severity} SEVERITY
                      </div>
                    </div>
                  </div>
                  
                  <div className="audit-finding-desc">
                    {finding.desc}
                  </div>
                  
                  <div className="audit-finding-impact">
                    <strong>Impact:</strong> {finding.impact}
                  </div>

                  {finding.action_type === 'FIX_FORMULA' && finding.suggested_formula && (
                    <div className="audit-finding-action-row">
                      <code className="audit-finding-code">
                        {finding.suggested_formula}
                      </code>
                      <button 
                        className="audit-finding-btn"
                        onClick={() => onApplyFix(finding)}>
                        Apply Fix
                      </button>
                    </div>
                  )}
                </div>
              </div>
            );
          })}

          {(!report.findings || report.findings.length === 0) && (
            <div className="audit-empty-state">
              No critical issues found. Your data is looking great! 🎉
            </div>
          )}
        </div>

      </div>
    </div>
  );
};
