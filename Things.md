# Project Development Roadmap

## ✅ Completed & Stable

_Core foundations that are established and do not require immediate changes._

1. **Groq Streaming Foundation**: Token streaming pipeline is implemented and stable enough as a base.
2. **Core Tab Structure**: Chat / Audit / Skills / Setup architecture is in place and reusable.
3. **Office Bridge Base**: Core `Excel.run()` integration for read/write/highlight/navigation exists and should be extended, not rewritten.

---

## 🟡 Partial Implementations

_Features that are functional but require further refinement or feature completion._

- **AI Chat Task Pane**: Works, but full conversation history persistence is not complete.
- **Groq Model Switcher**: Implemented, but model set does not fully match the planned lineup and routing policy is basic.
- **Selection Context Engine**: Implemented, but toggle controls are not fully wired and context policy is not granular enough.**Slash Command Palette**: Implemented for many commands, but not all planned commands/flows are complete end-to-end.
- **Formula Writer & Debugger**: Implemented with auto-apply and explain flow, but needs stronger validation/confidence and better edge-case handling.
- **Response Preview & Insert**: Partially present (some write confirmations), but not a unified accept/reject workflow for all actions.
- **Full Dataset Analyzer / Metrics Analysis**: Partially present via audit + metrics intents, but not yet a complete “analyst-grade” report engine.
- **Smart Pivot/Chart Generation**: Works in basic form, but reliability and schema robustness need improvement.
- **Data Cleaning Suite**: Partial (trim, some fix actions), but not full one-click cleaning coverage.
- **Column Classifier/Tagger**: Prompt layer exists, but not complete production feature flow/UI.
- **Conditional Format Intelligence**: Very limited/indirect; not complete feature.
- **Dashboard Auto-Builder**: Early autonomous dashboard generation exists, but template/mapping quality is incomplete.

---

## 📈 Optimization Required

_Features that are functionally complete but need UX/UI or engine upgrades._

- **Audit Engine UI/UX**: Complete enough to use, but faces practicality issues (small effective viewport, density, sticky space usage).
- **Audit Local-First Engine**: Substantial and functional, but scoring, confidence, cross-column intelligence, and explainability need V2 upgrades.
- **Settings and Preferences**: Functional UI exists, but security and enforceability of settings are not strong enough.
- **Error Handling around AI Outputs**: Present, but needs stricter schema enforcement and safer fallback behavior.

---

## 🚀 Planned (Untouched)

_Future modules and features scheduled for upcoming development cycles._

### Word & Cross-App Integration

- Word text rewriter full workflow
- Document Q&A with exact paragraph citation
- Word inconsistency detector
- Template-based Word report generator (8 templates)
- Excel-to-Word live bridge with captions
- Executive summary writer from dashboard to Word

### Advanced Analytics & Tools

- Multi-range comparison (fully productized)
- Advanced table builder (sortable/filterable/color-coded full pipeline)
- Anomaly detection with severity model beyond current baseline
- 12 professional dashboard templates system
- VBA macro generator with tested insertion
- AI forecasting engine with confidence intervals
- Multi-step agent orchestration flow

### Platform & Infrastructure

- Scheduled reports with snapshot versioning
- Custom prompt library
- Token usage dashboard + cost estimator
- AppSource publish pipeline
- Secure encrypted API key storage (current storage is plain local storage, so planned security target is untouched)

---

_Note: This roadmap can be converted into a strict checklist with P0/P1/P2 priority and owner-style execution order upon request._
