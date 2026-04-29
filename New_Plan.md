## Comprehensive Plan (New Architecture)

### Phase 0: Stabilization Baseline

1. Freeze current working behavior and create golden test sheets (**data for AI** scenarios).
2. Add error envelope format for all failures (**code**, **stage**, **message**, **recoverable**).
3. Add feature flags per intent so migrations are reversible.

### Phase 1: Agent Core Foundation

1. Intent Router
   - Centralized classifier for **FORMULA | EXPLAIN | AUDIT | PIVOT | CLEAN | CLASSIFY | COMPARE | METRICS**.
2. Typed Schemas (**zod**)
   - Request schema, context schema, per-intent output schemas, execution log schema.
3. Execution Engine
   - **plan -> run -> validate -> commit**.
4. Step Runner
   - Step timeouts, retries, cancellation, fallback path.
5. Validator Gates
   - Schema gate, formula safety gate, semantic gate, policy gate.
6. Observability

- Structured logs for every run and every gate result.

### Phase 2: Formula Pipeline (Production Hardening)

1. Planner for formula tasks
   - Identify lookup/aggregation/math pattern.
2. Generator + repair retries
   - Retry 1 strict output correction, Retry 2 reduced context.
3. Semantic validators
   - Header-intent matching, range alignment, function compatibility.
4. Action isolation
   - Preview -> apply -> overwrite confirm -> rollback support.
5. Confidence scoring

- Show confidence + assumptions before apply.

### Phase 3: Explain Pipeline (Analyst-grade)

1. Explain schema
   - **summary**, **facts**, **metrics**, **assumptions**, **followUps**.
2. Semantic truth checks
   - Header meaning preservation (**Unit_Price** != **Total Sales**).
3. Citation normalization
   - Only real cell/range references.
4. Contradiction checks
   - Reject self-contradictory output.
5. Presentation layer

- concise view + detailed view.

### Phase 4: Audit Pipeline v2

1. Local deterministic engine remains source of truth.
2. AI only enriches, cannot overwrite verified findings.
3. Finding schema
   - severity, evidence, impact, fix-plan, confidence.
4. Fix planner
   - previewable bulk fixes, per-fix rollback.
5. Audit history/diff

- new/resolved/regressed findings across runs.

### Phase 5: Data Cleaning Pipeline

1. Planner decomposes cleaning intent into operations.
2. Operation schemas
   - trim, dedupe, null strategy, casing, type coercion.
3. Dry-run mode
   - show rows impacted and sample diff.
4. Commit in transactions
   - apply selected operations only.
5. Rollback stack per operation set.

### Phase 6: Pivot/Chart/Metrics Pipeline

1. Structured dashboard spec schema.
2. Chart/pivot validators
   - header existence, type compatibility, aggregator validity.
3. Retry strategy for invalid specs.
4. Preview spec card before Office apply.
5. Commit layer inserts assets with deterministic naming.

### Phase 7: Classify/Compare/What-if Pipelines

1. Classify schema (labels, confidence, rationale).
2. Compare schema (delta map, trend summary, anomalies).
3. What-if schema (assumption, impacted ranges, projected deltas).
4. Policy gates for maximum affected range and write safety.

### Phase 8: Context Engine v2

1. Unified context contract for all intents.
2. Scope policies (**selection**, **table**, **sheet**) with caps.
3. Header map + column semantics always included when relevant.
4. Context budget manager
   - truncation + summary strategy.
5. Real-time selection synchronization (already started, harden).

### Phase 9: Commit & Safety Layer

1. Central **CommitService**
   - all writes go through it.
2. Write modes
   - **PreviewOnly**, **ConfirmWrites**, **AutoSafe**.
3. Guardrails
   - protected range checks, shape mismatch checks.
4. Rollback journal

- reversible operations with snapshot IDs.

### Phase 10: Prompt/Policy Governance

1. Prompt registry by intent + versioning.
2. Output contracts embedded in prompts.
3. Policy rules per intent/domain.
4. Automatic prompt fallback variants on retries.

### Phase 11: Testing & Quality

1. Unit tests
   - schemas, validators, planner logic, retries.
2. Integration tests
   - engine with mocked model responses.
3. Workbook scenario tests
   - **data for AI** fixed scenarios.
4. Regression suite
   - known failure cases (semantic drift, invalid formula syntax).
5. Performance tests

- latency and large-sheet behavior.

### Phase 12: UX Alignment

1. Unified action cards for preview/apply/discard.
2. Validation error UX with corrective guidance.
3. Inspector panel (optional)
   - plan steps, logs, validation outcomes.
4. Compact mode for dense productivity use.

### Phase 13: Security & Operations

1. Secure key storage upgrade path.
2. Sensitive-data minimization in prompts.
3. Telemetry redaction policy.
4. Rate-limit backoff + graceful degradation.

### Phase 14: Extensibility

1. Plugin-style intent modules.
2. Shared interfaces for planner/runner/validator/commit.
3. Versioned module contracts for safe upgrades.
4. Optional migration path to LangGraph/AutoGen orchestrators.
