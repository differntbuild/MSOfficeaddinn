# Target Architecture: Enterprise AI Office Copilot

This document outlines the target architecture for the AI Office Copilot add-in. It is designed to be highly reliable, enterprise-grade, and resilient against AI hallucinations and destructive actions.

## 1. Frontend & Host Integration
**Tech Stack:** React, Office.js, Zustand, TailwindCSS/FluentUI
*   **Thin UI Architecture:** The frontend is strictly a presentation layer. It manages user input, renders AI responses, and displays Action Cards (preview, confirm, discard).
*   **State Management:** Zustand manages local UI state (loading indicators, chat history, active view). 
*   **No Reasoning Logic:** The frontend does not construct complex prompts or handle validation logic. It only sends intents and context payloads to the Orchestration Core.

## 2. Agent Orchestration Core
**Tech Stack:** LangGraph (State Machine), TypeScript
*   **Deterministic Workflows:** Replaces single-shot prompting with a directed acyclic graph (DAG) of states: `Route -> Plan -> GatherContext -> Generate -> Validate -> Preview -> Commit -> Log`.
*   **Action Isolation:** Direct "model -> write" paths are strictly prohibited. All destructive or modifying actions halt at the `Preview` state awaiting user confirmation.

## 3. Typed Contracts & Schema Enforcement
**Tech Stack:** TypeScript, Zod
*   **Strict IO Boundaries:** Every intent (Formula, Explain, Audit, Pivot/Chart, Cleaning) has a strictly typed schema.
*   **Schema Gates:** The orchestration core rejects any LLM output that fails Zod validation, automatically triggering the fallback/repair chain.

## 4. Hybrid Intelligence Engine
*   **Deterministic First:** Local, rule-based engines (like `auditEngine.js`) run first to detect factual, structural, or statistical anomalies.
*   **AI Enrichment Second:** The LLM consumes the deterministic findings to provide natural language explanations, executive summaries, and prioritization. The AI is restricted from overwriting verified deterministic findings.

## 5. Model Routing & Fallback Chain
*   **Dynamic Routing:** 
    *   *Primary Model* (e.g., GPT-4o / Claude 3.5 Sonnet) for complex reasoning (Audit, Multi-step Planning).
    *   *Secondary Model* (e.g., Llama 3 / Haiku) for fast, low-risk tasks (Classification, formatting).
*   **Retry Chain:** If validation fails, the core attempts up to 2 retries with progressively constrained prompts and reduced context to force compliance.

## 6. RAG & Semantic Memory
**Tech Stack:** LlamaIndex, Qdrant/Weaviate (Vector Store)
*   **Context Scope Policies:** Retrieval is bounded by user-defined scopes (`selection`, `table`, `sheet`, `workbook`).
*   **Workbook Awareness:** Embeds headers, named ranges, and summaries of past AI runs. This allows the AI to answer "What did we do to this sheet last week?" and prevents it from losing context in massive workbooks.

## 7. Validation & Guardrails Framework
*   **Formula Validators:** Checks syntax validity, range alignment, and XLOOKUP vector shapes.
*   **Semantic Validators:** Ensures header meanings are preserved (e.g., verifying that "Unit Price" isn't misconstrued as "Total Revenue").
*   **Policy Guardrails:** Enforces permission boundaries and checks against modifying protected or locked ranges.

## 8. Safe Write & Commit Layer
*   **Central `CommitService`:** All Office.js `context.sync()` write operations funnel through a single service.
*   **Transaction Lifecycle:** `Preview -> Confirm -> Apply -> Record Snapshot`.
*   **Rollback Journal:** Maintains snapshot IDs for every write operation, allowing users to reliably undo complex AI operations (like bulk data cleaning).

## 9. Observability & Evaluation
**Tech Stack:** OpenTelemetry, Ragas/DeepEval
*   **Distributed Tracing:** Every node in the LangGraph is traced with OpenTelemetry to track latency and failure rates.
*   **Structured Logging:** Execution logs capture the exact prompt, retries, gate outcomes, and the final applied action.
*   **Eval Pipeline:** Changes to prompts or models are run against "Golden Workbooks" using DeepEval to ensure zero regressions in formula generation accuracy or audit findings.

## 10. Testing Strategy
*   **Unit Tests:** For Zod schemas, validators, routers, and planners.
*   **Integration Tests:** Testing the LangGraph workflows with mocked, deterministic LLM JSON responses.
*   **Scenario Tests:** End-to-end tests using fixed Office fixtures (`data for AI` sheets) to verify actual host behavior.

## 11. Security & Enterprise Readiness
*   **Key Management:** Encrypted, secure storage for API keys.
*   **Data Minimization:** PII filtering/redaction layer strips sensitive names or IDs before the context payload hits the LLM.
*   **Explicit Consent:** No hidden background writes. Every modification requires explicit user UI interaction.
*   **Content Security Policy (CSP):** Strict boundaries on what domains the add-in can communicate with.

## 12. Progressive Rollout Plan
*   **Phase A:** Perfecting Formula + Explain pipelines.
*   **Phase B:** Implementing the Hybrid Audit + Data Cleaning engines.
*   **Phase C:** Launching the Pivot, Chart, and Dashboard builder.
*   **Phase D:** Advanced multi-step autonomous agents (e.g., "Find the errors, fix them, and build me a summary deck").
