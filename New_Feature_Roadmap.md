# Enterprise AI Copilot: Feature Roadmap (New Architecture)

This roadmap reimagines the original feature set through the lens of the **Enterprise Target Architecture** (LangGraph, Zod Schemas, Central Commit Service, and Hybrid Intelligence).

---

## Phase 1: Core Agentic Foundation & Formula Engine
*Focus: Establishing the strict IO boundaries, state machines, and the most critical Excel skill.*

*   **Agent Orchestration UI:** 
    *   Instead of a simple chat window, the UI supports **Action Cards** (Preview, Confirm, Discard) powered by Zustand state.
*   **Encrypted Key Manager:** 
    *   Move from plain `localStorage` to an encrypted vault for API keys.
*   **Selection Context Engine (V2):** 
    *   RAG-lite implementation. The system strictly bounds context gathering using policies (`selection` vs `table` vs `sheet`) to prevent token window overflow.
*   **Formula Agent (Implemented ✅):** 
    *   Uses `ExecutionEngine.ts`. 
    *   Workflow: `Plan -> Zod Schema -> Generate -> Validate (Safety, Shape, Semantics) -> Repair -> Preview`.

---

## Phase 2: Hybrid Analysis & Cleaning Agents
*Focus: Upgrading unstructured analytical features into strict, deterministic Hybrid pipelines.*

*   **Audit Agent (V2 Upgrade):** 
    *   *Hybrid Engine:* Local deterministic rules run first to find `#REF!`, `#DIV/0!`, or hardcoded totals. 
    *   *LLM Role:* Enriches findings, assesses business impact, and maps to `AuditFindingSchema`.
*   **Data Cleaning Agent:** 
    *   LLM does **not** write data directly. It outputs a `DataCleaningPlanSchema` (e.g., `[{"action": "TRIM", "column": "A"}, {"action": "COERCE_TYPE", "type": "DATE", "column": "B"}]`).
    *   The `CommitService` executes the plan transactionally via Office.js.
*   **Column Classifier & Tagger Agent:** 
    *   Uses a fast secondary model (e.g., Llama 3 / Haiku).
    *   Outputs strict `CategoryTagSchema` array.
*   **Multi-Range Comparison & Anomaly Detection:**
    *   Agent generates a `ComparisonDeltaSchema` isolating exact cell discrepancies before rendering a UI summary card.

---

## Phase 3: Autonomous Dashboards & Word Reporting
*Focus: Structured asset generation and cross-application bridges.*

*   **Smart Pivot & Dashboard Agent:** 
    *   User asks for a "Sales Dashboard".
    *   LLM outputs a `DashboardConfigSchema` (defining charts, KPIs, and pivot rows/cols).
    *   Office.js reads the JSON config and deterministically builds the 12 templates (Sales, Finance, HR, etc.) without relying on the LLM to write Office.js code.
*   **Word Document Q&A & Inconsistency Agent:** 
    *   *RAG Layer:* Chunks the Word document into semantic blocks.
    *   Outputs `DocumentInsightSchema` with mandatory exact-paragraph citations.
*   **Excel-to-Word Reporting Bridge:** 
    *   Agent reads Excel `DashboardConfigSchema`, generates an `ExecutiveSummarySchema`, and uses Word.js to insert the narrative alongside linked charts.

---

## Phase 4: Multi-Step LangGraph Orchestration
*Focus: Complex workflows spanning multiple intents and long-running operations.*

*   **Macro / Script Generator Agent:** 
    *   Outputs `OfficeScriptSchema`. 
    *   Passes through a strict Security Validator gate before presenting the "Apply Script" Action Card.
*   **Multi-Step Autonomous Agents (LangGraph):** 
    *   "Clean the data, audit for errors, and build a finance dashboard."
    *   LangGraph manages state across multiple agents: `DataCleaner -> Auditor -> DashboardBuilder`, handling rollbacks if any node fails.
*   **Custom Prompt Library & Telemetry:** 
    *   Observability layer tracks token usage per agent node.
    *   Saved prompts map to specific Zod schemas for structured repeatability.
