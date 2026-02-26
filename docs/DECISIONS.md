# Architectural Decisions (ADR)

## 2026-01-29: Initial Audit
- **Context**: User reports input failure in specific cells.
- **Decision**: Perform a static analysis of the Apps Script code and use browser tools to verify spreadsheet state if possible.
- **Rationale**: Direct debugging of Apps Script environment requires understanding the trigger logic (e.g., `onEdit`).
