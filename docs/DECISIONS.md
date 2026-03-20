# Architectural Decisions (ADR)

## 2026-03-19: Documenting Item Creation Logic
- **Context**: The logic for item name and code generation in `door_v4.gs` was complex and undocumented, making maintenance difficult.
- **Decision**: Create a comprehensive logic document (`docs/품목생성로직.md`) using a conclusion-first, table-rich format.
- **Rationale**: Ensures that future developers (or AI assistants) can quickly understand the mapping between sheet columns and ERP codes without re-analyzing 1000+ lines of code.
