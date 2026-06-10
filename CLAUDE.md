# Excel AI Plugin — agent guide

Open-source, bring-your-own-model AI plugin for Microsoft Excel.
Architecture: thin Office.js add-in (`addin/`) + local Node.js sidecar (`sidecar/`)
sharing protocol types and tool schemas (`shared/`). Design doc: `docs/DESIGN.md`.

## Commands

| Command | What it does |
|---|---|
| `pnpm install` | Install all workspace deps (Node >= 22.13 required for `node:sqlite`) |
| `pnpm check` | Typecheck every package (builds `shared` first) |
| `pnpm test` | Vitest unit + integration suites (agent loop runs against the workbook emulator) |
| `pnpm build` | Build shared → sidecar → addin |
| `pnpm smoke` | Full-stack E2E: live sidecar + mock OpenAI-compatible model + fake add-in over real WS |
| `pnpm dev:sidecar` | Build & start the sidecar (prints pairing code, serves built add-in) |

## Rules

- A change is done only when `pnpm check && pnpm test && pnpm build && pnpm smoke` all pass.
- `shared/src/tools.ts` is the single source of truth for the tool surface. If you touch it,
  update BOTH implementations: `sidecar/src/workbook/emulator.ts` and `addin/src/officeBridge.ts`,
  plus the agent tests.
- All formulas in the protocol are canonical en-US (`formulas`, never `formulasLocal`).
- Write tools must stage into change-sets — never apply directly. Every write op carries a `reason`.
- Tests must not touch `~/.excelai` (the vitest setup redirects `EXCELAI_DATA_DIR`; keep it that way).
- The emulator's formula engine implements functions exactly or returns `#NAME?` — never approximate.
- No GPL dependencies (HyperFormula was rejected for this reason; the in-repo evaluator replaces it).

## Layout

- `shared/src/` — tool schemas (zod), WS protocol envelopes, change-set types, A1 helpers
- `sidecar/src/agent/` — tool-calling loop, schema validation, repair-retry, prompts
- `sidecar/src/changeset/` — stage → preview → approve → apply → verify (auto-rollback) → undo
- `sidecar/src/providers/` — Anthropic, OpenAI(+compatible: Ollama/LM Studio/vLLM), scripted (tests)
- `sidecar/src/workbook/` — executor interface, emulator, formula engine
- `sidecar/test/` — all tests; `sidecar/scripts/smoke.mjs` — E2E
- `addin/src/` — taskpane UI, Office.js bridge, custom functions (`=EXCELAI.AI()` family)

## Testing against real Excel (manual, needs a machine with Excel)

1. `pnpm build && pnpm dev:sidecar` (serves the add-in at http://localhost:8923)
2. Sideload `addin/manifest.xml` into Excel (web: Home → Add-ins → More Settings → Upload My Add-in)
3. Enter the pairing code from the sidecar console in the task pane, configure a model, chat.

Known platform caveat: Safari/Mac WKWebView does not exempt http://localhost from
mixed-content blocking; Windows + Excel on the web (Chromium) are the dev targets for now.
