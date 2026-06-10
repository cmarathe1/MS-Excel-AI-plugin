# MS Excel AI Plugin

**The last AI plugin you install in Excel.** Open source, bring-your-own-model, local-first.

Connect **any** AI model — Claude, OpenAI, Gemini, OpenRouter, any OpenAI-compatible endpoint, or fully local models via Ollama / LM Studio — and let it understand, analyze, and modify your workbooks as a real agent, with every change previewable and undoable.

## Why another Excel AI tool?

Because every existing one locks you to a single vendor, charges a subscription, writes into your sheets with no preview, forgets everything between sessions, and can't see anything outside the open workbook. This project is built around exactly those gaps:

- 🔑 **Bring your own model** — your keys, your endpoints, your local models. Works 100% offline with Ollama.
- 📦 **Optional built-in local model** — flip one switch and the plugin works with no API key, no Ollama, no setup (a small Apache-2.0 model like Qwen3-4B, run in-process). Strictly opt-in: it downloads nothing and consumes zero memory or compute until you enable it, loads lazily on first use, and unloads itself when idle. A project fine-tune specialized for spreadsheet work (**ExcelLM**) is on the roadmap.
- 🛡️ **Safe agentic edits** — every AI change is a reviewable change-set: preview → approve → one-click undo, with a full audit log.
- 🎯 **Works well with *any* model, weak or strong** — a model-agnostic harness does the heavy lifting: capability probing grades each connected model, scaffolding adapts to it, every tool call is schema-validated, every edit is verified after the fact, and ready-made playbooks cover the jobs Excel is actually used for (reconciliation, cleaning, budgeting, reporting, merging) so precision comes from the system, not the model.
- 📊 **Built for big workbooks** — a workbook indexer and formula dependency graph let the agent work on 50k-row, 20-tab models without blowing the context window.
- ⚡ **Bulk AI functions** — `=AI()`, `=AI_CLASSIFY()`, `=AI_EXTRACT()` and friends, batched and cached, routed to cheap or local models.
- 🧠 **Real memory** — per-workbook and per-user memory that makes every session smarter than the last. Local, visible, editable, exportable.
- 🔌 **Open data connectivity** — import PDFs, Word docs, and scanned tables into your sheet with zero setup (bundled converter), read other local workbooks and CSVs, and connect databases, SaaS tools, and APIs through the [Model Context Protocol](https://modelcontextprotocol.io) if you want to go further.
- 🔒 **Private by default** — a local companion app keeps keys in your OS keychain and all data on your machine. No accounts, no telemetry, no server.

## Architecture in one paragraph

A thin Office.js add-in (taskpane + custom functions) acts as the agent's "hands and eyes" inside Excel, and a local companion app (the *sidecar*) running on `localhost` hosts everything stateful and sensitive: the model provider layer, the agent loop, the memory store, the MCP client, the batch engine, and your API keys.

## 📐 Design document

The full design — problem analysis, competitive gaps, architecture, safety model, memory system, MCP integration, and the phased roadmap — lives in **[docs/DESIGN.md](docs/DESIGN.md)**.

## Status

🚧 **Phase 1 (Foundation) implemented** — see the [roadmap](docs/DESIGN.md#16-phased-roadmap):

- ✅ Monorepo: `shared/` (tool schemas + WS protocol) · `sidecar/` (Node daemon) · `addin/` (Office.js)
- ✅ Model layer: Anthropic, OpenAI, and any OpenAI-compatible endpoint (Ollama, LM Studio, vLLM, …)
- ✅ Agent loop with schema-validated tools, repair-retry for malformed calls, and emulated tool calling for models without native support
- ✅ Change-sets: stage → preview (with reasons) → approve → verify (auto-rollback on introduced errors) → one-click undo
- ✅ Sidecar: pairing-code auth, provider settings API, `=AI()` function batch engine with caching, cost ledger, FTS-backed memory store
- ✅ Headless workbook emulator with a real formula engine, so the entire stack is tested without Excel: 40 unit/integration tests + a full-stack E2E smoke test (`pnpm smoke`)
- ⬜ Workbook indexer, playbooks, built-in local model, MCP, document import — next phases

### Run it

```bash
pnpm install        # Node >= 22.13
pnpm check && pnpm test && pnpm smoke   # typecheck, tests, full-stack E2E
pnpm build && pnpm dev:sidecar          # start the sidecar (prints a pairing code)
```

Then sideload `addin/manifest.xml` into Excel (web: Home → Add-ins → More Settings → Upload My Add-in), enter the pairing code in the task pane, pick your model, and chat.

> Current dev targets: Excel on Windows and Excel on the web (Chromium). Excel for Mac needs an HTTPS-localhost story (WKWebView blocks mixed content) — tracked in the design risks.

## License

MIT (planned).
