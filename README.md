# MS Excel AI Plugin

**The last AI plugin you install in Excel.** Open source, bring-your-own-model, local-first.

Connect **any** AI model — Claude, OpenAI, Gemini, OpenRouter, any OpenAI-compatible endpoint, or fully local models via Ollama / LM Studio — and let it understand, analyze, and modify your workbooks as a real agent, with every change previewable and undoable.

## Why another Excel AI tool?

Because every existing one locks you to a single vendor, charges a subscription, writes into your sheets with no preview, forgets everything between sessions, and can't see anything outside the open workbook. This project is built around exactly those gaps:

- 🔑 **Bring your own model** — your keys, your endpoints, your local models. Works 100% offline with Ollama.
- 🛡️ **Safe agentic edits** — every AI change is a reviewable change-set: preview → approve → one-click undo, with a full audit log.
- 📊 **Built for big workbooks** — a workbook indexer and formula dependency graph let the agent work on 50k-row, 20-tab models without blowing the context window.
- ⚡ **Bulk AI functions** — `=AI()`, `=AI_CLASSIFY()`, `=AI_EXTRACT()` and friends, batched and cached, routed to cheap or local models.
- 🧠 **Real memory** — per-workbook and per-user memory that makes every session smarter than the last. Local, visible, editable, exportable.
- 🔌 **Open data connectivity** — connect databases, SaaS tools, and APIs through the [Model Context Protocol](https://modelcontextprotocol.io), plus read other local workbooks and CSVs.
- 🔒 **Private by default** — a local companion app keeps keys in your OS keychain and all data on your machine. No accounts, no telemetry, no server.

## Architecture in one paragraph

A thin Office.js add-in (taskpane + custom functions) acts as the agent's "hands and eyes" inside Excel, and a local companion app (the *sidecar*) running on `localhost` hosts everything stateful and sensitive: the model provider layer, the agent loop, the memory store, the MCP client, the batch engine, and your API keys.

## 📐 Design document

The full design — problem analysis, competitive gaps, architecture, safety model, memory system, MCP integration, and the phased roadmap — lives in **[docs/DESIGN.md](docs/DESIGN.md)**.

## Status

🚧 Design phase. No code yet — the design document above is the current deliverable. Feedback and contributions to the design are welcome via issues.

## License

MIT (planned).
