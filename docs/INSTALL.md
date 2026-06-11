# Install & test on your machine

Verified targets: **Excel on Windows (desktop)** and **Excel on the web** (Edge/Chrome).
You need: [Node.js ≥ 22.13](https://nodejs.org), pnpm (`corepack enable`), and a clone of this repo.

## 1. Build and start the sidecar

```bash
pnpm install
pnpm build
pnpm dev:sidecar
```

You should see:

```
Excel AI sidecar listening on http://127.0.0.1:8923
Pairing code (enter it in the add-in once): AB12CD
```

Leave this terminal running. The sidecar serves the add-in UI, talks to your model,
and stores everything (keys, history, cache) locally under `~/.excelai`.

Sanity check: open <http://localhost:8923/api/health> in a browser — you should see `{"ok":true,…}`.

## 2. Sideload the add-in into Excel

**Excel on the web** (fastest):
1. Open [office.com](https://office.com) → Excel → blank workbook.
2. **Home → Add-ins → More Settings** (or **Insert → Office Add-ins**) → **Upload My Add-in**.
3. Browse to `addin/manifest.xml` in this repo → **Upload**.

**Excel on Windows (desktop):**
1. Share a folder containing `addin/manifest.xml` (right-click folder → Properties → Sharing → Share), or use a local path via the registry-free method:
   **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, add the folder's network path (`\\YOURPC\share`), tick **Show in Menu**, OK, restart Excel.
2. **Insert → My Add-ins → Shared Folder → Excel AI → Add**.

An **Excel AI** button appears on the Home ribbon.

## 3. Pair and add your models

1. Click **Excel AI** on the ribbon — the task pane opens and reports "✓ Sidecar detected".
2. Enter the **pairing code** from the sidecar console. This is one-time: the pairing
   survives sidecar restarts (tokens live in `~/.excelai`). You'll only re-pair if you
   delete that folder — the pane detects this and returns to the pairing screen, and a
   **Re-pair** button is always available in the connection banner.
3. Settings opens automatically. You can save **multiple models** and switch any time:
   - **Ollama (free, local):** choose *OpenAI-compatible* → **Use Ollama defaults** →
     **List available models** shows exactly what's installed → pick one → **Add model**.
   - **Anthropic / OpenAI:** paste your API key, click **List available models** (or type
     a model id), **Add model**.
4. Each saved model is a card showing provider, endpoint and masked key, with
   **Use** (make active), **Test** (live round-trip — expect `✓ "OK"`), and **Delete**.
   The **active model** is marked on its card and shown as a chip in the chat header.

### Is Ollama ready out of the box?

Almost — two commands once: install Ollama from [ollama.com](https://ollama.com), then
`ollama pull llama3.1` (any tool-capable model works: `llama3.1`, `qwen3`, `mistral-nemo`).
The desktop app keeps the server running automatically. No CORS/`OLLAMA_ORIGINS` setup is
needed — the sidecar talks to Ollama server-to-server, not from the browser. **List
available models** in Settings doubles as the readiness check: if your installed models
appear, you're ready.

## 4. Try it

- Put numbers in `A1:A2`, then ask: *"Total A1:A2 into A3 and label it in B3."*
  The agent reads the sheet, then proposes a **change-set** — review the items
  (each has a reason), click **Apply**, and watch `=SUM(A1:A2)` land. **Undo** reverts it.
- In any cell: `=EXCELAI.AI("One word: the capital of France")`
  or `=EXCELAI.AI_CLASSIFY(A1, "positive,negative,neutral")`.
  Identical formulas are cached — filling a column of repeated prompts bills once.

## Troubleshooting

| Symptom | Fix |
|---|---|
| Task pane: "Sidecar not detected" | Is `pnpm dev:sidecar` running? Port 8923 free? Set `EXCELAI_PORT` and edit the manifest URLs if you must change it. |
| "Invalid or already-used pairing code" | Codes are single-use. Restart the sidecar for a fresh code (existing pairings stay valid). |
| Test: `✗ … 401` | Wrong/expired API key. Delete the model card and re-add with the right key. |
| Test: `✗ … unsupported parameter` | Should not happen — the OpenAI adapter auto-adapts `max_tokens`/`temperature` for newer models. If you still see it, report the exact message. |
| Test times out (Ollama) | Is Ollama running? Does `ollama list` show your model? Base URL must end in `/v1`. Use **List available models** to verify. |
| Chat: "No model configured" | Open Settings, add a model (the first one becomes active automatically). |
| `=EXCELAI.AI` shows `#NAME?` | The add-in isn't loaded in this workbook — open the task pane once, then re-enter the formula. |
| `=EXCELAI.AI` shows `#VALUE!` with "Not paired" | Open the task pane and complete pairing first. |
| Upload rejected / pane blank on another machine | The manifest points at `http://localhost:8923` — sidecar and Excel must run on the **same machine**. |
| Excel for Mac | Not yet supported: WKWebView blocks `http://localhost` mixed content. The sidecar supports TLS (`EXCELAI_TLS_CERT`/`EXCELAI_TLS_KEY`, e.g. via `npx office-addin-dev-certs install`); switch the manifest URLs to `https://localhost:8923`. Tracked in the design risks. |

Everything local: delete `~/.excelai` to reset pairing, keys, cache and memory.
