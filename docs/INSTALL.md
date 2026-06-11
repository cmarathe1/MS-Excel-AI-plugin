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

## 3. Pair and pick a model

1. Click **Excel AI** on the ribbon — the task pane opens and reports "✓ Sidecar detected".
2. Enter the **pairing code** from the sidecar console (one-time).
3. Settings opens automatically. Pick a provider:
   - **Ollama (free, local):** choose *OpenAI-compatible*, click **Use Ollama defaults**
     (`http://localhost:11434/v1`, model `llama3.1`). Run `ollama pull llama3.1` first.
   - **Anthropic:** model e.g. `claude-sonnet-4-6`, paste your API key.
   - **OpenAI:** model e.g. `gpt-5.2`, paste your API key.
4. Click **Test connection** — you should see `✓ Connected — model replied: "OK"`.
5. **Save**, switch to **Chat**.

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
| Test connection: `✗ … 401` | Wrong/expired API key. Re-paste it and test again. |
| Test connection times out (Ollama) | Is `ollama serve` running? Does `ollama list` show your model? Base URL must end in `/v1`. |
| Chat: "No model configured" | Open Settings, save a provider. |
| `=EXCELAI.AI` shows `#NAME?` | The add-in isn't loaded in this workbook — open the task pane once, then re-enter the formula. |
| `=EXCELAI.AI` shows `#VALUE!` with "Not paired" | Open the task pane and complete pairing first. |
| Upload rejected / pane blank on another machine | The manifest points at `http://localhost:8923` — sidecar and Excel must run on the **same machine**. |
| Excel for Mac | Not yet supported: WKWebView blocks `http://localhost` mixed content. The sidecar supports TLS (`EXCELAI_TLS_CERT`/`EXCELAI_TLS_KEY`, e.g. via `npx office-addin-dev-certs install`); switch the manifest URLs to `https://localhost:8923`. Tracked in the design risks. |

Everything local: delete `~/.excelai` to reset pairing, keys, cache and memory.
