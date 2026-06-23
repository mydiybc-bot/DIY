# CLAUDE.md

Guidance for AI assistants (and humans) working in this repository.

## What this project is

A self-hosted **employee dialog training site** for a Taiwanese dessert/bakery
chain ("DIY"), plus a set of internal **business dashboards** served from the
same lightweight Python web server. The UI text is almost entirely in
**Traditional Chinese (zh-TW)** — keep new user-facing strings and error
messages in Traditional Chinese to match the existing tone.

Core feature: employees log in, pick a training "section" (單元), and answer
customer-service questions in free text or voice. Each answer is scored 0–10,
primarily by an LLM with a deterministic rule-based fallback. Admins manage the
question bank / accounts and export reports to Excel.

There is **no framework** — the server is built directly on Python's stdlib
`http.server`. There are no third-party web/ORM dependencies; persistence is
plain JSON files written atomically.

## Running & developing

```bash
# Install deps (stdlib does most of the work; these cover Excel + Google APIs)
pip install -r requirements.txt

# Start the server (defaults to 0.0.0.0:8765, override with PORT)
python3 server.py
# -> http://127.0.0.1:8765
```

`start_training_site.command` is a macOS launcher with a hardcoded local path —
ignore it in this environment.

### Tests

```bash
# Pure unit tests — no server, no network needed
python3 -m unittest test_training_logic -v

# Integration tests — REQUIRE a running server on :8765 and hit live endpoints
python3 server.py &        # in another shell
python3 -m unittest test_integration -v
```

Notes:
- `test_training_logic.py` is the reliable, hermetic suite. Run it after any
  change to `training_logic.py`, `data_store.py`, or `excel_tools.py`.
- `test_integration.py` contains a **hardcoded macOS path**
  (`/Users/diybc/Desktop/...`) and assumes a live server; it will not pass
  as-is in CI/cloud. Treat it as a local smoke test, not a gate.
- `test_every_reference_answer_scores_ten` calls `score_answer`, which tries
  the LLM first. With no `OPENAI_API_KEY` set it transparently falls back to
  rule-based scoring, so the test still runs offline.

## Environment variables

| Variable | Used by | Purpose |
|---|---|---|
| `PORT` | `server.py` | Listen port (default `8765`). |
| `TRAINING_DATA_DIR` | `storage_paths.py` | Directory for all JSON state. Unset = repo root (ephemeral). Set to a persistent disk in production. |
| `OPENAI_API_KEY` | `training_logic.py`, `voice_transcription.py` | LLM answer scoring (GPT-4o-mini) and audio transcription (Whisper). Without it, scoring falls back to rules and voice replies fail gracefully. |
| `OPENAI_TRANSCRIBE_MODEL` / `OPENAI_TRANSCRIBE_LANGUAGE` | `voice_transcription.py` | Optional overrides (default `gpt-4o-mini-transcribe`, `zh`). |
| `ANTHROPIC_API_KEY` | `recipe_auth.py` | Backs the "recipe studio" Claude proxy (`/api/recipe/claude`). |
| `RECIPE_STUDIO_PASSWORD` | `recipe_auth.py` | Password gate for the recipe studio; if unset, recipe login is always denied. |
| `REIMBURSEMENT_SOURCE_DIR` | `reimbursement_dashboard.py` | Directory of `.xls` reimbursement files (defaults to a macOS path). |

## Architecture

### Request flow

`server.py` defines a single `TrainingHandler(BaseHTTPRequestHandler)` served by
a `ThreadingHTTPServer`. There is no router abstraction — `do_GET` / `do_POST`
are long `if parsed.path == ...:` chains. To add an endpoint, add a branch in
the appropriate method and reuse the existing helpers:

- `self.read_json()` / `read_body_bytes()` / `read_multipart_form()` for input.
- `self.send_json(payload, status=, cookies=, expire_cookies=)` for output (sets
  `no-store` cache headers and `HttpOnly` cookies).
- `self._require_auth(role=...)` returns the session dict or sends a 401/403 and
  returns `None` — the caller must `if not session: return`.

### Sessions & auth

- **Training login** (employees/admin): cookie `training_session` -> in-memory
  `SESSIONS` dict. Lost on restart (acceptable; it only holds login state and
  the active training run). Credentials live in `auth_store.py`.
- **Recipe studio**: a completely separate token system in `recipe_auth.py`
  (header `X-Recipe-Token`, 4-hour fixed TTL, in-memory). Deliberately isolated
  from the employee auth — keep it that way.

### Persistence layer (`storage_paths.py` + `*_store.py`)

All durable state is JSON under `data_dir()` (== `TRAINING_DATA_DIR` or repo
root). Conventions to follow when touching storage:

- **Always** go through `atomic_write_json()` (temp file + `os.replace`) — never
  write JSON in place.
- **Always** serialize read-modify-write sequences under the global
  `FILE_LOCK` (a reentrant lock; the server is one process with many threads).
- Stores: `TrainingContentStore` (question bank + rules → `training_content.json`),
  `AuthStore` (`auth_config.json`), `ReportStore` (append-only
  `training_reports.json`), `ProgressStore` (per-employee resume state
  `training_progress.json`).
- Each store `load()`s defaults from `training_data.py` when its file is absent,
  and `validate()`s on read/write — validation raises `ValueError` with a
  Chinese message, which handlers turn into HTTP 400.
- **Optimistic concurrency**: admin saves use a revision counter
  (`get_rev`/`bump_rev` in `revisions.json`). `/api/admin/save-all` checks
  `base_rev` and returns HTTP 409 on conflict. Preserve this when editing admin
  write paths.

### Scoring engine (`training_logic.py`)

This is the heart of the app and the most likely thing you'll be asked to tune.

- `score_answer(question, answer, rules)` is the entry point: it tries
  `_score_with_ai` (OpenAI GPT-4o-mini, JSON-mode prompt built from the
  admin-configured rules) and **falls back to `_score_with_rules`** (Chinese
  keyword/tone/expression heuristics) on any error. Both return the same dict
  shape — keep them in sync if you add fields.
- `protect_improved_attempt(previous, current)` prevents the score from dropping
  when a retry clearly improved (more points hit / profanity removed) — a UX
  guarantee, not a bug.
- `respond(session, answer)` drives the conversation: scores the answer, updates
  the per-question best score, reveals the reference answer after
  `max_attempts_before_answer` tries, advances on a 10/10, and ends on the
  configured `end_phrase`. Profanity caps the score at 2 (enforced server-side
  even if the LLM disagrees).
- All coaching/feedback wording is **driven by `rules`** (editable in the admin
  UI / `training_data.py`), not hardcoded. When changing tone, change the rules,
  not string literals in the logic.

### Other server modules

- `excel_tools.py` — round-trips the question bank to/from `.xlsx` (sheets
  `rules` + `questions`) and exports reports. Import re-validates through
  `TrainingContentStore.validate`.
- `voice_transcription.py` — multipart audio → OpenAI transcription; raises
  `TranscriptionError` (→ HTTP 502).
- `self_dashboard.py` — reads a large cached JSON snapshot
  (`data/self_dashboard_snapshot.json`, ~14 MB) of customer purchase data; falls
  back to parsing a source `.xlsx` if the snapshot is absent.
- `reimbursement_dashboard.py` — parses `.xls` reimbursement files from a
  directory into a dashboard payload.

### Frontend

Plain HTML/CSS/JS, no build step. Served from `static/` via `serve_file`. Key
routes: `/` (index), `/dashboard`, `/pnl`, `/self-dashboard`, `/google-reviews`,
and any `static/<file>`. The `static/` directory also holds many standalone
dashboard pages (`pos-dashboard*.html`, `dashboard-schedule*.html`,
`recipe-studio.html`, etc.). Note there are **two** `index.html` files: the repo
root one and `static/index.html` — the server serves the one in `static/`.

## Conventions

- **Python**: every module starts with `from __future__ import annotations` and
  uses modern type hints (`dict | None`). Match that. Stdlib-first — avoid adding
  web frameworks or heavy deps; the only third-party libs are for Excel/Google
  (`openpyxl`, `xlrd`, `google-api-python-client`) and `certifi`/`beautifulsoup4`.
- **User-facing strings stay Traditional Chinese.** Validation errors, API error
  messages, and UI copy are all zh-TW.
- **Errors**: raise `ValueError` with a Chinese message in validation/logic;
  handlers convert to `{"ok": false, "error": ...}` with an appropriate status.
- **JSON responses** use `ensure_ascii=False` so Chinese is human-readable.
- Don't print to stdout from request handling (`log_message` is silenced
  deliberately); use `stderr` for diagnostics (see the AI-fallback log line).

## Deployment

Configured for **Render** via `render.yaml` (see `RENDER_DEPLOY.md`, in Chinese):
build `pip install -r requirements.txt`, start `python3 server.py`, Python
3.13.2. The free plan has **no persistent disk** — set `TRAINING_DATA_DIR` to a
mounted Persistent Disk (`/var/data/employee-training`) so accounts, the
question bank, reports, and progress survive restarts. `OPENAI_API_KEY` (and, if
using the recipe studio, `ANTHROPIC_API_KEY` / `RECIPE_STUDIO_PASSWORD`) must be
set in the Render environment.

## Gotchas

- **State is ephemeral by default.** Without `TRAINING_DATA_DIR`, JSON files land
  in the repo root and are lost on redeploy. `.gitignore` already excludes
  `training_reports.json`, `training_progress.json`, `auth_config.json`.
- Default credentials live in `training_data.py` (`ADMIN_PASSWORD = "admin123"`,
  employee `service123`). Change in production; don't commit real ones.
- `data/self_dashboard_snapshot.json` is large (~14 MB) and committed — be
  mindful when editing/regenerating.
- Several modules default to **hardcoded macOS paths** (dashboards, the
  `.command` launcher, `test_integration.py`). Override via env vars rather than
  editing the defaults when running elsewhere.
- The server is single-process/multithreaded; correctness of concurrent writes
  relies entirely on `FILE_LOCK` + atomic writes. Honor both.
```
