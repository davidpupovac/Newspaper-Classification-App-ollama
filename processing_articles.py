from typing import List, Union
import requests
import time
import ollama
import json
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import asyncio
from pathlib import Path
import http.server
import socketserver
import threading
import webbrowser
import socket

# -----------------------------------------------------------------
# Load data

PAGE_TITLE = "Text Extractor"
DEFAULT_PORT = 8765

HTML_PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__PAGE_TITLE__</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  :root {
    --ink: #1c1b18;
    --paper: #f7f5f0;
    --rule: #d8d3c5;
    --accent: #8a3b2b;
    --accent-soft: #f1e3dc;
    --mono: "IBM Plex Mono", "SFMono-Regular", Consolas, monospace;
    --serif: "Source Serif 4", Georgia, serif;
  }

  * { box-sizing: border-box; }

  body {
    margin: 0;
    min-height: 100vh;
    background: var(--paper);
    color: var(--ink);
    font-family: var(--serif);
    display: flex;
    justify-content: center;
    padding: 48px 20px;
  }

  main {
    width: 100%;
    max-width: 640px;
  }

  .eyebrow {
    font-family: var(--mono);
    font-size: 12px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--accent);
    margin: 0 0 8px;
  }

  h1 {
    font-size: 28px;
    font-weight: 600;
    margin: 0 0 6px;
    letter-spacing: -0.01em;
  }

  .sub {
    font-size: 15px;
    color: #5b5648;
    margin: 0 0 32px;
    line-height: 1.5;
  }

  .panel {
    border: 1px solid var(--rule);
    border-radius: 3px;
    background: #fff;
    padding: 28px;
    margin-bottom: 20px;
  }

  .field-label {
    font-family: var(--mono);
    font-size: 11px;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #8a8472;
    display: block;
    margin-bottom: 10px;
  }

  .drop {
    border: 1.5px dashed var(--rule);
    border-radius: 3px;
    padding: 28px 16px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.15s ease, background 0.15s ease;
  }

  .drop:hover, .drop.drag {
    border-color: var(--accent);
    background: var(--accent-soft);
  }

  .drop p {
    margin: 0;
    font-size: 14px;
    color: #5b5648;
  }

  .drop .filename {
    font-family: var(--mono);
    font-size: 13px;
    color: var(--ink);
    margin-top: 6px;
    word-break: break-all;
  }

  input[type="file"] { display: none; }

  .row {
    display: flex;
    align-items: center;
    gap: 14px;
    margin-top: 24px;
  }

  input[type="number"] {
    width: 110px;
    font-family: var(--mono);
    font-size: 15px;
    padding: 9px 10px;
    border: 1px solid var(--rule);
    border-radius: 3px;
    background: var(--paper);
    color: var(--ink);
  }

  input[type="number"]:focus {
    outline: 2px solid var(--accent);
    outline-offset: 1px;
  }

  .count-hint {
    font-family: var(--mono);
    font-size: 12px;
    color: #8a8472;
  }

  select {
    font-family: var(--mono);
    font-size: 14px;
    padding: 8px 10px;
    border: 1px solid var(--rule);
    border-radius: 3px;
    background: var(--paper);
    color: var(--ink);
  }

  button {
    font-family: var(--mono);
    font-size: 13px;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    background: var(--ink);
    color: var(--paper);
    border: none;
    border-radius: 3px;
    padding: 12px 22px;
    cursor: pointer;
    margin-top: 24px;
  }

  button:disabled {
    background: #c9c4b6;
    cursor: not-allowed;
  }

  button:not(:disabled):hover {
    background: var(--accent);
  }

  .status {
    font-family: var(--mono);
    font-size: 13px;
    border-radius: 3px;
    padding: 14px 16px;
    margin-top: 4px;
    line-height: 1.6;
    white-space: pre-wrap;
  }

  .status.ok {
    background: #ecf3ea;
    color: #355c30;
    border: 1px solid #c3dabb;
  }

  .status.err {
    background: #f6e7e3;
    color: #8a3b2b;
    border: 1px solid #e0bdb1;
  }

  .status.idle {
    background: var(--paper);
    color: #8a8472;
    border: 1px solid var(--rule);
  }

  .col-row {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 18px;
  }

  .preview {
    margin-top: 18px;
    border-top: 1px solid var(--rule);
    padding-top: 16px;
  }

  .preview-item {
    font-size: 13px;
    color: #44403a;
    padding: 6px 0;
    border-bottom: 1px dotted var(--rule);
  }

  .preview-item:last-child { border-bottom: none; }

  .preview-item .idx {
    font-family: var(--mono);
    color: #8a8472;
    margin-right: 8px;
  }

  footer {
    font-family: var(--mono);
    font-size: 11.5px;
    color: #9a9684;
    text-align: center;
    margin-top: 28px;
    line-height: 1.6;
  }

  code {
    background: var(--accent-soft);
    padding: 1px 5px;
    border-radius: 2px;
    font-size: 0.95em;
  }
</style>
</head>
<body>
<main>
  <p class="eyebrow">Data prep</p>
  <h1>__PAGE_TITLE__</h1>
  <p class="sub">Load a spreadsheet, choose a column, and pick how many rows to keep. When you click Extract, the result is sent back to your Python session as <code>news_articles</code>.</p>

  <div class="panel">
    <label class="field-label" for="file-input">Spreadsheet</label>
    <div class="drop" id="drop-zone">
      <p>Click to choose a file, or drag one here</p>
      <p class="sub" style="margin:4px 0 0;font-size:12.5px;">.xlsx, .xls, or .csv</p>
      <p class="filename" id="filename"></p>
    </div>
    <input type="file" id="file-input" accept=".xlsx,.xls,.csv">

    <div class="col-row">
      <div>
        <label class="field-label" for="column-select">Column to extract</label>
        <select id="column-select" disabled>
          <option>Load a file first</option>
        </select>
      </div>
    </div>

    <div class="row">
      <div>
        <label class="field-label" for="row-count">Rows to keep</label>
        <input type="number" id="row-count" min="1" value="500" disabled>
      </div>
      <span class="count-hint" id="row-hint">— rows available</span>
    </div>

    <button id="extract-btn" disabled>Extract</button>
  </div>

  <div class="status idle" id="status">Waiting for a spreadsheet.</div>

  <div class="preview" id="preview" style="display:none;">
    <p class="field-label">Preview (first 5 of selection)</p>
    <div id="preview-list"></div>
  </div>

  <footer id="footer-note">Click Extract to send this column back to your Python session.</footer>
</main>

<script>
  const fileInput = document.getElementById('file-input');
  const dropZone = document.getElementById('drop-zone');
  const filenameEl = document.getElementById('filename');
  const columnSelect = document.getElementById('column-select');
  const rowCountInput = document.getElementById('row-count');
  const rowHint = document.getElementById('row-hint');
  const extractBtn = document.getElementById('extract-btn');
  const statusEl = document.getElementById('status');
  const previewEl = document.getElementById('preview');
  const previewList = document.getElementById('preview-list');
  const footerNote = document.getElementById('footer-note');

  let rows = [];          // full sheet as array of objects
  let columns = [];

  function setStatus(kind, msg) {
    statusEl.className = 'status ' + kind;
    statusEl.textContent = msg;
  }

  dropZone.addEventListener('click', () => fileInput.click());

  ['dragover', 'dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, e => {
      e.preventDefault();
      if (evt === 'dragover') dropZone.classList.add('drag');
      else dropZone.classList.remove('drag');
    });
  });

  dropZone.addEventListener('drop', e => {
    const f = e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  fileInput.addEventListener('change', e => {
    const f = e.target.files[0];
    if (f) handleFile(f);
  });

  function handleFile(file) {
    filenameEl.textContent = file.name;
    setStatus('idle', 'Reading file…');
    resetDownstream();

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

        if (!rows.length) {
          setStatus('err', 'Error: the sheet appears to be empty.');
          return;
        }

        columns = Object.keys(rows[0]);
        populateColumnSelect(columns);

        rowCountInput.disabled = false;
        rowCountInput.max = rows.length;
        rowCountInput.value = Math.min(500, rows.length);
        rowHint.textContent = `of ${rows.length} rows available`;

        extractBtn.disabled = false;
        setStatus('ok', `Loaded "${sheetName}" — ${rows.length} rows, ${columns.length} columns.`);
      } catch (err) {
        setStatus('err', 'Error loading spreadsheet: ' + err.message);
      }
    };
    reader.onerror = () => setStatus('err', 'Error: could not read the file.');
    reader.readAsArrayBuffer(file);
  }

  function populateColumnSelect(cols) {
    columnSelect.innerHTML = '';
    columnSelect.disabled = false;

    // Prefer a column literally named "Message" if present, falling back to first column.
    let preferredIndex = cols.findIndex(c => c.trim().toLowerCase() === 'message');
    if (preferredIndex === -1) preferredIndex = 0;

    cols.forEach((col, i) => {
      const opt = document.createElement('option');
      opt.value = col;
      opt.textContent = col;
      if (i === preferredIndex) opt.selected = true;
      columnSelect.appendChild(opt);
    });
  }

  function resetDownstream() {
    previewEl.style.display = 'none';
    previewList.innerHTML = '';
  }

  extractBtn.addEventListener('click', async () => {
    const column = columnSelect.value;
    const n = parseInt(rowCountInput.value, 10);

    if (!column) {
      setStatus('err', 'Error: choose a column first.');
      return;
    }
    if (!rows.length) {
      setStatus('err', 'Error: no data loaded.');
      return;
    }
    if (!n || n < 1) {
      setStatus('err', 'Error: row count must be at least 1.');
      return;
    }
    if (!(column in rows[0])) {
      setStatus('err', `Error: column "${column}" not found. Available columns: ${columns.join(', ')}`);
      return;
    }

    const extracted = rows.slice(0, n).map(r => r[column]);

    extractBtn.disabled = true;
    extractBtn.textContent = 'Sending…';

    try {
      const resp = await fetch('/submit', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ column: column, values: extracted })
      });
      if (!resp.ok) throw new Error('Server responded with ' + resp.status);

      setStatus('ok',
        `Sent ${extracted.length} value(s) from column "${column}" back to Python.\n` +
        `You can return to your IDE — "news_articles" is ready there now.`
      );
      footerNote.textContent = 'Done. This tab can be closed; your Python session has the data.';

      previewList.innerHTML = '';
      extracted.slice(0, 5).forEach((val, i) => {
        const div = document.createElement('div');
        div.className = 'preview-item';
        div.innerHTML = `<span class="idx">${i}</span>${escapeHtml(String(val))}`;
        previewList.appendChild(div);
      });
      previewEl.style.display = extracted.length ? 'block' : 'none';
    } catch (err) {
      setStatus('err', 'Error sending data to Python: ' + err.message);
      extractBtn.disabled = false;
      extractBtn.textContent = 'Extract';
    }
  });

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
</script>
</body>
</html>
""".replace("__PAGE_TITLE__", PAGE_TITLE)


def _find_open_port(preferred):
    """Use the preferred port if free, otherwise let the OS pick one."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", preferred))
            return preferred
        except OSError:
            s.bind(("127.0.0.1", 0))
            return s.getsockname()[1]


def get_news_articles(port=DEFAULT_PORT, open_browser=True):
    """
    Opens the Text Extractor page and BLOCKS until the user clicks
    "Extract" in the browser.

    Returns:
        list: the extracted column values, as a plain Python list.
    """
    result_holder = {}
    result_event = threading.Event()

    class Handler(http.server.BaseHTTPRequestHandler):
        def log_message(self, format, *args):
            pass  # keep the console quiet

        def do_GET(self):
            body = HTML_PAGE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_POST(self):
            if self.path != "/submit":
                self.send_response(404)
                self.end_headers()
                return

            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length)
            try:
                payload = json.loads(raw.decode("utf-8"))
                values = payload.get("values", [])
            except (json.JSONDecodeError, AttributeError) as exc:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(str(exc).encode("utf-8"))
                return

            result_holder["news_articles"] = values
            result_event.set()

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"status": "ok"}')

    actual_port = _find_open_port(port)
    server = socketserver.TCPServer(("127.0.0.1", actual_port), Handler)
    url = f"http://127.0.0.1:{actual_port}/"

    server_thread = threading.Thread(target=server.serve_forever, daemon=True)
    server_thread.start()

    print(f"{PAGE_TITLE} running at {url}")
    print("Waiting for you to click 'Extract' in the browser...")

    if open_browser:
        webbrowser.open(url)

    try:
        result_event.wait()  # blocks here until /submit is hit
    except KeyboardInterrupt:
        print("\nCancelled before extraction.")
        server.shutdown()
        return []

    server.shutdown()
    news_articles = result_holder.get("news_articles", [])
    print(f"Received {len(news_articles)} value(s). Server stopped.")
    return news_articles


news_articles = get_news_articles()

# -------------------------------------------------------------------------------
# Translation

def translate_to_english(
    news_articles: List[str],
    src_langs: Union[List[str], str] = "auto",
    chunk_size: int = 30,
    ollama_url: str = "http://localhost:11434",
    model: str = "llama3.2",
    num_gpu: int = -1,
    num_thread: int = None,
) -> List[str]:

    print(f"Using Ollama model '{model}' at {ollama_url} | GPU layers: {'all' if num_gpu == -1 else num_gpu}")

    model_options = {"num_gpu": num_gpu}
    if num_thread is not None:
        model_options["num_thread"] = num_thread

    def split_into_chunks(text, chunk_size):
        words = text.split()
        return [" ".join(words[i:i + chunk_size]) for i in range(0, len(words), chunk_size)]

    def ollama_translate(text: str, src_lang: str) -> str:
        lang_hint = f" (source language: {src_lang})" if src_lang != "auto" else ""
        prompt = (
            f"Translate the following text to English{lang_hint}. "
            f"Return ONLY the translated text, no explanations, no quotes:\n\n{text}"
        )
        response = requests.post(
            f"{ollama_url}/api/generate",
            json={
                "model": model,
                "prompt": prompt,
                "stream": False,
                "options": model_options,
            },
            timeout=120,
        )
        response.raise_for_status()
        return response.json()["response"].strip()

    def translate_text(text: str, lang: str) -> str:
        words = text.split()
        if len(words) <= chunk_size:
            return ollama_translate(text, lang)
        chunks = split_into_chunks(text, chunk_size)
        return " ".join(ollama_translate(chunk, lang) for chunk in chunks)

    try:
        # Resolve src_langs to a list
        if isinstance(src_langs, str):
            if src_langs == "auto":
                from langdetect import detect, LangDetectException
                src_langs_list = []
                for text in news_articles:
                    try:
                        src_langs_list.append(detect(text))
                    except LangDetectException:
                        src_langs_list.append("sr")
            else:
                src_langs_list = [src_langs] * len(news_articles)
        else:
            src_langs_list = src_langs

        translated_news_articles = []
        total = len(news_articles)
        total_start = time.perf_counter()

        for i, (news_article, lang) in enumerate(zip(news_articles, src_langs_list)):
            item_start = time.perf_counter()
            try:
                if lang == "en":
                    translated_news_articles.append(news_article)
                    elapsed = time.perf_counter() - item_start
                    print(f"[{i+1}/{total}] Skipped (English) - {elapsed:.3f}s")
                else:
                    translated = translate_text(news_article, lang)
                    translated_news_articles.append(translated)
                    elapsed = time.perf_counter() - item_start
                    print(f"[{i+1}/{total}] '{lang}' -> 'en' - {elapsed:.3f}s")

            except requests.RequestException as e:
                elapsed = time.perf_counter() - item_start
                print(f"[{i+1}/{total}] Request failed ({elapsed:.3f}s): {e} - keeping original")
                translated_news_articles.append(news_article)
            except Exception as e:
                elapsed = time.perf_counter() - item_start
                print(f"[{i+1}/{total}] Error ({elapsed:.3f}s): {e} - keeping original")
                translated_news_articles.append(news_article)

        total_elapsed = time.perf_counter() - total_start
        avg = total_elapsed / total if total > 0 else 0
        print(f"\nDone: {len(translated_news_articles)} news_articles | "
              f"Total: {total_elapsed:.2f}s | Avg: {avg:.3f}s/item")

        return translated_news_articles

    except Exception as e:
        print(f"Error during translation: {e}")
        return []


english_news_articles = translate_to_english(news_articles)

# -------------------------------------------------------------------------------
# Summarization

# -- Prompt  ------------------------------------------------------------
def build_prompt(news_article: str) -> str:
    return f"""You are a media analyst. Analyze the following news news_article and produce a concise summary (2–3 sentences) that captures:

Core Event: What is actually happening?
Tone: Is the framing alarming, celebratory, neutral, sympathetic, critical, hopeful, etc.? Be specific.
Narrative Angle: What story is being told — who are the heroes/villains/victims implied? What emotional or ideological lens is used?
Humanitarian Angle:

    a) Civilian harm and protection of civilians — who is affected and how; any reported casualties, injuries, or protection concerns?
    b) Attacks on healthcare, essential services and critical infrastructure — any mentions of hospitals, water, power, schools, or other essential services being damaged, disrupted, or targeted?
    c) Humanitarian access and operational constraints — are aid organizations able to operate? Any blockages, restrictions, security risks, or logistical barriers mentioned?
    d) Displacement and regional spillover — is there mention of people fleeing, being displaced, or effects spreading beyond the immediate area (e.g., refugee flows, cross-border impacts)?
    e) International humanitarian law concerns — any mentions of Geneva Conventions, POWs, detainees, the wounded, or other IHL-relevant issues?
    f) Response and gaps — how are humanitarian organizations or governments responding, and what gaps in aid, protection, or support are evident or implied?
    g) What are the immediate and long-term human consequences (e.g., displacement, health, access to food, safety)?

    Trend Tags: Close with a short sentence naming, in plain prose, the theme(s) this article relates to (from: civilian harm, infrastructure/essential services, humanitarian access, displacement/spillover, or other), the location(s) involved, and any key voice(s) quoted or referenced (e.g., named officials, agencies, witnesses).

News_article: "{news_article}"
Write the summary in plain prose. Do not use bullet points or headers. Keep it under 150 words.

Do not use intros such as: 'Here's a concise summary:' """


# -- Summarizer ----------------------------------------------------------------
def summarize_news_article(news_article: str, model: str = "llama3.2") -> str:
    """Send a single news_article to Ollama and return the summary string."""
    response = ollama.chat(
        model=model,
        messages=[
            {"role": "user", "content": build_prompt(news_article)},
        ],
    )
    return response["message"]["content"].strip()


def summarize_all(news_articles: list[str], model: str = "llama3.2") -> tuple[list[str], float]:
    """
    Iterate over a list of news_articles and return a parallel list of summaries.
    Each summary reflects the tone and narrative framing of its news article.
    """
    summaries = []
    t_start = time.perf_counter()

    for i, news_article in enumerate(news_articles, start=1):
        t0 = time.perf_counter()
        summary = summarize_news_article(news_article, model=model)
        summaries.append(summary)
        print(f"[{i}/{len(news_articles)}] {time.perf_counter() - t0:.2f}s")

    duration = time.perf_counter() - t_start
    print(f"Total: {duration:.2f}s")
    return summaries, duration


summaries, duration = summarize_all(english_news_articles)

# -------------------------------------------------------------------------------
# Named entity recognition

ENTITY_TYPES = ["PERSON", "ORGANIZATION", "LOCATION", "DATE", "EVENT", "PRODUCT", "MISC"]

SYSTEM_PROMPT = """You are a Named Entity Recognition (NER) expert.
Extract entities from the given text and return ONLY a valid JSON object.
The JSON must have these exact keys: PERSON, ORGANIZATION, LOCATION, DATE, EVENT, PRODUCT, MISC.
Each value must be a list of strings (entity mentions found in the text).
If no entities of a type are found, use an empty list [].
Do not include any explanation or text outside the JSON object."""

USER_PROMPT_TEMPLATE = """Extract all named entities from the following article text.

Article:
{article}

Return ONLY a JSON object with keys: PERSON, ORGANIZATION, LOCATION, DATE, EVENT, PRODUCT, MISC."""


def extract_entities(article: str, model: str = "llama3.2") -> dict:
    """Extract named entities from a single article using Ollama."""
    response = ollama.chat(
        model=model,
        messages=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": USER_PROMPT_TEMPLATE.format(article=article)},
        ],
        options={"temperature": 0},
    )
    raw = response["message"]["content"].strip()

    # Strip markdown code fences if present
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]

    entities = json.loads(raw.strip())

    # Normalise: guarantee all expected keys exist
    return {key: entities.get(key, []) for key in ENTITY_TYPES}


def run_ner_on_articles(
    articles: list[str],
    model: str = "llama3.2",
    verbose: bool = True,
) -> tuple[list[dict], float]:
    """
    Run NER on a list of article strings.

    Args:
        articles : list of article texts to process.
        model    : Ollama model name (default "llama3.2").
        verbose  : print per-article progress when True.

    Returns:
        results  : list[dict] - one entity dict per article,
                   keys: PERSON, ORGANIZATION, LOCATION, DATE, EVENT, PRODUCT, MISC.
        duration : total wall-clock execution time in seconds.
    """
    results: list[dict] = []
    start = time.perf_counter()

    for i, article in enumerate(articles, 1):
        t0 = time.perf_counter()
        try:
            entities = extract_entities(article, model=model)
        except (json.JSONDecodeError, KeyError) as exc:
            entities = {key: [] for key in ENTITY_TYPES}
            if verbose:
                print(f"  [Article {i}] Parse error: {exc}")

        elapsed = time.perf_counter() - t0
        results.append(entities)

        if verbose:
            total_found = sum(len(v) for v in entities.values())
            print(f"  [Article {i}/{len(articles)}] {total_found} entities found  ({elapsed:.2f}s)")

    duration = time.perf_counter() - start
    return results, duration


def print_summary(results: list[dict], duration: float) -> None:
    """Print a summary of NER results and timing statistics."""
    total_entities = sum(
        len(v) for doc in results for v in doc.values()
    )
    avg_per_article = total_entities / len(results) if results else 0
    avg_time = duration / len(results) if results else 0

    print("\n" + "=" * 50)
    print("NER SUMMARY")
    print("=" * 50)
    print(f"  Articles processed : {len(results)}")
    print(f"  Total entities     : {total_entities}")
    print(f"  Avg per article    : {avg_per_article:.1f}")
    print(f"  Total time         : {duration:.2f}s")
    print(f"  Avg time/article   : {avg_time:.2f}s")

    print("\nEntity type breakdown:")
    type_totals = {key: sum(len(doc[key]) for doc in results) for key in ENTITY_TYPES}
    for etype, count in sorted(type_totals.items(), key=lambda x: -x[1]):
        bar = "!" * min(count, 40)
        print(f"  {etype:<14} {count:>4}  {bar}")
    print("=" * 50)


named_entities = run_ner_on_articles(english_news_articles)

# -------------------------------------------------------------------------------
# Classification - based on summaries and named entities

ENTITY_CATEGORIES = ["PERSON", "ORGANIZATION", "LOCATION", "DATE", "EVENT", "PRODUCT", "MISC"]


def classify_summaries(
    summaries: list[str],
    named_entities: tuple[list[dict], list],
    model: str = "llama3.2"
) -> dict:
    """
    Unsupervised classification of summaries using Llama 3.2.
    Incorporates named entity data alongside summaries for richer classification.

    Args:
        summaries:      List of text summaries.
        named_entities: Tuple of two lists.
                        named_entities[0] - list of dicts (one per summary), each with
                                            optional keys: PERSON, ORGANIZATION, LOCATION,
                                            DATE, EVENT, PRODUCT, MISC (values: lists of strings).
                        named_entities[1] - second list (reserved / passed through as-is).
        model: Ollama model name.

    Returns:
        {
            "categories": ["Category A", ...],
            "classified": [
                {
                    "summary": "...",
                    "named_entities": {"PERSON": [...], "LOCATION": [...], ...},
                    "categories": ["Category A"],
                    "reasoning": "...",
                    "duration_seconds": 1.23
                },
                ...
            ],
            "total_duration_seconds": 12.34
        }
    """
    total_start = time.perf_counter()

    # -- Unpack the tuple -------------------------------------------------------
    entity_dicts: list[dict] = named_entities[0]

    if len(summaries) != len(entity_dicts):
        raise ValueError(
            f"Length mismatch: {len(summaries)} summaries vs "
            f"{len(entity_dicts)} entity dicts in named_entities[0]."
        )

    # -- Helpers ------------------------------------------------------------

    def format_entities(entity_dict: dict) -> str:
        """Render an entity dict as a compact readable string for prompts."""
        parts = []
        for cat in ENTITY_CATEGORIES:
            values = entity_dict.get(cat)
            if values:
                parts.append(f"{cat}: {', '.join(str(v) for v in values)}")
        return " | ".join(parts) if parts else "none"

    def build_corpus() -> str:
        lines = []
        for i, (s, e) in enumerate(zip(summaries, entity_dicts)):
            lines.append(f"{i+1}. Summary: {s}")
            lines.append(f"   Entities: {format_entities(e)}")
        return "\n".join(lines)

    corpus = build_corpus()

    # -- Step 1: Infer categories from the full corpus --------------------------
    category_prompt = f"""You are a text classification expert. Read the following summaries and their named entities, then infer a minimal set of meaningful thematic categories that best organizes them.

Rules:
- Propose between 3 and 8 categories
- Categories should be mutually distinct but may overlap in edge cases
- Base categories on both the content AND the named entities provided
- Named entities are grouped as: {', '.join(ENTITY_CATEGORIES)}
- Name each category with 2-4 words

Summaries with Named Entities:
{corpus}

Respond ONLY with a JSON array of category name strings. Example: ["Economic Policy", "Climate Science", "Human Interest"]
No explanation, no markdown."""

    resp1 = ollama.chat(model=model, messages=[{"role": "user", "content": category_prompt}])
    raw1 = resp1["message"]["content"].strip()
    raw1 = re.sub(r"^```(?:json)?\s*", "", raw1)
    raw1 = re.sub(r"\s*```$", "", raw1)

    try:
        categories = json.loads(raw1)
    except json.JSONDecodeError:
        categories = ["Uncategorized"]

    print(f"  Inferred {len(categories)} categories: {categories}\n")

    # -- Step 2: Assign each summary individually, measuring duration per summary
    classified = []
    for i, (summary, entities) in enumerate(zip(summaries, entity_dicts)):
        summary_start = time.perf_counter()

        single_corpus = f"1. Summary: {summary}\n   Entities: {format_entities(entities)}"

        classify_prompt = f"""You are a text classifier. Using ONLY the categories listed below, classify the summary below.

Categories: {json.dumps(categories)}

Rules:
- Assign 1-2 categories (use the most relevant)
- Use both the summary text AND the named entities to inform your decision
- Named entity types available: {', '.join(ENTITY_CATEGORIES)}
- Include a one-sentence reasoning for the assignment
- Return ONLY a JSON object (not an array)

The object must have:
  "categories": [list of assigned category strings]
  "reasoning": "one sentence"

Summary with Named Entities:
{single_corpus}

Respond with ONLY the JSON object. No markdown, no extra text."""

        resp = ollama.chat(model=model, messages=[{"role": "user", "content": classify_prompt}])
        raw = resp["message"]["content"].strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)

        try:
            entry = json.loads(raw)
        except json.JSONDecodeError:
            entry = {"categories": ["Uncategorized"], "reasoning": "Parse error"}

        summary_duration = round(time.perf_counter() - summary_start, 3)

        clean_entities = {cat: entities.get(cat, []) for cat in ENTITY_CATEGORIES}

        classified.append({
            "summary": summary,
            "named_entities": clean_entities,
            "categories": entry.get("categories", ["Uncategorized"]),
            "reasoning": entry.get("reasoning", ""),
            "duration_seconds": summary_duration,
        })

        print(f"  Classified summary {i+1}/{len(summaries)} in {summary_duration}s")

    total_duration = round(time.perf_counter() - total_start, 3)
    print(f"\n  Total classification duration: {total_duration}s")

    return {
        "categories": categories,
        "classified": classified,
        "total_duration_seconds": total_duration,
    }


classification = classify_summaries(summaries, named_entities)

# -------------------------------------------------------------------------------
# Export excel

ENTITY_CATEGORIES = ["PERSON", "ORGANIZATION", "LOCATION", "DATE", "EVENT", "PRODUCT", "MISC"]


def export_to_excel(result: dict, output_path: str = "./classification_results.xlsx") -> str:
    """
    Exports the dict returned by classify_summaries() to a formatted Excel file.

    Columns:
        - Summary
        - One column per inferred category (tick if assigned)
        - One column per named entity category (comma-separated values)
        - Reasoning
        - Duration (s)

    Args:
        result:      The dict returned by classify_summaries().
        output_path: Destination .xlsx path.

    Returns:
        The resolved output path.
    """
    # 1. Get the absolute path of the directory where THIS script is located
    script_dir = Path(__file__).parent.absolute()

    # 2. Join that directory with your filename to get the full save path
    output_path = script_dir / Path(output_path)

    categories = result["categories"]
    classified = result["classified"]
    total_dur = result.get("total_duration_seconds", "")

    wb = Workbook()
    ws = wb.active
    ws.title = "Classification Results"

    # -- Styles -----------------------------------------------------------------
    HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    SUMMARY_FILL = PatternFill("solid", start_color="1F4E79")   # dark blue
    CATEGORY_FILL = PatternFill("solid", start_color="2E75B6")  # mid blue
    ENTITY_FILL = PatternFill("solid", start_color="375623")    # dark green
    META_FILL = PatternFill("solid", start_color="595959")      # grey
    WRAP = Alignment(wrap_text=True, vertical="top")
    CENTER = Alignment(horizontal="center", vertical="top")
    TICK_FONT = Font(name="Arial", color="2E75B6", bold=True, size=11)
    ALT_FILL = PatternFill("solid", start_color="EBF3FB")       # light blue tint

    # -- Build header row -------------------------------------------------------
    headers = []

    # Group 1 - Summary
    headers.append(("Summary", SUMMARY_FILL))

    # Group 2 - Inferred classification categories
    for cat in categories:
        headers.append((cat, CATEGORY_FILL))

    # Group 3 - Named entity categories
    for ent in ENTITY_CATEGORIES:
        headers.append((ent, ENTITY_FILL))

    # Group 4 - Meta
    headers.append(("Reasoning", META_FILL))
    headers.append(("Duration (s)", META_FILL))

    for col_idx, (label, fill) in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = HEADER_FONT
        cell.fill = fill
        cell.alignment = CENTER

    # -- Data rows --------------------------------------------------------------
    for row_idx, item in enumerate(classified, start=2):
        assigned_cats = set(item.get("categories", []))
        entities = item.get("named_entities", {})
        col = 1

        # Summary
        cell = ws.cell(row=row_idx, column=col, value=item["summary"])
        cell.alignment = WRAP
        col += 1

        # Category tick columns
        for cat in categories:
            cell = ws.cell(row=row_idx, column=col, value="✓" if cat in assigned_cats else "")
            cell.alignment = CENTER
            if cat in assigned_cats:
                cell.font = TICK_FONT
            col += 1

        # Named entity columns
        for ent in ENTITY_CATEGORIES:
            values = entities.get(ent, [])
            cell = ws.cell(row=row_idx, column=col, value=", ".join(values) if values else "")
            cell.alignment = WRAP
            col += 1

        # Reasoning
        cell = ws.cell(row=row_idx, column=col, value=item.get("reasoning", ""))
        cell.alignment = WRAP
        col += 1

        # Duration
        ws.cell(row=row_idx, column=col, value=item.get("duration_seconds", ""))

        # Alternating row background
        if row_idx % 2 == 0:
            for c in range(1, len(headers) + 1):
                ws.cell(row=row_idx, column=c).fill = ALT_FILL

    # -- Total duration row -----------------------------------------------------
    total_row = len(classified) + 2
    label_col = len(headers) - 1          # "Reasoning" column
    value_col = len(headers)              # "Duration (s)" column

    label_cell = ws.cell(row=total_row, column=label_col, value="Total Duration (s)")
    label_cell.font = Font(name="Arial", bold=True, size=10)
    label_cell.alignment = Alignment(horizontal="right", vertical="top")

    value_cell = ws.cell(row=total_row, column=value_col, value=total_dur)
    value_cell.font = Font(name="Arial", bold=True, size=10)
    value_cell.alignment = Alignment(horizontal="center", vertical="top")

    # -- Column widths ----------------------------------------------------------
    ws.column_dimensions[get_column_letter(1)].width = 60          # Summary
    for i in range(2, 2 + len(categories)):                        # Category cols
        ws.column_dimensions[get_column_letter(i)].width = 18
    for i in range(2 + len(categories), 2 + len(categories) + len(ENTITY_CATEGORIES)):
        ws.column_dimensions[get_column_letter(i)].width = 25      # Entity cols
    ws.column_dimensions[get_column_letter(len(headers) - 1)].width = 50   # Reasoning
    ws.column_dimensions[get_column_letter(len(headers))].width = 14      # Duration

    # Freeze header row
    ws.freeze_panes = "A2"

    wb.save(str(output_path))
    return str(output_path)


export_to_excel(classification)

# -------------------------------------------------------------------------------
# Summarization of summaries

"""
Humanitarian Communications Director - Batched Summary Analyzer
Uses Ollama + llama3.2 with async batching for large summary sets.

Strategy:
  1. Chunk summaries into batches (default: 20 per batch)
  2. Process all batches concurrently with asyncio + semaphore (controls parallelism)
  3. Merge all batch analyses in a final synthesis call with enhanced merging logic
"""

# -- Configuration -------------------------------------------------------------
BATCH_SIZE = 20        # Summaries per batch (modified to 20)
MAX_CONCURRENT = 3     # Max parallel Ollama calls (tune to your RAM/GPU)
MODEL = "llama3.2"
TEMPERATURE = 0.3
MAX_TOKENS = 2048

SCRIPT_DIR = Path(__file__).parent.absolute()

# -- Prompts -------------------------------------------------------------------
BATCH_SYSTEM_PROMPT = """
You are a seasoned Humanitarian Communications Director with 20+ years of experience
working with organizations like ICRC, OCHA, and MSF. You analyze conflict-related
summaries and produce sharp, actionable intelligence for humanitarian advocacy teams.

Your analysis must be structured, evidence-based, and strictly neutral.
Use precise humanitarian and IHL terminology.

You will receive a batch of summaries from the same crisis context.
Analyze them TOGETHER as a unified situation - do NOT treat them separately.
Synthesize patterns, compounding factors, and trends across the batch.

Produce a consolidated analysis covering ALL seven sections:

## 1. HUMANITARIAN NEEDS & RESPONSE
- Needs: food, water, shelter, displacement, protection, health.
- Crisis phase (acute emergency, protracted, chronic).
- Access constraints and operational risks.

## 2. IHL / RULES OF WAR
- Violations of: distinction, proportionality, precautions in attack, siege warfare,
  starvation as method of warfare, protected persons/objects.
- Cite instruments: Geneva Conventions, Additional Protocols, Rome Statute.
- Note calls for investigations, accountability, sanctions.

## 3. IMPACT ON CIVILIANS
- Casualties, displacement, infrastructure damage.
- Dominant narratives (urban warfare, human shields, indiscriminate attacks).
- Coverage gaps: children, women, elderly, disabled, refugees.

## 4. WOUNDED, DETAINEES, POWs & SHIPWRECKED
- Healthcare access, medical evacuations, attacks on medical facilities.
- Detainee/POW treatment, ICRC access, exchange processes.

## 5. SPEAKING OPPORTUNITIES
- Where a neutral humanitarian voice adds value.
- Recommended formats: press statement, op-ed, UN briefing, social media thread.

## 6. TRENDS IN REPORTING
- Identify 3-4 main trends in reporting
- Describe the trends
- Terminology (e.g. changes in how parties, casualties, or events are labeled).
- Early signals of narrative shifts that could affect humanitarian access, perception, or messaging risk.

## 7. COMMUNICATION STRATEGIES
### A. Proactive Media Engagement
- Target journalists/outlets by beat.
- Pitching: exclusives, data visuals, field access.
- Newsjacking opportunities.

### B. Storytelling Tactics
- Human story + data combinations.
- Problem-solution narrative arcs.
- Visual content: infographics, 30-sec videos.

### C. Channel Prioritization
| Channel           | Best Use Case                         | Recommended Content               |
|-------------------|----------------------------------------|------------------------------------|
| Traditional media | Op-eds, embargoed reports             | IHL explainers, field testimonies |
| X/Twitter         | Rapid response, myth-busting          | Short threads, quote cards        |
| LinkedIn          | Policy influence, donor engagement    | Situation reports, briefs         |
| Instagram/TikTok  | Public awareness, youth audiences     | Human stories, 30-sec videos      |
| Podcasts          | Deep-dive analysis                    | Expert interviews, context        |

Be specific and actionable. Every recommendation must tie directly to evidence in the batch.
"""

SYNTHESIS_SYSTEM_PROMPT = """
You are a Humanitarian Communications Director synthesizing multiple batch analyses
from the same crisis into one final, unified strategic assessment.

Merge the batch analyses below into a single coherent report.
Eliminate redundancy, surface cross-cutting patterns, and escalate the most critical findings.
Maintain all six section headers. Prioritize actionability over comprehensiveness.

The final output should read as one authoritative document, not a list of batch summaries.

When merging, ensure you:
1. **Deduplicate information** - combine similar findings across batches
2. **Surface trends** - identify patterns that span multiple batches
3. **Quantify impact** - aggregate numbers, statistics, and scale across all summaries
4. **Highlight escalation** - note if situations worsen across batches
5. **Create unified recommendations** - consolidate communication strategies
6. **Maintain coherence** - ensure smooth flow without disjointed sections

The synthesis should feel like one comprehensive analysis, not multiple reports stitched together.
"""


# -- Helpers -------------------------------------------------------------------
def extract_text(item) -> str:
    """Extract text from a dict {'values': ...} or a plain string."""
    if isinstance(item, dict):
        return item.get("values", "").strip()
    return str(item).strip()


def chunk(lst: list, size: int) -> list:
    """Split a list into chunks of given size."""
    return [lst[i:i + size] for i in range(0, len(lst), size)]


def build_batch_prompt(texts: list, batch_num: int, total: int) -> str:
    numbered = "\n\n".join(f"[{i+1}] {t}" for i, t in enumerate(texts))
    return (
        f"BATCH {batch_num} of {total} - {len(texts)} summaries.\n"
        "Analyze these together as one unified situation.\n\n"
        f"---\n{numbered}\n---"
    )


def build_synthesis_prompt(batch_analyses: list) -> str:
    """Enhanced synthesis prompt with merging instructions."""
    total_batches = len(batch_analyses)

    parts = []
    for i, analysis in enumerate(batch_analyses):
        parts.append(f"=== BATCH {i+1} ANALYSIS (of {total_batches}) ===\n{analysis}")

    combined_prompt = (
        f"The following are {total_batches} batch analyses from the same crisis context.\n"
        f"Total batches processed: {total_batches}\n\n"
        "**MERGING INSTRUCTION:**\n"
        "Synthesize these analyses into ONE final unified strategic assessment.\n"
        "DO NOT simply repeat each batch separately. Instead:\n"
        "1. Combine all findings from sections 1-6 across all batches\n"
        "2. Eliminate redundant information while keeping the most critical points\n"
        "3. Aggregate statistics and scale across all batches\n"
        "4. Identify cross-cutting patterns and trends\n"
        "5. Prioritize the most urgent humanitarian concerns\n"
        "6. Create consolidated, actionable recommendations\n\n"
        "The final output should flow as a single, coherent document.\n\n"
        f"{chr(10).join(parts)}"
    )

    return combined_prompt


# -- Async Ollama call ---------------------------------------------------------
async def call_ollama(system: str, user: str, label: str, retry_count: int = 2) -> str:
    """Run a single Ollama chat call in a thread with retry logic."""
    loop = asyncio.get_event_loop()

    def _call():
        try:
            response = ollama.chat(
                model=MODEL,
                messages=[
                    {"role": "system", "content": system},
                    {"role": "user", "content": user},
                ],
                options={"temperature": TEMPERATURE, "num_predict": MAX_TOKENS},
            )
            return response["message"]["content"]
        except Exception as e:
            print(f"  Error in {label}: {e}")
            raise

    for attempt in range(retry_count + 1):
        try:
            if attempt > 0:
                print(f"  Retry {attempt}/{retry_count} for {label}...")
                await asyncio.sleep(2)  # Wait before retry
            print(f"  Starting {label}...")
            result = await loop.run_in_executor(None, _call)
            print(f"  Done    {label}")
            return result
        except Exception:
            if attempt == retry_count:
                print(f"  Failed {label} after {retry_count + 1} attempts")
                raise
            continue


# -- Core pipeline -------------------------------------------------------------
async def process_batches(batches: list) -> list:
    """Process all batches concurrently, bounded by MAX_CONCURRENT."""
    semaphore = asyncio.Semaphore(MAX_CONCURRENT)
    total = len(batches)

    async def bounded_call(texts, idx):
        async with semaphore:
            return await call_ollama(
                system=BATCH_SYSTEM_PROMPT,
                user=build_batch_prompt(texts, idx + 1, total),
                label=f"Batch {idx + 1}/{total} ({len(texts)} summaries)",
            )

    tasks = [bounded_call(batch, i) for i, batch in enumerate(batches)]
    return await asyncio.gather(*tasks)


async def merge_batch_analyses(batch_analyses: list) -> str:
    """Enhanced merging function for batch analyses."""
    print(f"\n{'=' * 70}")
    print(f"  MERGING {len(batch_analyses)} BATCH ANALYSES")
    print(f"{'=' * 70}")

    merged = await call_ollama(
        system=SYNTHESIS_SYSTEM_PROMPT,
        user=build_synthesis_prompt(batch_analyses),
        label="Final synthesis and merging",
    )

    return merged


async def run_analysis(summaries: list) -> dict:
    # 1. Extract and validate texts
    texts = [t for item in summaries if (t := extract_text(item))]
    if not texts:
        raise ValueError("No valid summaries found.")

    total_batches = -(-len(texts) // BATCH_SIZE)  # Ceiling division

    print(f"\n{'=' * 70}")
    print(f"  {len(texts)} summaries -> {total_batches} batch(es) "
          f"(batch size={BATCH_SIZE}, concurrency={MAX_CONCURRENT})")
    print(f"{'=' * 70}\n")

    t0 = time.perf_counter()

    # 2. Chunk into batches
    batches = chunk(texts, BATCH_SIZE)

    # Store batch metadata
    batch_metadata = []
    for i, batch in enumerate(batches):
        batch_metadata.append({
            "batch_num": i + 1,
            "size": len(batch),
            "summaries": batch
        })

    # 3. Process all batches concurrently
    batch_analyses = await process_batches(batches)

    # 4. Merge all results into one consistent report
    print(f"\n{'=' * 70}")
    print(f"  MERGING PHASE: Creating unified report from {len(batch_analyses)} batches")
    print(f"{'=' * 70}\n")

    final_analysis = await merge_batch_analyses(batch_analyses)

    elapsed = time.perf_counter() - t0

    print(f"\n{'=' * 70}")
    print(f"  FINAL ANALYSIS (completed in {elapsed:.1f}s)")
    print(f"{'=' * 70}\n")
    print(final_analysis)

    return {
        "summaries_analyzed": len(texts),
        "batches": len(batches),
        "batch_size_used": BATCH_SIZE,
        "max_concurrency": MAX_CONCURRENT,
        "elapsed_seconds": round(elapsed, 2),
        "input_summaries": texts,
        "batch_metadata": batch_metadata,
        "batch_analyses": batch_analyses,
        "final_analysis": final_analysis,
    }


def save_to_word(result: dict, output_file: str = "C:/Users/david/Desktop/analysis_output.docx") -> None:
    """Save consolidated analysis to Word format with all six sections."""
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from datetime import datetime

    # Create new document
    doc = Document()

    # Set document margins
    for section in doc.sections:
        section.top_margin = Cm(2.54)
        section.bottom_margin = Cm(2.54)
        section.left_margin = Cm(2.54)
        section.right_margin = Cm(2.54)

    # Title
    title = doc.add_paragraph()
    title_run = title.add_run('HUMANITARIAN SITUATION ANALYSIS')
    title_run.font.size = Pt(24)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0, 51, 102)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph('Consolidated Strategic Assessment')
    subtitle_run = subtitle.runs[0]
    subtitle_run.font.size = Pt(14)
    subtitle_run.font.italic = True
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # Metadata section
    doc.add_heading('Analysis Overview', level=1)
    metadata_table = doc.add_table(rows=4, cols=2)
    metadata_table.style = 'Light Grid Accent 1'

    metadata_items = [
        ("Total Summaries Analyzed", str(result['summaries_analyzed'])),
        ("Total Batches Processed", str(result['batches'])),
        ("Batch Size", str(result['batch_size_used'])),
        ("Processing Time", f"{result['elapsed_seconds']} seconds")
    ]

    for i, (key, value) in enumerate(metadata_items):
        row = metadata_table.rows[i]
        row.cells[0].text = key
        row.cells[1].text = value
        # Format first column as bold
        for paragraph in row.cells[0].paragraphs:
            for run in paragraph.runs:
                run.font.bold = True

    doc.add_paragraph()

    # Add separator
    doc.add_paragraph('─' * 70)

    # Parse and add the consolidated analysis with all six sections
    doc.add_heading('CONSOLIDATED ANALYSIS', level=1)

    final_analysis = result["final_analysis"]

    # Function to add formatted content
    def add_formatted_content(doc, content):
        lines = content.split('\n')

        for line in lines:
            if not line.strip():
                continue

            # Check for section headers (## 1., ## 2., etc.)
            if line.strip().startswith('##'):
                # This is a main section header
                header_text = line.strip().lstrip('#').strip()
                doc.add_heading(header_text, level=2)
            elif line.strip().startswith('###'):
                # This is a subsection header
                header_text = line.strip().lstrip('#').strip()
                doc.add_heading(header_text, level=3)
            elif line.strip().startswith('- ') or line.strip().startswith('* '):
                # Bullet points
                doc.add_paragraph(line.strip()[2:], style='List Bullet')
            elif line.strip().startswith('|') and line.strip().endswith('|'):
                # Skip markdown tables (they're handled separately)
                continue
            else:
                # Regular paragraph
                p = doc.add_paragraph(line.strip())

                # Format bold text
                bold_pattern = r'\*\*(.*?)\*\*'
                matches = re.findall(bold_pattern, line)
                if matches:
                    # Clear and re-add with formatting
                    p.clear()
                    parts = re.split(r'\*\*.*?\*\*', line)
                    bold_parts = re.findall(r'\*\*(.*?)\*\*', line)

                    for j, part in enumerate(parts):
                        if part:
                            p.add_run(part)
                        if j < len(bold_parts):
                            run = p.add_run(bold_parts[j])
                            run.bold = True

    # Add the consolidated analysis
    add_formatted_content(doc, final_analysis)

    # Add footer with timestamp
    footer = doc.sections[0].footer
    footer_para = footer.paragraphs[0]
    footer_para.text = f"Generated by Humanitarian Communications Director - Batch Analyzer | {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in footer_para.runs:
        run.font.size = Pt(8)

    # Save the document
    doc.save(output_file)
    print(f"\nConsolidated analysis saved to: {output_file}")

    # Also print a summary to console
    print("\n" + "=" * 70)
    print("CONSOLIDATED ANALYSIS COMPLETE - ALL SIX SECTIONS INCLUDED:")
    print("=" * 70)
    print("1. HUMANITARIAN NEEDS & RESPONSE")
    print("2. IHL / RULES OF WAR")
    print("3. IMPACT ON CIVILIANS")
    print("4. WOUNDED, DETAINEES, POWs & SHIPWRECKED")
    print("5. SPEAKING OPPORTUNITIES")
    print("6. COMMUNICATION STRATEGIES")
    print("=" * 70)


def print_summary_stats(result: dict) -> None:
    """Print summary statistics."""
    print(f"\n{'─' * 70}")
    print("  ANALYSIS COMPLETE")
    print(f"{'─' * 70}")
    print(f"  Summaries analyzed : {result['summaries_analyzed']}")
    print(f"  Batches processed  : {result['batches']} (size={result['batch_size_used']})")
    print(f"  Concurrency level  : {result['max_concurrency']}")
    print(f"  Processing time    : {result['elapsed_seconds']} seconds")
    print(f"  Avg time per batch : {result['elapsed_seconds'] / result['batches']:.1f} seconds")
    print(f"{'─' * 70}\n")


def get_or_create_event_loop():
    """Get the current event loop or create a new one if none exists."""
    try:
        loop = asyncio.get_event_loop()
        if loop.is_closed():
            raise RuntimeError("Event loop is closed")
        return loop
    except RuntimeError:
        # No event loop running in this thread, create a new one
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        return loop


def run_async_main(async_func, *args, **kwargs):
    """Run an async function, handling event loop properly for different environments."""
    try:
        # Try to get the current event loop
        loop = get_or_create_event_loop()

        # Check if we're already in a running loop (like in Jupyter)
        if loop.is_running():
            # Create a new loop in a separate thread for Jupyter environments
            import concurrent.futures

            def run_in_thread():
                new_loop = asyncio.new_event_loop()
                asyncio.set_event_loop(new_loop)
                try:
                    return new_loop.run_until_complete(async_func(*args, **kwargs))
                finally:
                    new_loop.close()

            with concurrent.futures.ThreadPoolExecutor() as executor:
                future = executor.submit(run_in_thread)
                return future.result()
        else:
            # Normal case - run in current loop
            return loop.run_until_complete(async_func(*args, **kwargs))
    except RuntimeError:
        # Fallback: create a new loop
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(async_func(*args, **kwargs))
        finally:
            loop.close()


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("╔══════════════════════════════════════════════════════════════════╗")
    print("║    Humanitarian Communications Director — Batched Analyzer       ║")
    print("║    Model: llama3.2 via Ollama                                     ║")
    print(f"║    Batch Size: {BATCH_SIZE} | Concurrency: {MAX_CONCURRENT}                                  ║")
    print("╚══════════════════════════════════════════════════════════════════╝")

    # Use the safe async runner that works in both scripts and interactive environments
    result = run_async_main(run_analysis, summaries)

    # Save only to Word format with consolidated analysis
    save_to_word(result, output_file="C:/Users/david/Desktop/analysis_output.docx")

    # Print summary statistics
    print_summary_stats(result)