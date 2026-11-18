"""
Council - ask Claude, Gemini and Copilot at once, from one desktop window.

No API keys and no pay-per-use: each model runs through its official tool,
signed in with a subscription you already have.
  Claude          -> Claude Code CLI      (claude)   - Claude Pro
  Gemini          -> Antigravity CLI      (agy)      - Google AI Pro
  GitHub Copilot  -> GitHub Copilot CLI   (copilot)  - GitHub Copilot plan (Free works)
  M365 Copilot    -> Microsoft 365 Copilot Chat API  - Business plan with Copilot license
  Local           -> Ollama on this PC (http://localhost:11434) - free, works offline

Only Python's standard library is used (tkinter ships with Python on Windows).
Run:  python council.py
"""

import datetime
import json
import os
import queue
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
from pathlib import Path
from tkinter import messagebox, scrolledtext, simpledialog, ttk

APP_DIR = Path.home() / ".council"
WORK_DIR = APP_DIR / "workspace"  # empty folder the CLIs run in, so they never touch your files
CHAT_DIR = APP_DIR / "chats"
CONFIG_PATH = APP_DIR / "config.json"

CHAT_PREAMBLE = (
    "You are answering as a chat assistant inside a desktop app. Reply directly in plain text. "
    "Do not use tools to read, create, edit or run files or commands - you can't access the "
    "user's computer, and any documents they share are pasted into this message. "
    "If you can search the web, you may."
)

DEFAULT_CONFIG = {
    "timeout_seconds": 300,
    "history_turns": 6,
    "synthesizer": "Claude",
    "models": [
        {
            "name": "Claude",
            "enabled": True,
            "color": "#d97757",
            # prompt is piped on stdin; -p with no value reads it from there
            "command": ["claude", "-p"],
            "prompt_via": "stdin",
            # make sure Claude Code uses your Pro login, never a metered API key
            "remove_env": ["ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN"],
        },
        {
            "name": "Gemini",
            "enabled": True,
            "color": "#4285f4",
            # agy's -p takes the prompt as its value (it does not read stdin)
            "command": ["agy", "-p", "{prompt}"],
            "prompt_via": "arg",
            "remove_env": ["GEMINI_API_KEY", "GOOGLE_API_KEY"],
        },
        {
            "name": "GitHub Copilot",
            "enabled": True,
            "color": "#8957e5",
            # GitHub docs: pipe the prompt in and do NOT pass -p
            "command": ["copilot", "-s", "--no-ask-user"],
            "prompt_via": "stdin",
            "remove_env": [],
        },
        {
            "name": "M365 Copilot",
            "type": "m365",
            "enabled": True,
            "color": "#0078d4",
            # From the 'Council' app registration in Voyager Publishing's Entra admin center
            "client_id": "088e5a23-de18-4145-b127-301c2a1e7dbc",
            "tenant_id": "8a2477f1-0127-433b-81b7-167696dc493c",
            "time_zone": "America/Los_Angeles",
            "web_search": True,
        },
        {
            "name": "Local",
            "type": "ollama",
            "enabled": True,
            "color": "#2f9e44",
            # any model from 'ollama list'; pick another one from the pane's dropdown
            "model": "gpt-oss:20b",
            "host": "http://localhost:11434",
            # unload the model from the GPU this long after the last question
            "keep_alive": "5m",
        },
    ],
}

# if the chosen local model isn't installed, the first of these that is gets used
OLLAMA_FALLBACKS = ["gpt-oss:20b", "qwen3:14b", "qwen3.8:27b", "gemma4:26b", "qwen2.5:14b", "qwen2.5:7b"]

# Delegated Microsoft Graph permissions the Copilot Chat API requires (all of them)
M365_SCOPES = (
    "Sites.Read.All Mail.Read People.Read.All OnlineMeetingTranscript.Read.All "
    "Chat.Read ChannelMessage.Read.All ExternalItem.Read.All offline_access"
)
GRAPH = "https://graph.microsoft.com/beta/copilot/conversations"


# ---------------------------------------------------------------- config

def load_config():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    WORK_DIR.mkdir(parents=True, exist_ok=True)
    CHAT_DIR.mkdir(parents=True, exist_ok=True)
    if not CONFIG_PATH.exists():
        CONFIG_PATH.write_text(json.dumps(DEFAULT_CONFIG, indent=2), encoding="utf-8")
        return json.loads(json.dumps(DEFAULT_CONFIG))
    try:
        cfg = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except Exception as e:  # broken config: fall back to defaults, keep the file for fixing
        messagebox.showwarning("Council", f"config.json could not be read ({e}). Using defaults.")
        return json.loads(json.dumps(DEFAULT_CONFIG))
    for k, v in DEFAULT_CONFIG.items():
        cfg.setdefault(k, v)
    # add any model introduced in a newer version of Council
    have = {m.get("name") for m in cfg["models"]}
    added = False
    for m in DEFAULT_CONFIG["models"]:
        if m["name"] not in have and not (m["name"] == "GitHub Copilot" and "Copilot" in have):
            cfg["models"].append(json.loads(json.dumps(m)))
            added = True
    # fill in missing Microsoft 365 IDs from the built-in defaults
    defaults = {m["name"]: m for m in DEFAULT_CONFIG["models"]}
    for m in cfg["models"]:
        if m.get("type") == "m365" and not (m.get("client_id") and m.get("tenant_id")):
            d = defaults.get(m["name"], {})
            if d.get("client_id"):
                m["client_id"], m["tenant_id"] = d["client_id"], d["tenant_id"]
                added = True
    if added:
        save_config(cfg)
    return cfg


def save_config(cfg):
    try:
        CONFIG_PATH.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


# ---------------------------------------------------------------- Microsoft 365 Copilot

class M365Error(Exception):
    pass


def _http(url, form=None, body=None, token=None, timeout=60):
    """Small JSON-over-HTTPS helper (standard library only)."""
    headers = {"Accept": "application/json"}
    data = None
    if form is not None:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    elif body is not None:
        data = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    if token:
        headers["Authorization"] = f"Bearer {token}"
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            raw = r.read().decode("utf-8", "replace")
            return r.status, (json.loads(raw) if raw.strip() else {})
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(raw)
        except ValueError:
            return e.code, {"error": {"message": raw[:500]}}
    except urllib.error.URLError as e:
        raise M365Error(f"Network error: {e.reason}")


class M365Client:
    """Signs in with the device-code flow and talks to the Copilot Chat API."""

    def __init__(self, model, on_signin):
        self.model = model
        self.on_signin = on_signin  # callback(message, code, url) - runs on the UI thread
        self.token_path = APP_DIR / "m365_token.json"
        self.lock = threading.Lock()

    @property
    def authority(self):
        tenant = self.model.get("tenant_id") or "organizations"
        return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0"

    def _save(self, tok):
        tok["expires_at"] = time.time() + int(tok.get("expires_in", 3600)) - 60
        try:
            self.token_path.write_text(json.dumps(tok), encoding="utf-8")
        except Exception:
            pass
        return tok["access_token"]

    def sign_out(self):
        try:
            self.token_path.unlink()
        except FileNotFoundError:
            pass

    def token(self):
        with self.lock:
            cid = self.model.get("client_id")
            if not cid:
                raise M365Error("No app (client) ID set. See README.md, 'Microsoft 365 Copilot setup'.")
            tok = {}
            if self.token_path.exists():
                try:
                    tok = json.loads(self.token_path.read_text(encoding="utf-8"))
                except ValueError:
                    tok = {}
            if tok.get("access_token") and tok.get("expires_at", 0) > time.time():
                return tok["access_token"]
            if tok.get("refresh_token"):
                status, res = _http(self.authority + "/token", form={
                    "client_id": cid, "grant_type": "refresh_token",
                    "refresh_token": tok["refresh_token"], "scope": M365_SCOPES,
                })
                if status == 200 and "access_token" in res:
                    return self._save(res)
            return self._device_flow(cid)

    def _device_flow(self, cid):
        status, res = _http(self.authority + "/devicecode", form={"client_id": cid, "scope": M365_SCOPES})
        if status != 200:
            raise M365Error("Sign-in could not start: " + _err_text(res)
                            + "\n\nTip: click 'Sign in' on this pane and re-enter both IDs.")
        self.on_signin(res.get("message", ""), res["user_code"], res["verification_uri"])
        interval = int(res.get("interval", 5))
        deadline = time.time() + int(res.get("expires_in", 900))
        while time.time() < deadline:
            time.sleep(interval)
            status, tok = _http(self.authority + "/token", form={
                "client_id": cid, "device_code": res["device_code"],
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            })
            if status == 200 and "access_token" in tok:
                return self._save(tok)
            code = tok.get("error")
            if code == "authorization_pending":
                continue
            if code == "slow_down":
                interval += 5
                continue
            raise M365Error("Sign-in failed: " + _err_text(tok))
        raise M365Error("Sign-in timed out. Ask again to get a new code.")

    def ask(self, conversation_id, text, timeout):
        """Returns (conversation_id, reply_text). Pass None to start a new conversation."""
        token = self.token()
        if not conversation_id:
            status, res = _http(GRAPH, body={}, token=token, timeout=timeout)
            if status not in (200, 201) or "id" not in res:
                raise M365Error(_m365_hint(status, res))
            conversation_id = res["id"]
        body = {
            "message": {"text": text},
            "locationHint": {"timeZone": self.model.get("time_zone") or "UTC"},
            "contextualResources": {"webContext": {"isWebEnabled": bool(self.model.get("web_search", True))}},
        }
        status, res = _http(f"{GRAPH}/{conversation_id}/chat", body=body, token=token, timeout=timeout)
        if status != 200:
            raise M365Error(_m365_hint(status, res))
        replies = [m.get("text", "") for m in res.get("messages", []) if m.get("text")]
        reply = next((t for t in reversed(replies) if t.strip() != text.strip()), "")
        if not reply:
            raise M365Error("Copilot returned an empty answer.")
        return conversation_id, reply


def _err_text(res):
    if isinstance(res.get("error"), dict):
        return res["error"].get("message") or res["error"].get("code") or str(res)
    return res.get("error_description") or str(res.get("error") or res)


def _m365_hint(status, res):
    msg = f"Copilot API error {status}: {_err_text(res)}"
    if status == 401:
        msg += "\n\nTip: use Settings > 'Sign out of Microsoft 365', then ask again."
    elif status == 403:
        msg += ("\n\nTip: check that (1) your user has the Copilot license assigned, and (2) all 7 "
                "Graph permissions were added with 'Grant admin consent' - then sign out and in again.")
    return msg


# ---------------------------------------------------------------- local model (Ollama)

def ollama_models(host, timeout=3):
    """Names of the models installed in Ollama, or None if Ollama isn't running."""
    try:
        with urllib.request.urlopen(host.rstrip("/") + "/api/tags", timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8", "replace"))
        return sorted(m["name"] for m in data.get("models", []))
    except Exception:
        return None


def pick_ollama_model(wanted, installed):
    if not installed:
        return wanted
    if wanted in installed:
        return wanted
    if wanted + ":latest" in installed:
        return wanted + ":latest"
    for name in OLLAMA_FALLBACKS:
        if name in installed:
            return name
    return installed[0]


def run_ollama(model, history, question, timeout, turns):
    """Ask a local Ollama model. Returns (ok, text)."""
    import re
    host = model.get("host") or "http://localhost:11434"
    name = model.get("model") or "gpt-oss:20b"
    messages = [{"role": "system", "content": CHAT_PREAMBLE}]
    for q, a in (history[-turns:] if turns > 0 else []):
        messages += [{"role": "user", "content": q}, {"role": "assistant", "content": a}]
    messages.append({"role": "user", "content": question})
    body = {"model": name, "messages": messages, "stream": False,
            "keep_alive": model.get("keep_alive", "5m")}
    req = urllib.request.Request(host.rstrip("/") + "/api/chat", data=json.dumps(body).encode(),
                                 headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            res = json.loads(r.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", "replace")
        try:
            raw = json.loads(raw).get("error", raw)
        except ValueError:
            pass
        tip = ""
        if "not found" in str(raw).lower():
            tip = (f"\n\nTip: '{name}' isn't downloaded. Pick another model from this pane's "
                   f"dropdown, or open a terminal and run:  ollama pull {name}")
        return False, f"Ollama error {e.code}: {raw}{tip}"
    except urllib.error.URLError as e:
        return False, (f"Can't reach Ollama at {host} ({e.reason}).\n\n"
                       "Tip: start Ollama (Start menu > Ollama, or the Ollama icon on your desktop), "
                       "then click the badge above to re-check.")
    except TimeoutError:
        return False, f"Timed out after {timeout} seconds (big local models can be slow to load)."
    text = (res.get("message") or {}).get("content", "")
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S).strip()  # hide reasoning
    if not text:
        return False, "The local model returned an empty answer."
    return True, text


# ---------------------------------------------------------------- running a CLI

def _cmd_safe(text):
    """Windows .cmd/.bat shims go through cmd.exe, which mangles some characters
    in arguments. Only used when a prompt must be passed as an argument to one."""
    return text.replace("\r", " ").replace("\n", " ").replace('"', "'").replace("%", " percent")


def run_model(model, prompt, timeout, proc_holder):
    """Run one CLI and return (ok, text)."""
    cmd = list(model["command"])
    exe = shutil.which(cmd[0])
    if not exe:
        return False, (
            f"'{cmd[0]}' was not found on your PATH.\n"
            f"Install it (see README.md), open a new terminal, and run '{cmd[0]}' once to sign in."
        )
    cmd[0] = exe
    is_shim = exe.lower().endswith((".cmd", ".bat"))

    via_stdin = model.get("prompt_via", "stdin") == "stdin"
    if not via_stdin:
        value = _cmd_safe(prompt) if is_shim else prompt
        cmd = [value if part == "{prompt}" else part for part in cmd]

    env = os.environ.copy()
    for key in model.get("remove_env", []):
        env.pop(key, None)
    env.setdefault("NO_COLOR", "1")

    flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    try:
        proc = subprocess.Popen(
            cmd,
            cwd=str(WORK_DIR),
            env=env,
            stdin=subprocess.PIPE if via_stdin else subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            creationflags=flags,
        )
    except Exception as e:
        return False, f"Could not start {model['name']}: {e}"

    proc_holder.append(proc)
    try:
        out, err = proc.communicate(input=prompt if via_stdin else None, timeout=timeout)
    except subprocess.TimeoutExpired:
        proc.kill()
        return False, f"Timed out after {timeout} seconds."

    out, err = (out or "").strip(), (err or "").strip()
    if getattr(proc, "council_stopped", False):
        return False, "Stopped."
    if proc.returncode != 0:
        msg = err or out or f"exited with code {proc.returncode}"
        hint = ""
        low = msg.lower()
        if any(w in low for w in ("login", "log in", "logged in", "auth", "sign in", "signed in")):
            hint = "\n\nTip: click the 'Sign in' button at the top of this pane."
        return False, msg + hint
    if not out:
        if "permission" in err.lower() or "denied" in err.lower():
            return False, (
                f"{model['name']} tried to use a tool on your PC (like opening a file), which "
                "Council blocks for safety, so it gave up without answering.\n\n"
                "Tip: if your question is about a document (like your résumé), use "
                "📎 Attach file so its text is sent with the question.")
        return False, (
            "The CLI finished but printed nothing."
            + (f"\n\nDetails: {err}" if err else "")
            + ("\n\nTip: older agy versions drop output when not in a terminal - run 'agy update'."
               if model["command"][0].lower().startswith("agy") and not err else "")
        )
    return True, out


# ---------------------------------------------------------------- attachments

MAX_ATTACH_CHARS = 40000


def read_attachment(path):
    """Return the text of a file the user attached. Raises ValueError with a friendly message."""
    p = Path(path)
    ext = p.suffix.lower()
    if ext == ".docx":
        import re
        import zipfile
        with zipfile.ZipFile(p) as z:
            xml = z.read("word/document.xml").decode("utf-8", "replace")
        xml = re.sub(r"</w:p>", "\n", xml)
        xml = re.sub(r"<w:tab/>", "\t", xml)
        text = re.sub(r"<[^>]+>", "", xml)
        import html
        text = html.unescape(text)
    elif ext == ".pdf":
        try:
            from pypdf import PdfReader  # optional
        except ImportError:
            raise ValueError("Council can't read PDFs on this PC yet. Open the PDF, "
                             "Save As / Export it as a Word (.docx) or text file, and attach that.")
        text = "\n".join((pg.extract_text() or "") for pg in PdfReader(str(p)).pages)
    elif ext in (".doc", ".pages", ".odt"):
        raise ValueError("Please save the document as .docx or .txt first, then attach it.")
    else:
        raw = p.read_bytes()
        if b"\x00" in raw[:2000]:
            raise ValueError("That file isn't text. Attach a .txt, .md, .csv, .json or .docx file.")
        text = raw.decode("utf-8", "replace")
    text = text.strip()
    if not text:
        raise ValueError("That file seems to be empty (or it's a scanned image with no text).")
    if len(text) > MAX_ATTACH_CHARS:
        text = text[:MAX_ATTACH_CHARS] + "\n[...file truncated...]"
    return text


# ---------------------------------------------------------------- UI

class ModelPane:
    def __init__(self, parent, model, on_signin=None):
        self.model = model
        self.is_m365 = model.get("type") == "m365"
        self.is_ollama = model.get("type") == "ollama"
        self.history = []  # list of (question, answer) for this model only
        self.procs = []
        self.last_answer = ""
        self.conversation_id = None  # M365 Copilot keeps its own multi-turn conversation
        self.stopped = False

        self.frame = ttk.Frame(parent, padding=4)
        head = ttk.Frame(self.frame)
        head.pack(fill="x")
        self.enabled = tk.BooleanVar(value=model.get("enabled", True))
        tk.Label(head, text="●", fg=model.get("color", "#888"), font=("Segoe UI", 12)).pack(side="left")
        ttk.Checkbutton(head, text=model["name"], variable=self.enabled).pack(side="left")
        self.status = ttk.Label(head, text="ready", foreground="#888")
        self.status.pack(side="right")
        ttk.Button(head, text="Copy", width=6, command=self.copy).pack(side="right", padx=4)
        self.signin_btn = None
        self.local_choice = None
        if self.is_ollama:
            self.local_choice = tk.StringVar(value=model.get("model", ""))
            self.local_box = ttk.Combobox(head, textvariable=self.local_choice, width=22, state="readonly")
            self.local_box.pack(side="right", padx=2)
        elif on_signin:
            self.signin_btn = ttk.Button(head, text="Sign in", width=10, command=lambda: on_signin(self))
            self.signin_btn.pack(side="right")

        # sign-in status bar under the header
        bar = tk.Frame(self.frame)
        bar.pack(fill="x", pady=(2, 3))
        self.badge = tk.Label(bar, text="", font=("Segoe UI", 9, "bold"), cursor="hand2", anchor="w")
        self.badge.pack(side="left")
        self.badge_detail = tk.Label(bar, text="", font=("Segoe UI", 8), fg="#888", anchor="w")
        self.badge_detail.pack(side="left", padx=6)
        self.on_check = None  # set by the app: click the badge to re-check
        self.badge.bind("<Button-1>", lambda e: self.on_check and self.on_check(self))
        self.state = None
        self.set_state("unknown")

        self.text = scrolledtext.ScrolledText(self.frame, wrap="word", font=("Segoe UI", 10), height=10)
        self.text.pack(fill="both", expand=True)
        self.text.tag_configure("you", foreground="#888", font=("Segoe UI", 9, "italic"))
        self.text.tag_configure("err", foreground="#c0392b")
        self.text.tag_configure("good", foreground="#1a7f37", font=("Segoe UI", 9, "bold"))
        self.text.configure(state="disabled")

    STATES = {
        "ok":      ("✔ Signed in",          "#1a7f37", "click to re-check"),
        "found":   ("✔ Sign-in found",      "#1a7f37", "click to verify"),
        "need":    ("✖ Sign-in needed",     "#c0392b", "use the Sign in button"),
        "setup":   ("✖ Setup needed",       "#c0392b", "use the Sign in button"),
        "check":   ("⟳ Checking sign-in…",  "#0078d4", ""),
        "waiting": ("⟳ Waiting for sign-in…", "#0078d4", "finish in the window that opened"),
        "unknown": ("● Not verified",       "#b08800", "click to check"),
        "local":   ("✔ Running on this PC", "#1a7f37", "free · works offline"),
        "nolocal": ("✖ Ollama not running", "#c0392b", "start Ollama, then click here"),
    }

    def set_state(self, state, detail=None):
        self.state = state
        text, color, hint = self.STATES[state]
        self.badge.configure(text=text, fg=color)
        self.badge_detail.configure(text=detail if detail is not None else hint)
        if self.signin_btn is not None:
            self.signin_btn.configure(text="Re-sign in" if state in ("ok", "found") else "Sign in")

    def write(self, content, tag=None):
        self.text.configure(state="normal")
        self.text.insert("end", content, tag)
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self):
        self.history.clear()
        self.last_answer = ""
        self.conversation_id = None
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
        self.status.configure(text="ready")

    def copy(self):
        if self.last_answer:
            self.frame.clipboard_clear()
            self.frame.clipboard_append(self.last_answer)
            self.status.configure(text="copied")

    def stop(self):
        self.stopped = True
        for p in self.procs:
            if p.poll() is None:
                p.council_stopped = True
                p.kill()

    def build_prompt(self, question, turns):
        if self.is_m365 or self.is_ollama:  # history is sent separately for these
            return question
        parts = [CHAT_PREAMBLE, ""]
        recent = self.history[-turns:] if turns > 0 else []
        if recent:
            parts.append("Conversation so far:")
            for q, a in recent:
                parts += [f"User: {q}", f"Assistant: {a}", ""]
            parts.append("New message:")
        parts.append(question)
        return "\n".join(parts)


class CouncilApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.events = queue.Queue()
        self.running = 0
        self.transcript = []
        self.chat_file = CHAT_DIR / f"chat-{datetime.datetime.now():%Y%m%d-%H%M%S}.md"

        self.m365 = {
            m["name"]: M365Client(m, self.request_signin)
            for m in self.cfg["models"] if m.get("type") == "m365"
        }

        root.title("Council - Claude · Gemini · Copilot · Local")
        root.geometry("1600x860")
        root.minsize(900, 500)

        # answers row
        panes = ttk.PanedWindow(root, orient="horizontal")
        panes.pack(fill="both", expand=True, padx=8, pady=(8, 4))
        self.panes = []
        for m in self.cfg["models"]:
            pane = ModelPane(panes, m, on_signin=self.sign_in)
            pane.on_check = self.verify
            panes.add(pane.frame, weight=1)
            self.panes.append(pane)

        # combined answer
        comb = ttk.LabelFrame(root, text="Combined answer", padding=4)
        comb.pack(fill="both", padx=8, pady=4)
        self.combined = scrolledtext.ScrolledText(comb, wrap="word", font=("Segoe UI", 10), height=7)
        self.combined.pack(fill="both", expand=True)
        self.combined.configure(state="disabled")

        # input row
        bottom = ttk.Frame(root, padding=(8, 4, 8, 8))
        bottom.pack(fill="x")
        self.entry = tk.Text(bottom, height=3, wrap="word", font=("Segoe UI", 11))
        self.entry.pack(side="left", fill="x", expand=True)
        self.entry.bind("<Return>", self.on_enter)
        self.entry.focus_set()

        self.attachments = []  # list of (filename, text)
        att = ttk.Frame(bottom)
        att.pack(side="left", padx=(6, 0))
        ttk.Button(att, text="📎 Attach file", command=self.attach_file).pack(fill="x")
        self.attach_label = tk.Label(att, text="no file", fg="#888", font=("Segoe UI", 8),
                                     wraplength=140, justify="left", cursor="hand2")
        self.attach_label.pack(fill="x", pady=(2, 0))
        self.attach_label.bind("<Button-1>", lambda e: self.clear_attachments())

        btns = ttk.Frame(bottom)
        btns.pack(side="left", padx=(8, 0))
        self.ask_btn = ttk.Button(btns, text="Ask all  (Enter)", command=self.ask)
        self.ask_btn.grid(row=0, column=0, sticky="ew")
        names = [m["name"] for m in self.cfg["models"]]
        self.synth_choice = tk.StringVar(value=self.cfg.get("synthesizer", names[0]))
        sf = ttk.Frame(btns)
        sf.grid(row=1, column=0, sticky="ew", pady=2)
        self.synth_btn = ttk.Button(sf, text="Combine with", command=self.synthesize)
        self.synth_btn.pack(side="left")
        ttk.Combobox(sf, textvariable=self.synth_choice, values=names, width=8, state="readonly").pack(side="left")
        bf = ttk.Frame(btns)
        bf.grid(row=2, column=0, sticky="ew")
        ttk.Button(bf, text="Stop", command=self.stop).pack(side="left", expand=True, fill="x")
        ttk.Button(bf, text="New chat", command=self.new_chat).pack(side="left", expand=True, fill="x")
        ttk.Button(bf, text="Settings", command=self.open_settings).pack(side="left", expand=True, fill="x")

        root.after(100, self.poll)
        root.after(300, self.initial_checks)

    # ---- helpers
    def on_enter(self, event):
        if event.state & 0x0001:  # Shift+Enter = newline
            return None
        self.ask()
        return "break"

    def set_busy(self, delta):
        self.running += delta
        busy = self.running > 0
        self.ask_btn.configure(state="disabled" if busy else "normal")
        self.synth_btn.configure(state="disabled" if busy else "normal")

    def log(self, text):
        self.transcript.append(text)
        try:
            self.chat_file.write_text("\n".join(self.transcript), encoding="utf-8")
        except Exception:
            pass

    def write_combined(self, content, clear=False):
        self.combined.configure(state="normal")
        if clear:
            self.combined.delete("1.0", "end")
        self.combined.insert("end", content)
        self.combined.see("end")
        self.combined.configure(state="disabled")

    # ---- Microsoft 365
    def request_signin(self, message, code, url):
        """Called from a worker thread when Microsoft needs you to sign in."""
        self.events.put(("signin", message, code, url))

    def ensure_m365_ids(self, pane, force=False):
        """Ask for the app registration IDs (both are required). Returns False if cancelled."""
        m = pane.model
        if m.get("client_id") and m.get("tenant_id") and not force:
            return True
        cid = simpledialog.askstring(
            "Microsoft 365 Copilot setup",
            "Paste the Application (client) ID of your 'Council' app registration\n"
            "(entra.microsoft.com > App registrations > Council > Overview):",
            parent=self.root, initialvalue=m.get("client_id", ""),
        )
        if not cid or not cid.strip():
            return False
        tid = simpledialog.askstring(
            "Microsoft 365 Copilot setup",
            "Paste the Directory (tenant) ID from the same Overview page\n(it's required):",
            parent=self.root, initialvalue=m.get("tenant_id", ""),
        )
        if not tid or not tid.strip():
            messagebox.showwarning("Council", "The Directory (tenant) ID is required for M365 Copilot.")
            return False
        m["client_id"] = cid.strip()
        m["tenant_id"] = tid.strip()
        save_config(self.cfg)
        return True

    def sign_in(self, pane):
        """'Sign in' button: opens the tool's own sign-in in a terminal window (or M365 setup)."""
        if pane.is_m365:
            if not self.ensure_m365_ids(pane, force=not (pane.model.get("client_id") and pane.model.get("tenant_id"))):
                return
            self.m365[pane.model["name"]].sign_out()
            pane.conversation_id = None
            # verifying triggers Microsoft's sign-in page (code copied for you) and then a test question
            self.verify(pane)
            return
        cmd0 = pane.model["command"][0]
        exe = shutil.which(cmd0)
        if not exe:
            messagebox.showerror(
                "Council", f"'{cmd0}' isn't installed or isn't on your PATH yet.\n\n"
                "Run setup.bat (or fix_claude.bat for Claude), then restart Council.")
            return
        steps = {
            "claude": "1. Your browser opens - pick your Claude Pro account.\n"
                      "2. Back in the black window, type /exit and press Enter.",
            "agy": "1. Your browser opens - sign in with your Google AI Pro account.\n"
                   "2. If the page shows a code: copy it, click in the black window,\n"
                   "   right-click to paste, and press Enter.\n"
                   "3. When it says you're signed in, close the black window.",
            "copilot": "1. In the black window, type /login and press Enter.\n"
                       "2. Sign in with GitHub in your browser.\n"
                       "3. Back in the black window, type /exit and press Enter.",
        }
        key = Path(cmd0).stem.lower()
        env = os.environ.copy()
        for k in pane.model.get("remove_env", []):
            env.pop(k, None)
        try:
            if sys.platform == "win32":
                # /c: the window closes when the tool exits, so Council knows you're done
                proc = subprocess.Popen(["cmd", "/c", exe], cwd=str(WORK_DIR), env=env,
                                        creationflags=subprocess.CREATE_NEW_CONSOLE)
            else:
                proc = subprocess.Popen([exe], cwd=str(WORK_DIR), env=env)
        except Exception as e:
            messagebox.showerror("Council", f"Couldn't open {cmd0}: {e}")
            return
        pane.set_state("waiting")
        pane.write(f"\n— Sign-in window opened for {pane.model['name']}. "
                   "Council will check automatically when you close it. —\n", "you")

        def watch():
            proc.wait()
            self.events.put(("verify", pane))

        threading.Thread(target=watch, daemon=True).start()
        messagebox.showinfo(f"Sign in to {pane.model['name']}",
                            steps.get(key, f"Sign in to {cmd0} in the window that opened.")
                            + "\n\nCouncil checks your sign-in automatically when that window closes.")

    # ---- sign-in status
    def quick_state(self, pane):
        """Fast startup guess from the tools' own credential files (no requests sent)."""
        home = Path.home()
        if pane.is_ollama:
            return self.refresh_local(pane)
        if pane.is_m365:
            m = pane.model
            if not (m.get("client_id") and m.get("tenant_id")):
                return "setup"
            try:
                tok = json.loads(self.m365[m["name"]].token_path.read_text(encoding="utf-8"))
                return "found" if tok.get("refresh_token") else "need"
            except Exception:
                return "need"
        key = Path(pane.model["command"][0]).stem.lower()
        if not shutil.which(pane.model["command"][0]):
            return "setup"
        try:
            if key == "claude":
                return "found" if (home / ".claude" / ".credentials.json").exists() else "unknown"
            if key == "agy":
                d = home / ".gemini" / "antigravity-cli"
                return "found" if any(d.glob("credentials*")) else "unknown"
            if key == "copilot":
                cfg = json.loads((home / ".copilot" / "config.json").read_text(encoding="utf-8"))
                users = cfg.get("logged_in_users") or cfg.get("last_logged_in_user")
                return "found" if users else "unknown"
        except Exception:
            pass
        return "unknown"

    def refresh_local(self, pane):
        """Fill the local pane's model dropdown from Ollama. Returns a badge state."""
        m = pane.model
        installed = ollama_models(m.get("host") or "http://localhost:11434")
        if installed is None:
            return "nolocal"
        pane.local_box.configure(values=installed)
        chosen = pick_ollama_model(m.get("model", ""), installed)
        if chosen != m.get("model"):
            pane.write(f"'{m.get('model')}' isn't downloaded in Ollama, so using '{chosen}'.\n"
                       "Pick another model from the dropdown above any time.\n", "you")
        pane.local_choice.set(chosen)
        m["model"] = chosen
        if not getattr(pane, "local_bound", False):
            pane.local_bound = True

            def changed(_e=None):
                m["model"] = pane.local_choice.get()
                save_config(self.cfg)
                pane.write(f"Local model switched to {m['model']}.\n", "you")
            pane.local_box.bind("<<ComboboxSelected>>", changed)
        return "local"

    def initial_checks(self):
        for p in self.panes:
            st = self.quick_state(p)
            detail = None
            if st == "setup" and not p.is_m365:
                detail = f"'{p.model['command'][0]}' not installed - run setup.bat"
            p.set_state(st, detail)

    def verify(self, pane):
        """Send a tiny test question to confirm the sign-in really works."""
        if pane.state == "check":
            return
        if pane.is_ollama:  # just re-read Ollama's model list; no test question needed
            pane.set_state(self.refresh_local(pane))
            return
        if pane.is_m365 and not self.ensure_m365_ids(pane):
            return
        pane.set_state("check")
        timeout = int(self.cfg.get("timeout_seconds", 300))

        def job():
            ok, text = self.call_model(pane, "Reply with only the word OK.", timeout, keep=False)
            self.events.put(("verified", pane, ok, text))

        threading.Thread(target=job, daemon=True).start()

    @staticmethod
    def looks_like_auth_error(text):
        low = text.lower()
        return any(w in low for w in ("login", "log in", "logged in", "auth", "sign in",
                                      "sign-in", "signed in", "credential", "401", "token"))

    def call_model(self, pane, prompt, timeout, conversation_id=None, keep=True):
        """Run any model. For M365 Copilot, keep=True continues the pane's conversation."""
        if pane.is_ollama:
            turns = int(self.cfg.get("history_turns", 6)) if keep else 0
            return run_ollama(pane.model, pane.history, prompt, timeout, turns)
        if not pane.is_m365:
            return run_model(pane.model, prompt, timeout, pane.procs)
        client = self.m365[pane.model["name"]]
        try:
            conv, reply = client.ask(conversation_id, prompt, timeout)
        except M365Error as e:
            return False, str(e)
        except Exception as e:
            return False, f"Unexpected error: {e}"
        if keep:
            pane.conversation_id = conv
        return True, reply

    # ---- attachments
    def attach_file(self):
        from tkinter import filedialog
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Attach files to your next question",
            filetypes=[("Documents", "*.docx *.txt *.md *.pdf *.csv *.json"), ("All files", "*.*")])
        for path in paths or []:
            try:
                self.attachments.append((Path(path).name, read_attachment(path)))
            except ValueError as e:
                messagebox.showwarning("Council", f"{Path(path).name}: {e}")
            except Exception as e:
                messagebox.showwarning("Council", f"Couldn't read {Path(path).name}: {e}")
        self.update_attach_label()

    def clear_attachments(self):
        self.attachments = []
        self.update_attach_label()

    def update_attach_label(self):
        if self.attachments:
            names = ", ".join(n for n, _ in self.attachments)
            self.attach_label.configure(text=f"📎 {names}\n(click to remove)", fg="#0078d4")
        else:
            self.attach_label.configure(text="no file", fg="#888")

    # ---- actions
    def ask(self):
        typed = self.entry.get("1.0", "end").strip()
        if not typed or self.running:
            return
        question = typed
        files = [n for n, _ in self.attachments]
        if self.attachments:
            question += "".join(f"\n\n--- Attached file: {n} ---\n{t}\n--- End of {n} ---"
                                for n, t in self.attachments)
        active = [p for p in self.panes if p.enabled.get()]
        for p in [p for p in active if p.is_m365]:
            if not self.ensure_m365_ids(p):
                active.remove(p)
        if not active:
            messagebox.showinfo("Council", "Tick at least one model.")
            return
        self.entry.delete("1.0", "end")
        self.write_combined("", clear=True)
        shown = typed + (f"   [📎 {', '.join(files)}]" if files else "")
        self.log(f"## You\n{shown}\n")
        self.clear_attachments()
        timeout = int(self.cfg.get("timeout_seconds", 300))
        turns = int(self.cfg.get("history_turns", 6))
        for pane in active:
            pane.write(("\n" if pane.history else "") + f"You: {shown}\n\n", "you")
            pane.status.configure(text="thinking…")
            pane.procs = []
            pane.stopped = False
            prompt = pane.build_prompt(question, turns)
            self.set_busy(+1)
            threading.Thread(
                target=self.worker, args=(pane, question, prompt, timeout), daemon=True
            ).start()

    def worker(self, pane, question, prompt, timeout):
        started = datetime.datetime.now()
        ok, text = self.call_model(pane, prompt, timeout, conversation_id=pane.conversation_id)
        if pane.stopped and (pane.is_m365 or pane.is_ollama):
            ok, text = False, "Stopped."
        secs = (datetime.datetime.now() - started).seconds
        self.events.put(("answer", pane, question, ok, text, secs))

    def synthesize(self):
        if self.running:
            return
        answered = [p for p in self.panes if p.history and p.enabled.get()]
        if len(answered) < 2:
            messagebox.showinfo("Council", "Ask a question first - at least two models need to answer.")
            return
        question = answered[0].history[-1][0]
        judge = next((p for p in self.panes if p.model["name"] == self.synth_choice.get()), self.panes[0])
        if judge.is_m365 and not self.ensure_m365_ids(judge):
            return
        parts = [
            CHAT_PREAMBLE, "",
            "Several AI assistants answered the same question. Write one best combined answer: "
            "keep what they agree on, resolve or flag where they disagree, and fix anything wrong. "
            "Start with the answer itself, then a short 'Where they differed' note if relevant.", "",
            f"Question: {question}", "",
        ]
        for p in answered:
            parts += [f"--- {p.model['name']} answered ---", p.history[-1][1], ""]
        self.write_combined(f"{judge.model['name']} is combining the answers…\n", clear=True)
        judge.procs = []
        self.set_busy(+1)
        timeout = int(self.cfg.get("timeout_seconds", 300))

        def job():
            # a fresh conversation, so the merge doesn't pollute the pane's own chat
            ok, text = self.call_model(judge, "\n".join(parts), timeout, conversation_id=None, keep=False)
            self.events.put(("combined", judge, ok, text))

        threading.Thread(target=job, daemon=True).start()

    def stop(self):
        for p in self.panes:
            p.stop()

    def new_chat(self):
        if self.running:
            self.stop()
        for p in self.panes:
            p.clear()
        self.write_combined("", clear=True)
        self.transcript = []
        self.chat_file = CHAT_DIR / f"chat-{datetime.datetime.now():%Y%m%d-%H%M%S}.md"
        # re-check every model's sign-in and show it, so the button visibly does something
        for p in self.panes:
            if p.state in ("ok", "found"):
                p.write("New chat started.\n", "you")
                continue
            p.write("New chat started - checking sign-in…\n", "you")
            if p.is_m365 and p.state == "need":
                p.set_state("need")  # avoid popping a Microsoft sign-in page unasked
            else:
                self.verify(p)
        self.entry.focus_set()

    def open_settings(self):
        if self.m365 and messagebox.askyesno(
            "Council", "Sign out of Microsoft 365 Copilot?\n\n(Choose No to open the settings file instead.)"
        ):
            for c in self.m365.values():
                c.sign_out()
            for p in self.panes:
                if p.is_m365:
                    p.conversation_id = None
            messagebox.showinfo("Council", "Signed out. You'll get a new sign-in code next time you ask.")
            return
        if sys.platform == "win32":
            os.startfile(str(CONFIG_PATH))
        messagebox.showinfo(
            "Council",
            f"Settings file:\n{CONFIG_PATH}\n\nEdit commands, timeout or history there, "
            "then restart Council.\nChats are saved in:\n" + str(CHAT_DIR),
        )

    def show_signin(self, message, code, url):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(code)
        except tk.TclError:
            pass
        webbrowser.open(url)
        messagebox.showinfo(
            "Sign in to Microsoft 365",
            f"A Microsoft sign-in page just opened in your browser.\n\n"
            f"Your code is:   {code}\n(already copied - just paste it)\n\n"
            "Sign in with your BUSINESS account (the one with the Copilot license) "
            "and accept the permissions. Council continues by itself afterwards.",
            parent=self.root,
        )

    # ---- event loop
    def poll(self):
        try:
            while True:
                ev = self.events.get_nowait()
                if ev[0] == "answer":
                    _, pane, question, ok, text, secs = ev
                    if ok:
                        pane.history.append((question, text))
                        pane.last_answer = text
                        pane.write(text + "\n")
                        pane.status.configure(text=f"done · {secs}s")
                        pane.set_state("local" if pane.is_ollama else "ok")
                    else:
                        pane.write(text + "\n", "err")
                        pane.status.configure(text="error")
                        if self.looks_like_auth_error(text):
                            pane.set_state("need")
                    self.log(f"## {pane.model['name']}\n{text}\n")
                    self.set_busy(-1)
                elif ev[0] == "signin":
                    _, message, code, url = ev
                    for p in self.panes:
                        if p.is_m365 and p.state == "check":
                            p.set_state("waiting", "paste the code in the Microsoft page that opened")
                    self.show_signin(message, code, url)
                elif ev[0] == "verify":
                    self.verify(ev[1])
                elif ev[0] == "verified":
                    _, pane, ok, text = ev
                    if ok:
                        pane.set_state("ok", "verified just now")
                        pane.write(f"✔ {pane.model['name']} is signed in and ready.\n", "good")
                    else:
                        pane.set_state("need" if self.looks_like_auth_error(text) else "unknown",
                                       "check failed - see message below")
                        pane.write(f"Sign-in check failed: {text[:400]}\n", "err")
                elif ev[0] == "combined":
                    _, judge, ok, text = ev
                    self.write_combined(text if ok else "Could not combine:\n" + text, clear=True)
                    self.log(f"## Combined ({judge.model['name']})\n{text}\n")
                    self.set_busy(-1)
        except queue.Empty:
            pass
        self.root.after(100, self.poll)


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if sys.platform == "win32" else "clam")
    except tk.TclError:
        pass
    CouncilApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
