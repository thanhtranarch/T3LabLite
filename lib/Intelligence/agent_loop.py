# -*- coding: utf-8 -*-
"""Provider-agnostic tool-calling agent loop for the T3Lab Assistant.

Runs on a BACKGROUND thread. Providers that set SUPPORTS_NATIVE_TOOLS expose:

    chat_agent(system_prompt, messages, tools, on_delta=None, max_tokens=...)
        -> {"text":        unicode,          # user-visible text ("" if none)
            "tool_calls":  [{"id","name","args"}],
            "assistant_msg": <provider-native message to append to transcript>,
            "stop_reason": str}
        or None on transport/parse failure.

    agent_tool_results(tool_calls, results)
        -> list of provider-native messages carrying the tool results.

Never touches WPF or the Revit API: UI feedback goes through `callbacks`
(caller marshals to the UI thread); tools run via the injected
`execute_tool(name, args)` (core.server's _execute_tool, which marshals write
tools onto Revit's main thread via ExternalEvent).

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Agent Loop"

import json
import re
import threading
import time

from core import jsonsafe
from Intelligence.tool_schema import (LAUNCHER_TOOL_NAME, MEMORY_TOOL_NAME,
                                      PAST_CHATS_TOOL_NAME,
                                      is_file_output_only, is_model_modifying)

# Pseudo-tools the pane runs itself (launcher, memory file, archived chats);
# none touches the model.
_LOCAL_PSEUDO_TOOLS = (LAUNCHER_TOOL_NAME, MEMORY_TOOL_NAME, PAST_CHATS_TOOL_NAME)


# ─── Result truncation ─────────────────────────────────────────────────────────
# Bound the JSON fed back to the model; the UI card shows the full result.
_MAX_RESULT_CHARS = 4000


def _json_safe(obj):
    """Recursively coerce byte strings to unicode. Never raises.

    Code-page bytes from Revit/OS break json (decodes as UTF-8) and the
    IronPython `u"{}".format` fallback alike. Delegates to core.jsonsafe so
    there is ONE implementation (a private copy once left the MCP server path
    without it and broke tools/list on non-ASCII view names).
    """
    return jsonsafe.coerce(obj)


def _dumps(obj):
    try:
        return json.dumps(_json_safe(obj), ensure_ascii=False)
    except Exception:
        try:
            return u"{}".format(_json_safe(obj))
        except Exception:
            return u"<unserializable {}>".format(type(obj).__name__)


def _shrink_result(result, limit, path=None):
    """Fit `result` into `limit` chars without cutting JSON mid-string.

    1. Dict: halve its longest list fields until it fits, reporting
       `omitted_items` ({field: dropped count}); scalars like `total_count`
       survive (Working rule 10: numbers come from tools).
    2. Fallback (huge scalar, list root): wrap as
       {"truncated", "chars", "preview", "full_result_path"}.
    `path` = saved full copy (result_store) or None. Returns (text, was_truncated).
    """
    s = _dumps(result)
    if len(s) <= limit:
        return s, False
    if isinstance(result, dict):
        trimmed = dict(result)
        omitted = {}
        lists = [(k, v) for k, v in trimmed.items()
                 if isinstance(v, (list, tuple)) and len(v) > 1]
        lists.sort(key=lambda kv: -len(_dumps(kv[1])))
        for k, v in lists:
            v = list(v)
            while len(v) > 1 and len(_dumps(trimmed)) > limit:
                keep = max(1, len(v) // 2)
                omitted[k] = omitted.get(k, 0) + (len(v) - keep)
                v = v[:keep]
                trimmed[k] = v
            if len(_dumps(trimmed)) <= limit:
                break
        if omitted:
            trimmed["truncated"] = True
            trimmed["omitted_items"] = omitted
            if path:
                trimmed["full_result_path"] = path
            s2 = _dumps(trimmed)
            if len(s2) <= limit:
                return s2, True
    head = max(200, limit - 300)
    wrap = {"truncated": True, "chars": len(s), "preview": s[:head]}
    if path:
        wrap["full_result_path"] = path
    return _dumps(wrap), True


# ─── In-turn compaction of old tool results ───────────────────────────────────
# Tool results older than the last KEEP_LAST are folded to a stub, in batches
# (once per KEEP_LAST iterations) so the cached prefix breaks once per window.
# Only for providers without server-side context editing: Claude sends
# context_management instead (5.5 models reject client-edited history with
# preserved thinking).
_FOLD_KEEP_LAST = 6
_FOLD_MIN_CHARS = 200
_FOLD_STUB = u"[omitted: earlier tool result, {} chars — re-run the tool if you need it]"


def _is_tool_result_msg(m):
    if not isinstance(m, dict):
        return False
    if m.get("role") == "tool":                       # OpenAI / Ollama shape
        return True
    c = m.get("content")
    return (m.get("role") == "user" and isinstance(c, list) and
            any(isinstance(b, dict) and b.get("type") == "tool_result" for b in c))


def _fold_text(text):
    if isinstance(text, (u"".__class__, "".__class__)) and \
            len(text) > _FOLD_MIN_CHARS and not text.startswith(u"[omitted:"):
        return _FOLD_STUB.format(len(text)), True
    return text, False


def _fold_old_tool_results(messages, keep_last=_FOLD_KEEP_LAST):
    """Stub all but the last `keep_last` tool results, in place; returns the
    count folded. Assistant/thinking/tool_use blocks are never touched."""
    idxs = [i for i, m in enumerate(messages) if _is_tool_result_msg(m)]
    old = idxs[:-keep_last] if keep_last > 0 else idxs
    folded = 0
    for i in old:
        m = messages[i]
        if m.get("role") == "tool":
            m["content"], did = _fold_text(m.get("content"))
            folded += int(did)
            continue
        new_content = []
        for blk in m.get("content") or []:
            if isinstance(blk, dict) and blk.get("type") == "tool_result":
                blk = dict(blk)
                blk["content"], did = _fold_text(blk.get("content"))
                folded += int(did)
            new_content.append(blk)
        m["content"] = new_content
    return folded


def _should_fold(iteration, keep_last=_FOLD_KEEP_LAST):
    """True on iterations keep_last+1, 2*keep_last+1, ..."""
    return iteration > keep_last and (iteration - 1) % keep_last == 0


def _result_to_json(result, limit=_MAX_RESULT_CHARS, spill=None, tool=None):
    """Bounded result JSON for the model; on overflow `spill(tool, result)`
    saves the full copy first so the shrunken JSON can point at it."""
    s = _dumps(result)
    if len(s) <= limit:
        return s
    path = None
    if spill is not None:
        try:
            path = spill(tool, result)
        except Exception:
            path = None
    text, _ = _shrink_result(result, limit, path)
    return text


# ─── Parameter-error self-correction ───────────────────────────────────────────
# Small local models repeat a call with bad args; the tool's schema is attached
# to the error result. Bounded per run, at most once per tool.
_MAX_PARAM_HINTS = 3


def _param_hint(tools, name):
    """Compact 'expected parameters' line for `name`, or u''.

    Reads Anthropic (`input_schema`) and OpenAI/Ollama (`function.parameters`)
    shapes.
    """
    schema = None
    for t in tools or []:
        if not isinstance(t, dict):
            continue
        if t.get("name") == name and isinstance(t.get("input_schema"), dict):
            schema = t["input_schema"]
            break
        fn = t.get("function")
        if isinstance(fn, dict) and fn.get("name") == name \
                and isinstance(fn.get("parameters"), dict):
            schema = fn["parameters"]
            break
    if not isinstance(schema, dict):
        return u""
    props = schema.get("properties") or {}
    if not isinstance(props, dict) or not props:
        return u""
    required = set(schema.get("required") or [])
    parts = []
    for pname in list(props.keys())[:12]:
        pspec = props[pname] if isinstance(props[pname], dict) else {}
        bits = [u"{}".format(pspec.get("type") or u"any")]
        if pspec.get("enum"):
            vals = u"|".join(u"{}".format(v) for v in list(pspec["enum"])[:8])
            bits.append(u"enum: " + vals)
        parts.append(u"{} ({}, {})".format(
            pname, u"required" if pname in required else u"optional",
            u", ".join(bits)))
    return u"Expected parameters for `{}`: {}".format(name, u"; ".join(parts))


# ─── Phantom tool calls ────────────────────────────────────────────────────────
# A {"name", "parameters"} blob in chat text is a tool call: executed if the
# name is real (_text_tool_call), else never shown — one correction per retry.
_MAX_PHANTOM_RETRIES = 2

# ─── Announced-but-not-performed turns ────────────────────────────────────────
# "Đang liệt kê các mức…" with no tool call is narration, not an answer (once
# ended "list levels" with no level shown); zero-tool prose is ungrounded.
_MAX_ANNOUNCE_NUDGES = 1
# Server-run tools (Claude's tool search, remote MCP servers) can end a turn
# with stop_reason "pause_turn": the API stopped a long server-side loop and
# resumes it when the assistant message is sent back as is. Bounded so a
# server that keeps pausing cannot spin the loop.
_MAX_PAUSE_RESUMES = 5

_ANNOUNCE_PATTERNS = (
    # Vietnamese
    'dang liet ke', 'dang lay', 'dang kiem tra', 'dang tim', 'dang tai',
    'dang truy xuat', 'dang thuc hien', 'dang tien hanh', 'dang xu ly',
    'dang chay', 'toi se', 'minh se', 'de toi', 'de minh', 'se tien hanh',
    'cho chut', 'cho mot chut', 'vui long doi', 'doi mot lat',
    # English
    'let me', "i'll ", 'i will ', 'i am going to', "i'm going to",
    'i am now', "i'm now", 'going to check', 'one moment', 'hold on',
    'give me a moment', 'let us ', "let's ",
)

_ANNOUNCE_FIXUP = (
    u"You described what you were about to do but called no tool, so nothing "
    u"happened and the user received no data. Do not narrate — act. Call the "
    u"tool that answers the request NOW, then report its real result. If no "
    u"available tool can do it, say plainly that you cannot and why."
)


def _announces_work(text):
    """True when `text` promises an action instead of delivering one.

    Narrow: a question or a long reply never counts.
    """
    if not text:
        return False
    body = u"{}".format(text).strip()
    if len(body) > 220 or body.endswith(u'?'):
        return False
    from Intelligence.knowledge import vi_text
    folded = vi_text.fold_diacritics(body).lower()
    for p in _ANNOUNCE_PATTERNS:
        if p in folded:
            return True
    # A trailing ellipsis on a short line is the same promise.
    return body.endswith(u'...') or body.endswith(u'…')


# ── "Taught the user to do it by hand" guard ─────────────────────────────────
# The long counterpart of _announces_work: a tutorial of Revit clicks instead of
# doing the work (for "tô đỏ tường" it walked through a nonexistent "Bulk Set
# Parameter" dialog) is a refusal with extra steps.
_HOWTO_MARKERS = (
    # Vietnamese (diacritics folded before matching)
    'huong dan', 'cac buoc', 'buoc 1', 'tung buoc', 'ban co the su dung',
    'ban hay', 'nhap chuot', 'nhan chuot', 'chon cong cu', 'thanh cong cu',
    'hop thoai', 'tren ribbon', 'menu ngu canh', 'lam theo cac buoc',
    # English
    'here are the steps', 'follow these steps', 'step 1', 'step-by-step',
    'you can use the', 'click on the', 'right-click', 'in the ribbon',
    'dialog box', 'from the menu', 'select the tool',
)

# The user ASKED for instructions: a walkthrough is then right.
_ASKS_HOWTO = (
    'lam the nao', 'lam sao', 'cach ', 'huong dan toi', 'chi toi cach',
    'how do i', 'how to', 'how can i', 'what are the steps', 'explain how',
)

_HOWTO_FIXUP = (
    u"You explained how the USER could do it by hand in Revit's interface. "
    u"That is a refusal — they asked YOU to do it, and you have tools that "
    u"perform this directly on the live model. Never describe ribbon buttons, "
    u"dialogs or click sequences. Call the tool that performs the request NOW "
    u"and report the real result. If genuinely no tool can do it, say that in "
    u"one sentence and stop — do not substitute a tutorial."
)


def _asks_for_instructions(text):
    """True when the USER asked how to do something themselves."""
    if not text:
        return False
    from Intelligence.knowledge import vi_text
    folded = vi_text.fold_diacritics(u"{}".format(text)).lower()
    return any(p in folded for p in _ASKS_HOWTO)


def _instructs_instead_of_acting(text, user_text=None):
    """True when the reply is a do-it-yourself walkthrough (any length).

    Needs a how-to marker AND a numbered list or second marker, so a passing
    mention ("click Reload in pyRevit") cannot trip it.
    """
    if not text or _asks_for_instructions(user_text):
        return False
    from Intelligence.knowledge import vi_text
    body = u"{}".format(text)
    folded = vi_text.fold_diacritics(body).lower()
    hits = sum(1 for p in _HOWTO_MARKERS if p in folded)
    if not hits:
        return False
    numbered = len(re.findall(r'(?m)^\s*\d+[\.\)]\s', body)) >= 2
    return hits >= 2 or numbered


_PHANTOM_FIXUP = (
    u"`{}` is not a tool, and JSON written as chat text is never executed. "
    u"An Active skill in your system prompt is a set of INSTRUCTIONS for YOU "
    u"to carry out — \"playbook\", \"skill\" and \"apply_playbook\" are not "
    u"tool names. Do the work NOW: use a real tool call from the tools list "
    u"for anything that reads or edits the model, and plain prose for "
    u"everything else. Do not output JSON as text again."
)


def _find_json_blob(text):
    """(obj, start, end) for the outermost {...} in `text`, else None.

    Fenced ```json``` is skipped: that is the model SHOWING json on request.
    """
    if not text or u"{" not in text or u"```" in text:
        return None
    try:
        import re
        m = re.search(r'\{[\s\S]*\}', text)
        if not m:
            return None
        return json.loads(m.group()), m.start(), m.end()
    except Exception:
        return None


def _tool_call_shape(obj):
    """(name, args) if `obj` is a tool call written as JSON, else None.

    Accepts Anthropic ({"name","parameters"|"input"}), OpenAI
    ({"function": {"name","arguments"}}) and legacy T3Lab ({"intent","params"}).
    """
    if not isinstance(obj, dict):
        return None
    fn = obj.get("function")
    if isinstance(fn, dict):
        obj = fn
    name = obj.get("name") or obj.get("tool") or obj.get("intent")
    args = (obj.get("parameters") or obj.get("arguments")
            or obj.get("params") or obj.get("input") or {})
    if not name or not isinstance(args, dict):
        return None
    try:
        name = name.strip()
    except Exception:
        return None          # name wasn't a string
    return (name, args) if name else None


# Keys a JSON tool call may carry; any other key means DATA with a "name"
# field (a level, a family type...), not a call.
_CALL_KEYS = frozenset((
    "name", "tool", "intent", "function", "type", "id",
    "parameters", "arguments", "params", "input",
    "message", "thought", "reasoning", "explanation",
))


def _only_call_keys(obj):
    try:
        return not (set(obj.keys()) - _CALL_KEYS)
    except Exception:
        return False


def is_bare_tool_call_text(content):
    """True when `content` is only a tool call written as JSON text, e.g.
    {"name": "apply_playbook", "parameters": {...}}.

    The Assistant's history loader prunes these from old transcripts: replayed,
    they teach the model to answer with a fake tool call.
    """
    if not content:
        return False
    try:
        stripped = content.strip()
    except Exception:
        return False
    if not (stripped.startswith(u"{") and stripped.endswith(u"}")):
        return False
    found = _find_json_blob(stripped)
    if not found:
        return False
    return bool(_tool_call_shape(found[0])) and _only_call_keys(found[0])


def _call_key(name, args):
    """Canonical identity of one tool call (duplicate guard + read prefetch
    share it, so they agree on "already done")."""
    try:
        return (name, json.dumps(_json_safe(args), sort_keys=True,
                                 ensure_ascii=False))
    except Exception:
        try:
            return (name, u"{}".format(_json_safe(args)))
        except Exception:
            return (name, u"<unkeyable>")


def _sanitize_history(history, limit=None, budget_tokens=None):
    """Reduce persisted chat history to plain-text user/assistant messages.

    Trimmed by token budget, not a fixed count: `budget_tokens` is that of the
    provider ACTUALLY answering (from AgentLoop.run), so switching to a smaller
    model mid-session cannot overflow it (the dialog folds the rest into the
    rolling summary, Intelligence.conversation). `limit` = explicit message count.
    """
    hist = list(history or [])
    if limit:
        hist = hist[-limit:]
    out = []
    for h in hist:
        role    = h.get("role", "")
        content = h.get("content", "")
        if role not in ("user", "assistant"):
            continue
        if isinstance(content, list):
            parts = [b.get("text", "") for b in content
                     if isinstance(b, dict) and b.get("type") == "text"]
            content = u"\n".join(parts)
        if content:
            out.append({"role": role, "content": content})
    if budget_tokens:
        try:
            from Intelligence.context_budget import fit_messages
            out, _dropped = fit_messages(out, budget_tokens)
        except Exception:
            pass
    # Provider APIs reject a transcript that starts with an assistant turn.
    while out and out[0]["role"] == "assistant":
        out.pop(0)
    return out


# ─── Tool gate: who has to approve a call before it runs ──────────────────────
# ONE policy for every path that runs a tool for the user: native AgentLoop and
# the pane's legacy JSON-intent loop both ask ToolGate (they once disagreed, and
# "Ask before edits" was enforced only by the prompt). The gate only DECIDES;
# the approval card is the window's job, so this half is pure and testable.

GATE_RUN         = 'run'           # execute now, nobody is asked
GATE_DESTRUCTIVE = 'destructive'   # always confirm, in every action mode
GATE_EDIT        = 'edit'          # "Ask before edits" is on: confirm first

ACTION_MODES = ('auto', 'confirm')

# Calls that only REPORT although the tool can write (a purge dry run touches
# nothing; ToolGate.prepare forces the first purge into one).
_REPORT_ONLY_CALLS = {
    'purge_unused': ('dry_run', True),
}


def _is_report_only(name, args):
    rule = _REPORT_ONLY_CALLS.get(name)
    if not rule:
        return False
    key, default = rule
    try:
        return bool((args or {}).get(key, default))
    except Exception:
        return False


class ToolGate(object):
    """Pre-execution policy for ONE user request.

    action_mode:       'auto' | 'confirm' (settings agents.action_mode).
    is_destructive:    callable(name, args) -> bool
                       (core.server T3LabAIServer.is_destructive).
    is_modifying:      callable(name) -> bool; default
                       tool_schema.is_model_modifying (unknown = modifying).
    preapproved_edits: True while an approved plan runs (edits were OK'd on
                       the plan card); destructive calls still confirm.

    Fails CLOSED: a classifier that raises means confirm, never a silent run.
    """

    def __init__(self, action_mode='auto', is_destructive=None,
                 is_modifying=None, preapproved_edits=False):
        self.action_mode = (action_mode if action_mode in ACTION_MODES
                            else 'auto')
        self._is_destructive = is_destructive
        self._is_modifying = is_modifying or is_model_modifying
        self.preapproved_edits = bool(preapproved_edits)
        self._purge_seen = False

    def verdict(self, name, args=None):
        """GATE_RUN | GATE_DESTRUCTIVE | GATE_EDIT for one call. No side
        effects."""
        if not name or name in _LOCAL_PSEUDO_TOOLS:
            return GATE_RUN
        args = args or {}
        if self._is_destructive is not None:
            try:
                if self._is_destructive(name, args):
                    return GATE_DESTRUCTIVE
            except Exception:
                return GATE_DESTRUCTIVE
        if (self.action_mode == 'confirm' and not self.preapproved_edits
                and self.modifies(name, args)
                and not is_file_output_only(name)):
            # Exports only write a new file and run anyway
            # (tool_schema.FILE_OUTPUT_TOOL_NAMES).
            return GATE_EDIT
        return GATE_RUN

    def modifies(self, name, args=None):
        """True when this call can change the model or write a file."""
        if not name or name in _LOCAL_PSEUDO_TOOLS:
            return False
        if _is_report_only(name, args or {}):
            return False
        try:
            return bool(self._is_modifying(name))
        except Exception:
            return True

    def prepare(self, name, args):
        """(args, verdict) for a call that is ABOUT to run.

        The first purge_unused of a request is forced to a dry run; only a
        second, explicit (destructive, so confirmed) call can delete.
        """
        args = dict(args or {})
        if name == 'purge_unused' and not self._purge_seen:
            self._purge_seen = True
            if not bool(args.get('dry_run', True)):
                args['dry_run'] = True
        return args, self.verdict(name, args)


def tool_def_name(tool):
    """Name of one provider-native tool definition (Anthropic `{"name"}`,
    OpenAI-style `{"function": {"name"}}`), or u''."""
    try:
        return (tool.get("name")
                or (tool.get("function") or {}).get("name") or u"")
    except Exception:
        return u""


def read_only_tools(tools, is_modifying=None):
    """`tools` minus every definition that can change the model or write a
    file (background read tasks have no approval card)."""
    is_modifying = is_modifying or is_model_modifying
    out = []
    for t in tools or []:
        name = tool_def_name(t)
        try:
            if name and not is_modifying(name):
                out.append(t)
        except Exception:
            pass
    return out


def declined_result(name):
    """The tool result handed back to the model when the user said no."""
    return {"cancelled": True,
            "note": u"User declined the '{}' action. Do not retry it; say "
                    u"what was skipped and ask what to change.".format(name)}


# ─── One answer from the UI thread to a waiting worker ────────────────────────

class PendingDecision(object):
    """A one-shot decision the UI thread hands to a blocked worker thread
    (plan / confirm cards).

    FIRST decision wins: a Run click just after Stop, or a late click after a
    timeout, must not overturn it.
    """

    STOPPED = 'stopped'
    EXPIRED = 'expired'

    def __init__(self):
        self._evt   = threading.Event()
        self._lock  = threading.Lock()
        self._value = None

    def decide(self, value):
        """Record `value` unless already decided; True if THIS call decided."""
        with self._lock:
            if self._evt.is_set():
                return False
            self._value = value
            self._evt.set()
            return True

    @property
    def value(self):
        return self._value

    def is_decided(self):
        return self._evt.is_set()

    def wait(self, timeout_sec, should_abort=None, poll=0.25):
        """Block until decided, aborted (STOPPED) or timed out (EXPIRED);
        returns the value. `should_abort` (the window's Stop flag) is polled
        between waits."""
        deadline = time.time() + max(0.0, float(timeout_sec or 0))
        while not self._evt.is_set():
            if should_abort is not None:
                try:
                    if should_abort():
                        self.decide(self.STOPPED)
                        break
                except Exception:
                    pass
            remaining = deadline - time.time()
            if remaining <= 0:
                self.decide(self.EXPIRED)
                break
            self._evt.wait(min(poll, remaining))
        return self._value


class AgentLoop(object):
    """One user request = one AgentLoop.run(). Cancellable between steps."""

    def __init__(self, provider, execute_tool, tools, callbacks=None,
                 max_iterations=10, max_tokens=1500, time_budget_sec=240,
                 turn_timer=None, execute_tools_batch=None,
                 is_write_tool=None, result_chars=None, spill_result=None,
                 compact_results=False):
        self._provider       = provider
        self._execute_tool   = execute_tool
        self._tools          = tools
        # Client-side folding (_fold_old_tool_results); off for Claude.
        self._compact        = bool(compact_results)
        # Per-result cap (cloud 32k, local 4k) + full-copy seam (_result_to_json).
        self._result_chars   = int(result_chars or _MAX_RESULT_CHARS)
        self._spill          = spill_result
        # Optional read-batching seam; engages only when both are supplied.
        self._execute_batch  = execute_tools_batch
        self._is_write_tool  = is_write_tool
        self._cb             = callbacks or {}
        self._max_iterations = max_iterations
        self._max_tokens     = max_tokens
        self._time_budget    = time_budget_sec
        self._cancelled      = False
        # User-wait seconds, excluded from the budget (pause() / resume()).
        self._paused_total   = 0.0
        self._paused_since   = None
        # Optional Intelligence.telemetry.TurnTimer; every use is guarded.
        self._timer          = turn_timer

    # ── Cancellation ──────────────────────────────────────────────────────────

    def cancel(self):
        """Request a stop after the current model turn / tool (a Transaction
        mid-flight cannot be safely aborted)."""
        self._cancelled = True

    def is_cancelled(self):
        return self._cancelled

    # ── Time budget ───────────────────────────────────────────────────────────

    def pause(self):
        """Stop the budget clock while a confirm card waits for the user."""
        if self._paused_since is None:
            self._paused_since = time.time()

    def resume(self):
        if self._paused_since is not None:
            self._paused_total += time.time() - self._paused_since
            self._paused_since = None

    def _elapsed(self, started):
        """Budget seconds used so far, user-wait time excluded."""
        paused = self._paused_total
        if self._paused_since is not None:
            paused += time.time() - self._paused_since
        return (time.time() - started) - paused

    def _guard_tripped(self):
        """True when the optional guard_check callback says the context is
        stale (e.g. the user switched Revit documents mid-request)."""
        guard = self._cb.get("guard_check")
        if guard is None:
            return False
        try:
            return bool(guard())
        except Exception:
            return False

    # ── Callback helpers (never raise into the loop) ──────────────────────────

    def _emit(self, name, *args):
        fn = self._cb.get(name)
        if fn:
            try:
                fn(*args)
            except Exception:
                pass

    # ── Telemetry (optional, never raises into the loop) ──────────────────────

    def _note_usage(self, usage, model=None):
        """Add one call's token usage to the turn timer and record the model
        the provider really used (may differ from the badge) for pricing."""
        if self._timer is None:
            return
        try:
            if model:
                self._timer.model = u"{}".format(model)
            if usage:
                self._timer.add_usage(usage)
        except Exception:
            pass

    def _note_roundtrip(self, n_calls=1):
        """Record one crossing to Revit's main thread carrying n tool calls."""
        if self._timer is None:
            return
        try:
            self._timer.tool_roundtrip(n_calls)
        except Exception:
            pass

    # ── Read batching ─────────────────────────────────────────────────────────

    def leading_read_run(self, calls, done_calls=None):
        """Indices of the LEADING run of batchable read-only calls ([] if < 2).

        Stops at the first write, pseudo-tool or duplicate, which preserves
        read-after-write ordering.
        """
        if self._execute_batch is None or self._is_write_tool is None:
            return []
        done = done_calls or set()
        idxs = []
        for i, tc in enumerate(calls or []):
            name = tc.get("name", "")
            if not name or name in _LOCAL_PSEUDO_TOOLS:
                break
            try:
                if self._is_write_tool(name):
                    break
            except Exception:
                break
            if _call_key(name, tc.get("args") or {}) in done:
                break
            idxs.append(i)
        return idxs if len(idxs) > 1 else []

    def _prefetch_reads(self, calls, done_calls):
        """Execute the leading read run in ONE Revit round-trip; {index: result}.

        The sequential loop still emits per-call events and applies its guards;
        only the crossing is shared. Any failure falls back to per-call runs.
        """
        idxs = self.leading_read_run(calls, done_calls)
        if not idxs:
            return {}
        batch = [(calls[i].get("name", ""), calls[i].get("args") or {})
                 for i in idxs]
        try:
            results = self._execute_batch(batch)
        except Exception:
            return {}
        if not isinstance(results, list) or len(results) != len(idxs):
            return {}
        self._note_roundtrip(len(idxs))
        return dict(zip(idxs, results))

    # ── Text tool-call rescue ─────────────────────────────────────────────────

    def _registered_tool_names(self):
        """Tool names from the provider-native schema list."""
        names = set()
        for t in (self._tools or []):
            try:
                n = t.get("name") or (t.get("function") or {}).get("name")
                if n:
                    names.add(n)
            except Exception:
                pass
        return names

    def _text_tool_call(self, text):
        """A tool call written as chat text instead of a native block (local
        models after a tool error; hosted ones "calling" a /slash playbook).

        Returns {"name", "args", "prose", "known"} or None. `prose` = text with
        the blob removed (never shown); `known` = registered tool (execute) vs
        phantom (correct the model).
        """
        found = _find_json_blob(text)
        if not found:
            return None
        obj, start, end = found
        shape = _tool_call_shape(obj)
        if not shape:
            return None
        name, args = shape
        known = name in self._registered_tool_names()
        # Unregistered name: suppress only if the object has ONLY call keys.
        if not known and not _only_call_keys(obj):
            return None
        return {"name": name, "args": args, "known": known,
                "prose": (text[:start] + u" " + text[end:]).strip()}

    # ── Main loop ─────────────────────────────────────────────────────────────

    def add_tools(self, tools):
        """Make more tools callable from the next round (find_tools on a
        provider without tool_reference); returns the names added."""
        have = set()
        for t in (self._tools or []):
            if isinstance(t, dict):
                have.add(t.get("name") or (t.get("function") or {}).get("name"))
        added = []
        for t in tools or []:
            if not isinstance(t, dict):
                continue
            n = t.get("name") or (t.get("function") or {}).get("name")
            if n and n not in have:
                self._tools = list(self._tools or []) + [t]
                have.add(n)
                added.append(n)
        return added

    def run(self, history, system_prompt, user_content):
        """Execute the agentic loop. Blocking — call from a worker thread.

        Returns:
            {"status": "done"|"cancelled"|"failed"|"max_iterations"|"timeout"
                       |"doc_changed",
             "text": <last user-visible text>,
             "launch_intent": <str|None>,   # open_t3lab_tool target, terminal
             "iterations": int, "tool_runs": int}
        """
        # History gets the window LEFT after the fixed part (system prompt +
        # tools: ~6k–18k tokens locally); recorded so the dialog folds to it.
        try:
            from Intelligence import context_budget as _cb
            _fixed = _cb.request_overhead(system_prompt, self._tools)
            _cb.note_overhead(getattr(self._provider, 'NAME', u''), _fixed)
            _budget = _cb.history_budget_for(
                self._provider,
                overhead=_fixed + _cb.message_tokens({'content': user_content}),
                reply_tokens=self._max_tokens)
        except Exception:
            _budget = None
        messages = _sanitize_history(history, budget_tokens=_budget)
        messages.append({"role": "user", "content": user_content})

        started    = time.time()
        last_text  = u""
        tool_runs  = 0
        iteration  = 0
        done_calls = set()   # (name, canonical args) that succeeded this run
        phantom_retries = 0  # JSON calls to tools that don't exist, corrected
        announce_nudges = 0  # "I'm about to…" replies that called nothing
        pause_resumes = 0    # pause_turn continuations of server-run tools
        hinted = set()       # tools whose schema we already attached to an error

        while iteration < self._max_iterations:
            iteration += 1

            if self._cancelled:
                return self._finish("cancelled", last_text, None, iteration, tool_runs)
            if self._guard_tripped():
                return self._finish("doc_changed", last_text, None, iteration, tool_runs)
            if self._elapsed(started) > self._time_budget:
                return self._finish("timeout", last_text, None, iteration, tool_runs)

            # Blocking providers (no deltas) still emit on_text_delta once.
            turn = {"streamed": False}

            def _on_delta(chunk, _turn=turn):
                _turn["streamed"] = True
                self._emit("on_text_delta", chunk)

            if self._compact and _should_fold(iteration):
                try:
                    _fold_old_tool_results(messages)
                except Exception:
                    pass

            try:
                resp = self._provider.chat_agent(
                    system_prompt, messages, self._tools,
                    on_delta=_on_delta, max_tokens=self._max_tokens)
            except Exception:
                resp = None

            if resp is None:
                # Later turns failing still leave earlier work worth reporting.
                status = "failed" if iteration == 1 else "done"
                return self._finish(status, last_text, None, iteration, tool_runs)

            self._note_usage(resp.get("usage"), resp.get("model"))

            text  = (resp.get("text") or u"").strip()
            calls = resp.get("tool_calls") or []

            if (not calls and resp.get("stop_reason") == "pause_turn"
                    and pause_resumes < _MAX_PAUSE_RESUMES):
                pause_resumes += 1
                if text:
                    last_text = text
                    if not turn["streamed"]:
                        self._emit("on_text_delta", text)
                    self._emit("on_turn_text", text, False)
                messages.append(resp["assistant_msg"])
                continue

            # Tool call written as JSON text: strip the blob from the reply;
            # a registered name is executed anyway.
            rescued = False
            phantom  = None
            if not calls and text:
                _tc = self._text_tool_call(text)
                if _tc and _tc["known"]:
                    calls   = [{"id": "text_rescue", "name": _tc["name"],
                                "args": _tc["args"]}]
                    rescued = True
                    text    = _tc["prose"]
                elif _tc:
                    phantom = _tc
                    text    = _tc["prose"]

            # Phantom (usually `apply_playbook` from a /slash skill): correct
            # and retry, bounded; falling through leaves an empty turn, which
            # script.py routes to the legacy JSON-intent path.
            if phantom is not None and phantom_retries < _MAX_PHANTOM_RETRIES:
                phantom_retries += 1
                if text:
                    last_text = text
                # Replaces (or, if empty, drops) the bubble that streamed the blob.
                self._emit("on_turn_text", text, False)
                messages.append(resp["assistant_msg"])
                messages.append({"role": "user",
                                 "content": _PHANTOM_FIXUP.format(phantom["name"])})
                continue

            # Narrated instead of acted (no tool in this run): nudge once,
            # then let the text stand.
            if (not calls and tool_runs == 0
                    and announce_nudges < _MAX_ANNOUNCE_NUDGES
                    and _announces_work(text)):
                announce_nudges += 1
                last_text = text
                self._emit("on_turn_text", text, False)
                messages.append(resp["assistant_msg"])
                messages.append({"role": "user", "content": _ANNOUNCE_FIXUP})
                continue

            # Taught Revit clicks instead of acting: shares the nudge budget.
            if (not calls and tool_runs == 0
                    and announce_nudges < _MAX_ANNOUNCE_NUDGES
                    and _instructs_instead_of_acting(text, user_content)):
                announce_nudges += 1
                last_text = text
                self._emit("on_turn_text", text, False)
                messages.append(resp["assistant_msg"])
                messages.append({"role": "user", "content": _HOWTO_FIXUP})
                continue

            if text:
                last_text = text
                if not turn["streamed"]:
                    self._emit("on_text_delta", text)
                self._emit("on_turn_text", text, not calls)
            elif rescued or phantom is not None:
                # The turn was only the blob: drop its live bubble.
                self._emit("on_turn_text", u"", False)

            messages.append(resp["assistant_msg"])

            if not calls:
                return self._finish("done", last_text, None, iteration, tool_runs)

            # ── Execute tool calls sequentially (Revit is single-threaded) ──
            # except the LEADING read run, prefetched in one crossing (each
            # costs an ExternalEvent raise + wait for idle); cards and guards
            # below are unchanged.
            prefetched    = self._prefetch_reads(calls, done_calls)
            results       = []
            launch_intent = None
            doc_changed   = False
            for _idx, tc in enumerate(calls):
                name = tc.get("name", "")
                args = tc.get("args") or {}

                if self._guard_tripped():
                    doc_changed = True
                    results.append({"cancelled": True,
                                    "note": "Active document changed — request aborted."})
                    break

                if name == LAUNCHER_TOOL_NAME:
                    # Terminal: the window opens AFTER the loop (ShowDialog blocks).
                    launch_intent = args.get("tool_intent") or ""
                    results.append({"success": True,
                                    "note": "T3Lab tool window will open now."})
                    break

                if self._cancelled:
                    results.append({"cancelled": True,
                                    "note": "User stopped the request."})
                    break

                # Duplicate guard: an identical call that succeeded in THIS run
                # is never re-run (error result, no tool card).
                call_key = _call_key(name, args)
                if call_key in done_calls:
                    results.append({
                        "error": u"Duplicate call: `{}` with identical "
                                 u"arguments already succeeded in this "
                                 u"request. Do not repeat completed calls — "
                                 u"continue with the next step or give the "
                                 u"final answer.".format(name)})
                    continue

                self._emit("on_tool_start", name, args, iteration)
                t0 = time.time()
                if _idx in prefetched:
                    res = prefetched[_idx]
                else:
                    self._note_roundtrip(1)
                    try:
                        res = self._execute_tool(name, args)
                    except Exception as ex:
                        res = {"error": u"{}".format(ex), "tool": name}
                if not isinstance(res, dict):
                    res = {"result": res}
                dt = time.time() - t0
                ok = "error" not in res
                # Attach the tool schema to the error (see _MAX_PARAM_HINTS).
                if not ok and name not in hinted and len(hinted) < _MAX_PARAM_HINTS:
                    _ph = _param_hint(self._tools, name)
                    if _ph:
                        res["expected_parameters"] = _ph
                        hinted.add(name)
                tool_runs += 1
                self._emit("on_tool_done", name, res, ok, dt)
                results.append(res)
                if ok:
                    done_calls.add(call_key)

            if launch_intent:
                return self._finish("done", last_text, launch_intent,
                                    iteration, tool_runs)
            if doc_changed:
                return self._finish("doc_changed", last_text, None,
                                    iteration, tool_runs)
            if self._cancelled:
                return self._finish("cancelled", last_text, None,
                                    iteration, tool_runs)

            # Feed results back in the provider's native format, bounded here.
            try:
                res_strs = [_result_to_json(r, self._result_chars, self._spill,
                                            c.get("name"))
                            for c, r in zip(calls, results)]
                if rescued:
                    # No native tool_call block to reference: plain user text.
                    messages.append({
                        "role": "user",
                        "content": u"Tool `{}` returned: {}\nContinue the "
                                   u"task. Use REAL tool calls, never JSON "
                                   u"as text.".format(calls[0]["name"],
                                                      res_strs[0])})
                else:
                    messages.extend(
                        self._provider.agent_tool_results(calls, res_strs))
            except Exception:
                return self._finish("failed", last_text, None, iteration, tool_runs)

        return self._finish("max_iterations", last_text, None,
                            self._max_iterations, tool_runs)

    def _finish(self, status, text, launch_intent, iterations, tool_runs):
        if self._timer is not None:
            try:
                self._timer.iterations = iterations
            except Exception:
                pass
        result = {
            "status":        status,
            "text":          text,
            "launch_intent": launch_intent,
            "iterations":    iterations,
            "tool_runs":     tool_runs,
        }
        self._emit("on_finish", result)
        return result


# ─── System prompt (native tool-calling mode) ──────────────────────────────────
# No tool schemas (they go in the API `tools` param) and no forced JSON output.

_AGENT_PROMPT = u"""You are T3Lab Assistant, an AI agent embedded in Autodesk Revit via the T3Lab pyRevit extension. You can read and modify the live Revit model through the tools provided.

## Language & formatting
{language}Keep replies short and practical — one or two sentences between tool calls, a compact summary at the end.
Use markdown when it helps: **bold**, `code`, bullet lists, and pipe tables (| a | b |) for numeric summaries — the chat renders them natively. Do NOT use emoji.

## Units
All tool coordinates and dimensions are in METERS. Convert user input: 5000mm = 5.0, 3m = 3.0. Element ids are integers.

## You DO the work — you never teach it
This is the most important rule. The user is asking you to change their model, not to learn Revit. NEVER answer an action request with instructions, a numbered walkthrough, or the name of a ribbon button, dialog or menu. Do not write "you can use...", "click...", "select the tool...", "here are the steps". You are not a manual; you hold tools that act on the live model, and a tutorial is a refusal dressed up as help. "tô đỏ tường" means CALL revit_override_color, not explain how the user could paint walls themselves. Explain steps ONLY when the user explicitly asked how to do it by hand ("làm thế nào", "how do I"). If no tool can do what was asked, say so in one sentence and stop.

Do not invent safety rules either. Asking "are you sure?" for work that is not on the destructive list in Working rule 3 is the same refusal in a different costume. Graphic overrides (`revit_override_color`, `color_elements`, `operate_element` hide/isolate/halftone), tagging, creating views/sheets/levels/grids and setting parameters are ORDINARY edits: just do them, however many elements they touch. "toàn bộ", "tất cả", "all" describes the SCOPE the user wants — it is not a reason to stop and confirm.

## Plan, then act
For a multi-step request, decide the full tool sequence BEFORE the first call and state it in one short sentence. Chain the steps yourself — never ask the user to run intermediate steps you can do with tools. For a simple request, skip the plan and just act.

## Efficiency (each model turn is expensive — minimize turns)
- BATCH independent read-only calls into ONE turn: when you need e.g. levels + selection + view info, emit all those tool calls together instead of one per turn.
- Prefer ONE bulk tool over many single calls: `bulk_set_parameter` over repeated `set_parameter`, `ai_element_filter` over per-element inspection, `tag_all_*` over per-element tags.
- ONE `ai_element_filter` call, fully narrowed, beats a broad call plus your own sifting. It filters by `category`/`categories`, `group_name`, `workset`, `level_name`, `type_name`, `name_contains` and a parameter — all at once, all server-side. Never pull a whole category down and pick through it yourself: that is slower, costs thousands of tokens, and the list you sift may be a truncated page.
- NEVER FERRY IDS. To act on the elements a filter found, re-call it with `store_as="sel"` and pass `element_ids=["@sel"]` to the acting tool — the whole matched set, uncapped, at the cost of one word. Ids belong in the reply only when the user asked for specific ones.
- COUNTS COME FROM `group_by`, NOT FROM ROWS. "How many…", "thống kê", any breakdown: call `ai_element_filter` with `group_by` (e.g. `["name","level"]`, `["workset"]`) and read the `summary` counts. Never fetch 1000 element rows to count them.
- Trim what you do list: `fields=["id"]` returns a bare id array; add only the fields you will actually use.
- A result with `truncated: true` was cut to fit: `omitted_items` says how many rows were dropped per field and `full_result_path` names the saved full copy. Narrow the query (`fields`, `group_by`, `limit`/`offset`, `store_as`) instead of asking for the file.
- Do not re-query data that is already in this conversation (earlier tool results, the context block in the latest user message) unless the model may have changed since.
- Never call a tool "just to check" after a success result already told you the outcome; verify with a read tool only after LARGE modifications (20+ elements) or when a result looks suspicious.

## Working rules
1. Query before you modify: when element ids or names are unknown, use read tools (get_current_view_elements, ai_element_filter, list_levels, ...) first.
2. After finishing, summarize WHAT changed (counts + element ids) in English.
3. Destructive actions — `delete_element`, `purge_unused`, `edit_elements` ungroup, `manage_view_template` delete, `manage_links` delete, `manage_document` save_as / sync_with_central, and anything else removing or publishing model data: unless the user's current message already explicitly requested it, ask for confirmation in text FIRST and stop — do not call the tool in the same turn. (Revit cannot undo an ungroup by re-grouping, deleting a link takes its tags and dimensions with it, and syncing pushes your work to the whole team.)
4. If a tool returns an error, diagnose from the error text and retry ONCE with corrected arguments; if it fails again, report exactly what is missing — never loop on the same failing call.
5. If the request is ambiguous about WHICH elements to change, ask ONE precise clarifying question instead of guessing — a wrong bulk edit is worse than a question.
6. `open_t3lab_tool` opens a T3Lab window and ENDS your turn — only ever call it last, and never together with other tools. OPENING A WINDOW IS NOT DOING THE WORK: call it only when the user asked to OPEN/see a tool by name ("mở BatchOut", "open View Manager"). When they asked for a RESULT ("create 3d views for each room", "export the A sheets"), do it with the real tools and report what you produced — handing them a window to fill in themselves is a silent refusal dressed up as success. If no tool can do it, say exactly that.
7. When the user refers to the current selection ("these elements", "the selected walls", "các element này", "đang chọn"), call `revit_get_selected_elements` FIRST and operate on those element ids — never guess ids.
8. Trust tool results over assumptions: counts, names and ids come from the model, not from memory of typical projects.
9. Scope resolution when the request names no explicit target: for CHECKING / auditing / statistics / spell-check requests the default scope is the ENTIRE PROJECT — never silently limit to the current view (limit only when the user says so, or ask ONE question when unsure). For MODIFY actions, use the "Current Revit context" block in the latest user message — the user's selection first, then the active view. Never report "nothing found" from one filtered query: re-check with corrected arguments or wider scope first.
10. NUMBERS MUST COME FROM TOOLS, NEVER FROM YOU. Tool results may be truncated before you see them, so counting or summing rows/elements yourself gives wrong answers. Use the exact fields the tools compute: `total_count` (ai_element_filter), `row_count` / `column_totals` (get_schedule_data), `element_counts` (analyze_model_statistics). If a statistic you need has no tool-computed field, say so — do not estimate.
11. Color / highlight requests — "tô màu", "tô đỏ X", "bôi xanh X", "đổi màu X", "highlight X", "color X red": the color word is the OVERRIDE color to APPLY, never a parameter value to filter by (Revit elements have no "Color" or "Fill Color" parameter). For ALL elements of a category ("tô vàng sàn", "color the walls red") call `revit_override_color` with `category` + `color` directly — the server collects every matching element itself, no ids needed and NO count limit. Pass `element_ids` only for a SUBSET: the user's selection (via revit_get_selected_elements), or `["@sel"]` after `ai_element_filter(..., store_as="sel")` — never a hand-copied id list, which is capped at one page. Use `color_elements` only when the user wants elements colored BY a parameter's values (one distinct color per value, e.g. "tô màu tường theo Type") — and its parameter_name must be a REAL parameter, never empty, never a color name.
12. NEVER fabricate tool activity. Every action happens ONLY through a real tool call in this conversation — if you say you will use a tool, emit that tool call in the same turn, and only report results that came back from an actual tool result. Writing an invented "Result:" for a call you never made is a critical failure.
13. HISTORY IS NOT A TO-DO LIST: act only on the LATEST user message. Earlier turns are background context — an action already reported successful there is DONE; never re-execute it unless the user explicitly asks again. Each new command fully replaces the previous one: after "tô đỏ tường" (walls red) succeeded, "tô vàng sàn" targets FLOORS + YELLOW only — walls and red are finished business.
14. NEVER SAY A FILTER IS IMPOSSIBLE — `ai_element_filter` combines category/categories + `group_name` + `workset` + `level_name` + `type_name` + `name_contains` + parameter in ONE server-side call. "Elements inside the groups named _F1_" is `group_name="_F1_"` (the only tool that can answer it); "elements on workset X" is `workset="X"`. Do not fall back to listing everything and sifting, and do not tell the user no such filter exists.
    Then act on the result WITHOUT ids:
    - WHOLE-CATEGORY actions need no ids at all: `revit_override_color`, `operate_element`, `edit_elements`, `select_elements`, `set_element_workset`, `bulk_set_parameter` and `tag_elements` all accept `category` and collect every match themselves — one call, no count limit.
    - ANY OTHER subset: re-call `ai_element_filter` with the same filters plus `store_as="sel"`, then pass `element_ids=["@sel"]` to the acting tool. The full set goes through, uncapped, and no id enters the conversation. Pass literal ids only for a handful you already have (the user's selection, a couple picked out by name).
    Category names are the standard Revit ones (Walls, StructuralFoundations, Ducts, Areas, Rebar, ...) — if one is rejected, the error lists `did_you_mean`; retry with one of those instead of giving up or substituting another category.
15. NEVER SUBSTITUTE A DIFFERENT ACTION. Match the VERB the user asked for, not the nearest tool you happen to have.
   - `operate_element` (changes only how the ACTIVE VIEW looks): pin/ghim/khoá → "pin" (unpin/bỏ ghim → "unpin"); hide/ẩn → "hide"; isolate/cô lập → "isolate"; unhide/hiện lại → "unhide"; "bỏ isolate" / "hiện lại hết" → "reset_temporary" (needs no target); ẩn cả category trong view → "hide_category"; halftone/làm mờ → "halftone"; độ trong suốt → "transparency"; select/chọn → "select"; "chọn tương tự" → "select_similar".
   - `edit_elements` (changes the MODEL): mirror/lật/đối xứng → "mirror"; "đổi type/đổi sang loại X" → "change_type"; group/nhóm lại → "group"; ungroup/rã nhóm → "ungroup".
   - `manage_view`: đổi tỉ lệ/scale → "set_scale"; mức độ chi tiết → "set_detail_level"; discipline → "set_discipline"; crop/vùng cắt → "set_crop".
   - `manage_view_template` manages the templates themselves (list/rename/duplicate/delete); applying one to views is `apply_view_template`.
   - `manage_links` for linked models and CAD links (list/reload/unload/delete/pin); `manage_revision` for revisions (list/create/assign_to_sheets/set_issued); `manage_sheet` for duplicate/renumber/list_sets.
   - `export_model` for IFC / NWC / DWF / DGN — PDF stays `export_sheets_pdf`, DWG `export_dwg`, PNG `export_image`.
   - `check_bad_geometry` when an export crashes or geometry looks wrong; `manage_material` to read materials; `create_detail_annotation` for filled regions and detail lines; `manage_document` for save / save as / sync.
   - Only tô màu / color / highlight means an override color.
   If NO tool performs the requested action, say plainly that the assistant cannot do it yet and stop — doing a different thing to the model and reporting it as done is a critical failure, worse than doing nothing.

## Multiple open models
Several documents can be open in this Revit session; every tool operates on the ACTIVE one. To work across models: `list_open_documents` → `switch_active_document` (exact title) → read/modify there → switch again as needed, then combine everything into ONE final answer stating which model each number comes from. Element ids are only valid inside their own document — never reuse ids across models. When the user names a model that is not in `list_open_documents`, it may live in a different Revit window (separate process): say so and point them to the assistant in that window — do not guess.

## Live context
The latest user message carries a "Current Revit context" block (active view, selection, project) and, when relevant, "Reference from project knowledge" excerpts. Read them as the state of the model RIGHT NOW — they are not something the user typed.
"""


# Few-shot for LOCAL models on the default (no-specialist) path; small models
# pick tools far better with concrete traces. Specialists carry their own.
_LOCAL_GENERAL_FEWSHOT = u"""
## Examples (follow this tool-calling style)
User: "có bao nhiêu cửa?" -> call ai_element_filter(category="Doors"); read total_count; reply "Model có 84 cửa."
User: "chọn các element đang chọn và ẩn đi" -> call revit_get_selected_elements FIRST; then operate_element(operation="hide", element_ids=[...]); reply what was hidden.
User: "tô đỏ tường" -> call revit_override_color(category="Walls", color="red") ONE time (server collects every wall, no ids, no limit); reply "Đã tô đỏ 128 tường."
User: "pin toàn bộ tường trong view" (= LOCK the walls, NOT a color) -> call operate_element(operation="pin", category="Walls") ONE time; reply "Đã pin 54 tường trong view."
User: "cô lập cột" then later "bỏ isolate đi" -> first operate_element(operation="isolate", category="Columns"); then operate_element(operation="reset_temporary") with NO category and NO ids — that is the only way back out of isolate.
User: "tạo section mới" -> call create_view(view_type="section"); never fall back to a 3D view.
User: "mở batchout" -> call open_t3lab_tool(tool_intent="open_batchout") LAST and stop.
Rules shown above still apply: query before you modify, numbers come from tool fields (total_count/...), never invent a tool or a result.
"""


# Language directives; must match the language the UI renders.
_LANG_EN = (u"Always reply in English, regardless of the language the user "
            u"writes in. ")
_LANG_VI = (u"Always reply in Vietnamese, regardless of the language the user "
            u"writes in. ")
_LANG_AUTO = (u"Reply in the SAME language the user wrote in — Vietnamese in, "
              u"Vietnamese out; English in, English out. ")


def build_agent_system_prompt(revit_context=u"", local=False, lang="auto"):
    """STATIC system prompt for the native tool-calling agent path.

    Static so the prompt cache hits across turns (the system block is a cache
    breakpoint before `messages`); live state goes with the user turn via
    build_context_block() (measured in Intelligence/telemetry.py).
    `revit_context` is accepted and ignored, for old callers.
    local=True appends the few-shot traces; lang is 'auto' | 'vi' | 'en' and
    must match what the UI renders.
    """
    language = {'vi': _LANG_VI, 'en': _LANG_EN}.get(lang, _LANG_AUTO)
    prompt = _AGENT_PROMPT.format(language=language)
    if local:
        prompt += u"\n" + _LOCAL_GENERAL_FEWSHOT
        prompt += teacher_exemplar_block()
    return prompt


# ─── Portable teacher exemplars (few-shot distilled from teaching) ───────────────

def teacher_exemplar_block():
    """Teacher-exemplar few-shot block (leading newline), or u"". Never raises.

    From git-tracked lib/Intelligence/config/teacher_exemplars.json, so any
    machine's local model (qwen3:14b) gets the taught style without the
    non-portable fine-tune. Static per session (cache-safe); local models only.
    """
    try:
        from Intelligence.learning import exemplars as _ex
        block = _ex.build_exemplar_block(_ex.load_exemplars())
        return (u"\n" + block) if block else u""
    except Exception:
        return u""


# Fences live state off from the user's words; strip_context_block removes it.
CONTEXT_FENCE_OPEN  = u"<<<T3LAB_LIVE_CONTEXT"
CONTEXT_FENCE_CLOSE = u"T3LAB_LIVE_CONTEXT>>>"


def build_context_block(revit_context=u"", knowledge_ref=u"",
                        analysis_hint=u""):
    """The VOLATILE half of the prompt, to prepend to the current user turn
    (kept out of the cached system block).

    `analysis_hint` = `Intelligence.language.Utterance.to_prompt_hint()`
    (prohibition, scope, named categories, confirm need); goes LAST, closest
    to the user's words. Returns u"" when there is nothing live.
    """
    parts = []
    ctx = (revit_context or u"").strip()
    if ctx:
        parts.append(u"## Current Revit context\n" + ctx)
    ref = (knowledge_ref or u"").strip()
    if ref:
        parts.append(ref)
    hint = (analysis_hint or u"").strip()
    if hint:
        parts.append(hint)
    if not parts:
        return u""
    return u"{}\n{}\n{}".format(CONTEXT_FENCE_OPEN,
                                u"\n\n".join(parts),
                                CONTEXT_FENCE_CLOSE)


def apply_context_block(user_content, context_block):
    """Prepend `context_block` to a turn's user content: a string, or the
    Anthropic block list used for images (rag_processor.build_vision_content_blocks),
    where it becomes a leading text block.
    """
    if not context_block:
        return user_content
    if isinstance(user_content, list):
        return [{"type": "text", "text": context_block}] + list(user_content)
    return u"{}\n\n{}".format(context_block, user_content or u"")


def strip_context_block(text):
    """Remove any live-context fence from `text` (belt-and-braces: history
    stores the raw user text)."""
    if not text or CONTEXT_FENCE_OPEN not in text:
        return text
    out = []
    depth = 0
    for line in text.split(u"\n"):
        stripped = line.strip()
        if stripped == CONTEXT_FENCE_OPEN:
            depth += 1
            continue
        if stripped == CONTEXT_FENCE_CLOSE:
            depth = max(0, depth - 1)
            continue
        if depth == 0:
            out.append(line)
    return u"\n".join(out).strip()
