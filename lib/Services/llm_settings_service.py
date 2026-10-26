# -*- coding: utf-8 -*-
"""
Non-UI work behind LLMs Setting (GUI/LLMSettingDialog.py): key save, the
connection test, project-context reads and the context-digest rebuild.

Pure Python, no WPF: the dialog keeps only marshalling to the UI thread.
Background-thread safe unless a function says otherwise.
Test: python3 dev/test_assistant_llm.py (llm_settings_service section).
"""
from __future__ import unicode_literals

import os

_SETTINGS_HINT = u"%APPDATA%\\T3LabAI\\settings.json"


# ─── API key ──────────────────────────────────────────────────────────────────

def mask_key(key):
    """Stored-key mask for the key box: only the last 4 characters show."""
    if not key:
        return u""
    return u"•" * 8 + (key[-4:] if len(key) > 8 else u"")


def persist_api_key(settings_key, key, provider=None, settings=None):
    """Save an already validated key and make `provider` re-read it.

    Returns None on success, else (short hint, full tooltip) in English —
    set_api_key returns False when settings.json is locked or unreadable.
    """
    if settings is None:
        from config.settings import T3LabAISettings
        settings = T3LabAISettings()
    try:
        saved = settings.set_api_key(settings_key, key)
    except Exception:
        saved = False
    if not saved:
        where = getattr(settings, '_settings_file', _SETTINGS_HINT)
        return (u"Key not saved: settings.json locked — close other Revit, retry",
                u"The key works but could not be saved to {} — the file is "
                u"locked or read-only. Close other Revit sessions, then "
                u"click Save Key again.".format(where))
    if provider is not None and hasattr(provider, "reload_credentials"):
        provider.reload_credentials()
    elif provider is not None and hasattr(provider, "invalidate_models_cache"):
        provider.invalidate_models_cache()
    return None


# ─── Test Connection ──────────────────────────────────────────────────────────

def connection_test(router=None):
    """Probe the active provider (default: the LLMRouter singleton) with a
    one-sentence chat. Returns (ok, label, message). Never raises.
    """
    ok, label, msg = False, u"Result", u""
    try:
        if router is None:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
        name = router.get_active_name()
        provider = router.get_active_provider()

        if provider is None:
            msg = u"Provider '{}' not loaded.".format(name)
        elif not provider.check_health():
            if name == "ollama":
                # Recommend the tool-capable sweet spot, not a tiny NLU-only
                # model — the agentic path needs reliable tool calls.
                rec = u"qwen3:14b"
                try:
                    from Intelligence import local_llm as _ll
                    rec = _ll.recommended_tool_model()
                except Exception:
                    pass
                msg = (u"Ollama not available or no models installed.\n"
                       u"1. Make sure Ollama is running.\n"
                       u"2. Run: ollama pull {}".format(rec))
            elif name == "lmstudio":
                msg = (u"LM Studio not available or no model loaded.\n"
                       u"1. Open LM Studio.\n"
                       u"2. Load a model in LM Studio first.")
            elif name == "custom":
                msg = (u"Custom server not reachable.\n"
                       u"Check the Base URL and API key — Azure needs "
                       u"the api-key option and a model (deployment) "
                       u"name.")
            else:
                msg = u"Provider not reachable.\nCheck API key or service status."
        else:
            active_model = None
            try:
                active_model = provider.get_active_model()
            except Exception:
                pass
            if not active_model:
                msg = u"No model selected. Choose a model from the Model dropdown."
            else:
                resp = provider.chat(
                    [],
                    u"You are a concise assistant. Do not think. Reply in one short sentence only.",
                    u"Reply with exactly this sentence: 'Connected OK'",
                    max_tokens=120,
                )
                if resp and resp.strip():
                    ok, label, msg = True, u"Connected", resp.strip()[:120]
                else:
                    msg = u"Provider responded but returned an empty reply."
    except Exception as ex:
        msg = u"Error: {}".format(str(ex)[:100])
    return ok, label, msg


# ─── Project context ──────────────────────────────────────────────────────────

def collect_project_context(pid, on_error=None):
    """The document counter + one digest-stats dict per linked folder.

    on_error(ex) is told when the project store cannot be read.
    """
    info = {'pid': pid, 'summary': u'', 'rows': []}
    ps = None
    dirs = []
    try:
        from config.project_store import ProjectStore
        ps = ProjectStore()
        info['summary'] = ps.describe_documents(pid, cap=99)
        dirs = ps.get_knowledge_dirs(pid)
    except Exception as ex:
        if on_error is not None:
            on_error(ex)
    for path in dirs:
        row = {'path': path, 'stats': None,
               'missing': not os.path.isdir(path)}
        # linked_dir_stats is the same TTL-cached read describe_documents
        # just did, so the rows cost no extra share round trips.
        if ps is not None and not row['missing']:
            try:
                row['stats'] = ps.linked_dir_stats(path)
            except Exception:
                pass
        info['rows'].append(row)
    return info


def build_context_digests(todo, pid=None, project_ctx=True, set_status=None,
                          on_error=None, debug=None):
    """Rebuild <folder>/context/CONTEXT.md for each folder in `todo`, then
    (pid + project_ctx) the project-wide PROJECT_CONTEXT. Long-running.

    set_status(text) reports progress; on_error(ex) gets the failures that
    stop a step; debug(msg) gets per-file failures. Returns (results, ctx_info).
    """
    def _status(text):
        if set_status is not None:
            set_status(text)

    results = []
    ctx_info = []
    try:
        from Intelligence.knowledge import context_digest
        for i, folder in enumerate(todo, start=1):
            if not os.path.isdir(folder):
                continue
            label = os.path.basename(folder.rstrip(u'\\/')) or folder
            prefix = (u"[{}/{}] ".format(i, len(todo))
                      if len(todo) > 1 else u"")

            def _prog(name, _p=prefix, _l=label):
                _status(u"{}{}: {}".format(_p, _l[:18], name[:26]))
            _status(u"{}Reading {}…".format(prefix, label[:24]))
            try:
                res = context_digest.build_context_file(
                    folder, progress_cb=_prog)
            except Exception as bex:
                if debug is not None:
                    debug(u"digest {} error: {}".format(folder, bex))
                res = None
            if res:
                results.append(res)
    except Exception as wex:
        if on_error is not None:
            on_error(wex)

    # Project-wide BIM context (steps 3+4 of the RAG cycle): per-folder
    # digests summarise ONE file at a time; this pass answers a fixed set of
    # BIM-standard questions from the best passages across ALL linked
    # folders, with file+page citations.
    if pid and project_ctx:
        try:
            _status(u"Indexing for project context…")
            from config.project_store import ProjectStore as _PS
            from Intelligence.knowledge import project_context as _pc
            ps = _PS()
            store = ps.knowledge_store_for(pid)
            if store is not None:
                store.scan()
                embedder = None
                try:
                    from Intelligence.knowledge.embeddings import (
                        get_default_embedder)
                    emb = get_default_embedder()
                    if emb is not None and emb.is_available():
                        store.embed_pending(emb, budget_sec=120)
                        embedder = emb
                except Exception:
                    embedder = None
                chat = None
                try:
                    from Intelligence.knowledge.context_digest import (
                        _default_chat_fn)
                    chat = _default_chat_fn()
                except Exception:
                    chat = None

                def _tprog(topic):
                    _status(u"Context: {}".format(topic[:34]))
                ctx = _pc.build_project_context(
                    store, chat, embedder=embedder,
                    progress_cb=_tprog)
                out_dir = os.path.join(ps.project_dir(pid), 'files')
                ctx_path = _pc.write_project_context(out_dir, ctx)
                if ctx_path:
                    ctx_info.append(ctx)
        except Exception as pex:
            if on_error is not None:
                on_error(pex)
    return results, ctx_info


def digest_summary(results, ctx_info, todo):
    """Status line after a rebuild, or None when nothing was asked."""
    if results:
        docs = sum(r.get('files') or 0 for r in results)
        n_llm = sum(r.get('llm') or 0 for r in results)
        n_pg = sum(r.get('pages') or 0 for r in results)
        n_bad = sum(r.get('skipped') or 0 for r in results)
        n_re = sum(r.get('reused') or 0 for r in results)
        msg = (u"Context rebuilt: {} doc(s), {} page(s) read "
               u"in {} folder(s){}"
               .format(docs, n_pg, len(results),
                       u" · LLM {}".format(n_llm) if n_llm
                       else u" · excerpt only (no LLM)"))
        if n_re:
            msg += u" · {} unchanged (reused)".format(n_re)
        if n_bad:
            msg += u" · {} unreadable (see CONTEXT.md)".format(n_bad)
        if ctx_info:
            c = ctx_info[0]
            msg += u" · PROJECT_CONTEXT {}/{} topics".format(
                c.get('topics_found') or 0, c.get('topics_total') or 0)
        return msg
    if todo:
        return u"Nothing indexable found in the linked folder(s)."
    return None


# ─── Skills installer ─────────────────────────────────────────────────────────

def plain_report(report):
    """Installer report as plain text for a dialog (no markdown)."""
    report = report or {}
    lines = []
    for entry in (report.get('installed') or []):
        lines.append(u"installed  /{}".format(entry['id']))
    for entry in (report.get('updated') or []):
        lines.append(u"updated    /{}".format(entry['id']))
    for sid, why in (report.get('skipped') or [])[:6]:
        lines.append(u"skipped    {} ({})".format(sid, why))
    for err in (report.get('errors') or [])[:4]:
        lines.append(u"error      {}".format(err))
    if not lines:
        lines.append(u"Nothing changed.")
    return u"\n".join(lines)
