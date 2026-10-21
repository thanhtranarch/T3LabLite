# -*- coding: utf-8 -*-
"""The find_tools agent tool: query -> the tools to put in front of the model.

run(query, limit, current_names, registry_rows, deferred, client)
    -> (result, names_to_add)

  deferred=True  (Claude with defer_loading): every registry tool is already in
                 tools[] as a deferred schema, so the result carries
                 {"__blocks__": [tool_reference…]} and claude_provider turns it
                 into tool_result content; the API loads those schemas.
  deferred=False (OpenAI / Ollama / LM Studio, or Claude without tool
                 search): a compact list of matches, and `names_to_add` names
                 registry tools the caller should add with AgentLoop.add_tools.

Only names that really exist (in tools[] or the registry) ever come back: a
tool_reference to an unknown name is a 400 from the API, and an invented
tool is the worst failure this assistant has.
"""
from __future__ import unicode_literals

from Intelligence.toolhub import search

KINDS = ['mcp_tool', 'recipe']


def _tool_name(hit):
    if hit.get('kind') == 'recipe':
        return 'recipe_' + hit['slug']
    return hit['slug']


def run(query, limit=5, current_names=(), registry_rows=None, deferred=False,
        client=None):
    query = (query or u'').strip()
    if not query:
        return {"error": "find_tools needs a query: what do you want to do?"}, []
    try:
        limit = max(1, min(int(limit or 5), 10))
    except (TypeError, ValueError):
        limit = 5
    current = set(current_names or [])
    registry = {r['slug']: r for r in (registry_rows or [])}
    hits = search.search(query, k=limit * 3, kinds=KINDS,
                         tools=[{'name': r['slug'], 'description': r.get('description', '')}
                                for r in registry.values()],
                         client=client)
    found, info = [], {}
    for h in hits:
        name = _tool_name(h)
        if name in found:
            continue
        if name not in current and name not in registry:
            continue                      # nothing callable by that name here
        found.append(name)
        info[name] = h
        if len(found) >= limit:
            break
    if not found:
        return {"matches": [], "note": u"No tool matches '{}'. Try other words, or "
                u"use the tools you already have.".format(query)}, []
    if deferred:
        return {"__blocks__": [{"type": "tool_reference", "tool_name": n} for n in found],
                "found": found}, []
    matches = []
    for n in found:
        row = registry.get(n, {})
        matches.append({"name": n,
                        "title_vi": row.get('title_vi', '') or info[n].get('title_vi', ''),
                        "description": (row.get('description') or '')[:240]})
    to_add = [n for n in found if n not in current]
    note = (u"These tools are now available to call." if to_add
            else u"These tools are already available to call.")
    return {"matches": matches, "note": note}, to_add
