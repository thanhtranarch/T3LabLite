# -*- coding: utf-8 -*-
"""T3Lab Tool Hub client — the team catalog of tools, scripts and recipes.

Server: Supabase project t3lab-space, schema `toolhub`, reached only through
the public.toolhub_* RPCs (dev/sql/toolhub_v1.sql). Modules:

  config    project URL / publishable key, local folder
  auth      e-mail OTP sign-in, refresh, session stored with DPAPI
  client    ToolHubClient.rpc() over the shared Supabase HTTP helper
  catalog   local cache of the catalog (delta sync by updated_at)
  search    offline BM25 search over the cache; online RPC when signed in
  events    uploads audit-log summaries (no args, no code, no file names)
  recipes   recipe format, renderer and runner (Phase 3)

Runs under CPython 3 and IronPython 2.7; no dependency beyond the stdlib.
"""
