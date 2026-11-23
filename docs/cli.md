# T3Lab CLI — Revit for agents with a shell (ported from t3lab_dev, 2026-11-23)

Ideas from [HKUDS/CLI-Anything](https://github.com/HKUDS/CLI-Anything) (Apache-2.0, ideas only, no
code copied): the agent discovers the tools itself (`tools`, `help`), calls them with typed flags,
gets compact output, and checks exported files instead of trusting exit 0. The CLI is the bridge
(`lib/core/bridge.py`, Lite's own, edited by hand); the shared helpers are `lib/core/tool_cli.py`
(identical to t3lab_dev's). **Not run in Revit yet.**

## Commands

`B` = `python -I "%APPDATA%\T3LabAI\bridge.py"`

| Command | Does |
|---|---|
| `B tools [words] [--limit=15] [--refresh]` | ranked tool names + one sentence; no words = every name on one line. Read from the cache, so it works with Revit closed |
| `B help <tool>` | usage, parameters, type, required, enum, default |
| `B call <tool> --a=1 --ids=1,2 --flag --no-flag` | flags typed by the schema (integer, number, boolean, array, JSON object); `--json '{...}'` / `--json-file=args.json` still work, flags win |
| `B instances` | the Revit windows and their open models (`list_revit_instances`) |
| `B verify <file\|folder\|pattern>... [--ext=pdf,dwg] [--since=30] [--expect=N] [--min-bytes=N]` | each file exists, is not empty and has the right header (`%PDF`, `AC10xx`, `ISO-10303-21`, OLE for RVT/RFA, ZIP for XLSX…); `--expect` = the count the tool reported |
| `B skill [--install=<skills folder>]` | prints / writes `t3lab-cli/SKILL.md` (AgentSkills) for Claude Code, Codex |
| `B hub …`, `B harness …` | unchanged: `docs/data-hub.md`, `docs/harness.md` |

Output: `--out=pretty` (default) `| compact | tsv` (header once, a row per line), `--pick=a,b` keeps
those columns of a list. A tool parameter named like a global flag (`json`, `agent`, `port`, `out`,
`pick`, also `tier`, `turbo`) goes in `--json`.

Exit codes: **0** ok · **1** tool or Revit error (`error_code`, `retryable`; also a harness or
read-only refusal) · **2** bad syntax or bad arguments, nothing ran.

## One path for every caller

`call` (and `instances`, every recipe step) builds one `tools/call` and sends it through `_answer`,
the function that answers the AI apps over stdio. So the bridge argument check (`bad_arguments`,
"did you mean"), the harness write check, the Agent Cursors stamp, and in Revit Lite's policy
(`core/mcp_policy.py`: read-only mode, toolsets, the confirmation of destructive calls) apply
exactly as for Claude Desktop. The CLI only adds a stricter check of its own flags (unknown flag,
wrong type, value not in the enum) before that.

**Who the CLI is:** `--agent=<id>` or `T3LAB_AGENT`, else `cli` — the person at this terminal,
person-facing like `mcp-client` (`HUMAN_AGENTS`), shown as "CLI" on Agent Cursors. It runs as
transport `cli`, like `bridge.py harness`. An AI agent driving the CLI with a harness task passes
its own name (`--agent=claude-code`), so the write check finds its task.

## Recipes (`t3lab_recipe`)

Team recipes from Tool Hub (`lib/Intelligence/toolhub/recipes.py`, deployed as
`toolhub_recipes.py`), read from the signed-in Assistant's cache
`%APPDATA%\T3LabAI\toolhub\catalog.json`; only published recipes whose sha256 matches the review.
`action=list | show | run` (`slug`, `inputs`, `confirm`). A recipe that changes the model runs only
with `confirm=true` after the person saw the steps; each step is one call through `_answer`, so a
step refused there (arguments, harness, read-only, a "No" in Revit) stops the recipe. A step can
never call `t3lab_recipe`; a launched harness agent cannot run one (it only reads). Listed in
`tools/list` only while the catalog holds a published recipe (Lite lists every bridge tool, so an
empty catalog costs no tokens); callable either way, e.g.
`B call t3lab_recipe --action=run --slug=<slug> --inputs={"level":"L1"} --confirm`.

## Not ported from t3lab_dev

`capabilities`, `jobs`, `approvals` / `approve` / `reject`, `token`, `serve-http`, `turbo`,
`copilot`, tiers / profiles, `t3lab_call` and the `core` / `full` toolsets: Lite's bridge has none
of them. The SKILL.md text in `tool_cli.py` names some (tier, `t3lab_result`, `--revit_document`);
the bridge swaps those lines (`SKILL_LITE`) instead of editing the shared file.

## Tests

```
python lite_guard/test_tool_cli.py        # contract, flags, output, verify, CLI gate + identity, skill
python lite_guard/test_bridge_recipes.py  # t3lab_recipe over MCP and the CLI, deploy
```

## NEEDS VERIFICATION

- With Revit open: `call` through flags, `verify` on a real BatchOut folder, a recipe synced by the
  Assistant, a destructive step asking in Revit.
