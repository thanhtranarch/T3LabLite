# -*- coding: utf-8 -*-
"""The MCP tool registry as T3LabAIServer builds it, without Revit.

Lite's registry literal names module constants (mcp_wiki.TOOL_NAME, ...), so it
is built by running _register_default_tools on a bare server object instead of
reading the literal by AST (t3lab_dev: dev/server_registry.py).
"""
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
SERVER = os.path.join(LIB, 'core', 'server.py')


def spec_entries():
    if LIB not in sys.path:
        sys.path.insert(0, LIB)
    from core import tool_specs
    return tool_specs.registry_entries()


def registry(src=None):
    """{name: {name, description, inputSchema}} — hand-written + spec tools."""
    if LIB not in sys.path:
        sys.path.insert(0, LIB)
    import core.server as S
    srv = object.__new__(S.T3LabAIServer)
    srv._tools = {}
    S.T3LabAIServer._register_default_tools(srv)
    return dict(srv._tools)


def literal(src=None):
    """The hand-written entries (no spec tools)."""
    specs = spec_entries()
    return dict((k, v) for k, v in registry().items() if k not in specs)
