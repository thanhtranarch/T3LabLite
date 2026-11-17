# -*- coding: utf-8 -*-
"""Agent cursors "Glow" (experimental, off by default): outline the elements
the agents touched inside Revit's own 3D views, drawn by Revit through a
DirectContext3D server. The overlay only floats above the view; this sits in
the model and follows orbit / zoom.

update() / stop() run on Revit's main thread (agent_cursors.notify, inside
the MCP ExternalEvent). RenderScene re-checks the setting every frame, so
turning Glow or the cursors off stops the drawing at once even though the
server is only removed on the next tool call. Engine-agnostic; NOT verified
in Revit yet.
"""

from __future__ import unicode_literals

MAX_BOXES = 200
COLOR = (234, 104, 12)          # T3Lab orange
_SERVER_GUID = '6a0f9c1e-7b1d-4c39-9f3e-2b8d5e4a7c11'
# The 12 edges of a box whose 8 corners are ordered x-major (see _corners).
EDGES = ((0, 1), (2, 3), (4, 5), (6, 7), (0, 2), (1, 3), (4, 6), (5, 7),
         (0, 4), (1, 5), (2, 6), (3, 7))

_state = {'server': None, 'cls': None, 'boxes': [], 'doc': None, 'dirty': True}


def corners(lo, hi):
    """8 corners of an axis box, x-major: (x, y, z) for x, y, z in (lo, hi)."""
    return [(x, y, z) for x in (lo[0], hi[0]) for y in (lo[1], hi[1]) for z in (lo[2], hi[2])]


def line_indices(n_boxes):
    """Vertex index pairs for n boxes of 8 vertices each."""
    out = []
    for b in range(n_boxes):
        base = 8 * b
        out.extend((base + i, base + j) for i, j in EDGES)
    return out


def _wanted():
    try:
        from core import agent_cursors
        return agent_cursors.is_enabled() and agent_cursors.options().get('glow')
    except Exception:
        return False


def _server_class():
    if _state['cls'] is not None:
        return _state['cls']
    from System import Guid
    from Autodesk.Revit.DB import XYZ, Outline, View3D
    from Autodesk.Revit.DB.ExternalService import ExternalServices
    from Autodesk.Revit.DB.DirectContext3D import IDirectContext3DServer

    class _GlowServer(IDirectContext3DServer):
        __namespace__ = 'T3Lab.AgentCursorsGlow'

        def GetServerId(self):
            return Guid(_SERVER_GUID)

        def GetServiceId(self):
            return ExternalServices.BuiltInExternalServices.DirectContext3DService

        def GetName(self):
            return 'T3Lab Agent Cursors glow'

        def GetVendorId(self):
            return 'T3Lab'

        def GetDescription(self):
            return 'Outlines the elements AI agents touched (Agent Cursors).'

        def GetApplicationId(self):
            return ''

        def GetSourceId(self):
            return ''

        def UsesHandles(self):
            return False

        def UseInTransparentPass(self, view):
            return False

        def CanExecute(self, view):
            try:
                return (isinstance(view, View3D) and bool(_state['boxes'])
                        and view.Document.Title == _state['doc'] and _wanted())
            except Exception:
                return False

        def GetBoundingBox(self, view):
            try:
                boxes = _state['boxes']
                if not boxes:
                    return None
                lo = [min(b[0][k] for b in boxes) for k in range(3)]
                hi = [max(b[1][k] for b in boxes) for k in range(3)]
                return Outline(XYZ(*lo), XYZ(*hi))
            except Exception:
                return None

        def RenderScene(self, view, display_style):
            try:
                if _wanted() and _state['boxes']:
                    _render(_state['boxes'])
            except Exception:
                pass

    _state['cls'] = _GlowServer
    return _GlowServer


def _render(boxes):
    """One line list of every box edge, built per frame (few hundred lines)."""
    from Autodesk.Revit.DB import XYZ, ColorWithTransparency
    from Autodesk.Revit.DB.DirectContext3D import (
        DrawContext, VertexBuffer, IndexBuffer, VertexFormat, VertexFormatBits,
        VertexPositionColored, IndexLine, EffectInstance, PrimitiveType)
    n_vert = 8 * len(boxes)
    pairs = line_indices(len(boxes))
    vfloats = n_vert * VertexPositionColored.GetSizeInFloats()
    vb = VertexBuffer(vfloats)
    vb.Map(vfloats)
    stream = vb.GetVertexStreamPositionColored()
    color = ColorWithTransparency(COLOR[0], COLOR[1], COLOR[2], 0)
    for lo, hi in boxes:
        for c in corners(lo, hi):
            stream.AddVertex(VertexPositionColored(XYZ(*c), color))
    vb.Unmap()
    ishorts = len(pairs) * IndexLine.GetSizeInShortInts()
    ib = IndexBuffer(ishorts)
    ib.Map(ishorts)
    istream = ib.GetIndexStreamLine()
    for a, b in pairs:
        istream.AddLine(IndexLine(a, b))
    ib.Unmap()
    fmt = VertexFormat(VertexFormatBits.PositionColored)
    effect = EffectInstance(VertexFormatBits.PositionColored)
    DrawContext.FlushBuffer(vb, n_vert, ib, ishorts, fmt, effect,
                            PrimitiveType.LineList, 0, len(pairs))


def _service():
    from Autodesk.Revit.DB.ExternalService import ExternalServiceRegistry, ExternalServices
    return ExternalServiceRegistry.GetService(
        ExternalServices.BuiltInExternalServices.DirectContext3DService)


def _register():
    if _state['server'] is not None:
        return
    service = _service()
    server = _server_class()()
    service.AddServer(server)
    active = service.GetActiveServerIds()
    active.Add(server.GetServerId())
    service.SetActiveServers(active)
    _state['server'] = server


def stop(app=None):
    """Remove the server (main thread). Never raises."""
    server = _state['server']
    _state['server'] = None
    _state['boxes'] = []
    if server is None:
        return
    try:
        service = _service()
        active = service.GetActiveServerIds()
        if active.Contains(server.GetServerId()):
            active.Remove(server.GetServerId())
            service.SetActiveServers(active)
        service.RemoveServer(server.GetServerId())
    except Exception:
        pass
    _refresh(app)


def update(app, ids):
    """Outline `ids` of the active document (main thread). Never raises."""
    try:
        from Snippets._compat import make_eid
        uidoc = app.ActiveUIDocument
        doc = uidoc.Document
        boxes = []
        for i in ids[:MAX_BOXES]:
            elem = doc.GetElement(make_eid(i))
            bb = elem.get_BoundingBox(None) if elem is not None else None
            if bb is not None:
                boxes.append(((bb.Min.X, bb.Min.Y, bb.Min.Z), (bb.Max.X, bb.Max.Y, bb.Max.Z)))
        _state['boxes'] = boxes
        _state['doc'] = doc.Title
        if boxes:
            _register()
        _refresh(app)
    except Exception:
        pass


def _refresh(app):
    try:
        if app is not None:
            app.ActiveUIDocument.RefreshActiveView()
    except Exception:
        pass
