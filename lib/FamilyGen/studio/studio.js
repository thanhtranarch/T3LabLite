/* FamiGen Studio page script. Talks to Python (FamilyGen/studio_controller.py)
   through pyRevit 7.1's WebView2 messages; renders the preview with three.js
   r158 (global THREE). Geometry arrives already tessellated in millimetres by
   FamilyGen/preview_mesh.py, so the page and the Revit builder read the same
   schema the same way. Protocol: docs/famigen-studio-spec.md section 6. */
(function () {
  'use strict';

  // ── bridge ────────────────────────────────────────────────────────────
  // pyRevit: window.chrome.webview.postMessage(obj) -> output.on_message;
  // output.post_message(obj) -> 'message' event (data may be a JSON string).
  // Outside WebView2 (tests, a browser) window.__studio is a stand-in.
  var bridge = (function () {
    var wv = window.chrome && window.chrome.webview;
    if (wv && typeof wv.postMessage === 'function') {
      return {
        send: function (m) { wv.postMessage(m); },
        listen: function (fn) {
          wv.addEventListener('message', function (e) {
            var d = e.data;
            if (typeof d === 'string') { try { d = JSON.parse(d); } catch (err) { return; } }
            fn(d);
          });
        }
      };
    }
    var listeners = [];
    window.__studio = { sent: [], receive: function (m) { listeners.forEach(function (f) { f(m); }); } };
    return {
      send: function (m) { window.__studio.sent.push(m); },
      listen: function (fn) { listeners.push(fn); }
    };
  })();

  function $(id) { return document.getElementById(id); }
  // Stroke icons in the page's own style (svg rule in studio.css), never emoji or text glyphs.
  var ICONS = {
    clip: 'M13 7.5l-5 5a3 3 0 0 1-4.2-4.2l5.3-5.3a2 2 0 0 1 2.8 2.8L6.6 11a1 1 0 0 1-1.4-1.4L10 4.8',
    doc: 'M4 1.5h5l3 3v10H4zM9 1.5v3h3',
    text: 'M4 1.5h5l3 3v10H4zM6 8h4M6 10.5h4M6 5.5h2',
    x: 'M4.5 4.5l7 7M11.5 4.5l-7 7',
    right: 'M6 4l4 4-4 4',
    down: 'M4 6l4 4 4-4'
  };
  function icon(name) {
    var ns = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(ns, 'svg'); svg.setAttribute('viewBox', '0 0 16 16'); svg.setAttribute('aria-hidden', 'true');
    var path = document.createElementNS(ns, 'path'); path.setAttribute('d', ICONS[name]); svg.appendChild(path);
    return svg;
  }
  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = text;
    return n;
  }
  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function fmt(n) { return Math.round(n).toLocaleString('en-US'); }
  function each(sel, fn) { Array.prototype.forEach.call(document.querySelectorAll(sel), fn); }

  var CURVED = { Cylinder: true, Revolution: true, Sweep: true, SweptBlend: true };
  var MODES = [['low', 'Fast'], ['medium', 'Balanced'], ['high', 'Best'], ['xhigh', 'Deep']];
  var PLOT = ['#2563EB', '#D97706', '#059669', '#DC2626', '#7C3AED', '#0891B2', '#C026D3', '#65A30D']; // T3.Plot.1-8

  var state = {
    busy: false, claudeReady: false, preview: null, selected: -1, type: '',
    json: '', picked: [], collapsed: { 'props:Geometry': true }, folder: '', style: 'shaded', detail: 'fine',
    templates: [], groups: [], typingNode: null, inited: false, attachCount: 0
  };

  // ── 3D scene (Revit axes: Z up, millimetres) ─────────────────────────
  var view = (function () {
    var host = $('viewport');
    var renderer = new THREE.WebGLRenderer({ antialias: true, alpha: true, preserveDrawingBuffer: true });
    renderer.setPixelRatio(Math.min(window.devicePixelRatio || 1, 2));
    host.insertBefore(renderer.domElement, host.firstChild);
    var scene = new THREE.Scene();
    var camera = new THREE.PerspectiveCamera(40, 1, 1, 1e6);
    camera.up.set(0, 0, 1);
    // Front / Right / Plan are drawings, as in Revit: parallel projection, no vanishing lines.
    var ortho = new THREE.OrthographicCamera(-1, 1, 1, -1, 1, 1e6);
    ortho.up.set(0, 0, 1);
    function cam() { return orbit.ortho ? ortho : camera; }
    scene.add(new THREE.HemisphereLight(0xffffff, 0x8a8a92, 1.15));
    var key = new THREE.DirectionalLight(0xffffff, 1.4); key.position.set(-0.6, -1.0, 1.4); scene.add(key);
    var fill = new THREE.DirectionalLight(0xffffff, 0.45); fill.position.set(1.0, 0.8, 0.4); scene.add(fill);
    var root = new THREE.Group(); scene.add(root);
    var helpers = new THREE.Group(); scene.add(helpers);
    var planeGroup = new THREE.Group(); scene.add(planeGroup);
    var box = null, grid = null;
    var meshes = [];          // {part, mesh, edges, lit, flat}
    var orbit = { target: new THREE.Vector3(0, 0, 0), radius: 3000, yaw: -0.9, pitch: 0.45, ortho: false };
    var ORTHO = { front: true, right: true, top: true };
    var bboxCache = null;
    var labelSpecs = [];      // {node, pos:[x,y,z]}

    function resize() {
      var w = host.clientWidth, h = host.clientHeight;
      if (!w || !h) return;
      renderer.setSize(w, h, false);
      camera.aspect = w / h; camera.updateProjectionMatrix();
      render();
    }
    function place() {
      var cp = Math.cos(orbit.pitch), sp = Math.sin(orbit.pitch);
      var dir = new THREE.Vector3(Math.cos(orbit.yaw) * cp, Math.sin(orbit.yaw) * cp, sp);
      var c = cam();
      if (orbit.ortho) {
        // Same framing as the perspective camera at the target's distance, so zoom and pan carry over.
        var halfH = orbit.radius * Math.tan(THREE.MathUtils.degToRad(camera.fov / 2));
        ortho.top = halfH; ortho.bottom = -halfH; ortho.left = -halfH * camera.aspect; ortho.right = halfH * camera.aspect;
      }
      c.position.copy(orbit.target).addScaledVector(dir, orbit.radius);
      c.near = Math.max(1, orbit.radius / 500); c.far = orbit.radius * 50;
      c.updateProjectionMatrix();
      c.lookAt(orbit.target);
    }
    function render() { place(); renderer.render(scene, cam()); updateLabels(); }

    var VIEWS = { iso: [-0.9, 0.45], front: [-Math.PI / 2, 0.0001], right: [0, 0.0001], top: [-Math.PI / 2, Math.PI / 2 - 0.001] };
    function fit(viewName) {
      resize();
      var b = bboxCache;
      var min = b ? new THREE.Vector3().fromArray(b[0]) : new THREE.Vector3(-500, -500, 0);
      var max = b ? new THREE.Vector3().fromArray(b[1]) : new THREE.Vector3(500, 500, 1000);
      orbit.target.copy(min).add(max).multiplyScalar(0.5);
      var r = Math.max(50, min.distanceTo(max) / 2);
      // The bounding sphere must fit the narrower of the two view angles.
      var vHalf = THREE.MathUtils.degToRad(camera.fov / 2);
      var hHalf = Math.atan(Math.tan(vHalf) * camera.aspect);
      orbit.radius = r / Math.sin(Math.min(vHalf, hHalf)) * 1.12;
      var v = viewName === 'ref' && state.refCam
        ? [state.refCam.azimuth_deg * Math.PI / 180, Math.max(-1.4, Math.min(1.5, state.refCam.elevation_deg * Math.PI / 180))]
        : (VIEWS[viewName || 'iso'] || VIEWS.iso);
      orbit.yaw = v[0]; orbit.pitch = v[1];
      orbit.ortho = !!ORTHO[viewName];
      orbit.hideAxis = { front: 'y', right: 'x', top: 'z' }[viewName] || null;
      // A floor grid seen edge-on is only a line through the elevation.
      if (grid) grid.visible = !(viewName === 'front' || viewName === 'right');
      render();
    }

    function disposeGroup(group) {
      while (group.children.length) {
        var c = group.children[0];
        if (c.geometry) c.geometry.dispose();
        if (c.material) c.material.dispose();
        group.remove(c);
      }
    }
    function clear() {
      meshes.forEach(function (m) { m.lit.dispose(); m.flat.dispose(); });
      meshes = [];
      disposeGroup(root); disposeGroup(helpers); disposeGroup(planeGroup);
      box = null; grid = null;
      $('labels').innerHTML = ''; labelSpecs = [];
    }

    function buildGrid(b) {
      var size = 1000, step = 100;
      if (b) {
        var span = Math.max(b[1][0] - b[0][0], b[1][1] - b[0][1], 1);
        size = Math.pow(10, Math.ceil(Math.log10(span * 2)));
        step = size / 20;
      }
      var g = new THREE.GridHelper(size, Math.round(size / step), new THREE.Color(css('--grid-major')), new THREE.Color(css('--grid-minor')));
      g.rotation.x = Math.PI / 2;
      if (b) g.position.set((b[0][0] + b[1][0]) / 2, (b[0][1] + b[1][1]) / 2, Math.min(0, b[0][2]));
      return g;
    }

    // axis: the size the label measures; Front / Right / Plan hide the one along the view
    // direction (D in Front, W in Right, H in Plan), which they cannot show.
    function addLabel(cls, text, pos, title, axis) {
      var n = el('div', cls, text);
      if (title) n.title = title;
      $('labels').appendChild(n);
      labelSpecs.push({ node: n, pos: pos, cls: cls, axis: axis || null });
    }

    // Reference planes as dashed outlines a little larger than the model,
    // labelled dimensions as a parameter tag between their two planes.
    function buildPlanes(p) {
      var b = p.bbox || [[-500, -500, 0], [500, 500, 1000]];
      var pad = Math.max(60, 0.06 * Math.max(b[1][0] - b[0][0], b[1][1] - b[0][1], b[1][2] - b[0][2]));
      var lo = [b[0][0] - pad, b[0][1] - pad, Math.min(0, b[0][2]) - pad], hi = [b[1][0] + pad, b[1][1] + pad, b[1][2] + pad];
      var color = new THREE.Color(css('--plane'));
      // Template planes stay out of the picture; only the family's own planes show.
      (p.planes || []).filter(function (pl) { return !pl.template; }).forEach(function (pl) {
        var o = pl.offset, pts;
        if (pl.axis === 'x') pts = [[o, lo[1], lo[2]], [o, hi[1], lo[2]], [o, hi[1], hi[2]], [o, lo[1], hi[2]]];
        else if (pl.axis === 'y') pts = [[lo[0], o, lo[2]], [hi[0], o, lo[2]], [hi[0], o, hi[2]], [lo[0], o, hi[2]]];
        else pts = [[lo[0], lo[1], o], [hi[0], lo[1], o], [hi[0], hi[1], o], [lo[0], hi[1], o]];
        var g = new THREE.BufferGeometry().setFromPoints(pts.concat([pts[0]]).map(function (q) { return new THREE.Vector3(q[0], q[1], q[2]); }));
        var line = new THREE.Line(g, new THREE.LineDashedMaterial({ color: color, dashSize: pad / 3, gapSize: pad / 4,
          transparent: true, opacity: 0.55 }));
        line.computeLineDistances();
        planeGroup.add(line);
        addLabel('plane-label', pl.name, pts[3], pl.name + ': ' + pl.axis + ' = ' + fmt(pl.offset) + ' mm', pl.axis);
      });
      (p.dimensions || []).forEach(function (d) {
        var mid = (d.from + d.to) / 2, pos;
        if (d.axis === 'x') pos = [mid, b[0][1] - pad / 2, b[0][2]];
        else if (d.axis === 'y') pos = [b[1][0] + pad / 2, mid, b[0][2]];
        else pos = [b[1][0] + pad / 2, b[0][1] - pad / 2, mid];
        addLabel('dim-param', d.parameter + ' = ' + fmt(d.value), pos, 'Labelled dimension: ' + d.parameter, d.axis);
      });
    }

    function load(preview) {
      clear();
      bboxCache = preview.bbox;
      var fallback = new THREE.Color(css('--border-strong'));
      (preview.parts || []).forEach(function (part, i) {
        var geo = new THREE.BufferGeometry();
        geo.setAttribute('position', new THREE.Float32BufferAttribute(part.pos, 3));
        geo.setIndex(part.idx);
        geo.computeVertexNormals();
        var common = { side: THREE.DoubleSide, polygonOffset: true, polygonOffsetFactor: 1, polygonOffsetUnits: 1 };
        var lit = new THREE.MeshStandardMaterial(Object.assign({ color: fallback.clone(), roughness: 0.75, metalness: 0.05, flatShading: true }, common));
        var flat = new THREE.MeshBasicMaterial(Object.assign({ color: fallback.clone() }, common));
        var mesh = new THREE.Mesh(geo, lit);
        mesh.userData.index = i;
        root.add(mesh);
        var edges = new THREE.LineSegments(new THREE.EdgesGeometry(geo, CURVED[part.kind] ? 50 : 28),
          new THREE.LineBasicMaterial({ color: new THREE.Color(css('--ink')), transparent: true }));
        root.add(edges);
        meshes.push({ part: part, mesh: mesh, edges: edges, lit: lit, flat: flat });
      });
      grid = buildGrid(preview.bbox); helpers.add(grid);
      if (preview.bbox) {
        var b3 = new THREE.Box3(new THREE.Vector3().fromArray(preview.bbox[0]), new THREE.Vector3().fromArray(preview.bbox[1]));
        box = new THREE.Box3Helper(b3, new THREE.Color(css('--muted')));
        helpers.add(box);
        // Plain W / D / H only where no labelled dimension already tells that size.
        var a = preview.bbox[0], c = preview.bbox[1], s = preview.size_mm;
        var labelled = {};
        (preview.dimensions || []).forEach(function (d) { labelled[d.axis] = true; });
        if (!labelled.x) addLabel('dim-label', 'W ' + fmt(s[0]), [(a[0] + c[0]) / 2, a[1], a[2]], null, 'x');
        if (!labelled.y) addLabel('dim-label', 'D ' + fmt(s[1]), [c[0], (a[1] + c[1]) / 2, a[2]], null, 'y');
        if (!labelled.z) addLabel('dim-label', 'H ' + fmt(s[2]), [c[0], a[1], (a[2] + c[2]) / 2], null, 'z');
      }
      buildPlanes(preview);
      apply();
    }

    // Visual style + detail level + type materials + selection, like Revit's view control bar.
    function apply() {
      var p = state.preview; if (!p) { render(); return; }
      var colors = {};
      (p.materials || []).forEach(function (m) { if (m.color) colors[m.name] = m; });
      var swap = (state.type && p.type_materials && p.type_materials[state.type]) || {};
      var fallback = css('--border-strong'), voidColor = css('--bad-accent');
      var paper = css('--surface'), ink = css('--ink'), sel = css('--select');
      var style = state.style;
      meshes.forEach(function (m, i) {
        // Void Forms are hidden unless asked for: in Revit they cut, they are not seen.
        var shown = (m.part.levels || ['coarse', 'medium', 'fine']).indexOf(state.detail) >= 0
                    && (m.part.solid || $('chk-voids').checked);
        var picked = state.picked.indexOf(i) >= 0;
        var name = swap[m.part.material] || m.part.material;
        var def = colors[name];
        var color = !m.part.solid ? voidColor : (def ? def.color : fallback);
        var t = m.part.solid ? (def ? Math.min(90, def.transparency || 0) : 0) : 72;
        var mat;
        if (style === 'hidden') { mat = m.flat; mat.color.set(m.part.solid ? paper : voidColor); t = m.part.solid ? 0 : 80; }
        else if (style === 'consistent') { mat = m.flat; mat.color.set(color); }
        // Curved parts (cylinders, tubes, turned and swept forms) shade smoothly;
        // prismatic ones keep crisp faces, like Revit's shaded view.
        else { mat = m.lit; mat.color.set(color); mat.flatShading = style !== 'realistic' && !CURVED[m.part.kind];
               mat.roughness = style === 'realistic' ? 0.45 : 0.75; mat.metalness = style === 'realistic' ? 0.15 : 0.05; }
        if (picked) mat.color.lerp(new THREE.Color(sel), 0.55);
        mat.transparent = t > 0; mat.opacity = 1 - t / 100; mat.depthWrite = t === 0;
        mat.needsUpdate = true;
        m.mesh.material = mat;
        m.mesh.visible = shown && style !== 'wireframe';
        m.edges.visible = shown && style !== 'realistic';
        m.edges.material.color.set(picked ? sel : ink);
        m.edges.material.opacity = style === 'wireframe' ? 0.9 : (m.part.solid ? 0.6 : 0.3);
      });
      if (box) box.visible = $('chk-dims').checked;
      planeGroup.visible = $('chk-planes').checked;
      render();
    }

    function project(pos) {
      var v = new THREE.Vector3(pos[0], pos[1], pos[2]).project(cam());
      if (v.z > 1 || v.z < -1) return null;
      return [(v.x + 1) / 2 * host.clientWidth, (1 - v.y) / 2 * host.clientHeight];
    }
    function updateLabels() {
      var dims = $('chk-dims').checked, planes = $('chk-planes').checked;
      labelSpecs.forEach(function (l) {
        var on = (l.cls === 'dim-label' ? dims : planes) && !(l.axis && l.axis === orbit.hideAxis);
        var p = on ? project(l.pos) : null;
        l.node.classList.toggle('hidden', !p);
        if (p) { l.node.style.left = p[0] + 'px'; l.node.style.top = p[1] + 'px'; }
      });
    }

    // Render the current model from named views for the reference check.
    function snapshot(views) {
      var keep = { yaw: orbit.yaw, pitch: orbit.pitch, radius: orbit.radius, target: orbit.target.clone(), ortho: orbit.ortho, hideAxis: orbit.hideAxis,
                   grid: grid ? grid.visible : true };
      var helpersOn = helpers.visible, planesOn = planeGroup.visible, styleKeep = state.style, selKeep = state.picked;
      helpers.visible = false; planeGroup.visible = false; state.style = 'shaded'; state.picked = []; apply();
      var out = [];
      views.forEach(function (v) { fit(v); out.push(renderer.domElement.toDataURL('image/png')); });
      helpers.visible = helpersOn; planeGroup.visible = planesOn; state.style = styleKeep; state.picked = selKeep;
      orbit.yaw = keep.yaw; orbit.pitch = keep.pitch; orbit.radius = keep.radius; orbit.target.copy(keep.target);
      orbit.ortho = keep.ortho; orbit.hideAxis = keep.hideAxis; if (grid) grid.visible = keep.grid;
      apply();
      return out;
    }

    // pointer: left = orbit, right / middle / shift = pan, wheel = zoom, click = pick
    var drag = null, ray = new THREE.Raycaster();
    host.addEventListener('contextmenu', function (e) { if (!onControl(e)) e.preventDefault(); });
    // Floating controls (views, Planes/Size, detail level, style) sit on top of the
    // viewport: a press that starts on them must stay theirs - capturing it here
    // swallowed their clicks.
    function onControl(e) { return e.target && e.target.closest && e.target.closest('.float, .busy'); }
    host.addEventListener('pointerdown', function (e) {
      if (onControl(e)) return;
      drag = { x: e.clientX, y: e.clientY, pan: e.button === 2 || e.button === 1 || e.shiftKey, moved: false };
      host.setPointerCapture(e.pointerId); host.classList.add('dragging');
    });
    host.addEventListener('pointermove', function (e) {
      if (!drag) return;
      var dx = e.clientX - drag.x, dy = e.clientY - drag.y;
      if (Math.abs(dx) + Math.abs(dy) > 2) drag.moved = true;
      drag.x = e.clientX; drag.y = e.clientY;
      // Drawings (Front / Right / Plan) pan with any button, as a Revit 2D view; 3D orbits.
      if (drag.pan || orbit.ortho) {
        var scale = orbit.radius * 2 * Math.tan(THREE.MathUtils.degToRad(camera.fov / 2)) / host.clientHeight;
        var right = new THREE.Vector3().setFromMatrixColumn(cam().matrix, 0);
        var up = new THREE.Vector3().setFromMatrixColumn(cam().matrix, 1);
        orbit.target.addScaledVector(right, -dx * scale).addScaledVector(up, dy * scale);
      } else {
        orbit.yaw -= dx * 0.008;
        orbit.pitch = Math.max(-1.55, Math.min(1.5699, orbit.pitch + dy * 0.008));
      }
      render();
    });
    host.addEventListener('pointerup', function (e) {
      host.classList.remove('dragging');
      var d = drag; drag = null;
      if (d && !d.moved && e.button === 0) {
        var r = host.getBoundingClientRect();
        ray.setFromCamera(new THREE.Vector2((e.clientX - r.left) / r.width * 2 - 1, -(e.clientY - r.top) / r.height * 2 + 1), cam());
        var hit = ray.intersectObjects(meshes.filter(function (m) { return m.mesh.visible; }).map(function (m) { return m.mesh; }), false)[0];
        select(hit ? [hit.object.userData.index] : [], true);
      }
    });
    host.addEventListener('dblclick', function (e) { if (!onControl(e)) fit(currentView()); });
    host.addEventListener('wheel', function (e) {
      if (onControl(e)) return;
      e.preventDefault();
      orbit.radius = Math.max(20, orbit.radius * Math.pow(1.0015, e.deltaY));
      render();
    }, { passive: false });
    window.addEventListener('resize', resize);
    if (window.ResizeObserver) new ResizeObserver(resize).observe(host);

    return { load: load, clear: function () { clear(); bboxCache = null; render(); }, fit: fit, apply: apply,
             resize: resize, render: render, snapshot: snapshot };
  })();

  function currentView() {
    var a = document.querySelector('[data-view].active');
    return a ? a.getAttribute('data-view') : 'iso';
  }

  // ── preview → panels ──────────────────────────────────────────────────
  function swatch(color, isVoid) {
    var s = el('span', 'sw' + (isVoid ? ' void' : ''));
    if (color && !isVoid) s.style.background = color;
    return s;
  }
  function materialColor(name) {
    var p = state.preview; if (!p) return null;
    for (var i = 0; i < p.materials.length; i++) if (p.materials[i].name === name) return p.materials[i].color;
    return null;
  }

  // Selection is a list of part indexes: one row of the Project Browser tree can
  // stand for several repeated parts (Leg 1-4 -> "Leg x4").
  function select(indexes, fromView) {
    state.picked = indexes || [];
    var part = state.preview && state.preview.parts[state.picked[0]];
    bridge.send({ type: 'select', label: part ? part.label : null });
    view.apply();
    var key = state.picked.join(',');
    each('#parts-list .row', function (r) { r.classList.toggle('selected', !!key && r.getAttribute('data-i') === key); });
    if (fromView && state.picked.length) {
      showTab('parts');
      var row = null;
      each('#parts-list .row', function (r) {
        if (!row && r.getAttribute('data-i').split(',').indexOf(String(state.picked[0])) >= 0) row = r;
      });
      if (row) {
        var group = row.getAttribute('data-group');
        if (state.collapsed[group]) { state.collapsed[group] = false; renderParts(state.preview); return select(state.picked, true); }
        row.classList.add('selected');
        row.scrollIntoView({ block: 'nearest' });
      }
    }
  }

  function renderPreview(p, fit) {
    if (p.parse_error) { status('The family data cannot be read; the 3D view keeps the last model.', 'danger'); return; }
    var hadModel = state.preview && state.preview.parts && state.preview.parts.length;
    var sameParts = hadModel && p.parts && p.parts.length === hadModel;
    state.preview = p;
    if (!sameParts) state.picked = [];          // a parameter edit keeps the selection
    if (p.preview_type) state.type = p.preview_type;   // the TYPE shown = the type Python flexed
    var has = p.parts && p.parts.length;
    $('viewport-empty').classList.toggle('hidden', !!has);
    var types = (p.table && p.table.columns) || [];
    var sel = $('sel-type');
    sel.innerHTML = '';
    types.forEach(function (t) { var o = el('option', null, t); o.value = t; sel.appendChild(o); });
    $('type-field').classList.toggle('hidden', types.length === 0);
    state.type = types.indexOf(state.type) >= 0 ? state.type : (types[0] || '');
    sel.value = state.type;
    if (has) {
      view.load(p);
      if (fit || !hadModel) view.fit(currentView());
    } else {
      view.clear();
    }
    $('part-count').textContent = (has ? p.parts.length : 0) + (has && p.parts.length === 1 ? ' item' : ' items');
    var badge = $('size-badge');
    badge.classList.toggle('hidden', !has);
    if (has) badge.textContent = 'W ' + fmt(p.size_mm[0]) + ' · D ' + fmt(p.size_mm[1]) + ' · H ' + fmt(p.size_mm[2]) + ' mm';
    var focused = document.activeElement && document.activeElement.classList.contains('cell-input')
      ? document.activeElement.getAttribute('aria-label') : null;
    renderSummary(p); renderParts(p); renderMaterials(p); renderTypes(p); renderPlanes(p); renderIssues(p);
    if (focused) {      // the grid was rebuilt: keep typing in the same cell
      var again = document.querySelector('.cell-input[aria-label="' + focused.replace(/"/g, '\\"') + '"]');
      if (again) { again.focus(); again.select(); }
    }
  }

  // Revit Properties palette: the family like the type selector, then grouped
  // name / value rows.
  function renderSummary(p) {
    var box = $('summary'); box.innerHTML = '';
    var s = p.summary || {};
    if (!s.family_name && !s.parts) return;
    var head = el('div', 'prop-family');
    head.appendChild(el('b', null, s.family_name || '(no name)'));
    head.appendChild(el('span', null, (s.family_category || '(no category)') + ' · Loadable Family'));
    box.appendChild(head);
    var size = p.size_mm || [0, 0, 0];
    var concept = s.concept || {};
    var groups = [
      ['Concept', [['What', concept.what], ['Use', concept.use], ['Placement', concept.placement]]
        .filter(function (kv) { return kv[1]; })],
      ['Identity Data', [['Category', s.family_category || '—'], ['Types', (s.types || []).length || 1],
                         ['Parameters', s.parameters || 0], ['Materials', (s.materials || []).length]]],
      ['Geometry', [['Parts', s.parts || 0], ['Solids', s.solids || 0], ['Voids', s.voids || 0],
                    ['Joints checked', s.joints || 0]]],
      ['Dimensions', [['Width', fmt(size[0]) + ' mm'], ['Depth', fmt(size[1]) + ' mm'], ['Height', fmt(size[2]) + ' mm'],
                      ['Labelled dimensions', s.dimensions || 0], ['Reference planes', s.reference_planes || 0]]],
    ];
    groups.forEach(function (g) {
      var key = 'props:' + g[0];
      if (!g[1].length) return;
      box.appendChild(band(g[0], key, null));
      if (state.collapsed[key]) return;
      g[1].forEach(function (kv) {
        var row = el('div', 'prop');
        row.appendChild(el('span', null, kv[0]));
        var value = el('span', null, String(kv[1]));
        value.title = String(kv[1]);              // long concept text: the full sentence on hover
        row.appendChild(value);
        box.appendChild(row);
      });
    });
  }

  // A collapsible group band (Revit Properties / Project Browser).
  function band(title, key, count, rerender) {
    var b = el('div', 'band');
    var caret = el('span', 'caret'); caret.appendChild(icon(state.collapsed[key] ? 'right' : 'down')); b.appendChild(caret);
    b.appendChild(el('span', null, title));
    if (count !== null && count !== undefined) b.appendChild(el('span', 'n', String(count)));
    b.addEventListener('click', function () {
      state.collapsed[key] = !state.collapsed[key];
      (rerender || renderSummary)(state.preview);
    });
    return b;
  }

  // Project Browser: parts grouped by subcategory; repeated parts ("Leg 1".."Leg 4")
  // fold into one row "Leg x4" that selects all of them.
  function renderParts(p) {
    var list = $('parts-list'); list.innerHTML = '';
    var parts = p.parts || [];
    if (!parts.length) { list.appendChild(el('div', 'empty-row', 'No parts yet. Describe a family in the chat.')); return; }
    var groups = {}, order = [];
    parts.forEach(function (part, i) {
      var g = part.subcategory || (part.solid ? 'Other' : 'Voids');
      if (!groups[g]) { groups[g] = []; order.push(g); }
      var base = part.label.replace(/ \d+$/, '').replace(/ joint$/, '');
      var items = groups[g], last = items[items.length - 1];
      if (last && last.base === base && last.material === part.material && last.solid === part.solid) last.idx.push(i);
      else items.push({ base: base, label: part.label, idx: [i], kind: part.kind, material: part.material,
                        solid: part.solid, levels: part.levels });
    });
    order.forEach(function (g) {
      var key = 'parts:' + g;
      var total = groups[g].reduce(function (n, it) { return n + it.idx.length; }, 0);
      list.appendChild(band(g, key, total, renderParts));
      if (state.collapsed[key]) return;
      groups[g].forEach(function (it) {
        var row = el('div', 'row clickable');
        row.setAttribute('data-i', it.idx.join(','));
        row.setAttribute('data-group', key);
        var name = el('span');
        name.appendChild(swatch(materialColor(it.material), !it.solid));
        name.appendChild(document.createTextNode(it.idx.length > 1 ? it.base : it.label));
        if (it.idx.length > 1) name.appendChild(el('span', 'x', ' ×' + it.idx.length));
        row.title = it.label + (it.levels && it.levels.length < 3 ? ' · shown at ' + it.levels.join(' / ') : '');
        row.appendChild(name);
        row.appendChild(el('span', 'sub', it.solid ? it.kind : 'Void ' + it.kind));
        row.appendChild(el('span', 'sub', it.solid ? (it.material || 'by category') : 'cuts'));
        row.addEventListener('click', function () {
          var same = state.picked.join(',') === it.idx.join(',');
          select(same ? [] : it.idx, false);
        });
        list.appendChild(row);
      });
    });
  }

  function renderMaterials(p) {
    var list = $('materials-list'); list.innerHTML = '';
    (p.materials || []).forEach(function (m) {
      var row = el('div', 'row mats');
      row.title = 'Family parameter: ' + m.parameter + (m.transparency ? ' · ' + m.transparency + '% transparent' : '');
      var name = el('span');
      name.appendChild(swatch(m.color));
      name.appendChild(document.createTextNode(m.name));
      row.appendChild(name);
      row.appendChild(el('span', 'sub', m.parameter));
      row.appendChild(el('span', 'sub num', String(m.parts)));
      list.appendChild(row);
    });
    if (!list.children.length) list.appendChild(el('div', 'empty-row', 'No materials. Parts keep the category default.'));
  }

  // Length / number values are edited in place: Enter, Tab, arrows (±10 mm, Shift ±100) or
  // leaving the cell sends set_param; Python re-flexes the preview (FamilyGen/flex.py).
  var NUMERIC = { length: true, number: true, integer: true };
  function paramInput(row, cell, typeName) {
    var wrap = el('span', 'cell-edit');
    var input = el('input'); input.type = 'number'; input.className = 'cell-input'; input.disabled = state.busy;
    input.value = cell.value === null || cell.value === undefined ? '' : cell.value;
    input.step = row.kind === 'length' ? 10 : (row.kind === 'integer' ? 1 : 0.1);
    if (row.kind === 'length') input.min = 1;
    input.title = 'Type a value' + (row.kind === 'length' ? ' in mm' : '') + ', then Enter. Arrow keys step '
      + (row.kind === 'length' ? '10 mm (Shift: 100 mm)' : 'by one') + '.';
    input.setAttribute('aria-label', row.name + (typeName ? ' for ' + typeName : ''));
    var sent = String(input.value);
    function commit() {
      var v = String(input.value).trim();
      if (v === '' || v === sent || isNaN(Number(v))) { if (v === '') input.value = sent; return; }
      if (row.kind === 'length' && Number(v) <= 0) {
        input.value = sent; status(row.name + ': a length must be more than 0 mm.', 'warning'); return;
      }
      if (state.busy) { input.value = sent; status('Wait until the current step finishes, then change the value.', 'warning'); return; }
      sent = v;
      bridge.send({ type: 'set_param', param: row.name, type_name: typeName, value: Number(v) });
    }
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); commit(); }
      if (e.key === 'Escape') { input.value = sent; input.blur(); }
      if ((e.key === 'ArrowUp' || e.key === 'ArrowDown') && e.shiftKey && row.kind === 'length') {
        e.preventDefault(); input.value = Number(input.value || 0) + (e.key === 'ArrowUp' ? 100 : -100); commit();
      }
    });
    input.addEventListener('keyup', function (e) { if (!e.shiftKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) commit(); });
    input.addEventListener('change', commit);
    wrap.appendChild(input);
    if (row.kind === 'length') wrap.appendChild(el('span', 'unit', 'mm'));
    return wrap;
  }

  function renderTypes(p) {
    var t = $('types-table'); t.innerHTML = '';
    var table = p.table || { columns: [], rows: [] };
    if (!table.rows.length) { t.appendChild(el('caption', null, 'No parameters or materials.')); return; }
    var driven = {};
    (p.dimensions || []).forEach(function (d) { driven[String(d.parameter).toLowerCase()] = true; });
    var cols = table.columns.length ? table.columns : ['Default'];
    var head = el('tr'); head.appendChild(el('th', null, 'Parameter'));
    cols.forEach(function (c) { head.appendChild(el('th', null, c)); });
    t.appendChild(head);
    table.rows.forEach(function (row) {
      var tr = el('tr');
      var name = el('td', null, row.name + (row.instance ? ' (instance)' : '') + (row.shared ? ' · shared' : ''));
      if (driven[row.name.toLowerCase()]) { name.appendChild(el('span', 'tag', 'resizes')); name.title = 'Drives a labelled dimension'; }
      tr.appendChild(name);
      var cells = table.columns.length ? row.cells : [{ value: row['default'], override: false }];
      cells.forEach(function (c, ci) {
        var td = el('td', c.override ? 'override' : '');
        if (row.kind === 'material' && c.value) td.appendChild(swatch(materialColor(c.value)));
        if (NUMERIC[row.kind]) {
          td.appendChild(paramInput(row, c, table.columns.length ? table.columns[ci] : ''));
        } else {
          td.appendChild(document.createTextNode(c.value === null || c.value === undefined ? '—' :
            (row.kind === 'length' ? fmt(c.value) + ' mm' : String(c.value))));
        }
        tr.appendChild(td);
      });
      t.appendChild(tr);
    });
  }

  // Parameter fit: the category's standard sizes and whether each one resizes the family.
  var FIT_WORDS = { driven: 'resizes', present: 'not linked', missing: 'missing', unused: 'not used' };
  function renderFit(list, fit) {
    if (!fit || !fit.rows) return;
    list.appendChild(band('Parameter fit · ' + fit.category + ' · ' + fit.driven + ' of ' + fit.required + ' sizes resize', 'sizes:fit', null, renderPlanes));
    if (state.collapsed['sizes:fit']) return;
    fit.rows.forEach(function (r) {
      var row = el('div', 'row kv fit');
      row.setAttribute('data-state', r.state);
      var name = el('span');
      name.appendChild(el('span', 'fit-dot'));
      name.appendChild(document.createTextNode(r.name + (r.required ? '' : ' (optional)') + (r.as && r.as !== r.name ? ' = ' + r.as : '')));
      row.appendChild(name);
      // Scope and shared status in words; a wrong scope says what it should be.
      var bits = [FIT_WORDS[r.state] + (r.axis ? ' · ' + r.axis.toUpperCase() : '')];
      if (r.as) bits.push(r.scope_ok === false ? '→ ' + r.scope : r.scope);
      if (r.shared) bits.push('shared');
      else if (r.in_file) bits.push('shareable');
      var sub = el('span', 'sub', bits.join(' · '));
      if (r.scope_ok === false) sub.classList.add('warn-text');
      row.appendChild(sub);
      row.title = r.why + (r.as ? ' · ' + r.scope + ' parameter recommended' : '') + ': ' + { driven: 'a labelled dimension uses it, so changing it resizes the family',
        present: 'the parameter exists but no labelled dimension uses it yet', missing: 'not in the family',
        unused: 'optional; this family does not need it' }[r.state];
      list.appendChild(row);
    });
    var sf = state.sharedFile || {};
    list.appendChild(el('div', 'fit-note', sf.path
      ? 'Shared parameter file: ' + sf.name + ' (' + sf.count + ' parameters)'
      : 'No shared parameter file is set in Revit (Manage > Shared Parameters): parameters stay family parameters.'));
    var actions = el('div', 'fit-action');
    if (fit.scope_issues || fit.shareable) {
      var sb = el('button', 'btn sm', 'Apply standards');
      sb.title = [fit.scope_issues ? fit.scope_issues + ' parameter(s) get the standard type / instance scope' : '',
                  fit.shareable ? fit.shareable + ' parameter(s) become shared parameters from ' + sf.name : '']
                 .filter(Boolean).join('; ') + '. No AI needed.';
      sb.disabled = state.busy;
      sb.addEventListener('click', function () { bridge.send({ type: 'apply_standards' }); });
      actions.appendChild(sb);
    }
    if (fit.driven < fit.required) {
      var b = el('button', 'btn sm', 'Make it parametric');
      b.title = 'Ask Claude for the missing ' + fit.category + ' sizes: reference planes, face locks and labelled dimensions';
      b.disabled = state.busy || !canSend();
      b.addEventListener('click', function () { bridge.send({ type: 'make_parametric' }); });
      actions.appendChild(b);
    }
    if (actions.children.length) list.appendChild(actions);
  }

  function renderPlanes(p) {
    var list = $('planes-list'); list.innerHTML = '';
    var planes = p.planes || [], dims = p.dimensions || [];
    renderFit(list, p.param_fit);
    list.appendChild(band('Reference planes', 'sizes:planes', planes.length, renderPlanes));
    if (!state.collapsed['sizes:planes']) planes.forEach(function (pl) {
      var row = el('div', 'row kv');
      var name = el('span');
      name.appendChild(el('span', 'sw plane' + (pl.template ? ' template' : '')));
      name.appendChild(document.createTextNode(pl.name + (pl.template ? ' (template)' : '')));
      row.appendChild(name);
      row.appendChild(el('span', 'sub num', pl.axis.toUpperCase() + ' = ' + fmt(pl.offset)));
      list.appendChild(row);
    });
    list.appendChild(band('Labelled dimensions', 'sizes:dims', dims.length, renderPlanes));
    if (!state.collapsed['sizes:dims']) dims.forEach(function (d) {
      var row = el('div', 'row kv');
      row.appendChild(el('span', null, d.parameter + ' (' + d.axis.toUpperCase() + ')'));
      row.appendChild(el('span', 'sub num', fmt(d.value) + ' mm'));
      list.appendChild(row);
    });
    if (!dims.length && !state.collapsed['sizes:dims']) list.appendChild(el('div', 'empty-row', 'No labelled dimensions: the family does not resize yet. Ask Claude "make it parametric: Width, Depth, Height".'));
  }

  // Checks: every row names its severity in words (status never by colour alone).
  function renderIssues(p) {
    var box = $('issues'); box.innerHTML = '';
    var friendly = p.issues || { errors: p.errors || [], warnings: p.warnings || [] };
    function add(cls, sev, text) {
      var row = el('div', 'issue ' + cls);
      row.appendChild(el('span', 'sev', sev));
      row.appendChild(el('span', null, text));
      box.appendChild(row);
    }
    friendly.errors.forEach(function (e) { add('error', 'Error', e); });
    friendly.warnings.forEach(function (w) { add('warning', 'Warning', w); });
    if (!box.children.length) add('ok', 'OK', 'All checks passed. Revit checks each shape once more when it builds the family.');
    else if (friendly.errors.length) box.insertBefore(el('div', 'note', 'Fix the errors before Create Family - ask Claude to fix them.'), box.firstChild);
    var count = $('issue-count');
    var n = friendly.errors.length, w = friendly.warnings.length;
    count.classList.toggle('hidden', !(n || w));
    count.classList.toggle('warn', !n && !!w);
    count.textContent = n || w;
  }



  // ── chat ──────────────────────────────────────────────────────────────
  function ago(t) {
    var s = Math.round((Date.now() - t) / 1000);
    if (s < 45) return 'just now';
    if (s < 3600) return Math.round(s / 60) + ' min ago';
    return new Date(t).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  }
  setInterval(function () { each('.msg .time', function (n) { n.textContent = ago(+n.getAttribute('data-t')); }); }, 30000);

  function addMessage(m) {
    $('chat-empty').classList.add('hidden');
    removeTyping();
    var role = m.role || 'system';
    var box = el('div', 'msg ' + role + (m.problems ? ' problems' : ''));
    if (role === 'user' || role === 'assistant') box.appendChild(el('div', 'who', role === 'user' ? 'You' : 'Claude'));
    if (role === 'analysis') m.assumptionsTitle = 'Components';
    box.appendChild(document.createTextNode(m.text || ''));
    if (m.mcp_id) {
      // MCP mode: Sent → Read by the app → Answered (mcp_status).
      var rc = el('div', 'receipt');
      rc.setAttribute('data-mcp-id', m.mcp_id); rc.setAttribute('data-state', 'sent');
      rc.appendChild(el('span', 'receipt-text', m.mcp_state || 'Sent'));
      if (state.claudeReady) {
        var fb = el('button', 'link', 'Use built-in Claude');
        fb.title = 'Take the message back from the AI app and send it with your Claude API key';
        fb.addEventListener('click', function () { bridge.send({ type: 'mcp_fallback', id: m.mcp_id }); });
        rc.appendChild(fb);
      }
      box.appendChild(rc);
    }
    if (m.thumbs && m.thumbs.length) {
      var th = el('div', 'thumbs');
      m.thumbs.forEach(function (src) { var i = el('img'); i.src = src; i.alt = 'reference'; th.appendChild(i); });
      box.appendChild(th);
    }
    if (m.files && m.files.length) { var fl = el('div', 'files'); fl.appendChild(icon('clip')); fl.appendChild(document.createTextNode(m.files.join(', '))); box.appendChild(fl); }
    if (m.assumptions && m.assumptions.length) {
      var det = el('details'); det.open = m.assumptions.length <= 4;
      det.appendChild(el('summary', null, (m.assumptionsTitle || 'Assumptions') + ' (' + m.assumptions.length + ')'));
      var ul = el('ul'); m.assumptions.forEach(function (a) { ul.appendChild(el('li', null, a)); });
      det.appendChild(ul); box.appendChild(det);
    }
    if (m.meta) box.appendChild(el('div', 'meta', m.meta));
    if (m.listen) {
      var cl = el('button', 'btn sm', 'Copy listen command');
      cl.title = 'Paste it once in your AI app: it then answers every Studio message by itself';
      cl.addEventListener('click', function () { bridge.send({ type: 'copy_listen' }); });
      box.appendChild(cl);
    }
    if (role !== 'system') { var t = el('div', 'time', 'just now'); t.setAttribute('data-t', Date.now()); box.appendChild(t); }
    $('messages').appendChild(box);
    $('messages').scrollTop = $('messages').scrollHeight;
  }

  function removeTyping() { if (state.typingNode) { state.typingNode.remove(); state.typingNode = null; } }
  function showTyping(text) {
    if (!state.typingNode) {
      state.typingNode = el('div', 'msg assistant typing');
      var dots = el('span', 'dots'); dots.appendChild(el('i')); dots.appendChild(el('i')); dots.appendChild(el('i'));
      state.typingNode.appendChild(dots);
      state.typingNode.appendChild(el('span', 'typing-text'));
      $('messages').appendChild(state.typingNode);
    }
    state.typingNode.querySelector('.typing-text').textContent = text || 'Claude is working…';
    $('messages').scrollTop = $('messages').scrollHeight;
  }

  // Before init arrives Send stays on: Python answers with the reason if Claude is off.
  function canSend() {
    return !state.inited || state.claudeReady || !!(state.engine && state.engine.effective === 'mcp');
  }

  function setBusy(on, text, stoppable) {
    state.busy = on;
    $('busy').classList.toggle('hidden', !on);
    $('busy-text').textContent = text || 'Working…';
    // Revit's build cannot be interrupted: no Stop then.
    $('btn-stop').classList.toggle('hidden', !on || stoppable === false);
    ['btn-send', 'btn-create', 'btn-attach', 'sel-category', 'btn-templates'].forEach(function (id) {
      $(id).disabled = on || (id === 'btn-send' && !canSend());
    });
    $('prompt').readOnly = on;
    if (!on) removeTyping();                  // the timeline's running node shows the work now
    each('.cell-input', function (i) { i.disabled = on; });
    if (!on && state.preview) { renderTypes(state.preview); renderPlanes(state.preview); }
  }

  function status(text, tone) {
    var s = $('status'); s.textContent = text; s.title = text; s.setAttribute('data-tone', tone || 'info');
    $('status-dot').setAttribute('data-tone', tone || 'info');
  }

  function setFolder(path) {
    state.folder = path || '';
    var b = $('btn-folder');
    b.textContent = !path ? 'Choose output folder…' : (path.length > 42 ? '…' + path.slice(-41) : path);
    b.title = (path ? path + '\n' : '') + 'Folder the .rfa is saved in. Click to change.';
    b.classList.toggle('set', !!path);
  }

  function renderAttachments(items) {
    var box = $('attachments'); box.innerHTML = '';
    state.attachCount = items.length;
    items.forEach(function (a) {
      var c = el('span', 'chip');
      if (a.thumb) { var img = el('img'); img.src = a.thumb; img.alt = ''; c.appendChild(img); }
      else c.appendChild(icon(a.kind === 'text' ? 'text' : 'doc'));
      c.appendChild(el('span', null, a.name + ' · ' + a.size));
      var x = el('button'); x.appendChild(icon('x')); x.title = 'Remove ' + a.name; x.setAttribute('aria-label', 'Remove ' + a.name);
      x.addEventListener('click', function () { bridge.send({ type: 'remove_attachment', id: a.id }); });
      c.appendChild(x); box.appendChild(c);
    });
  }

  function fillSelect(sel, values, current, labels) {
    sel.innerHTML = '';
    values.forEach(function (v, i) { var o = el('option', null, labels ? labels[i] : v); o.value = v; sel.appendChild(o); });
    if (current) sel.value = current;
  }

  // Built-in Claude or the AI app over MCP (auto-switches while an AI app works with FamiGen).
  function applyEngine(e) {
    state.engine = e;
    $('sel-engine').value = e.pref || 'auto';
    var mcp = e.effective === 'mcp', app = e.client || 'your AI app';
    $('sel-engine').querySelector('option[value="auto"]').textContent = 'Auto' + (e.pref === 'auto' ? (mcp ? ' → ' + app : ' → built-in Claude') : '');
    $('ai-text').textContent = mcp ? 'MCP · ' + app + (e.listening ? ' · listening' : '') : (state.claudeReady ? 'AI ready' : 'AI off');
    // Over MCP and not listening yet: the AI badge copies the listen command when clicked.
    state.listenAction = mcp && !e.listening;
    $('claude-pill').classList.toggle('claude-pill-action', state.listenAction);
    $('claude-pill').title = state.listenAction
      ? 'Click to copy the listen command; paste it once in ' + app + ' and it answers every Studio message'
      : (state.badgeTitle || '');
    // The model list belongs to the built-in Claude: empty and useless over MCP or without a key.
    $('sel-model').classList.toggle('hidden', mcp || !state.claudeReady);
    $('claude-pill').setAttribute('data-state', mcp || state.claudeReady ? 'ready' : 'off');
    $('prompt').placeholder = mcp ? 'Message for ' + app + (e.listening ? ' (listening)…' : '…')
                                  : 'Describe a family, or the change you want…';
    $('sel-mode').classList.toggle('hidden', mcp);
    $('chk-check').parentNode.classList.toggle('hidden', mcp);
    $('btn-send').disabled = state.busy || !canSend();
  }

  // Auto-detect shows the category Claude chose inside its own option.
  function setDetected(category) {
    var opt = $('sel-category').querySelector('option[value="Auto"]');
    if (!opt) return;
    opt.textContent = category ? 'Auto → ' + category : 'Auto-detect';
    $('sel-category').title = category ? 'Auto-detect chose ' + category + '. Pick a category to fix it.'
      : 'Family category. Auto-detect lets Claude pick it from what you describe.';
  }

  // ── templates (library + example chips) ──────────────────────────────
  function initials(text) { return text.split(/\s+/).filter(Boolean).slice(0, 2).map(function (w) { return w[0]; }).join('').toUpperCase(); }
  function useTemplate(t) {
    if ($('sel-category').value !== 'Auto' && $('sel-category').value !== t.category) {
      $('sel-category').value = t.category;
      bridge.send({ type: 'set_category', category: t.category });
    }
    $('prompt').value = t.prompt;
    closeModal('modal-templates');
    $('prompt').focus();
    status('Template "' + t.title + '" in the message box: edit sizes and finishes, then send.', 'info');
  }
  function renderExamples() {
    var box = $('examples'); box.innerHTML = '';
    var cat = $('sel-category').value;
    var list = state.templates.filter(function (t) { return cat === 'Auto' || t.category === cat; });
    (list.length ? list : state.templates).slice(0, 3).forEach(function (t) {
      var b = el('button', 'example');
      b.appendChild(el('b', null, t.title));
      b.appendChild(document.createTextNode(t.prompt));
      b.addEventListener('click', function () { useTemplate(t); });
      box.appendChild(b);
    });
  }
  function renderTemplates(query) {
    var body = $('tpl-body'); body.innerHTML = '';
    var q = (query || '').toLowerCase().trim();
    var shown = 0;
    state.groups.forEach(function (g, gi) {
      var items = state.templates.filter(function (t) {
        return t.group === g && (!q || (t.title + ' ' + t.prompt + ' ' + t.category + ' ' + g).toLowerCase().indexOf(q) >= 0);
      });
      if (!items.length) return;
      body.appendChild(el('div', 'tpl-group', g));
      var grid = el('div', 'tpl-grid');
      items.forEach(function (t) {
        var card = el('button', 'tpl'); card.title = t.prompt;
        var glyph = el('div', 'tpl-glyph', initials(t.title)); glyph.style.background = PLOT[gi % PLOT.length];
        card.appendChild(glyph);
        var txt = el('div');
        txt.appendChild(el('div', 'tpl-title', t.title));
        txt.appendChild(el('div', 'tpl-sub', t.category + ' · ' + t.prompt));
        card.appendChild(txt);
        card.addEventListener('click', function () { useTemplate(t); });
        grid.appendChild(card); shown++;
      });
      body.appendChild(grid);
    });
    if (!shown) body.appendChild(el('div', 'tpl-empty', 'No template matches "' + query + '". Describe the family in the chat instead.'));
    $('tpl-count').textContent = shown + ' templates';
  }
  function openModal(id) { $(id).classList.remove('hidden'); }
  function closeModal(id) { $(id).classList.add('hidden'); }
  each('[data-close]', function (b) { b.addEventListener('click', function () { closeModal(b.getAttribute('data-close')); }); });
  each('.modal', function (m) { m.addEventListener('click', function (e) { if (e.target === m) closeModal(m.id); }); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') each('.modal', function (m) { closeModal(m.id); }); });

  function renderHistory(items) {
    var list = $('hist-list'); list.innerHTML = '';
    $('hist-count').textContent = items.length ? items.length + ' recent' : '';
    items.forEach(function (it) {
      var row = el('div', 'row clickable');
      var main = el('span');
      main.appendChild(el('div', null, it.family_name || it.id));
      main.appendChild(el('div', 'sub', (it.category || '?') + ' · ' + it.parts + ' part(s) · ' + (it.source === 'claude' ? 'Claude' : it.source || '') +
        (it.created ? ' · ' + new Date(it.created * 1000).toLocaleString() : '')));
      row.appendChild(main);
      row.appendChild(el('span', 'sub', 'Open'));
      row.title = it.note || it.id;
      row.addEventListener('click', function () { closeModal('modal-history'); bridge.send({ type: 'load_history', id: it.id }); });
      list.appendChild(row);
    });
    if (!items.length) list.appendChild(el('div', 'empty-row', 'Nothing yet. Families Claude or an MCP client proposes appear here.'));
  }


  // ── reference files: paste, drop ──────────────────────────────────────
  var TEXT_EXT = /\.(txt|md|csv|json|xml)$/i;
  function sendFile(file) {
    var media = file.type || (TEXT_EXT.test(file.name) ? 'text/plain' : '');
    if (TEXT_EXT.test(file.name)) media = 'text/plain';
    if (!/^image\/(png|jpeg|gif|webp)$/.test(media) && media !== 'application/pdf' && media !== 'text/plain') {
      status(file.name + ': use an image, a PDF or a text file.', 'warning'); return;
    }
    var reader = new FileReader();
    reader.onload = function () {
      var data = String(reader.result).split(',', 2)[1] || '';
      bridge.send({ type: 'attach_data', name: file.name || ('pasted-' + Date.now() + '.png'), media_type: media, data: data });
    };
    reader.readAsDataURL(file);
  }
  document.addEventListener('paste', function (e) {
    var items = (e.clipboardData && e.clipboardData.items) || [];
    var used = false;
    Array.prototype.forEach.call(items, function (it) {
      if (it.kind === 'file') {
        var f = it.getAsFile();
        if (f) { used = true; if (!f.name || f.name === 'image.png') f = new File([f], 'pasted-' + Date.now() + '.png', { type: f.type }); sendFile(f); }
      }
    });
    if (used) e.preventDefault();
  });
  ['dragover', 'drop'].forEach(function (t) { document.addEventListener(t, function (e) { e.preventDefault(); }); });
  var pane = $('chat-pane');
  pane.addEventListener('dragenter', function () { pane.classList.add('dropping'); $('drop-hint').classList.remove('hidden'); });
  pane.addEventListener('dragleave', function (e) { if (!pane.contains(e.relatedTarget)) { pane.classList.remove('dropping'); $('drop-hint').classList.add('hidden'); } });
  pane.addEventListener('drop', function (e) {
    pane.classList.remove('dropping'); $('drop-hint').classList.add('hidden');
    Array.prototype.forEach.call((e.dataTransfer && e.dataTransfer.files) || [], sendFile);
  });

  // ── progress strip (studio_controller._progress) ─────────────────────
  // The stage card shows the current step; each new step text becomes a timeline node, so
  // the column fills in at the pace the AI works. Earlier running nodes turn done.
  var STATE_WORDS = { todo: 'not started', active: 'in progress', done: 'done', wait: 'waiting for you',
    warn: 'done, with problems', error: 'failed', skip: 'skipped' };
  function renderProgress(m) {
    var steps = m.steps || [], reached = 0, current = null;
    steps.forEach(function (st, i) {
      var li = document.querySelector('.step[data-step="' + st.id + '"]');
      if (li) { li.setAttribute('data-state', st.state); li.title = st.label + ': ' + (STATE_WORDS[st.state] || st.state); }
      if (st.state !== 'todo' && st.state !== 'skip') { reached = i + 1; current = st; }
    });
    var stage = $('progress-card');
    stage.setAttribute('data-state', current ? current.state : 'todo');
    $('stage-num').textContent = reached + '/' + steps.length;
    $('stage-label').textContent = current ? current.label : 'Ready';
    $('progress-fill').style.width = (steps.length ? Math.round(100 * reached / steps.length) : 0) + '%';
    var t = $('progress-text'); t.textContent = m.text || ''; t.title = m.text || '';
    if (!current || !m.text) { state.lastStep = null; return; }
    var key = current.id + '|' + current.state + '|' + m.text;
    if (key === state.lastStep) return;
    state.lastStep = key;
    // A step that ends (done / warn / error) rewrites its own running node when that node is
    // still the newest one, so "Revit is building…" becomes "Saved …" instead of a second node.
    var last = $('messages').lastElementChild;
    if (last && last.classList.contains('progress-node') && last.getAttribute('data-step') === current.id &&
        /^(done|warn|error)$/.test(current.state)) {
      last.setAttribute('data-state', current.state);
      var dd = last.querySelector('.dots'); if (dd) dd.remove();
      last.lastChild.textContent = m.text;
      return;
    }
    each('#messages .msg.progress-node[data-state="active"], #messages .msg.progress-node[data-state="wait"]', function (n) {
      n.setAttribute('data-state', 'done'); var d = n.querySelector('.dots'); if (d) d.remove();
    });
    $('chat-empty').classList.add('hidden');
    var node = el('div', 'msg progress-node');
    node.setAttribute('data-state', current.state);
    node.setAttribute('data-step', current.id);
    var name = el('span', 'step-name', current.label);
    if (current.state === 'active') {
      var dots = el('span', 'dots'); dots.appendChild(el('i')); dots.appendChild(el('i')); dots.appendChild(el('i'));
      name.appendChild(dots);
    }
    node.appendChild(name);
    node.appendChild(document.createTextNode(m.text));
    $('messages').appendChild(node);
    $('messages').scrollTop = $('messages').scrollHeight;
  }

  // ── messages from Python ──────────────────────────────────────────────
  var handlers = {
    init: function (m) {
      // The host could not hide its own title bar: drop ours (logo, title, window buttons).
      document.body.classList.toggle('host-chrome', m.frameless === false);
      fillSelect($('sel-category'), m.categories, m.category, m.categories.map(function (c) { return c === 'Auto' ? 'Auto-detect' : c; }));
      state.templates = m.templates || []; state.groups = m.template_groups || [];
      renderExamples();
      fillSelect($('sel-mode'), MODES.map(function (x) { return x[0]; }), m.effort, MODES.map(function (x) { return x[1]; }));
      var c = m.claude || {};
      var first = !state.inited;
      state.inited = true;
      state.claudeReady = !!c.ready;
      if (m.progress) renderProgress(m.progress);
      state.sharedFile = m.shared_file || null;
      fillSelect($('sel-model'), c.models || [], c.model, c.labels);
      $('sel-model').disabled = !c.ready; $('sel-mode').disabled = !c.ready;
      $('sel-model').classList.toggle('hidden', !c.ready);
      // Same wording as every T3 tool's AI badge (init_ai_badge): 'AI ready' / 'AI off'.
      var pill = $('claude-pill');
      pill.setAttribute('data-state', c.ready ? 'ready' : 'off');
      $('ai-text').textContent = c.ready ? 'AI ready' : 'AI off';
      state.badgeTitle = c.ready ? 'AI Mode: Claude, direct connection with your LLMs Setting key' : c.reason;
      pill.title = state.badgeTitle;
      if (!c.ready && first) addMessage({ role: 'system', text: c.reason });
      $('btn-send').disabled = !c.ready;
      $('chk-check').checked = m.check_refs !== false;
      applyEngine(m.engine || { pref: 'auto', effective: 'claude' });
      setFolder(m.folder); $('chk-load').checked = !!m.load;
      if (m.into_family) {
        $('btn-folder').classList.add('hidden'); $('chk-load').parentNode.classList.add('hidden');
        $('btn-create').textContent = 'Add to Family';
      }
      $('ver').textContent = 'Studio ' + m.studio_version + (m.pyrevit ? ' · pyRevit ' + m.pyrevit : '');
      status(c.ready ? 'Describe the family or pick a template.' : 'Claude is off: add a Claude API key in LLMs Setting, or work from an AI app over MCP.', 'info');
    },
    busy: function (m) { setBusy(!!m.on, m.text, m.stoppable); },
    progress: function (m) { renderProgress(m); },
    mcp_status: function (m) {
      each('.receipt', function (r) {
        var st = r.getAttribute('data-state');
        var mine = m.id ? r.getAttribute('data-mcp-id') === m.id : (st === 'sent' || st === 'read');
        if (!mine) return;
        r.setAttribute('data-state', m.state);
        r.innerHTML = ''; r.appendChild(el('span', 'receipt-text', m.text));     // the fallback link goes once read
      });
    },
    chat: function (m) { addMessage(m); },
    chat_reset: function () {
      removeTyping();
      state.refCam = null; $('view-ref').classList.add('hidden');
      each('#messages .msg', function (n) { n.remove(); });
      state.lastStep = null;
      $('chat-empty').classList.remove('hidden');
    },
    schema: function (m) {
      state.json = m.json;          // kept for Claude, Create and MCP; not shown
      state.pendingFit = true;
    },
    preview: function (m) { var fit = !!state.pendingFit; state.pendingFit = false; renderPreview(m, fit); },
    usage: function (m) {
      var parts = [fmt(m.input + m.cache_read + m.cache_write) + ' in', fmt(m.output) + ' out'];
      if (m.usd !== null && m.usd !== undefined) parts.push('≈ $' + (m.usd < 0.01 ? m.usd.toFixed(4) : m.usd.toFixed(2)));
      // Cost details stay out of sight: hover the AI badge to see them.
      state.usage = 'This conversation: ' + parts.join(' · ') + ' (' + m.requests + ' request(s))';
      $('claude-pill').title = state.badgeTitle + '\n' + state.usage;
    },
    attachments: function (m) { renderAttachments(m.items || []); },
    status: function (m) { status(m.text, m.tone); },
    folder: function (m) { setFolder(m.path); },
    category: function (m) { $('sel-category').value = m.category; setDetected(''); renderExamples(); },
    detected: function (m) { setDetected(m.category); },
    ref_camera: function (m) {
      state.refCam = { azimuth_deg: m.azimuth_deg, elevation_deg: m.elevation_deg };
      $('view-ref').classList.remove('hidden');
    },
    engine: function (m) { applyEngine(m); },
    history: function (m) { renderHistory(m.items || []); openModal('modal-history'); },
    snapshot_request: function (m) {
      var images = [];
      if (m.camera) state.refCam = m.camera;
      try { images = view.snapshot(m.views || ['iso']); } catch (err) { images = []; }
      bridge.send({ type: 'snapshot', token: m.token, views: m.views, images: images });
    },
    created: function (m) { addMessage({ role: 'created', problems: m.problems, text: (m.lines || []).join('\n') }); },
    error: function (m) { addMessage({ role: 'error', text: m.text }); }
  };
  bridge.listen(function (m) { if (m && handlers[m.type]) handlers[m.type](m); });

  // ── user actions → Python ─────────────────────────────────────────────
  function send() {
    if (state.busy) return;
    if (!canSend()) { status('Claude is off: add a Claude API key in LLMs Setting, or connect an AI app over MCP.', 'warning'); return; }
    var text = $('prompt').value.trim();
    // A picture alone is a request too.
    if (!text && state.attachCount) text = 'Create this family from the attached reference.';
    if (!text) { status('Describe the family, or the change you want.', 'warning'); return; }
    bridge.send({ type: 'send', text: text, json: state.json || '' });
    $('prompt').value = '';
    grow();
  }
  // The box starts at two lines and grows with the text (CSS max-height scrolls the rest).
  function grow() {
    var t = $('prompt');
    t.style.height = 'auto';
    t.style.height = t.scrollHeight + 'px';
  }
  $('prompt').addEventListener('input', grow);
  $('btn-send').addEventListener('click', send);
  $('prompt').addEventListener('keydown', function (e) {
    // Enter sends, Shift+Enter breaks the line; never while an IME (Vietnamese, Chinese) is composing.
    if (e.key === 'Enter' && !e.shiftKey && !e.altKey && !e.isComposing && e.keyCode !== 229) {
      e.preventDefault(); send();
    }
    if (e.key === '/' && !$('prompt').value) { e.preventDefault(); $('btn-templates').click(); }
  });
  // The T3 title bar moves the window: press on its empty part to drag, double-click to maximize.
  (function () {
    var bar = document.querySelector('.titlebar');
    var plain = function (e) { return !e.target.closest('button, select, input, a, .ai-badge'); };
    bar.addEventListener('mousedown', function (e) {
      if (e.button === 0 && e.detail === 1 && plain(e)) bridge.send({ type: 'window', action: 'drag' });
    });
    bar.addEventListener('dblclick', function (e) {
      if (plain(e)) bridge.send({ type: 'window', action: 'maximize' });
    });
  })();
  $('btn-min').addEventListener('click', function () { bridge.send({ type: 'window', action: 'minimize' }); });
  $('btn-max').addEventListener('click', function () { bridge.send({ type: 'window', action: 'maximize' }); });
  $('btn-close').addEventListener('click', function () { bridge.send({ type: 'close' }); });
  $('btn-stop').addEventListener('click', function () { bridge.send({ type: 'cancel' }); });
  $('btn-new').addEventListener('click', function () { bridge.send({ type: 'new_chat' }); setDetected(''); });
  $('btn-history').addEventListener('click', function () { bridge.send({ type: 'history' }); });
  $('btn-attach').addEventListener('click', function () { bridge.send({ type: 'attach' }); });
  $('btn-templates').addEventListener('click', function () {
    $('tpl-search').value = ''; renderTemplates(''); openModal('modal-templates'); $('tpl-search').focus();
  });
  $('tpl-search').addEventListener('input', function (e) { renderTemplates(e.target.value); });
  $('claude-pill').addEventListener('click', function () {
    if (state.listenAction) bridge.send({ type: 'copy_listen' });
  });
  $('btn-copy-prompt').addEventListener('click', function () { bridge.send({ type: 'copy_prompt', text: $('prompt').value }); });
  $('sel-engine').addEventListener('change', function (e) { bridge.send({ type: 'set_engine', engine: e.target.value }); });
  $('chk-check').addEventListener('change', function (e) { bridge.send({ type: 'set_check', on: e.target.checked }); });
  $('sel-category').addEventListener('change', function (e) {
    bridge.send({ type: 'set_category', category: e.target.value });
    setDetected(''); renderExamples();
  });
  $('sel-model').addEventListener('change', function (e) { bridge.send({ type: 'set_model', model: e.target.value }); });
  $('sel-mode').addEventListener('change', function (e) { bridge.send({ type: 'set_effort', effort: e.target.value }); });
  $('btn-folder').addEventListener('click', function () { bridge.send({ type: 'pick_folder' }); });
  $('chk-load').addEventListener('change', function (e) { bridge.send({ type: 'set_load', load: e.target.checked }); });
  $('btn-create').addEventListener('click', function () {
    if (!state.json) { status('Describe a family in the chat or pick a template first.', 'warning'); return; }
    bridge.send({ type: 'create', json: state.json, load: $('chk-load').checked });
  });
  $('sel-type').addEventListener('change', function (e) {
    state.type = e.target.value; view.apply();
    bridge.send({ type: 'set_type', name: state.type });     // Python flexes the preview to this type's sizes
  });
  $('sel-style').addEventListener('change', function (e) { state.style = e.target.value; view.apply(); });
  $('btn-fit').addEventListener('click', function () { view.fit(currentView()); });
  // Below 1100 px the details panel is a drawer over the 3D view (never just gone).
  $('btn-inspect').addEventListener('click', function () {
    var on = !document.body.classList.contains('show-inspect');
    document.body.classList.toggle('show-inspect', on);
    $('btn-inspect').classList.toggle('active', on);
    view.resize();
  });
  $('chk-dims').addEventListener('change', function () { view.apply(); });
  $('chk-planes').addEventListener('change', function () { view.apply(); });
  $('chk-voids').addEventListener('change', function () { view.apply(); });
  each('[data-view]', function (b) {
    b.addEventListener('click', function () {
      each('[data-view]', function (x) { x.classList.toggle('active', x === b); });
      view.fit(b.getAttribute('data-view'));
    });
  });
  each('[data-detail]', function (b) {
    b.addEventListener('click', function () {
      each('[data-detail]', function (x) { x.classList.toggle('active', x === b); });
      state.detail = b.getAttribute('data-detail'); view.apply();
    });
  });
  function showTab(name) {
    each('.tab', function (t) { t.classList.toggle('active', t.getAttribute('data-tab') === name); });
    each('.tab-panel', function (p) { p.classList.toggle('active', p.id === 'tab-' + name); });
  }
  each('.tab', function (t) { t.addEventListener('click', function () { showTab(t.getAttribute('data-tab')); }); });

  view.resize();
  view.fit('iso');
  // Say ready until init arrives: the first one can reach Python before the window can answer.
  (function hello(tries) {
    if (state.inited) return;
    bridge.send({ type: 'ready' });
    if (tries > 1) setTimeout(function () { hello(tries - 1); }, 1500);
  })(20);
})();
