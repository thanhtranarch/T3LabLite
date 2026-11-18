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
  function el(tag, cls, text) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (text !== undefined && text !== null) n.textContent = text;
    return n;
  }
  function css(name) { return getComputedStyle(document.documentElement).getPropertyValue(name).trim(); }
  function fmt(n) { return Math.round(n).toLocaleString('en-US'); }
  function each(sel, fn) { Array.prototype.forEach.call(document.querySelectorAll(sel), fn); }

  var MODES = [['low', 'Fast'], ['medium', 'Balanced'], ['high', 'Best'], ['xhigh', 'Deep']];
  var PLOT = ['#2563EB', '#D97706', '#059669', '#DC2626', '#7C3AED', '#0891B2', '#C026D3', '#65A30D']; // T3.Plot.1-8

  var state = {
    busy: false, claudeReady: false, preview: null, selected: -1, type: '',
    json: '', folder: '', style: 'shaded', detail: 'fine',
    templates: [], groups: [], typingNode: null
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
    scene.add(new THREE.HemisphereLight(0xffffff, 0x8a8a92, 1.15));
    var key = new THREE.DirectionalLight(0xffffff, 1.4); key.position.set(-0.6, -1.0, 1.4); scene.add(key);
    var fill = new THREE.DirectionalLight(0xffffff, 0.45); fill.position.set(1.0, 0.8, 0.4); scene.add(fill);
    var root = new THREE.Group(); scene.add(root);
    var helpers = new THREE.Group(); scene.add(helpers);
    var planeGroup = new THREE.Group(); scene.add(planeGroup);
    var box = null, grid = null;
    var meshes = [];          // {part, mesh, edges, lit, flat}
    var orbit = { target: new THREE.Vector3(0, 0, 0), radius: 3000, yaw: -0.9, pitch: 0.45 };
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
      camera.position.copy(orbit.target).addScaledVector(dir, orbit.radius);
      camera.near = Math.max(1, orbit.radius / 500); camera.far = orbit.radius * 50;
      camera.updateProjectionMatrix();
      camera.lookAt(orbit.target);
    }
    function render() { place(); renderer.render(scene, camera); updateLabels(); }

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
      var v = VIEWS[viewName || 'iso'] || VIEWS.iso;
      orbit.yaw = v[0]; orbit.pitch = v[1];
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

    function addLabel(cls, text, pos, title) {
      var n = el('div', cls, text);
      if (title) n.title = title;
      $('labels').appendChild(n);
      labelSpecs.push({ node: n, pos: pos, cls: cls });
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
        addLabel('plane-label', pl.name, pts[3], pl.name + ': ' + pl.axis + ' = ' + fmt(pl.offset) + ' mm');
      });
      (p.dimensions || []).forEach(function (d) {
        var mid = (d.from + d.to) / 2, pos;
        if (d.axis === 'x') pos = [mid, b[0][1] - pad / 2, b[0][2]];
        else if (d.axis === 'y') pos = [b[1][0] + pad / 2, mid, b[0][2]];
        else pos = [b[1][0] + pad / 2, b[0][1] - pad / 2, mid];
        addLabel('dim-param', d.parameter + ' = ' + fmt(d.value), pos, 'Labelled dimension: ' + d.parameter);
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
        var edges = new THREE.LineSegments(new THREE.EdgesGeometry(geo, 28),
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
        if (!labelled.x) addLabel('dim-label', 'W ' + fmt(s[0]), [(a[0] + c[0]) / 2, a[1], a[2]]);
        if (!labelled.y) addLabel('dim-label', 'D ' + fmt(s[1]), [c[0], (a[1] + c[1]) / 2, a[2]]);
        if (!labelled.z) addLabel('dim-label', 'H ' + fmt(s[2]), [c[0], a[1], (a[2] + c[2]) / 2]);
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
        var shown = (m.part.levels || ['coarse', 'medium', 'fine']).indexOf(state.detail) >= 0;
        var picked = i === state.selected;
        var name = swap[m.part.material] || m.part.material;
        var def = colors[name];
        var color = !m.part.solid ? voidColor : (def ? def.color : fallback);
        var t = m.part.solid ? (def ? Math.min(90, def.transparency || 0) : 0) : 72;
        var mat;
        if (style === 'hidden') { mat = m.flat; mat.color.set(m.part.solid ? paper : voidColor); t = m.part.solid ? 0 : 80; }
        else if (style === 'consistent') { mat = m.flat; mat.color.set(color); }
        else { mat = m.lit; mat.color.set(color); mat.flatShading = style !== 'realistic';
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
      var v = new THREE.Vector3(pos[0], pos[1], pos[2]).project(camera);
      if (v.z > 1 || v.z < -1) return null;
      return [(v.x + 1) / 2 * host.clientWidth, (1 - v.y) / 2 * host.clientHeight];
    }
    function updateLabels() {
      var dims = $('chk-dims').checked, planes = $('chk-planes').checked;
      labelSpecs.forEach(function (l) {
        var on = l.cls === 'dim-label' ? dims : planes;
        var p = on ? project(l.pos) : null;
        l.node.classList.toggle('hidden', !p);
        if (p) { l.node.style.left = p[0] + 'px'; l.node.style.top = p[1] + 'px'; }
      });
    }

    // Render the current model from named views for the reference check.
    function snapshot(views) {
      var keep = { yaw: orbit.yaw, pitch: orbit.pitch, radius: orbit.radius, target: orbit.target.clone() };
      var helpersOn = helpers.visible, planesOn = planeGroup.visible, styleKeep = state.style, selKeep = state.selected;
      helpers.visible = false; planeGroup.visible = false; state.style = 'shaded'; state.selected = -1; apply();
      var out = [];
      views.forEach(function (v) { fit(v); out.push(renderer.domElement.toDataURL('image/png')); });
      helpers.visible = helpersOn; planeGroup.visible = planesOn; state.style = styleKeep; state.selected = selKeep;
      orbit.yaw = keep.yaw; orbit.pitch = keep.pitch; orbit.radius = keep.radius; orbit.target.copy(keep.target);
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
      if (drag.pan) {
        var scale = orbit.radius * 2 * Math.tan(THREE.MathUtils.degToRad(camera.fov / 2)) / host.clientHeight;
        var right = new THREE.Vector3().setFromMatrixColumn(camera.matrix, 0);
        var up = new THREE.Vector3().setFromMatrixColumn(camera.matrix, 1);
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
        ray.setFromCamera(new THREE.Vector2((e.clientX - r.left) / r.width * 2 - 1, -(e.clientY - r.top) / r.height * 2 + 1), camera);
        var hit = ray.intersectObjects(meshes.filter(function (m) { return m.mesh.visible; }).map(function (m) { return m.mesh; }), false)[0];
        select(hit ? hit.object.userData.index : -1, true);
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

  function select(index, fromView) {
    state.selected = index;
    var part = state.preview && state.preview.parts[index];
    bridge.send({ type: 'select', label: part ? part.label : null });
    view.apply();
    each('#parts-list li', function (li) { li.classList.toggle('selected', +li.getAttribute('data-i') === index); });
    if (fromView && index >= 0) {
      showTab('parts');
      var li = $('parts-list').querySelector('[data-i="' + index + '"]');
      if (li) li.scrollIntoView({ block: 'nearest' });
    }
  }

  function renderPreview(p, fit) {
    if (p.parse_error) { status('The family data cannot be read; the 3D view keeps the last model.', 'danger'); return; }
    var hadModel = state.preview && state.preview.parts && state.preview.parts.length;
    state.preview = p;
    state.selected = -1;
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
    if (has) badge.textContent = fmt(p.size_mm[0]) + ' × ' + fmt(p.size_mm[1]) + ' × ' + fmt(p.size_mm[2]) + ' mm';
    renderSummary(p); renderParts(p); renderMaterials(p); renderTypes(p); renderPlanes(p); renderIssues(p);
  }

  function renderSummary(p) {
    var box = $('summary'); box.innerHTML = '';
    var s = p.summary || {};
    if (!s.family_name && !s.parts) return;
    box.appendChild(el('div', 'summary-title', s.family_name || '(no name)'));
    box.appendChild(el('div', 'summary-sub', s.family_category || '(no category)'));
    var stats = el('div', 'stats');
    [[s.parts, 'parts'], [(s.materials || []).length, 'materials'], [s.parameters, 'parameters'],
     [(s.types || []).length || 1, 'types'], [s.dimensions, 'dimensions'], [s.voids, 'voids']].forEach(function (x) {
      if (!x[0]) return;
      var st = el('span', 'stat'); st.appendChild(el('b', null, String(x[0]))); st.appendChild(document.createTextNode(' ' + x[1]));
      stats.appendChild(st);
    });
    box.appendChild(stats);
  }

  function renderParts(p) {
    var list = $('parts-list'); list.innerHTML = '';
    (p.parts || []).forEach(function (part, i) {
      var li = el('li', 'clickable'); li.setAttribute('data-i', i);
      li.appendChild(swatch(materialColor(part.material), !part.solid));
      var main = el('div', 'li-main');
      main.appendChild(el('div', 'li-name', part.label));
      var levels = part.levels && part.levels.length < 3 ? ' · ' + part.levels.join('/') : '';
      main.appendChild(el('div', 'li-sub', part.kind + (part.solid ? '' : ' · void') + (part.subcategory ? ' · ' + part.subcategory : '') + levels));
      li.appendChild(main);
      li.appendChild(el('span', 'li-side', part.solid ? (part.material || 'default') : 'cuts'));
      li.addEventListener('click', function () { select(state.selected === i ? -1 : i, false); });
      list.appendChild(li);
    });
    if (!list.children.length) list.appendChild(el('li', null, 'No parts yet.'));
  }

  function renderMaterials(p) {
    var list = $('materials-list'); list.innerHTML = '';
    (p.materials || []).forEach(function (m) {
      var li = el('li'); li.title = 'Family parameter: ' + m.parameter;
      li.appendChild(swatch(m.color));
      var main = el('div', 'li-main');
      main.appendChild(el('div', 'li-name', m.name));
      main.appendChild(el('div', 'li-sub', m.parameter + (m.transparency ? ' · ' + m.transparency + '% transparent' : '')));
      li.appendChild(main);
      li.appendChild(el('span', 'li-side', m.parts + ' part(s)'));
      list.appendChild(li);
    });
    if (!list.children.length) list.appendChild(el('li', null, 'No materials. Parts keep the category default.'));
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
      var name = el('td', null, row.name + (row.instance ? ' (instance)' : ''));
      if (driven[row.name.toLowerCase()]) { name.appendChild(el('span', 'tag', 'resizes')); name.title = 'Drives a labelled dimension'; }
      tr.appendChild(name);
      var cells = table.columns.length ? row.cells : [{ value: row['default'], override: false }];
      cells.forEach(function (c) {
        var td = el('td', c.override ? 'override' : '');
        if (row.kind === 'material' && c.value) td.appendChild(swatch(materialColor(c.value)));
        td.appendChild(document.createTextNode(c.value === null || c.value === undefined ? '—' :
          (row.kind === 'length' ? fmt(c.value) + ' mm' : String(c.value))));
        tr.appendChild(td);
      });
      t.appendChild(tr);
    });
  }

  function renderPlanes(p) {
    var list = $('planes-list'); list.innerHTML = '';
    (p.planes || []).forEach(function (pl) {
      var li = el('li');
      li.appendChild(el('span', 'sw plane' + (pl.template ? ' template' : '')));
      var main = el('div', 'li-main');
      main.appendChild(el('div', 'li-name', pl.name));
      main.appendChild(el('div', 'li-sub', (pl.template ? 'template plane · ' : '') + pl.axis + ' = ' + fmt(pl.offset) + ' mm'));
      li.appendChild(main);
      li.appendChild(el('span', 'li-side', pl.axis.toUpperCase()));
      list.appendChild(li);
    });
    (p.dimensions || []).forEach(function (d) {
      var li = el('li');
      li.appendChild(el('span', null, '⇔'));
      var main = el('div', 'li-main');
      main.appendChild(el('div', 'li-name', d.parameter));
      main.appendChild(el('div', 'li-sub', 'labelled dimension along ' + d.axis.toUpperCase()));
      li.appendChild(main);
      li.appendChild(el('span', 'li-side', fmt(d.value) + ' mm'));
      list.appendChild(li);
    });
  }

  function renderIssues(p) {
    var ul = $('issues'); ul.innerHTML = '';
    var friendly = p.issues || { errors: p.errors || [], warnings: p.warnings || [] };
    if ((p.errors || []).length) ul.appendChild(el('li', 'error', 'Fix these before Create Family (ask Claude to fix them):'));
    friendly.errors.forEach(function (e) { ul.appendChild(el('li', 'error', e)); });
    friendly.warnings.forEach(function (w) { ul.appendChild(el('li', 'warning', w)); });
    if (!ul.children.length) ul.appendChild(el('li', 'ok', 'All checks passed. Revit checks each shape once more when it builds the family.'));
    var count = $('issue-count');
    var n = (p.errors || []).length, w = (p.warnings || []).length;
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
    box.appendChild(document.createTextNode(m.text || ''));
    if (m.thumbs && m.thumbs.length) {
      var th = el('div', 'thumbs');
      m.thumbs.forEach(function (src) { var i = el('img'); i.src = src; i.alt = 'reference'; th.appendChild(i); });
      box.appendChild(th);
    }
    if (m.files && m.files.length) box.appendChild(el('div', 'files', '📎 ' + m.files.join(', ')));
    if (m.assumptions && m.assumptions.length) {
      var det = el('details'); det.open = m.assumptions.length <= 4;
      det.appendChild(el('summary', null, 'Assumptions (' + m.assumptions.length + ')'));
      var ul = el('ul'); m.assumptions.forEach(function (a) { ul.appendChild(el('li', null, a)); });
      det.appendChild(ul); box.appendChild(det);
    }
    if (m.meta) box.appendChild(el('div', 'meta', m.meta));
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

  function setBusy(on, text) {
    state.busy = on;
    $('busy').classList.toggle('hidden', !on);
    $('busy-text').textContent = text || 'Working…';
    $('btn-stop').classList.toggle('hidden', !on);
    ['btn-send', 'btn-create', 'btn-attach', 'sel-category', 'btn-templates'].forEach(function (id) {
      $(id).disabled = on || (id === 'btn-send' && !state.claudeReady && !(state.engine && state.engine.effective === 'mcp'));
    });
    $('prompt').readOnly = on;
    if (on && /Claude/.test(text || '')) showTyping(text); else if (!on) removeTyping();
  }

  function status(text, tone) { var s = $('status'); s.textContent = text; s.title = text; s.setAttribute('data-tone', tone || 'info'); }

  function setFolder(path) {
    state.folder = path || '';
    var b = $('btn-folder');
    b.textContent = !path ? 'Choose output folder…' : (path.length > 42 ? '…' + path.slice(-41) : path);
    b.title = (path ? path + '\n' : '') + 'Folder the .rfa is saved in. Click to change.';
    b.classList.toggle('set', !!path);
  }

  function renderAttachments(items) {
    var box = $('attachments'); box.innerHTML = '';
    items.forEach(function (a) {
      var c = el('span', 'chip');
      if (a.thumb) { var img = el('img'); img.src = a.thumb; img.alt = ''; c.appendChild(img); }
      else c.appendChild(el('span', null, a.kind === 'text' ? '📝' : '📄'));
      c.appendChild(el('span', null, a.name + ' · ' + a.size));
      var x = el('button', null, '×'); x.title = 'Remove ' + a.name; x.setAttribute('aria-label', 'Remove ' + a.name);
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
    $('ai-text').textContent = mcp ? 'MCP · ' + app : (state.claudeReady ? 'AI ready' : 'AI off');
    $('claude-pill').setAttribute('data-state', mcp || state.claudeReady ? 'ready' : 'off');
    $('prompt').placeholder = mcp ? 'Message for ' + app + '… (it reads it when you say "continue in FamiGen")'
                                  : 'Describe a family, or the change you want…';
    $('sel-mode').classList.toggle('hidden', mcp);
    $('chk-check').parentNode.classList.toggle('hidden', mcp);
    $('btn-send').disabled = state.busy || (!mcp && !state.claudeReady);
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
      var li = el('li');
      var main = el('div', 'li-main');
      main.appendChild(el('div', 'li-name', it.family_name || it.id));
      main.appendChild(el('div', 'li-sub', (it.category || '?') + ' · ' + it.parts + ' part(s) · ' + (it.source === 'claude' ? 'Claude' : it.source || '') +
        (it.created ? ' · ' + new Date(it.created * 1000).toLocaleString() : '')));
      li.appendChild(main);
      li.appendChild(el('span', 'li-side', 'Open'));
      li.title = it.note || it.id;
      li.addEventListener('click', function () { closeModal('modal-history'); bridge.send({ type: 'load_history', id: it.id }); });
      list.appendChild(li);
    });
    if (!items.length) list.appendChild(el('li', null, 'Nothing yet. Families Claude or an MCP client proposes appear here.'));
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

  // ── messages from Python ──────────────────────────────────────────────
  var handlers = {
    init: function (m) {
      fillSelect($('sel-category'), m.categories, m.category, m.categories.map(function (c) { return c === 'Auto' ? 'Auto-detect' : c; }));
      state.templates = m.templates || []; state.groups = m.template_groups || [];
      renderExamples();
      fillSelect($('sel-mode'), MODES.map(function (x) { return x[0]; }), m.effort, MODES.map(function (x) { return x[1]; }));
      var c = m.claude || {};
      state.claudeReady = !!c.ready;
      fillSelect($('sel-model'), c.models || [], c.model, c.labels);
      $('sel-model').disabled = !c.ready; $('sel-mode').disabled = !c.ready;
      // Same wording as every T3 tool's AI badge (init_ai_badge): 'AI ready' / 'AI off'.
      var pill = $('claude-pill');
      pill.setAttribute('data-state', c.ready ? 'ready' : 'off');
      $('ai-text').textContent = c.ready ? 'AI ready' : 'AI off';
      state.badgeTitle = c.ready ? 'AI Mode: Claude, direct connection with your LLMs Setting key' : c.reason;
      pill.title = state.badgeTitle;
      if (!c.ready) addMessage({ role: 'system', text: c.reason });
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
    busy: function (m) { setBusy(!!m.on, m.text); },
    chat: function (m) { addMessage(m); },
    chat_reset: function () {
      removeTyping();
      each('#messages .msg', function (n) { n.remove(); });
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
    engine: function (m) { applyEngine(m); },
    history: function (m) { renderHistory(m.items || []); openModal('modal-history'); },
    snapshot_request: function (m) {
      var images = [];
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
    var text = $('prompt').value.trim();
    if (!text) { status('Describe the family, or the change you want.', 'warning'); return; }
    bridge.send({ type: 'send', text: text, json: state.json || '' });
    $('prompt').value = '';
  }
  $('btn-send').addEventListener('click', send);
  $('prompt').addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); send(); }
    if (e.key === '/' && !$('prompt').value) { e.preventDefault(); $('btn-templates').click(); }
  });
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
  $('sel-type').addEventListener('change', function (e) { state.type = e.target.value; view.apply(); });
  $('sel-style').addEventListener('change', function (e) { state.style = e.target.value; view.apply(); });
  $('btn-fit').addEventListener('click', function () { view.fit(currentView()); });
  $('chk-dims').addEventListener('change', function () { view.apply(); });
  $('chk-planes').addEventListener('change', function () { view.apply(); });
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
  bridge.send({ type: 'ready' });
})();
