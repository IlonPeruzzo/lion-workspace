// ExtendScript for Adobe Premiere Pro — Lion Workspace Plugin

// ════════════════════════════════════════════════════════════════════
// MOTION TRACKER — JSX functions
// ════════════════════════════════════════════════════════════════════

// Pega info do clip selecionado: media path absoluto, fps, in/out, dimensões.
// Retorna JSON ou erro literal (NO_PROJECT, NO_SEQUENCE, NO_SELECTION, NO_MEDIA, ERR:...)
function lwTrackerGetClipInfo() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        // Strict selection: pega só clips de VIDEO (filtra audio) e
        // escolhe o que está na track MAIS ALTA (topmost video clip).
        var videoClips = [];
        // Itera todas as video tracks (de cima pra baixo no Premiere — V_top é o último índice)
        // Acumula com info de qual track tá pra escolher o topmost depois.
        for (var t = 0; t < seq.videoTracks.numTracks; t++) {
            for (var c = 0; c < seq.videoTracks[t].clips.numItems; c++) {
                try {
                    if (seq.videoTracks[t].clips[c].isSelected()) {
                        videoClips.push({ clip: seq.videoTracks[t].clips[c], trackIdx: t });
                    }
                } catch(e2) {}
            }
        }
        if (videoClips.length === 0) return 'NO_SELECTION';
        // Pega o topmost (maior trackIdx em Premiere = track de cima visualmente)
        var topmost = videoClips[0];
        for (var ti = 1; ti < videoClips.length; ti++) {
            if (videoClips[ti].trackIdx > topmost.trackIdx) topmost = videoClips[ti];
        }
        var clip = topmost.clip;
        var pi = clip.projectItem;
        if (!pi) return 'NO_MEDIA';
        var mediaPath = '';
        try { mediaPath = pi.getMediaPath(); } catch(eMP) {}
        if (!mediaPath) return 'NO_MEDIA';
        var f = new File(mediaPath);
        if (!f.exists) return 'NO_MEDIA';

        var fps = 30;
        try {
            // sequence.getSettings().videoFrameRate (Time obj)
            var st = seq.getSettings();
            if (st && st.videoFrameRate) {
                fps = 1 / parseFloat(st.videoFrameRate.seconds);
            }
        } catch(eF) {}

        // In/out e start ticks (clip-relative no source media)
        var inSec = 0, outSec = 0, durSec = 0, clipStartTicks = '0', inTicks = '0';
        try { inSec = parseFloat(clip.inPoint.seconds); } catch(e3) {}
        try { outSec = parseFloat(clip.outPoint.seconds); } catch(e4) {}
        try { durSec = parseFloat(clip.duration.seconds); } catch(e5) {}
        try { clipStartTicks = String(clip.start.ticks); } catch(e6) {}
        try { inTicks = String(clip.inPoint.ticks); } catch(e7) {}

        // Dimensões via projectItem
        var width = 0, height = 0;
        try {
            var dims = pi.getFootageInterpretation();
            // Fallback: usa videoSettings da sequence
            width = parseInt(pi.getMediaPath() ? 0 : 0, 10);
        } catch(eD) {}

        var out = '{"mediaPath":"' + f.fsName.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
        out += ',"fps":' + fps;
        out += ',"inPointSec":' + inSec;
        out += ',"outPointSec":' + outSec;
        out += ',"duration":' + durSec;
        out += ',"clipStartTicks":"' + clipStartTicks + '"';
        out += ',"inPointTicks":"' + inTicks + '"';
        out += ',"trackIdx":' + topmost.trackIdx;
        out += '}';
        return out;
    } catch(e) { return 'ERR:' + e.toString(); }
}

// Alvo do COLAR track: só precisa de seleção + ticks — SEM validar mídia no disco.
// Assim dá pra colar em adjustment layer, texto/MOGRT, color matte, clip offline etc
// (todos têm Motion component com Position, que é o que lwTrackerApplyKeyframes usa).
function lwTrackerGetPasteTarget() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        var videoClips = [];
        for (var t = 0; t < seq.videoTracks.numTracks; t++) {
            for (var c = 0; c < seq.videoTracks[t].clips.numItems; c++) {
                try {
                    if (seq.videoTracks[t].clips[c].isSelected()) {
                        videoClips.push({ clip: seq.videoTracks[t].clips[c], trackIdx: t });
                    }
                } catch(e2) {}
            }
        }
        if (videoClips.length === 0) return 'NO_SELECTION';
        var topmost = videoClips[0];
        for (var ti = 1; ti < videoClips.length; ti++) {
            if (videoClips[ti].trackIdx > topmost.trackIdx) topmost = videoClips[ti];
        }
        var clip = topmost.clip;
        var clipStartTicks = '0', inTicks = '0';
        try { clipStartTicks = String(clip.start.ticks); } catch(e6) {}
        try { inTicks = String(clip.inPoint.ticks); } catch(e7) {}
        var out = '{"clipStartTicks":"' + clipStartTicks + '"';
        out += ',"inPointTicks":"' + inTicks + '"';
        out += ',"trackIdx":' + topmost.trackIdx;
        out += ',"name":"' + String(clip.name || '').replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
        out += '}';
        return out;
    } catch(e) { return 'ERR:' + e.toString(); }
}

// Aplica keyframes de Position do tracking no clip selecionado ou cria adjustment layer.
// jsonStr: { applyMode: 'selected'|'null', clipStartTicks, clipInPointTicks, fps,
//            videoWidth, videoHeight, history: [{frame, time, x, y, lost}] }
function _trDbg(msg) {
    try {
        var f = new File(Folder.desktop.fsName + '/lion-tracker-apply.txt');
        if (f.open('a')) { f.encoding = 'UTF-8'; f.write('[' + (new Date()).toString() + '] ' + msg + '\n'); f.close(); }
    } catch(e) {}
}

function lwTrackerApplyKeyframes(jsonStr) {
    _trDbg('═══ lwTrackerApplyKeyframes called, jsonStr.length=' + (jsonStr ? jsonStr.length : 0));
    try {
        var data = null;
        try { data = eval('(' + jsonStr + ')'); } catch(eJ) { _trDbg('JSON parse err: ' + eJ); return 'ERR:json_parse:' + eJ; }
        if (!data || !data.history || data.history.length === 0) { _trDbg('no_data'); return 'ERR:no_data'; }
        _trDbg('history.length=' + data.history.length + ' fps=' + data.fps + ' videoWidth=' + data.videoWidth + ' videoHeight=' + data.videoHeight + ' clipInPointTicks=' + data.clipInPointTicks);
        _trDbg('history[0]=' + JSON.stringify(data.history[0]) + ' history[end]=' + JSON.stringify(data.history[data.history.length-1]));
        if (!app.project) { _trDbg('NO_PROJECT'); return 'NO_PROJECT'; }
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        try { app.enableQE(); } catch(eQE) {}

        // Acha clip selecionado — STRICT: so video clips (filtra audio).
        // getSelection() retorna video+audio quando user clica em parte sincronizada.
        // Iteramos so videoTracks pra garantir que sao clips de video com Motion component.
        var clips = [];
        for (var t = 0; t < seq.videoTracks.numTracks; t++) {
            for (var c = 0; c < seq.videoTracks[t].clips.numItems; c++) {
                try { if (seq.videoTracks[t].clips[c].isSelected()) clips.push(seq.videoTracks[t].clips[c]); } catch(e2) {}
            }
        }
        // Fallback: seleção foi perdida durante o tracking — usa clipStartTicks + trackIdx
        // pra achar o clip exato que o usuario tinha quando abriu o tracker.
        if (clips.length === 0 && data.clipStartTicks && data.trackIdx != null && data.trackIdx >= 0) {
            _trDbg('selection lost — buscando por trackIdx=' + data.trackIdx + ' startTicks=' + data.clipStartTicks);
            try {
                var tgtTrack = data.trackIdx;
                var tgtTicks = String(data.clipStartTicks);
                if (tgtTrack < seq.videoTracks.numTracks) {
                    var trk = seq.videoTracks[tgtTrack];
                    for (var ck = 0; ck < trk.clips.numItems; ck++) {
                        try {
                            if (String(trk.clips[ck].start.ticks) === tgtTicks) {
                                clips.push(trk.clips[ck]);
                                _trDbg('found clip by position');
                                break;
                            }
                        } catch(eF) {}
                    }
                }
            } catch(eL) { _trDbg('fallback lookup err: ' + eL); }
        }
        if (clips.length === 0) { _trDbg('no_selection'); return 'ERR:no_selection'; }
        var clip = clips[0];
        _trDbg('clip OK, finding Motion component');

        // Encontra Motion component e Position prop
        var motion = null;
        var comps = clip.components;
        _trDbg('clip has ' + comps.numItems + ' components');
        for (var ci = 0; ci < comps.numItems; ci++) {
            var cmp = comps[ci];
            var nm = '';
            try { nm = String(cmp.matchName || ''); } catch(eN) {}
            _trDbg('  comp[' + ci + '] matchName=' + nm);
            if (nm === 'AE.ADBE Motion') { motion = cmp; break; }
        }
        if (!motion) { _trDbg('NO MOTION COMPONENT'); return 'ERR:no_motion_component'; }
        _trDbg('Motion found');

        var posProp = null;
        var props = motion.properties;
        for (var pi = 0; pi < props.numItems; pi++) {
            var pp = props[pi];
            var pn = '';
            try { pn = String(pp.displayName || ''); } catch(eP) {}
            if (pn === 'Position' || pn === 'Posicao' || pn === 'Posicion') { posProp = pp; break; }
        }
        if (!posProp) { _trDbg('NO POSITION PROP'); return 'ERR:no_position_prop'; }
        _trDbg('Position prop found');

        // Habilita time-varying
        try { posProp.setTimeVarying(true); _trDbg('setTimeVarying OK'); } catch(eTv) { _trDbg('setTimeVarying err: ' + eTv); }

        // Premiere Position é NORMALIZED [0..1] em algumas versões, ou PIXEL em outras.
        // Auto-detect: lê valor atual, se ambos < 2 = normalized.
        var current = posProp.getValue();
        var isNormalized = false;
        if (current && current.length >= 2) {
            isNormalized = (Math.abs(current[0]) < 2.0 && Math.abs(current[1]) < 2.0);
        }
        // Em pixel mode, Position é em coords da SEQUENCE (não da source).
        // Tracking foi feito no proxy (data.videoWidth × data.videoHeight).
        // Precisamos reescalonar: trackPx * seqDim / proxyDim.
        var seqW = 1920, seqH = 1080;
        try { seqW = parseInt(seq.frameSizeHorizontal, 10) || 1920; } catch(eSw) {}
        try { seqH = parseInt(seq.frameSizeVertical, 10) || 1080; } catch(eSh) {}
        _trDbg('isNormalized=' + isNormalized + ' current=' + (current ? current.join(',') : 'null') + ' seqW=' + seqW + ' seqH=' + seqH);

        // Source media tick of each keyframe = clipInPointTicks + (frame - startFrame) * (TICKS_PER_SEC / fps)
        var TPS = 254016000000;
        var clipInTicks = parseFloat(data.clipInPointTicks || '0');
        var fps = data.fps || 30;
        var ticksPerFrame = TPS / fps;
        // history[0].frame é o startFrame
        var startFrame = data.history[0].frame;
        // MODE: 'track' (segue ponto, padrao) ou 'stabilize' (trava ponto, posiçao inversa)
        var mode = (data.mode === 'stabilize') ? 'stabilize' : 'track';
        _trDbg('mode=' + mode);
        // Ponto de referencia no primeiro frame (usado tanto pra track quanto stabilize)
        var refX = data.history[0].x;
        var refY = data.history[0].y;
        // Posicao base atual do clip (usada pra stabilize)
        var baseX = current && current.length >= 2 ? current[0] : (isNormalized ? 0.5 : seqW/2);
        var baseY = current && current.length >= 2 ? current[1] : (isNormalized ? 0.5 : seqH/2);
        var added = 0;
        var skipped = 0;

        for (var hi = 0; hi < data.history.length; hi++) {
            var h = data.history[hi];
            if (h.lost || h.x == null || h.y == null) { skipped++; continue; }
            // Source media time = in-point + frame ABSOLUTO do proxy (proxy frame 0 = in-point).
            // Absoluto (não relativo ao startFrame) → track "só pra frente" deixa os frames
            // ANTES do ponto marcado sem keyframe, no lugar certo da timeline.
            var srcTicks = clipInTicks + h.frame * ticksPerFrame;
            var kfTime = new Time();
            kfTime.ticks = String(Math.round(srcTicks));

            // Valor Position:
            // - track: position = tracked location (clip se posiciona onde o ponto esta)
            // - stabilize: position = base - (tracked - ref), clip se move oposto ao ponto
            //   pra que o ponto trackeado fique TRAVADO no lugar
            var val;
            if (mode === 'stabilize') {
                var dx = h.x - refX;
                var dy = h.y - refY;
                if (isNormalized) {
                    val = [baseX - dx / data.videoWidth, baseY - dy / data.videoHeight];
                } else {
                    val = [baseX - dx * seqW / data.videoWidth, baseY - dy * seqH / data.videoHeight];
                }
            } else {
                if (isNormalized) {
                    val = [h.x / data.videoWidth, h.y / data.videoHeight];
                } else {
                    val = [h.x * seqW / data.videoWidth, h.y * seqH / data.videoHeight];
                }
            }
            try {
                posProp.addKey(kfTime);
                var setT = new Time(); setT.ticks = String(Math.round(srcTicks));
                posProp.setValueAtKey(setT, val);
                added++;
                if (hi < 3 || hi === data.history.length - 1) {
                    _trDbg('  kf[' + hi + '] frame=' + h.frame + ' srcTicks=' + Math.round(srcTicks) + ' val=' + val.join(','));
                }
            } catch(eKf) { _trDbg('  kf[' + hi + '] err: ' + eKf); }
        }

        _trDbg('DONE mode=' + mode + ' added=' + added + ' skipped=' + skipped);
        try { forceUIRefresh(); } catch(eR) {}
        return 'OK: ' + added + ' keyframes aplicados (' + skipped + ' frames perdidos)';
    } catch(e) { return 'ERR:' + e.toString(); }
}

// Get current project folder path
function getProjectFolder() {
    try {
        if (app.project && app.project.path) {
            var projFile = new File(app.project.path);
            return projFile.parent.fsName;
        }
    } catch (e) {}
    return '';
}

// Find or create a bin by name
function findOrCreateBin(binName) {
    var root = app.project.rootItem;
    for (var i = 0; i < root.children.numItems; i++) {
        if (root.children[i].name === binName && root.children[i].type === 2) {
            return root.children[i];
        }
    }
    return root.createBin(binName);
}

// Import a file into a specific bin (create if needed)
function importFileToBin(filePath, binName) {
    try {
        if (!app.project) return 'NO_PROJECT';

        // Verify file exists
        var f = new File(filePath);
        if (!f.exists) return 'FILE_NOT_FOUND: ' + filePath;

        // Find or create target bin
        var targetBin = findOrCreateBin(binName);

        // Count items before import
        var root = app.project.rootItem;
        var countBefore = root.children.numItems;

        // Strategy 1: Import with target bin (Premiere 2020+)
        try {
            var ok = app.project.importFiles([filePath], false, targetBin, false);
            if (ok) {
                // Verify it actually landed in the bin
                $.sleep(800);
                for (var i = 0; i < targetBin.children.numItems; i++) {
                    var child = targetBin.children[i];
                    if (child.name === f.displayName || child.name === f.displayName.replace(/\.[^.]+$/, '')) {
                        return 'OK';
                    }
                }
                // If it imported somewhere, still OK
                if (root.children.numItems > countBefore) return 'OK';
                return 'OK';
            }
        } catch (e1) {}

        // Strategy 2: Import to root first, then move to bin
        try {
            app.project.importFiles([filePath]);
            $.sleep(1000);

            // Find the newly imported item
            var newItem = null;
            for (var i = root.children.numItems - 1; i >= 0; i--) {
                var child = root.children[i];
                if (child.type !== 2) { // not a bin
                    // Match by filename
                    var childName = child.name;
                    var fileName = f.displayName.replace(/\.[^.]+$/, '');
                    if (childName === f.displayName || childName === fileName) {
                        newItem = child;
                        break;
                    }
                }
            }

            // Fallback: just grab the last non-bin item if new items were added
            if (!newItem && root.children.numItems > countBefore) {
                for (var i = root.children.numItems - 1; i >= 0; i--) {
                    if (root.children[i].type !== 2) {
                        newItem = root.children[i];
                        break;
                    }
                }
            }

            if (newItem) {
                newItem.moveBin(targetBin);
                return 'OK';
            }
            return 'IMPORTED_BUT_NOT_FOUND';
        } catch (e2) {
            return 'ERROR2: ' + e2.toString();
        }
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// Cubic bezier evaluation: given control points and t (0-1), returns y value
function cubicBezier(t, cp1x, cp1y, cp2x, cp2y) {
    // Attempt to find t-parameter for given x using Newton's method
    // For simplicity, we just evaluate the standard cubic bezier parametrically
    var u = 1 - t;
    var y = 3 * u * u * t * cp1y + 3 * u * t * t * cp2y + t * t * t;
    return y;
}

// Solve for bezier t given x (iterative)
function solveBezierX(x, cp1x, cp2x) {
    var t = x;
    for (var i = 0; i < 8; i++) {
        var u = 1 - t;
        var currentX = 3 * u * u * t * cp1x + 3 * u * t * t * cp2x + t * t * t;
        var dx = 3 * (1 - t) * (1 - t) * cp1x + 6 * (1 - t) * t * (cp2x - cp1x) + 3 * t * t * (1 - cp2x);
        if (Math.abs(dx) < 1e-6) break;
        t = t - (currentX - x) / dx;
        t = Math.max(0, Math.min(1, t));
    }
    return t;
}

// Evaluate bezier curve: given normalized x (0-1), returns normalized y (0-1)
function evalBezierCurve(x, cp1x, cp1y, cp2x, cp2y) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    var t = solveBezierX(x, cp1x, cp2x);
    return cubicBezier(t, cp1x, cp1y, cp2x, cp2y);
}

// Find selected clips in the active sequence
function getSelectedClips() {
    var seq = app.project.activeSequence;
    if (!seq) return [];
    var clips = [];

    // Strategy 1: getSelection() (Premiere 2022+)
    try {
        var sel = seq.getSelection();
        if (sel && sel.length > 0) {
            for (var i = 0; i < sel.length; i++) clips.push(sel[i]);
            return clips;
        }
    } catch (e1) {}

    // Strategy 2: isSelected() on all tracks
    var allTracks = [];
    for (var t = 0; t < seq.videoTracks.numTracks; t++) allTracks.push(seq.videoTracks[t]);
    for (var t = 0; t < seq.audioTracks.numTracks; t++) allTracks.push(seq.audioTracks[t]);

    for (var t = 0; t < allTracks.length; t++) {
        for (var c = 0; c < allTracks[t].clips.numItems; c++) {
            try {
                if (allTracks[t].clips[c].isSelected()) clips.push(allTracks[t].clips[c]);
            } catch (e2) {}
        }
    }
    if (clips.length > 0) return clips;

    // Strategy 3: Clip at playhead
    var playhead = seq.getPlayerPosition();
    for (var t = 0; t < seq.videoTracks.numTracks; t++) {
        var track = seq.videoTracks[t];
        for (var c = 0; c < track.clips.numItems; c++) {
            var clip = track.clips[c];
            try {
                if (playhead.ticks >= clip.start.ticks && playhead.ticks <= clip.end.ticks) {
                    clips.push(clip);
                }
            } catch (e3) {}
        }
    }
    return clips;
}

// Apply bezier easing using "baking" approach:
// Instead of setting bezier handles (not supported in Premiere ExtendScript),
// we calculate intermediate values along the bezier curve and create
// multiple linear keyframes that simulate the curve shape.
// Helper: create a Time object at given ticks (cross-platform safe)
function createTimeAt(ticksNum) {
    var t = new Time();
    t.ticks = String(Math.round(ticksNum));
    return t;
}

// Force Premiere UI to refresh (nudge playhead 1 tick forward and back)
function forceUIRefresh() {
    try {
        var seq = app.project.activeSequence;
        if (!seq) return;
        var pos = seq.getPlayerPosition();
        var cur = Number(pos.ticks);
        // Precisa cutucar o playhead por um FRAME INTEIRO (nao +1 tick) — 1 tick e uma
        // fracao infima de frame, entao o frame exibido nao muda e o Premiere nao
        // re-renderiza a alteracao (anchor/scale/etc). Um frame inteiro forca o re-render.
        var frame = 4233583066;
        try { if (seq.timebase && Number(seq.timebase) > 0) frame = Number(seq.timebase); } catch (e) {}
        var nudged = (cur >= frame) ? (cur - frame) : (cur + frame);
        seq.setPlayerPosition(String(nudged));
        seq.setPlayerPosition(String(cur));
    } catch (e) {}
}

// ── Easing functions paramétricas (REAL bounce/elastic, não bezier) ──
// Bezier não bounce porque é monotônico no eixo Y. Pra REAL bounce/elastic
// (múltiplas oscilações), usamos funções matemáticas clássicas (Robert Penner).
function _easeBounceOut(t) {
    if (t < 1/2.75) return 7.5625 * t * t;
    if (t < 2/2.75) { t -= 1.5/2.75; return 7.5625 * t * t + 0.75; }
    if (t < 2.5/2.75) { t -= 2.25/2.75; return 7.5625 * t * t + 0.9375; }
    t -= 2.625/2.75; return 7.5625 * t * t + 0.984375;
}
function _easeBounceIn(t) { return 1 - _easeBounceOut(1 - t); }
function _easeBounceInOut(t) {
    if (t < 0.5) return _easeBounceIn(t * 2) * 0.5;
    return _easeBounceOut(t * 2 - 1) * 0.5 + 0.5;
}
function _easeElasticOut(t) {
    if (t === 0 || t === 1) return t;
    var c4 = (2 * Math.PI) / 0.4;  // period
    return Math.pow(2, -10 * t) * Math.sin((t - 0.1) * c4) + 1;
}
function _easeElasticIn(t) {
    if (t === 0 || t === 1) return t;
    var c4 = (2 * Math.PI) / 0.4;
    return -Math.pow(2, 10 * (t - 1)) * Math.sin((t - 1 - 0.1) * c4);
}
function _easeElasticInOut(t) {
    if (t === 0 || t === 1) return t;
    var c5 = (2 * Math.PI) / 0.45;
    if (t < 0.5) return -(Math.pow(2, 20 * t - 10) * Math.sin((20 * t - 11.125) * c5)) / 2;
    return (Math.pow(2, -20 * t + 10) * Math.sin((20 * t - 11.125) * c5)) / 2 + 1;
}
// Retorna função paramétrica pra (preset, dir) ou null se deve usar bezier normal
function _getParametricEasing(presetTag) {
    if (!presetTag) return null;
    var parts = String(presetTag).split(':');
    var preset = parts[0]; var dir = parts[1] || 'in';
    if (preset === 'bounce') {
        if (dir === 'out') return _easeBounceOut;
        if (dir === 'inout') return _easeBounceInOut;
        return _easeBounceIn;
    }
    if (preset === 'elastic') {
        if (dir === 'out') return _easeElasticOut;
        if (dir === 'inout') return _easeElasticInOut;
        return _easeElasticIn;
    }
    return null;
}

function applyBezierEasing(cp1x, cp1y, cp2x, cp2y, mode, presetTag) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var totalApplied = 0;
        var totalKeys = 0;
        var errors = [];
        // Bounce/elastic precisam de MAIS amostras (oscilações múltiplas)
        // Bezier comum: 10 steps é suficiente; bounce: 40 pra capturar os pulos.
        var paramFn = _getParametricEasing(presetTag);
        var STEPS = paramFn ? 40 : 10;

        for (var ci = 0; ci < clips.length; ci++) {
            var clip = clips[ci];
            try {
                var components = clip.components;
                for (var comp = 0; comp < components.numItems; comp++) {
                    var component = components[comp];
                    for (var p = 0; p < component.properties.numItems; p++) {
                        var prop = component.properties[p];
                        try {
                            if (!prop.isTimeVarying()) continue;
                            var keys = prop.getKeys();
                            if (!keys || keys.length < 2) continue;

                            // Collect key pairs to process
                            var pairs = [];
                            for (var k = 0; k < keys.length - 1; k++) {
                                pairs.push([k, k + 1]);
                            }

                            // Process each pair (in reverse to not mess up indices)
                            for (var pi = pairs.length - 1; pi >= 0; pi--) {
                                var kA = pairs[pi][0];
                                var kB = pairs[pi][1];
                                var timeA = keys[kA];
                                var timeB = keys[kB];
                                var valA = prop.getValueAtKey(timeA);
                                var valB = prop.getValueAtKey(timeB);

                                // Use explicit Number() to parse ticks (string in Premiere)
                                var ticksA = Number(timeA.ticks);
                                var ticksB = Number(timeB.ticks);
                                if (isNaN(ticksA) || isNaN(ticksB)) continue;
                                var durationTicks = ticksB - ticksA;
                                if (durationTicks <= 0) continue;

                                // First pass: collect all new keyframe data
                                var newKeys = [];
                                for (var s = 1; s < STEPS; s++) {
                                    var fraction = s / STEPS;
                                    // Usa função paramétrica (bounce/elastic) ou bezier
                                    var easedFraction = paramFn
                                        ? paramFn(fraction)
                                        : evalBezierCurve(fraction, cp1x, cp1y, cp2x, cp2y);
                                    var midTicks = ticksA + durationTicks * fraction;

                                    var midVal;
                                    if (typeof valA === 'number') {
                                        midVal = valA + (valB - valA) * easedFraction;
                                    } else if (valA instanceof Array) {
                                        midVal = [];
                                        for (var d = 0; d < valA.length; d++) {
                                            midVal.push(valA[d] + (valB[d] - valA[d]) * easedFraction);
                                        }
                                    } else {
                                        continue;
                                    }
                                    newKeys.push({ ticks: midTicks, val: midVal });
                                }

                                // Second pass: add keys and set values
                                for (var nk = 0; nk < newKeys.length; nk++) {
                                    try {
                                        var midTime = createTimeAt(newKeys[nk].ticks);
                                        prop.addKey(midTime);
                                        // Re-create time for setValueAtKey to ensure exact match
                                        var setTime = createTimeAt(newKeys[nk].ticks);
                                        prop.setValueAtKey(setTime, newKeys[nk].val);
                                        totalKeys++;
                                    } catch (ek) {
                                        errors.push(prop.displayName + ':' + ek.toString());
                                    }
                                }
                                totalApplied++;
                            }
                        } catch (ep) {
                            errors.push('prop:' + ep.toString());
                        }
                    }
                }
            } catch (ec) {
                errors.push('clip:' + ec.toString());
            }
        }

        if (totalApplied === 0) {
            if (errors.length > 0) return 'ERRORS:' + errors.join('|');
            return 'NO_KEYS';
        }
        forceUIRefresh();
        return 'OK:' + totalApplied + ':' + totalKeys;
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// Simple import to root (no bin)
function importFileToProject(filePath) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var f = new File(filePath);
        if (!f.exists) return 'FILE_NOT_FOUND';
        app.project.importFiles([filePath]);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// ============================================
// ANCHOR POINT PRO v2
// 25-point grid (5x5), target modes (clip/sequence), Vector Motion support
// Uses normalized coordinates (0-1) for sub-anchor precision
// ============================================

// Localized property name variants (short prefixes for accent-safe matching)
var _anchorNames = ['Anchor Point', 'Ponto de ancoragem', 'Ponto de Ancoragem', 'Anchor', 'Ankerpunkt', "Point d'ancrage", 'Punto de anclaje'];
var _positionNames = ['Position', 'Posicao', 'Posicion'];
var _scaleNames = ['Scale', 'Escala', 'Skalierung', 'Echelle'];
var _motionNames = ['Motion', 'Movimento', 'Movement', 'Bewegung', 'Mouvement', 'Movimiento'];
var _vectorMotionNames = ['Vector Motion', 'Movimento Vetorial', 'Vectorielle Bewegung'];
var _transformNames = ['Transform', 'Transformar', 'Transformieren', 'Transformer', 'Trasformare'];

// Short prefixes for accent-safe fallback matching (handles Posicao vs Posição etc.)
var _anchorPrefixes = ['anchor', 'ancor', 'anker', 'ancla'];
var _positionPrefixes = ['posi'];
var _scalePrefixes = ['scal', 'esca', 'skal'];

// Find a named component on a clip
function findComponentByName(clip, names) {
    try {
        var comps = clip.components;
        for (var i = 0; i < comps.numItems; i++) {
            var dn = comps[i].displayName;
            for (var n = 0; n < names.length; n++) {
                if (dn === names[n]) return comps[i];
            }
        }
        // Partial match fallback (case-insensitive)
        for (var i = 0; i < comps.numItems; i++) {
            var dnLow = comps[i].displayName.toLowerCase();
            for (var n = 0; n < names.length; n++) {
                if (dnLow.indexOf(names[n].toLowerCase()) >= 0) return comps[i];
            }
        }
    } catch (e) {}
    return null;
}

// Find Motion component — first by name, then by property count heuristic
function findMotionComponent(clip) {
    var c = findComponentByName(clip, _motionNames);
    if (c) return c;
    try {
        for (var i = 1; i < clip.components.numItems; i++) {
            try {
                if (clip.components[i].properties && clip.components[i].properties.numItems >= 4)
                    return clip.components[i];
            } catch (e2) {}
        }
    } catch (e) {}
    return null;
}

// Find Vector Motion component
function findVectorMotionComponent(clip) {
    return findComponentByName(clip, _vectorMotionNames);
}

// Find Transform effect component on clip
function findTransformComponent(clip) {
    return findComponentByName(clip, _transformNames);
}

// Find a property in a component by name (exact, then partial, then prefix)
function findProp(component, names, prefixes) {
    if (!component) return null;
    try {
        var props = component.properties;
        // Pass 1: exact match
        for (var p = 0; p < props.numItems; p++) {
            var pn = props[p].displayName;
            for (var n = 0; n < names.length; n++) {
                if (pn === names[n]) return props[p];
            }
        }
        // Pass 2: partial match (handles localized names)
        for (var p = 0; p < props.numItems; p++) {
            var pnL = props[p].displayName.toLowerCase();
            for (var n = 0; n < names.length; n++) {
                if (pnL.indexOf(names[n].toLowerCase()) >= 0) return props[p];
            }
        }
        // Pass 3: short prefix match (handles accented chars like Posição, Posición)
        if (prefixes) {
            for (var p = 0; p < props.numItems; p++) {
                var pnL2 = props[p].displayName.toLowerCase();
                for (var n = 0; n < prefixes.length; n++) {
                    if (pnL2.indexOf(prefixes[n]) >= 0) return props[p];
                }
            }
        }
    } catch (e) {}
    return null;
}

// Get source clip pixel dimensions (multiple fallback methods)
function getSourceDimensions(clip) {
    var pi = clip.projectItem;
    if (!pi) return null;
    // Method 1: getMetadataValue (Premiere 2020+)
    try {
        var vw = parseInt(pi.getMetadataValue('Column.Intrinsic.VideoWidth'), 10);
        var vh = parseInt(pi.getMetadataValue('Column.Intrinsic.VideoHeight'), 10);
        if (vw > 0 && vh > 0) return { width: vw, height: vh };
    } catch (e) {}
    // Method 2: Parse project metadata XML
    try {
        var md = pi.getProjectMetadata();
        if (md) {
            var w = 0, h = 0;
            var wi = md.indexOf('Column.Intrinsic.VideoWidth');
            if (wi < 0) wi = md.indexOf('VideoWidth');
            if (wi >= 0) { var ws = md.indexOf('>', wi) + 1; w = parseInt(md.substring(ws, md.indexOf('<', ws)), 10); }
            var hi = md.indexOf('Column.Intrinsic.VideoHeight');
            if (hi < 0) hi = md.indexOf('VideoHeight');
            if (hi >= 0) { var hs = md.indexOf('>', hi) + 1; h = parseInt(md.substring(hs, md.indexOf('<', hs)), 10); }
            if (w > 0 && h > 0) return { width: w, height: h };
        }
    } catch (e) {}
    // Method 3: XMP
    try {
        var xmp = pi.getXMPMetadata();
        if (xmp) {
            var wi2 = xmp.indexOf('stDim:w="');
            var hi2 = xmp.indexOf('stDim:h="');
            if (wi2 >= 0 && hi2 >= 0) {
                var w2 = parseInt(xmp.substring(wi2 + 9, xmp.indexOf('"', wi2 + 9)), 10);
                var h2 = parseInt(xmp.substring(hi2 + 9, xmp.indexOf('"', hi2 + 9)), 10);
                if (w2 > 0 && h2 > 0) return { width: w2, height: h2 };
            }
        }
    } catch (e) {}
    return null;
}

// Apply anchor to a single component (Motion or Vector Motion)
// anchorNorm: [0-1, 0-1], compensate: bool
// Auto-detects if Premiere uses normalized (0-1) or pixel values for anchor/position
// Returns: { ok: bool, debug: string }
function _applyAnchorToComponent(comp, clip, anchorNorm, compensate, seqW, seqH) {
    if (!comp) return { ok: false, debug: 'no-comp' };
    var anchorProp = findProp(comp, _anchorNames, _anchorPrefixes);
    var posProp = findProp(comp, _positionNames, _positionPrefixes);
    var scaleProp = findProp(comp, _scaleNames, _scalePrefixes);

    if (!anchorProp) return { ok: false, debug: 'no-anchor-prop' };
    if (!posProp) return { ok: false, debug: 'no-pos-prop(comp=' + comp.displayName + ')' };

    var oldAnchor = anchorProp.getValue();
    var oldPos = posProp.getValue();
    if (!oldAnchor || !oldPos) return { ok: false, debug: 'no-values' };

    // Get current scale
    var scale = 1;
    try {
        var sv = scaleProp ? scaleProp.getValue() : 100;
        if (typeof sv === 'number') scale = sv / 100;
        else if (sv instanceof Array) scale = sv[0] / 100;
    } catch (e) {}

    // MODE: detecta pela combinacao de anchor + position.
    // Premiere 2025+ usa NORMALIZED [0..1] tambem no Motion. Versoes antigas
    // e Transform sempre normalized. Pixel mode: valores >> 1 (centro = srcW/2, srcH/2).
    // Heuristica robusta: se AMBOS anchor[0] AND position[0] estao em [0, 5], normalized.
    // Pixel-(0,0) num clip pixel mode nao acontece junto com position-(0,0) — em pixel
    // mode position default eh (seqW/2, seqH/2), tipicamente > 100.
    var anchorSmall = Math.abs(oldAnchor[0]) < 5.0 && Math.abs(oldAnchor[1]) < 5.0;
    var posSmall = Math.abs(oldPos[0]) < 5.0 && Math.abs(oldPos[1]) < 5.0;
    var isNormalized = anchorSmall && posSmall;

    if (isNormalized) {
        // ═══ NORMALIZED MODE ═══
        // Anchor and Position are both 0-1 (Premiere Pro 2025+)
        // anchorNorm is already in 0-1 space — use directly
        var newAx = anchorNorm[0];
        var newAy = anchorNorm[1];

        if (compensate) {
            var dAx = newAx - oldAnchor[0];
            var dAy = newAy - oldAnchor[1];
            // For compensation: need to account for source/sequence size ratio
            var srcDims = getSourceDimensions(clip);
            var ratioW = 1, ratioH = 1;
            if (srcDims && seqW > 0 && seqH > 0) {
                ratioW = srcDims.width / seqW;
                ratioH = srcDims.height / seqH;
            }
            var newPx = oldPos[0] + dAx * scale * ratioW;
            var newPy = oldPos[1] + dAy * scale * ratioH;
            anchorProp.setValue([newAx, newAy], true);
            posProp.setValue([newPx, newPy], true);
        } else {
            anchorProp.setValue([newAx, newAy], true);
        }

        return { ok: true, debug: 'NORM a[' + newAx.toFixed(3) + ',' + newAy.toFixed(3) + '] pos[' + oldPos[0].toFixed(3) + ',' + oldPos[1].toFixed(3) + ']' };

    } else {
        // ═══ PIXEL MODE ═══
        // Anchor is in source pixels, Position is normalized to sequence
        var srcDims = getSourceDimensions(clip);
        var srcW, srcH;
        if (srcDims) { srcW = srcDims.width; srcH = srcDims.height; }
        else if (oldAnchor[0] > 10 && oldAnchor[1] > 10) {
            srcW = Math.round(oldAnchor[0] * 2);
            srcH = Math.round(oldAnchor[1] * 2);
        } else { srcW = seqW; srcH = seqH; }

        var newAx = anchorNorm[0] * srcW;
        var newAy = anchorNorm[1] * srcH;

        if (compensate) {
            var dAx = newAx - oldAnchor[0];
            var dAy = newAy - oldAnchor[1];
            var newPx = oldPos[0] + (dAx * scale) / seqW;
            var newPy = oldPos[1] + (dAy * scale) / seqH;
            anchorProp.setValue([newAx, newAy], true);
            posProp.setValue([newPx, newPy], true);
        } else {
            anchorProp.setValue([newAx, newAy], true);
        }

        return { ok: true, debug: 'PX ' + srcW + 'x' + srcH + ' a[' + Math.round(newAx) + ',' + Math.round(newAy) + ']' };
    }
}

// Main entry: set anchor with normalized coords
// nx, ny: 0-1 normalized position
// compensate: boolean
// target: 'clip' or 'sequence'
// useVector: boolean — also apply to Vector Motion
// useTransform: boolean — apply to Transform effect instead of/alongside Motion
function setClipAnchorPointPro(nx, ny, compensate, target, useVector, useTransform) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var seqW = parseInt(seq.frameSizeHorizontal, 10) || 1920;
        var seqH = parseInt(seq.frameSizeVertical, 10) || 1080;
        var totalApplied = 0;
        var debug = '';

        var anchorNorm = [nx, ny];

        for (var ci = 0; ci < clips.length; ci++) {
            var clip = clips[ci];
            try {
                if (useTransform) {
                    // Apply to Transform effect
                    var transform = findTransformComponent(clip);
                    if (transform) {
                        var tResult = _applyAnchorToComponent(transform, clip, anchorNorm, compensate, seqW, seqH);
                        if (tResult.ok) { debug = 'Transform:' + tResult.debug; totalApplied++; }
                        else { debug = 'Transform:' + tResult.debug; }
                    } else {
                        debug = 'Transform not found — add the effect first';
                    }
                } else {
                    // Apply to standard Motion
                    var motion = findMotionComponent(clip);
                    var result = _applyAnchorToComponent(motion, clip, anchorNorm, compensate, seqW, seqH);
                    if (result.ok) { debug = result.debug; totalApplied++; }
                    else { debug = result.debug; }
                }

                // Apply to Vector Motion if requested (alongside Motion or Transform)
                if (useVector) {
                    var vMotion = findVectorMotionComponent(clip);
                    if (vMotion) {
                        _applyAnchorToComponent(vMotion, clip, anchorNorm, compensate, seqW, seqH);
                    }
                }
            } catch (ec) {
                debug = 'err:' + ec.toString();
            }
        }

        if (totalApplied === 0) return 'NO_APPLIED|' + debug;
        forceUIRefresh();
        return 'OK:' + totalApplied + '|' + debug;
    } catch (e) {
        return 'ERROR:' + e.toString();
    }
}

// DIAGNOSTIC: dump all raw values from a selected clip for debugging
function debugAnchorValues() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var seqW = parseInt(seq.frameSizeHorizontal, 10) || 0;
        var seqH = parseInt(seq.frameSizeVertical, 10) || 0;
        var clip = clips[0];
        var motion = findMotionComponent(clip);
        if (!motion) return 'NO_MOTION|comp_count=' + clip.components.numItems;

        var anchorProp = findProp(motion, _anchorNames, _anchorPrefixes);
        var posProp = findProp(motion, _positionNames, _positionPrefixes);
        var scaleProp = findProp(motion, _scaleNames, _scalePrefixes);

        var info = 'seq=' + seqW + 'x' + seqH;
        info += '|motion=' + motion.displayName;

        // List ALL properties in Motion
        var propList = [];
        for (var p = 0; p < motion.properties.numItems; p++) {
            try {
                var pr = motion.properties[p];
                var v = pr.getValue();
                propList.push(pr.displayName + '=' + JSON.stringify(v));
            } catch (e) {
                propList.push(motion.properties[p].displayName + '=ERR');
            }
        }
        info += '|props=[' + propList.join(';') + ']';

        // Source dimensions
        var srcDims = getSourceDimensions(clip);
        info += '|srcDims=' + (srcDims ? srcDims.width + 'x' + srcDims.height : 'null');

        // Raw anchor + position values
        if (anchorProp) {
            var av = anchorProp.getValue();
            info += '|anchor=' + anchorProp.displayName + ':' + JSON.stringify(av);
        } else { info += '|anchor=NOT_FOUND'; }
        if (posProp) {
            var pv = posProp.getValue();
            info += '|pos=' + posProp.displayName + ':' + JSON.stringify(pv);
        } else { info += '|pos=NOT_FOUND'; }
        if (scaleProp) {
            var sv = scaleProp.getValue();
            info += '|scale=' + scaleProp.displayName + ':' + JSON.stringify(sv);
        } else { info += '|scale=NOT_FOUND'; }

        // Vector Motion too
        var vm = findVectorMotionComponent(clip);
        if (vm) {
            var vmProps = [];
            for (var p = 0; p < vm.properties.numItems; p++) {
                try {
                    var pr2 = vm.properties[p];
                    vmProps.push(pr2.displayName + '=' + JSON.stringify(pr2.getValue()));
                } catch (e) { vmProps.push(vm.properties[p].displayName + '=ERR'); }
            }
            info += '|vectorMotion=[' + vmProps.join(';') + ']';
        } else { info += '|vectorMotion=NONE'; }

        return info;
    } catch (e) { return 'ERROR:' + e.toString(); }
}

// Legacy compat — keep old function name working
function setClipAnchorPoint(position, compensate) {
    var map = {
        'tl': [0, 0], 'tc': [0.5, 0], 'tr': [1, 0],
        'cl': [0, 0.5], 'cc': [0.5, 0.5], 'cr': [1, 0.5],
        'bl': [0, 1], 'bc': [0.5, 1], 'br': [1, 1]
    };
    var c = map[position] || [0.5, 0.5];
    return setClipAnchorPointPro(c[0], c[1], compensate, 'clip', false);
}

// Add keyframe at playhead for Position + Anchor Point
// useVector: boolean — also add keyframes on Vector Motion
function addAnchorKeyframe(useVector) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var playhead = seq.getPlayerPosition();
        var totalApplied = 0;

        for (var ci = 0; ci < clips.length; ci++) {
            var clip = clips[ci];
            try {
                // Calculate clip-relative time
                var clipTime = new Time();
                try {
                    clipTime.seconds = playhead.seconds - clip.start.seconds + clip.inPoint.seconds;
                } catch (et) { clipTime = playhead; }

                var compsToProcess = [];
                var motion = findMotionComponent(clip);
                if (motion) compsToProcess.push(motion);
                if (useVector) {
                    var vm = findVectorMotionComponent(clip);
                    if (vm) compsToProcess.push(vm);
                }

                var didSomething = false;
                for (var cpi = 0; cpi < compsToProcess.length; cpi++) {
                    var comp = compsToProcess[cpi];
                    var ancProp = findProp(comp, _anchorNames, _anchorPrefixes);
                    var posProp = findProp(comp, _positionNames, _positionPrefixes);

                    // Fallback scan
                    if (!ancProp || !posProp) {
                        for (var fp = 0; fp < comp.properties.numItems; fp++) {
                            try {
                                var fn = comp.properties[fp].displayName.toLowerCase();
                                if (!ancProp && (fn.indexOf('anchor') >= 0 || fn.indexOf('ancor') >= 0))
                                    ancProp = comp.properties[fp];
                                else if (!posProp && fn.indexOf('posi') >= 0 && fn.indexOf('compo') < 0)
                                    posProp = comp.properties[fp];
                            } catch (e) {}
                        }
                    }

                    // Enable time-varying + add keyframes
                    try { if (ancProp && !ancProp.isTimeVarying()) ancProp.setTimeVarying(true); } catch (e) {}
                    try { if (posProp && !posProp.isTimeVarying()) posProp.setTimeVarying(true); } catch (e) {}
                    try { if (ancProp) { ancProp.addKey(clipTime); didSomething = true; } } catch (e) {}
                    try { if (posProp) { posProp.addKey(clipTime); didSomething = true; } } catch (e) {}
                }

                if (didSomething) totalApplied++;
            } catch (ec) {}
        }

        if (totalApplied === 0) return 'NO_APPLIED';
        forceUIRefresh();
        return 'OK:' + totalApplied;
    } catch (e) {
        return 'ERROR:' + e.toString();
    }
}

// ============================================
// COPY / PASTE (Lion Copy-Pasta)
// Paste: importa arquivo(s)/imagem do clipboard do SO pra timeline
// Copy:  pega arquivo do clip selecionado e poe no clipboard do SO
// ============================================

// Abre dialog nativo pra escolher pasta de salvamento de prints
function cpSelectFolder() {
    try {
        // Abre a pasta do projeto por padrao, se existir
        try {
            if (app.project && app.project.path) {
                var pf = new File(app.project.path);
                if (pf.parent && pf.parent.exists) Folder.current = pf.parent;
            }
        } catch (e) {}
        // Folder.selectDialog e o metodo CORRETO (selectDlg nao existe em ExtendScript)
        var folder = Folder.selectDialog('Escolha a pasta de destino');
        if (!folder) return 'CANCEL';
        return folder.fsName;
    } catch (e) { return 'NO_FOLDER'; }
}

// Retorna o caminho do media do primeiro clip selecionado na timeline
function cpGetSelectedClipMediaPath() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        // Seleção ESTRITA: só aceita clips de fato selecionados, sem fallback
        // pro clip no playhead (que confunde — copia sem ter clicado)
        var clips = [];
        try {
            var sel = seq.getSelection();
            if (sel && sel.length > 0) {
                for (var i = 0; i < sel.length; i++) clips.push(sel[i]);
            }
        } catch (e1) {}
        if (clips.length === 0) {
            // Tenta isSelected() em cada track como fallback (NÃO usa playhead)
            for (var t = 0; t < seq.videoTracks.numTracks; t++) {
                for (var c = 0; c < seq.videoTracks[t].clips.numItems; c++) {
                    try {
                        if (seq.videoTracks[t].clips[c].isSelected()) clips.push(seq.videoTracks[t].clips[c]);
                    } catch (e2) {}
                }
            }
            for (var t2 = 0; t2 < seq.audioTracks.numTracks; t2++) {
                for (var c2 = 0; c2 < seq.audioTracks[t2].clips.numItems; c2++) {
                    try {
                        if (seq.audioTracks[t2].clips[c2].isSelected()) clips.push(seq.audioTracks[t2].clips[c2]);
                    } catch (e3) {}
                }
            }
        }
        if (clips.length === 0) return 'NO_SELECTION';
        var clip = clips[0];
        try {
            var pi = clip.projectItem;
            if (!pi) return 'NO_MEDIA';
            var p = null;
            try { p = pi.getMediaPath(); } catch(e) {}
            if (!p) return 'NO_MEDIA';
            var f = new File(p);
            if (!f.exists) return 'FILE_NOT_FOUND';
            return 'OK:' + f.fsName;
        } catch (e) { return 'ERR:' + e.toString(); }
    } catch (e) { return 'ERR:' + e.toString(); }
}

// Procura item importado no projeto por caminho
function _cpFindItemByPath(filePath) {
    try {
        var target = new File(filePath);
        function search(bin) {
            try {
                for (var i = bin.children.numItems - 1; i >= 0; i--) {
                    var child = bin.children[i];
                    if (child.type === 2) {
                        var found = search(child);
                        if (found) return found;
                    } else {
                        try {
                            var p = child.getMediaPath ? child.getMediaPath() : '';
                            if (p) {
                                var cf = new File(p);
                                if (cf.fsName === target.fsName) return child;
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) {}
            return null;
        }
        return search(app.project.rootItem);
    } catch (e) { return null; }
}

// Importa lista de arquivos pra Premiere, opcionalmente em um bin, opcionalmente insere no playhead
// filePathsJson: string JSON de array de caminhos
function cpImportFiles(filePathsJson, insertAtPlayhead, useBin, binName) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var filePaths = null;
        try { filePaths = eval('(' + filePathsJson + ')'); } catch (e) { return 'PARSE_ERROR'; }
        if (!filePaths || filePaths.length === 0) return 'NO_FILES';

        var seq = app.project.activeSequence;
        var targetBin = null;
        if (useBin) {
            try { targetBin = findOrCreateBin(binName || 'Copy-Pasta'); } catch (e) {}
        }

        var imported = 0;
        var inserted = 0;
        var failed = 0;

        for (var i = 0; i < filePaths.length; i++) {
            var fp = filePaths[i];
            try {
                var f = new File(fp);
                if (!f.exists) { failed++; continue; }

                // Premiere às vezes retorna undefined (sucesso sem return value),
                // então NÃO confia no return — verifica via _cpFindItemByPath depois.
                try {
                    if (targetBin) {
                        app.project.importFiles([fp], false, targetBin, false);
                    } else {
                        app.project.importFiles([fp]);
                    }
                } catch (e1) {}
                $.sleep(300);
                var okImport = !!_cpFindItemByPath(fp);
                // Fallback SÓ se item realmente não está no projeto
                if (!okImport) {
                    try { app.project.importFiles([fp]); } catch (e2) {}
                    $.sleep(300);
                    okImport = !!_cpFindItemByPath(fp);
                }

                if (!okImport) { failed++; continue; }
                imported++;
                $.sleep(500);

                // Insere no playhead (smart track: encontra track LIVRE pelo período
                // INTEIRO do clip — não só no ponto do playhead. Antes verificava só
                // o playhead, e o insertClip empurrava clips ao lado pra abrir espaço.
                // Agora verifica todo o range e usa overwriteClip (não empurra).
                if (insertAtPlayhead && seq) {
                    var clipItem = _cpFindItemByPath(fp);
                    if (clipItem) {
                        try {
                            var insertTime = seq.getPlayerPosition();
                            var insertTicks = Number(insertTime.ticks);
                            var mType = _getMediaType(fp);
                            // Duração do clip a inserir — pega via projectItem
                            var clipDurTicks = _cpGetItemDurationTicks(clipItem, mType);
                            var smartTrack = _cpFindFreeTrackForRange(
                                seq, mType, insertTicks, insertTicks + clipDurTicks
                            );
                            if (smartTrack) {
                                // overwriteClip NÃO empurra clips — se range já é livre,
                                // não destrói nada. Se nenhuma track estava livre, criamos
                                // uma nova (smartTrack já vem garantida livre).
                                smartTrack.overwriteClip(clipItem, insertTime);
                                inserted++;
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) { failed++; }
        }

        try { _cpFlushDbg(); } catch(eFD) {}
        return 'OK:' + imported + ':' + inserted + ':' + failed;
    } catch (e) {
        try { _cpFlushDbg(); } catch(eFD) {}
        return 'ERR:' + e.toString();
    }
}

// ============================================// ============================================

// Smart track finder (LEGACY — só checa playhead). Mantida pra compat.
function findEmptyTrackAtPlayhead(seq, mediaType) {
    return _cpFindFreeTrackForRange(seq, mediaType,
        Number(seq.getPlayerPosition().ticks),
        Number(seq.getPlayerPosition().ticks) + 1);
}

// Pega duração esperada do clip a inserir (em ticks).
// Pra vídeos/áudios: usa outPoint - inPoint do projectItem.
// Pra imagens: usa duração padrão do Premiere (geralmente 5s) ou 5s default.
function _cpGetItemDurationTicks(item, mediaType) {
    var TPS = 254016000000;
    try {
        if (item.getOutPoint && item.getInPoint) {
            var inT = Number(item.getInPoint().ticks);
            var outT = Number(item.getOutPoint().ticks);
            if (outT > inT) return outT - inT;
        }
    } catch(e) {}
    // Fallback: duração default do Premiere pra still images é configurada nas
    // preferências (default 5s = 30 frames @ 6fps still). Usamos 5s como margem.
    return 5 * TPS;
}

// Debug log dedicado pra paste (separa do audio log)
var _cpPasteDbgBuf = [];
function _cpDbg(msg) { _cpPasteDbgBuf.push('[' + (new Date()).toString() + '] ' + msg); }
function _cpFlushDbg() {
    if (_cpPasteDbgBuf.length === 0) return;
    var blob = _cpPasteDbgBuf.join('\n') + '\n';
    _cpPasteDbgBuf = [];
    try {
        var f = new File(Folder.desktop.fsName + '/lion-paste-debug.txt');
        if (f.open('a')) { f.encoding = 'UTF-8'; f.write(blob); f.close(); }
    } catch(e) {}
}

// Acha track LIVRE pra o range [startTicks, endTicks). Se nenhuma livre,
// cria nova track no topo. NUNCA retorna track ocupada — antes retorna null.
function _cpFindFreeTrackForRange(seq, mediaType, startTicks, endTicks) {
    var tracks = (mediaType === 'audio') ? seq.audioTracks : seq.videoTracks;
    var beforeNumTracks = tracks.numTracks;
    _cpDbg('FindFreeTrack: type=' + mediaType + ' range=[' + startTicks + ',' + endTicks + '] numTracks=' + beforeNumTracks);

    // Pra cada track existente (V1 → topo), vê se o range [start, end) está LIVRE
    for (var t = 0; t < beforeNumTracks; t++) {
        var track = tracks[t];
        var occupied = false;
        var why = '';
        try {
            if (track.isLocked && track.isLocked()) { occupied = true; why = 'locked'; }
            for (var c = 0; !occupied && c < track.clips.numItems; c++) {
                var clip = track.clips[c];
                var clipStart = Number(clip.start.ticks);
                var clipEnd = Number(clip.end.ticks);
                if (clipStart < endTicks && clipEnd > startTicks) {
                    occupied = true;
                    why = 'clip-overlap@' + clipStart + '-' + clipEnd;
                }
            }
        } catch (e) { occupied = true; why = 'ex:' + e; }
        _cpDbg('  track[' + t + '] occupied=' + occupied + ' (' + why + ')');
        if (!occupied) { _cpDbg('  -> using existing track[' + t + ']'); return track; }
    }

    // Nenhuma track existente livre — tenta criar uma nova.
    // Snapshot das referências de track ANTES, pra achar a track NOVA depois
    // (QE pode inserir no topo OU no fim, dependendo da versão do Premiere).
    _cpDbg('  All tracks occupied — trying to ADD new track');
    var beforeRefs = [];
    for (var bi = 0; bi < beforeNumTracks; bi++) {
        try { beforeRefs.push(tracks[bi]); } catch(eBR) {}
    }

    var attempts = [];
    if (mediaType === 'audio') {
        attempts = [
            // 5-arg signature (numV, vidIdx, numA, audCh, audIdx) — alguns QE aceitam
            { name: 'QE.addTracks(0,0,1,1,numA)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 0, 1, 1, seq.audioTracks.numTracks); } },
            // 3-arg signature (numV, numA, audIdx)
            { name: 'QE.addTracks(0,1,numA)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 1, seq.audioTracks.numTracks); } },
            { name: 'QE.addAudioTrack(numA)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addAudioTrack(seq.audioTracks.numTracks); } },
            { name: 'QE.addAudioTrack()', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addAudioTrack(); } },
            // 2-arg fallback (adiciona embaixo)
            { name: 'QE.addTracks(0,1)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 1); } },
            { name: 'addTracks(0,0,1,1,-1)', fn: function(){ seq.addTracks(0, 0, 1, 1, -1); } },
            { name: 'audioTracks.addTrack(1)', fn: function(){ seq.audioTracks.addTrack(1); } },
            { name: 'audioTracks.addTrack()', fn: function(){ seq.audioTracks.addTrack(); } }
        ];
    } else {
        attempts = [
            // 5-arg signature (numV, vidIdx, numA, audCh, audIdx) — passa numTracks como posição = topo
            { name: 'QE.addTracks(1,numV,0,0,0)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(1, seq.videoTracks.numTracks, 0, 0, 0); } },
            // 3-arg signature (numV, numA, vidIdx)
            { name: 'QE.addTracks(1,0,numV)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(1, 0, seq.videoTracks.numTracks); } },
            { name: 'QE.addVideoTrack(numV)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addVideoTrack(seq.videoTracks.numTracks); } },
            { name: 'QE.addVideoTrack()', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addVideoTrack(); } },
            // 2-arg fallback (adiciona embaixo — vou tratar isso depois com swap)
            { name: 'QE.addTracks(1,0)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(1, 0); } },
            { name: 'addTracks(1,numVid,0,0,0)', fn: function(){ seq.addTracks(1, seq.videoTracks.numTracks, 0, 0, 0); } },
            { name: 'addTracks(1,-1,0,0,0)', fn: function(){ seq.addTracks(1, -1, 0, 0, 0); } },
            { name: 'videoTracks.addTrack(1)', fn: function(){ seq.videoTracks.addTrack(1); } },
            { name: 'videoTracks.addTrack()', fn: function(){ seq.videoTracks.addTrack(); } }
        ];
    }
    for (var a = 0; a < attempts.length; a++) {
        var att = attempts[a];
        try {
            att.fn();
            // Re-fetch tracks collection — cached ref pode estar stale após addTracks
            var freshTracks = (mediaType === 'audio') ? seq.audioTracks : seq.videoTracks;
            // Premiere pode demorar pra registrar — poll até 8 ciclos × 20ms
            for (var pt = 0; pt < 8; pt++) {
                if (freshTracks.numTracks > beforeNumTracks) break;
                $.sleep(20);
                freshTracks = (mediaType === 'audio') ? seq.audioTracks : seq.videoTracks;
            }
            var afterNum = freshTracks.numTracks;
            _cpDbg('  attempt[' + att.name + ']: numTracks ' + beforeNumTracks + ' -> ' + afterNum);
            if (afterNum > beforeNumTracks) {
                // ACHA o índice da track EMPTY (sem clips) — método mais confiável
                // que ref-compare em ExtendScript (refs de C++ wrappers podem
                // não comparar corretamente com ===).
                // Premiere SEMPRE adiciona track NOVA empty, então sweep por empty.
                var emptyTracks = [];
                for (var ei = 0; ei < afterNum; ei++) {
                    try {
                        var clipCount = freshTracks[ei].clips.numItems;
                        _cpDbg('    track[' + ei + '].clips.numItems=' + clipCount);
                        if (clipCount === 0) emptyTracks.push(ei);
                    } catch(eEI) { _cpDbg('    track[' + ei + '] err: ' + eEI); }
                }
                _cpDbg('  emptyTracks: ' + emptyTracks.join(','));
                if (emptyTracks.length > 0) {
                    // Pega a empty no índice mais ALTO (topo da timeline em Premiere = numTracks-1)
                    var chosenIdx = emptyTracks[emptyTracks.length - 1];
                    _cpDbg('  -> using empty track at index ' + chosenIdx + ' (topmost empty)');
                    return freshTracks[chosenIdx];
                }
                // Sem empty? Tenta ref-compare como fallback
                _cpDbg('  no empty track — trying ref-compare');
                for (var ni = 0; ni < afterNum; ni++) {
                    var candidate = freshTracks[ni];
                    var isOld = false;
                    for (var bj = 0; bj < beforeRefs.length; bj++) {
                        if (beforeRefs[bj] === candidate) { isOld = true; break; }
                    }
                    if (!isOld) {
                        _cpDbg('  -> NEW track via ref-compare at index ' + ni);
                        return candidate;
                    }
                }
                _cpDbg('  could not identify new track — bailing');
                return null;
            }
        } catch (e) {
            _cpDbg('  attempt[' + att.name + '] threw: ' + e.toString());
        }
    }

    _cpDbg('  ALL ATTEMPTS FAILED — returning null (no overwrite)');
    return null;
}

// Detect media type from file extension
function _getMediaType(filePath) {
    var ext = filePath.replace(/^.*\./, '').toLowerCase();
    if (/^(mp3|wav|aac|flac|ogg|m4a|wma|aiff)$/.test(ext)) return 'audio';
    return 'video';
}

// Insert a clip into the active sequence at the playhead position
function insertClipToTimeline(filePath) {
    try {
        if (!app.project) return 'NO_PROJECT';

        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var f = new File(filePath);
        var clipItem = null;
        var fileName = f.displayName.replace(/\.[^.]+$/, '');

        // Match by file path (most accurate)
        function matchByPath(item) {
            try {
                var treePath = item.getMediaPath ? item.getMediaPath() : '';
                if (treePath) {
                    var itemFile = new File(treePath);
                    if (itemFile.fsName === f.fsName) return true;
                }
            } catch (e) {}
            return false;
        }

        // Search a specific bin for the clip
        function searchBin(bin, usePathMatch) {
            for (var i = bin.children.numItems - 1; i >= 0; i--) {
                var child = bin.children[i];
                if (child.type === 2) {
                    var found = searchBin(child, usePathMatch);
                    if (found) return found;
                } else {
                    if (usePathMatch && matchByPath(child)) return child;
                    if (!usePathMatch && (child.name === f.displayName || child.name === fileName)) return child;
                }
            }
            return null;
        }

        // Strategy 1: Find "YouTube Downloads" bin and search by path (most accurate)
        var ytBin = findOrCreateBin('YouTube Downloads');
        if (ytBin) {
            clipItem = searchBin(ytBin, true);
            if (!clipItem) clipItem = searchBin(ytBin, false);
        }

        // Strategy 2: Search entire project by path
        if (!clipItem) clipItem = searchBin(app.project.rootItem, true);

        // Strategy 3: Search entire project by name (last resort, search from end = newest first)
        if (!clipItem) clipItem = searchBin(app.project.rootItem, false);

        if (!clipItem) return 'CLIP_NOT_FOUND';

        var insertTime = seq.getPlayerPosition();
        var insertTicks = Number(insertTime.ticks);
        var mType = _getMediaType(filePath);

        // Smart track: encontra track livre pra TODA duração do clip
        // (não só ponto do playhead — antes sobreescrevia textos/clips adjacentes).
        var clipDurTicks = _cpGetItemDurationTicks(clipItem, mType);
        var smartTrack = _cpFindFreeTrackForRange(
            seq, mType, insertTicks, insertTicks + clipDurTicks
        );
        if (!smartTrack) return 'NO_TRACK';

        // overwriteClip NÃO empurra clips ao lado (insertClip empurrava).
        // _cpFindFreeTrackForRange já garante range livre, então overwrite é seguro.
        smartTrack.overwriteClip(clipItem, insertTime);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}



// ============================================
// BACKGROUND REMOVER — pega clip selecionado
// ============================================
function lwBgGetSelected() {
    try {
        if (!app.project) return "NO_PROJECT";
        var seq = app.project.activeSequence;
        var info = null;

        // 1) Tenta pegar do timeline (clip selecionado em alguma track)
        if (seq) {
            var foundClip = null, foundTrackIdx = -1, foundTrackTyp = "video";
            for (var v = 0; v < seq.videoTracks.numTracks; v++) {
                var trk = seq.videoTracks[v];
                for (var c = 0; c < trk.clips.numItems; c++) {
                    var cl = trk.clips[c];
                    if (cl.isSelected && cl.isSelected()) {
                        foundClip = cl;
                        foundTrackIdx = v;
                        break;
                    }
                }
                if (foundClip) break;
            }
            if (foundClip) {
                // Rejeita NEST (nested sequence) — não tem mídia real, é virtual
                try {
                    var pItem = foundClip.projectItem;
                    if (pItem && pItem.isSequence && pItem.isSequence()) return "NEST_UNSUPPORTED";
                    if (pItem && pItem.type === ProjectItemType.SEQUENCE) return "NEST_UNSUPPORTED";
                } catch(e) {}
                var path = "";
                try { path = foundClip.projectItem.getMediaPath(); } catch(e) {}
                // Se path vazio ou aponta pra .prproj (nest sem media path), rejeita
                if (!path || String(path).toLowerCase().indexOf('.prproj') >= 0) return "NO_MEDIA_PATH";
                // Valida extensão — só aceita imagens
                var extMatch = String(path).toLowerCase().match(/\.([a-z0-9]+)$/);
                var ext = extMatch ? extMatch[1] : '';
                var IMG_EXTS = { png:1, jpg:1, jpeg:1, webp:1, tif:1, tiff:1, bmp:1, gif:1, heic:1, heif:1 };
                if (!IMG_EXTS[ext]) return "NOT_AN_IMAGE:" + ext;
                var TPS = 254016000000;
                var startT = Number(foundClip.start.ticks);
                var endT = Number(foundClip.end.ticks);
                var inT = 0;
                try { inT = Number(foundClip.inPoint.ticks); } catch(e) {}
                // Duração do trecho cortado na timeline (in→out do clip)
                var clipDurTicks = endT - startT;
                var outT = inT + clipDurTicks;
                info = {
                    source: "timeline",
                    path: path,
                    name: foundClip.name,
                    startTicks: String(startT),
                    endTicks: String(endT),
                    inPointTicks: String(inT),
                    outPointTicks: String(outT),
                    // segundos (mais conveniente pro client)
                    inSec: inT / TPS,
                    outSec: outT / TPS,
                    durationSec: clipDurTicks / TPS,
                    trackIdx: foundTrackIdx,
                    trackTyp: foundTrackTyp
                };
            }
        }

        // 2) Senão, tenta selection do project panel
        if (!info) {
            var sel = app.project.getSelection ? app.project.getSelection() : null;
            if (sel && sel.length > 0) {
                var item = sel[0];
                var p = "";
                try { p = item.getMediaPath(); } catch(e) {}
                if (p) {
                    info = { source: "project", path: p, name: item.name };
                }
            }
        }

        if (!info) return "NO_SELECTION";
        if (!info.path) return "NO_MEDIA_PATH";
        return "OK:" + JSON.stringify(info);
    } catch (e) {
        return "ERROR:" + e.toString();
    }
}

// Importa o PNG e insere no timeline na MESMA posicao do clip original,
// numa track ACIMA (cria nova track se preciso). Se origem não e timeline,
// soh importa pra raiz do projeto.
function lwBgImportAndPlace(filePath, srcInfoJson) {
    try {
        if (!app.project) return "NO_PROJECT";
        var f = new File(filePath);
        if (!f.exists) return "FILE_NOT_FOUND";

        var info = null;
        try { info = eval("(" + srcInfoJson + ")"); } catch(e) {}

        // Importa o PNG
        var beforeIds = {};
        for (var i = 0; i < app.project.rootItem.children.numItems; i++) {
            beforeIds[app.project.rootItem.children[i].nodeId] = true;
        }
        app.project.importFiles([filePath]);
        $.sleep(400);

        // Acha o item recem importado
        var newItem = null;
        for (var j = app.project.rootItem.children.numItems - 1; j >= 0; j--) {
            var ch = app.project.rootItem.children[j];
            if (!beforeIds[ch.nodeId]) { newItem = ch; break; }
        }
        if (!newItem) {
            // fallback: pega o ultimo
            newItem = app.project.rootItem.children[app.project.rootItem.children.numItems - 1];
        }

        // Move pro bin "Lion BG Remover"
        try {
            var bin = findOrCreateBin("Lion BG Remover");
            if (bin && newItem) try { newItem.moveBin(bin); } catch(e) {}
        } catch(e) {}

        if (!info || info.source !== "timeline") {
            return "OK:imported";
        }

        // Insere no timeline na MESMA posicao do clip original.
        // ANTES: sempre na track imediatamente acima (info.trackIdx + 1) — se
        // ela tinha algo, dava overwrite/insert que destruía/empurrava clips.
        // AGORA: usa _cpFindFreeTrackForRange que verifica o range INTEIRO do
        // clip e procura a 1ª track LIVRE acima do original. Se nenhuma livre,
        // CRIA nova track no topo. Sempre overwriteClip (que não empurra).
        var seq = app.project.activeSequence;
        if (!seq) return "OK:imported";

        var startTicksNum = Number(info.startTicks);
        var endTicksNum = Number(info.endTicks);
        // Fallback de duração se endTicks inválido
        if (!endTicksNum || endTicksNum <= startTicksNum) {
            var TPS = 254016000000;
            endTicksNum = startTicksNum + 5 * TPS; // 5s default pra imagens
        }

        var startTrackIdx = (info.trackIdx >= 0 ? info.trackIdx + 1 : 0);
        var dstTrack = _findFreeVideoTrackFromIndex(
            seq, startTrackIdx, startTicksNum, endTicksNum
        );

        if (!dstTrack) return "OK:imported";

        // Cria Time pro start
        var startTime = new Time();
        startTime.ticks = info.startTicks;
        try {
            // overwriteClip: range já garantidamente livre, não destrói nada
            dstTrack.overwriteClip(newItem, startTime);
        } catch (eOv) {
            // Fallback raro: insertClip
            try { dstTrack.insertClip(newItem, startTime); } catch(e2) {}
        }
        return "OK:placed";
    } catch (e) {
        return "ERROR:" + e.toString();
    }
}

// Procura primeira track de VÍDEO livre no range [startTicks, endTicks)
// começando do índice fromIdx. Se nenhuma livre, CRIA uma nova track no topo.
// Mesma lógica do _cpFindFreeTrackForRange mas começando de um índice especificado
// (pra preferir track logo acima do original).
function _findFreeVideoTrackFromIndex(seq, fromIdx, startTicks, endTicks) {
    var tracks = seq.videoTracks;

    // Tenta cada track existente A PARTIR de fromIdx
    for (var t = fromIdx; t < tracks.numTracks; t++) {
        var track = tracks[t];
        var occupied = false;
        try {
            if (track.isLocked && track.isLocked()) { occupied = true; }
            for (var c = 0; !occupied && c < track.clips.numItems; c++) {
                var clip = track.clips[c];
                var clipStart = Number(clip.start.ticks);
                var clipEnd = Number(clip.end.ticks);
                // Overlap: clipStart < endTicks E clipEnd > startTicks
                if (clipStart < endTicks && clipEnd > startTicks) {
                    occupied = true;
                }
            }
        } catch (e) { occupied = true; }
        if (!occupied) return track;
    }

    // Nenhuma track existente livre — cria nova no topo
    try {
        try { seq.videoTracks.addTrack(1); } catch(e) {
            try { seq.addTracks(1, 0); } catch(e2) {}
        }
        return seq.videoTracks[seq.videoTracks.numTracks - 1];
    } catch (e) {
        return tracks[fromIdx] || tracks[0];
    }
}

// ─── Recarrega o PNG no Premiere após edição da máscara ────────────
// Quando o user edita a máscara no app desktop e salva, o conteúdo do
// arquivo no disco mudou (mesmo path). Premiere mantém cache da media
// do projectItem — chamamos changeMediaPath pra forçar reler do disco.
function lwBgRefreshFootage(filePath) {
    try {
        if (!app.project) return "NO_PROJECT";
        var f = new File(filePath);
        if (!f.exists) return "FILE_NOT_FOUND";
        var fsTarget = f.fsName;

        // Procura todos os projectItems que apontam pra esse path
        var matches = [];
        function search(bin) {
            try {
                for (var i = 0; i < bin.children.numItems; i++) {
                    var ch = bin.children[i];
                    if (ch.type === 2) { // bin
                        search(ch);
                    } else {
                        try {
                            var p = ch.getMediaPath ? ch.getMediaPath() : '';
                            if (p) {
                                var cf = new File(p);
                                if (cf.fsName === fsTarget) matches.push(ch);
                            }
                        } catch(e) {}
                    }
                }
            } catch(e) {}
        }
        search(app.project.rootItem);

        if (matches.length === 0) return "NOT_IN_PROJECT";

        // Força refresh em cada match
        var refreshed = 0;
        for (var k = 0; k < matches.length; k++) {
            var item = matches[k];
            try {
                if (typeof item.refreshMedia === 'function') {
                    item.refreshMedia();
                    refreshed++;
                } else if (typeof item.changeMediaPath === 'function') {
                    // Truque: re-aponta pro mesmo path → Premiere relê do disco
                    item.changeMediaPath(filePath);
                    refreshed++;
                }
            } catch(e) {}
        }

        // Nudge playhead — força timeline repintar com a media nova
        try {
            var seq = app.project.activeSequence;
            if (seq) {
                var pos = seq.getPlayerPosition();
                var pT = Number(pos.ticks);
                seq.setPlayerPosition(String(pT + 1));
                $.sleep(40);
                seq.setPlayerPosition(String(pT));
            }
        } catch(e) {}

        return "OK:" + refreshed;
    } catch (e) {
        return "ERR:" + e.toString();
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — Excalibur-style command palette
// Lista todos os efeitos / presets / transições disponíveis e
// permite aplicar via matchName ao clip selecionado.
// ═══════════════════════════════════════════════════════════════════

// Helper: lê propriedade safely retornando string descritiva
function _safeStr(obj, prop) {
    try {
        var v = obj[prop];
        if (v === null) return 'null';
        if (v === undefined) return 'undefined';
        if (typeof v === 'function') return 'fn()';
        return String(v).substr(0, 80);
    } catch(e) { return 'ERR:' + e; }
}

// Helper: escapa string pra JSON-compatible (evita break em title/name com aspas/backslash)
function _lwEscapeJson(s) {
    if (s == null) return '';
    s = String(s);
    s = s.replace(/\\/g, '\\\\');
    s = s.replace(/"/g, '\\"');
    s = s.replace(/\n/g, '\\n');
    s = s.replace(/\r/g, '');
    s = s.replace(/\t/g, ' ');
    return s;
}

// Lista todos os effects + transitions + presets disponíveis no Premiere.
// Retorna JSON: { items: [...], debug: "...", error: "..." }
function lwSearchListAll() {
    var items = [];
    var seen = {};
    var debug = [];
    var errors = [];

    function pushItem(kind, name, matchName, category) {
        if (!matchName) return;
        var key = kind + '|' + matchName;
        if (seen[key]) return;
        seen[key] = 1;
        items.push({
            kind: kind,
            name: name || matchName,
            matchName: matchName,
            category: category || '',
        });
    }

    // Tenta MUITOS jeitos de extrair length/iteração — QE varia muito por versão
    function _qeLen(obj) {
        if (!obj) return 0;
        try { if (typeof obj.numItems === 'number') return obj.numItems; } catch(e) {}
        try { if (typeof obj.length === 'number') return obj.length; } catch(e) {}
        try { if (typeof obj.getItemAt === 'function') {
            // Tenta achar o tamanho via getItemAt indo até falhar
            var n = 0;
            while (n < 500) { try { if (!obj.getItemAt(n)) break; n++; } catch(e2) { break; } }
            return n;
        }} catch(e) {}
        // Fallback: brute-force até falhar
        var i = 0;
        while (i < 500) { try { if (obj[i] === undefined || obj[i] === null) break; i++; } catch(e3) { break; } }
        return i;
    }
    function _qeAt(obj, i) {
        try { var v = obj[i]; if (v) return v; } catch(e) {}
        try { return obj.getItemAt(i); } catch(e2) {}
        return null;
    }

    // Lê propriedade tentando vários acessos (alguns QE expõem como method, outros prop)
    function _qeProp(obj, names) {
        if (!obj) return '';
        for (var i = 0; i < names.length; i++) {
            var n = names[i];
            try {
                var v = obj[n];
                if (typeof v === 'function') {
                    try { v = v.call(obj); } catch(e1) {}
                }
                if (v && typeof v === 'string') return v;
            } catch(e) {}
        }
        return '';
    }
    var NAME_KEYS = ['name', 'title', 'displayName', 'effectName', 'getName'];
    var MATCH_KEYS = ['matchName', 'matchname', 'getMatchName'];

    // Recursivo — enumera todos effects em qualquer profundidade.
    function _scanEffects(node, kind, parentCat, depth, sample) {
        if (!node || depth > 4) return 0;
        // CASO NOVO (PR 2024+ / 25.x): list é Array de strings (nomes só)
        if (typeof node === 'string') {
            // Em Premiere 2024+, name é o "matchName" — apply usa getVideoEffectByName(name)
            pushItem(kind, node, node, parentCat || '');
            if (sample && sample.length < 5) sample.push(node);
            return 1;
        }
        var added = 0;
        // CASO ANTIGO: nodes são objetos com matchName/name
        var nMatch = _qeProp(node, MATCH_KEYS);
        var nName = _qeProp(node, NAME_KEYS);
        if (nMatch) {
            pushItem(kind, nName, nMatch, parentCat || '');
            if (sample && sample.length < 5 && nName) sample.push(nName);
            return 1;
        }
        // Container — itera filhos
        var len = _qeLen(node);
        if (len === 0) return 0;
        var thisCat = nName || parentCat || '';
        for (var i = 0; i < len; i++) {
            var child = _qeAt(node, i);
            if (child === null || child === undefined) continue;
            added += _scanEffects(child, kind, thisCat, depth + 1, sample);
        }
        return added;
    }

    function iterateList(qeList, kind, label) {
        try {
            if (!qeList) { debug.push(label + '=null'); return 0; }
            var sampleNames = [];
            var added = _scanEffects(qeList, kind, '', 0, sampleNames);
            var lenTop = _qeLen(qeList);
            debug.push(label + '=' + lenTop + 'top/' + added + 'fx[' + sampleNames.join('|') + ']');
            return added;
        } catch (e) {
            errors.push(label + ':' + e.toString().slice(0, 80));
            return 0;
        }
    }

    try {
        // 1) Garante app
        if (typeof app === 'undefined' || !app) return _lwSearchJson([], 'no-app', ['app indefinido']);

        // 2) Habilita QE — retentativas
        var qeReady = false;
        // QE quase sempre tá pronto na 1ª tentativa — sleep só se falhar
        for (var tries = 0; tries < 3; tries++) {
            try { if (typeof qe === 'undefined') app.enableQE(); } catch(eEnable) {}
            if (typeof qe !== 'undefined' && qe.project) { qeReady = true; break; }
            $.sleep(30);
        }
        debug.push('qeReady=' + qeReady);
        if (!qeReady) {
            errors.push('qe.project undefined — nenhum projeto aberto?');
            return _lwSearchJson([], debug.join(','), errors);
        }

        // 3) Effects + Transitions — tenta múltiplas APIs
        // Versões do Premiere expõem essas listas de jeitos diferentes:
        // - PR antigo: qe.project.getVideoEffectList()
        // - PR 2024+: pode estar em qe.project ou qe.app ou via property
        var listCalls = [
            ['getVideoEffectList',     'video-fx', 'vfx'],
            ['getAudioEffectList',     'audio-fx', 'afx'],
            ['getVideoTransitionList', 'video-tx', 'vtx'],
            ['getAudioTransitionList', 'audio-tx', 'atx'],
        ];
        // Lista de "containers" pra tentar — qe.project, qe.app, qe
        var qeContainers = [];
        try { if (qe && qe.project) qeContainers.push({ obj: qe.project, label: 'qe.project' }); } catch(eQp) {}
        try { if (qe && qe.app) qeContainers.push({ obj: qe.app, label: 'qe.app' }); } catch(eQa) {}
        try { if (qe) qeContainers.push({ obj: qe, label: 'qe' }); } catch(eQ) {}
        // ExtendScript é ES3 — não tem .map(). Loop manual.
        var _ctrLabels = '';
        for (var _lc = 0; _lc < qeContainers.length; _lc++) {
            if (_lc > 0) _ctrLabels += '+';
            _ctrLabels += qeContainers[_lc].label;
        }
        debug.push('containers=' + _ctrLabels);

        for (var l = 0; l < listCalls.length; l++) {
            var fnName = listCalls[l][0], kind = listCalls[l][1], lbl = listCalls[l][2];
            var lst = null, found = false;
            for (var ci = 0; ci < qeContainers.length && !lst; ci++) {
                var container = qeContainers[ci].obj;
                // Tenta como method
                try {
                    if (typeof container[fnName] === 'function') {
                        lst = container[fnName]();
                        if (lst) { found = true; debug.push(lbl + '@' + qeContainers[ci].label + 'fn'); }
                    }
                } catch(eL1) {}
                // Tenta como property
                if (!lst) {
                    try {
                        var propName = fnName.replace(/^get/, '').charAt(0).toLowerCase() + fnName.replace(/^get/, '').slice(1);
                        if (container[propName]) {
                            lst = container[propName];
                            found = true;
                            debug.push(lbl + '@' + qeContainers[ci].label + 'prop:' + propName);
                        }
                    } catch(eL2) {}
                }
            }
            if (!found) { errors.push(fnName + ':not-found'); continue; }
            iterateList(lst, kind, lbl);
        }

        // 4) Effect Presets — scan rootItem por bins .prfpset
        try {
            if (app.project && app.project.rootItem) {
                _lwScanPresets(app.project.rootItem, items, seen, 0);
            }
        } catch(ePr) { errors.push('presets:' + ePr.toString().slice(0, 60)); }
    } catch(eAll) {
        errors.push('outer:' + eAll.toString().slice(0, 80));
    }

    // Escreve resultado num arquivo temp pra debug — usuário pode mandar o conteúdo
    try {
        var fdbg = new File(Folder.temp.fsName + '/lion-search-debug.txt');
        if (fdbg.open('w')) {
            fdbg.encoding = 'UTF-8';
            fdbg.writeln('items=' + items.length);
            fdbg.writeln('debug=' + debug.join(','));
            fdbg.writeln('errors=' + (errors.length ? errors.join(' | ') : '(none)'));
            fdbg.writeln('');
            fdbg.writeln('--- typeof checks ---');
            fdbg.writeln('typeof app=' + (typeof app));
            fdbg.writeln('typeof qe=' + (typeof qe));
            try { fdbg.writeln('app.version=' + app.version); } catch(e1) { fdbg.writeln('app.version=ERR:' + e1); }
            try { fdbg.writeln('qe.project=' + (qe.project ? 'YES' : 'no')); } catch(e2) { fdbg.writeln('qe.project=ERR:' + e2); }

            // Inspeciona a estrutura do PRIMEIRO item da getVideoEffectList()
            try {
                if (qe && qe.project && typeof qe.project.getVideoEffectList === 'function') {
                    var vfxL = qe.project.getVideoEffectList();
                    fdbg.writeln('');
                    fdbg.writeln('--- vfxList top-level inspection ---');
                    fdbg.writeln('numItems=' + (vfxL.numItems || 'undefined'));
                    fdbg.writeln('length=' + (vfxL.length || 'undefined'));
                    var item0 = null;
                    try { item0 = vfxL[0]; } catch(e0) {}
                    if (!item0) { try { item0 = vfxL.getItemAt(0); } catch(e0b) {} }
                    if (item0) {
                        fdbg.writeln('item[0] type=' + (typeof item0));
                        fdbg.writeln('item[0] toString=' + (item0.toString ? item0.toString().substr(0, 100) : 'no toString'));
                        var props = '';
                        for (var pp in item0) { props += pp + ','; }
                        fdbg.writeln('item[0] enum props: ' + (props.length > 800 ? props.substr(0, 800) + '...' : props));
                        // Tenta acessar várias props comuns
                        fdbg.writeln('item[0].name=' + _safeStr(item0, 'name'));
                        fdbg.writeln('item[0].matchName=' + _safeStr(item0, 'matchName'));
                        fdbg.writeln('item[0].title=' + _safeStr(item0, 'title'));
                        fdbg.writeln('item[0].displayName=' + _safeStr(item0, 'displayName'));
                        fdbg.writeln('item[0].numItems=' + _safeStr(item0, 'numItems'));
                        fdbg.writeln('item[0].length=' + _safeStr(item0, 'length'));
                    } else {
                        fdbg.writeln('item[0] = null/undefined');
                    }
                }
            } catch(eIns) { fdbg.writeln('inspection err: ' + eIns); }

            fdbg.writeln('');
            fdbg.writeln('--- first 10 items achados ---');
            for (var ii = 0; ii < Math.min(10, items.length); ii++) {
                fdbg.writeln(items[ii].kind + ' / cat=' + items[ii].category + ' / ' + items[ii].name + ' (' + items[ii].matchName + ')');
            }
            fdbg.close();
        }
    } catch(eFw) {}

    return _lwSearchJson(items, debug.join(','), errors);
}

// Builda JSON estruturado: { items, debug, error } — preserva TODAS as flags
function _lwSearchJson(items, debug, errors) {
    var out = '{"items":[';
    for (var i = 0; i < items.length; i++) {
        if (i > 0) out += ',';
        var it = items[i];
        out += '{"kind":"' + _lwEscapeJson(it.kind) + '"';
        out += ',"name":"' + _lwEscapeJson(it.name) + '"';
        out += ',"matchName":"' + _lwEscapeJson(it.matchName) + '"';
        out += ',"category":"' + _lwEscapeJson(it.category) + '"';
        if (it.audioOnly) out += ',"audioOnly":true';
        if (it.videoOnly) out += ',"videoOnly":true';
        if (it.isContainer) out += ',"isContainer":true';
        if (it.projName) out += ',"projName":"' + _lwEscapeJson(it.projName) + '"';
        out += '}';
    }
    out += '],"debug":"' + _lwEscapeJson(debug || '') + '"';
    out += ',"error":"' + _lwEscapeJson((errors && errors.length) ? errors.join(' | ') : '') + '"}';
    return out;
}

function _lwScanPresets(item, items, seen, depth) {
    depth = depth || 0;
    if (depth > 10) return;
    try {
        if (item.type === 1 && item.children && item.children.numItems > 0) {
            for (var i = 0; i < item.children.numItems; i++) {
                _lwScanPresets(item.children[i], items, seen, depth + 1);
            }
        }
    } catch(e) {}
}

// Helper: itera todas as QE tracks e retorna a lista de QE clips selecionados.
// Tenta 3 estratégias:
//   1) seq.getSelection() — API moderna mais confiável
//   2) QE isSelected() — direto no QE clip
//   3) std isSelected() + mapear pra QE por start time (fallback)
function _lwFindSelectedQEClips(qeSeq, isVideo, seq) {
    var found = [];
    var debugInfo = [];

    // ─── Estratégia 1: seq.getSelection() (API mais nova e confiável) ───
    var stdSelected = [];
    try {
        if (seq && typeof seq.getSelection === 'function') {
            var selectedItems = seq.getSelection();
            if (selectedItems && selectedItems.length !== undefined) {
                debugInfo.push('getSelection=' + selectedItems.length);
                for (var gs = 0; gs < selectedItems.length; gs++) {
                    var si = selectedItems[gs];
                    // Filtra video vs audio
                    var isVid = false;
                    try {
                        // mediaType property: 1=video, 4=audio (varia por versão)
                        // Tenta detectar por nodeId/parent track
                        var p = si.projectItem || null;
                        if (p) {
                            try { isVid = p.hasVideo && p.hasVideo(); } catch(eHV) {}
                            if (!isVid) {
                                try {
                                    var path = p.getMediaPath && p.getMediaPath();
                                    if (path) isVid = !/\.(mp3|wav|aac|m4a|flac|ogg|aif|aiff)$/i.test(path);
                                } catch(eMP) {}
                            }
                        }
                    } catch(eMT) {}
                    // Se não conseguiu detectar, assume baseado em isVideo solicitado
                    if (isVid === undefined) isVid = isVideo;
                    if ((isVideo && isVid) || (!isVideo && !isVid)) {
                        stdSelected.push(si);
                    }
                }
            }
        }
    } catch(eGS) { debugInfo.push('getSelection-err:' + eGS.toString().slice(0,40)); }

    // Se não temos getSelection ou retornou vazio, fallback pra iterar tracks
    if (stdSelected.length === 0 && seq) {
        var trackList0 = isVideo ? seq.videoTracks : seq.audioTracks;
        for (var ti0 = 0; ti0 < trackList0.numTracks; ti0++) {
            var tk0 = trackList0[ti0];
            for (var ci0 = 0; ci0 < tk0.clips.numItems; ci0++) {
                try {
                    if (tk0.clips[ci0].isSelected()) stdSelected.push(tk0.clips[ci0]);
                } catch(eIs0) {}
            }
        }
        debugInfo.push('stdIterate=' + stdSelected.length);
    }

    // ─── Estratégia 2: QE isSelected() ───
    var numTracks = isVideo ? qeSeq.numVideoTracks : qeSeq.numAudioTracks;
    for (var i = 0; i < numTracks; i++) {
        var t = null;
        try { t = isVideo ? qeSeq.getVideoTrackAt(i) : qeSeq.getAudioTrackAt(i); } catch(e0) {}
        if (!t) continue;
        var n = 0;
        try { n = t.numItems; } catch(eN) { n = 0; }
        for (var j = 0; j < n; j++) {
            var item = null;
            try { item = t.getItemAt(j); } catch(eI) {}
            if (!item) continue;
            try { if (item.type === 1) continue; } catch(eT) {}
            try { if (item.isSelected && item.isSelected()) { found.push(item); } } catch(eS) {}
        }
    }
    if (found.length > 0) {
        debugInfo.push('qeIsSelected=' + found.length);
        _lwDbgPreset && _lwDbgPreset('  _lwFindSelectedQEClips(' + (isVideo?'video':'audio') + '): ' + debugInfo.join(','));
        return found;
    }

    // ─── Estratégia 3: mapeia stdSelected → QE clips por start time ───
    if (stdSelected.length > 0) {
        for (var ssi = 0; ssi < stdSelected.length; ssi++) {
            var stdClip = stdSelected[ssi];
            // Tenta achar a track index do clip selecionado
            var trackIdx = -1;
            var trackListM = isVideo ? seq.videoTracks : seq.audioTracks;
            for (var tim = 0; tim < trackListM.numTracks && trackIdx < 0; tim++) {
                var tkm = trackListM[tim];
                for (var cim = 0; cim < tkm.clips.numItems; cim++) {
                    if (tkm.clips[cim] === stdClip || (tkm.clips[cim].nodeId && tkm.clips[cim].nodeId === stdClip.nodeId)) {
                        trackIdx = tim; break;
                    }
                }
            }
            if (trackIdx < 0) continue;
            var qeT = null;
            try { qeT = isVideo ? qeSeq.getVideoTrackAt(trackIdx) : qeSeq.getAudioTrackAt(trackIdx); } catch(eGT2) {}
            if (!qeT) continue;
            var startSec = 0, startTicks = '';
            try { startSec = parseFloat(stdClip.start.seconds); } catch(eSS) {}
            try { startTicks = String(stdClip.start.ticks); } catch(eST) {}
            var matched = null;
            var qN2 = 0;
            try { qN2 = qeT.numItems; } catch(eQN2) {}
            for (var qj2 = 0; qj2 < qN2; qj2++) {
                var qit2 = null;
                try { qit2 = qeT.getItemAt(qj2); } catch(eGI2) {}
                if (!qit2) continue;
                try { if (qit2.type === 1) continue; } catch(eQT2) {}
                try {
                    if (String(qit2.start.ticks) === startTicks) { matched = qit2; break; }
                } catch(eQTk2) {}
                try {
                    if (Math.abs(qit2.start.seconds - startSec) < 0.5) { matched = qit2; break; }
                } catch(eQS2) {}
            }
            if (matched) found.push(matched);
        }
        debugInfo.push('mappedQE=' + found.length);
    }

    if (typeof _lwDbgPreset === 'function') {
        _lwDbgPreset('  _lwFindSelectedQEClips(' + (isVideo?'video':'audio') + '): ' + debugInfo.join(',') + ' → ' + found.length);
    }
    return found;
}

// Aplica um effect/transition num clip selecionado.
// kind = "video-fx" | "audio-fx" | "video-tx" | "audio-tx"
// matchName = nome do effect (em PR 2024+, é o display name retornado por getVideoEffectList)
// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — APPLY PRESET WITH FULL DATA (effects + props + keyframes)
// Inspired by Excalibur's approach: receives parsed JSON from plugin client,
// applies each effect via addVideoEffect + setValue/addKey.
// ═══════════════════════════════════════════════════════════════════
function lwApplyPresetData(presetName, dataJsonStr) {
    try {
        _lwDbgPreset('═══ ApplyData: ' + presetName + ' ═══');
        if (!dataJsonStr) return 'ERR:no_data';
        var data;
        try { data = JSON.parse(dataJsonStr); }
        catch(eJ) { _lwDbgPreset('  JSON parse err: ' + eJ); return 'ERR:json_parse:' + eJ; }
        if (!data || !data.length) { _lwDbgPreset('  empty data'); return 'ERR:empty_data'; }
        _lwDbgPreset('  data has ' + data.length + ' effects');

        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        var seq = app.project.activeSequence;
        try { app.enableQE(); } catch(eQE) {}
        if (typeof qe === 'undefined' || !qe.project) return 'ERR:qe_not_available';
        var qeSeq = qe.project.getActiveSequence();
        if (!qeSeq) return 'ERR:no_qe_sequence';

        // Lista de NOMES de efeitos registrados (1x por apply) — usada pra montar
        // candidatos de lookup sem cair em colisao de nome (ex: "Gaussian Blur"
        // do Impact roubou o nome do nativo; o nativo legacy fica em "(Legacy)")
        var _fxNameList = null;
        try { _fxNameList = qe.project.getVideoEffectList(); } catch(eFl) {}
        var _fxListHas = function(name) {
            if (!_fxNameList || !name) return false;
            for (var fli = 0; fli < _fxNameList.length; fli++) {
                if (String(_fxNameList[fli]) === name) return true;
            }
            return false;
        };

        // Acha selection — com recovery via snapshot
        var sel = _lwFindSelectedQEClips(qeSeq, true, seq);
        var selKind = 'video';
        if (sel.length === 0) { sel = _lwFindSelectedQEClips(qeSeq, false, seq); selKind = 'audio'; }
        var stdSel = [];
        if (sel.length === 0) {
            var rec = _lwRecoverFromSnapshot(seq, qeSeq);
            if (rec.qeClips.length > 0) {
                sel = rec.qeClips;
                stdSel = rec.stdClips;
                selKind = rec.kind;
                _lwDbgPreset('  recovery: ' + sel.length + ' clips');
            } else {
                return 'NO_CLIP_SELECTED';
            }
        }
        if (stdSel.length === 0) {
            var trackList = (selKind === 'video') ? seq.videoTracks : seq.audioTracks;
            for (var ti = 0; ti < trackList.numTracks; ti++) {
                var tk = trackList[ti];
                for (var ci = 0; ci < tk.clips.numItems; ci++) {
                    if (tk.clips[ci].isSelected()) stdSel.push(tk.clips[ci]);
                }
            }
        }
        _lwDbgPreset('  ' + sel.length + ' QE + ' + stdSel.length + ' std (kind=' + selKind + ')');

        // ── HELPERS ──
        // Built-in components que TODO clip de vídeo tem (não chamar addVideoEffect — já existem)
        var builtinMatches = {
            'AE.ADBE Motion': 'Motion',
            'AE.ADBE Opacity': 'Opacity',
            'AE.ADBE Timecode': 'Time Remapping',
            'AE.ADBE MarkerParam': 'Marker'
        };

        // Acha component existente no clip por matchName ou displayName
        function _findComponent(clip, matchName, displayName) {
            try {
                var comps = clip.components;
                if (!comps) return null;
                for (var i = 0; i < comps.numItems; i++) {
                    var c = comps[i];
                    var mn = ''; var dn = '';
                    try { mn = String(c.matchName || ''); } catch(e1) {}
                    try { dn = String(c.displayName || ''); } catch(e2) {}
                    if (matchName && mn === matchName) return c;
                    if (displayName && dn === displayName) return c;
                }
            } catch(e) {}
            return null;
        }

        // Acha property no component por displayName
        function _findProp(component, propName) {
            try {
                var props = component.properties;
                if (!props) return null;
                for (var i = 0; i < props.numItems; i++) {
                    var p = props[i];
                    var pn = '';
                    try { pn = String(p.displayName || p.name || ''); } catch(e) {}
                    if (pn === propName) return p;
                }
            } catch(e) {}
            return null;
        }

        // Parse keyframe value baseado em ParameterControlType
        // type 2/3 = float, 4 = bool, 6 = point (x:y), 7 = enum
        function _parseVal(rawVal, ptype) {
            if (ptype === 4) return (String(rawVal).toLowerCase() === 'true');
            if (ptype === 6) {
                // Format "x:y" → array
                var parts = String(rawVal).split(':');
                return [parseFloat(parts[0]), parseFloat(parts[1] || 0)];
            }
            if (ptype === 7) return parseInt(rawVal, 10);
            return parseFloat(rawVal);
        }

        // Ticks → seconds (Premiere usa 254016000000 ticks/sec)
        var TICKS_PER_SEC = 254016000000;
        function _ticksToSec(ticks) {
            return parseFloat(ticks) / TICKS_PER_SEC;
        }

        // ── APLICA cada efeito ──
        var totalApplied = 0;
        var errors = [];
        for (var fxi = 0; fxi < data.length; fxi++) {
            var effect = data[fxi];
            if (!effect) continue;
            var mn = effect.matchName || '';
            var nm = effect.name || '';
            var props = effect.props || [];
            _lwDbgPreset('  fx[' + fxi + ']: ' + mn + ' (' + nm + ') props=' + props.length);

            var isBuiltin = !!builtinMatches[mn];
            var displayName = builtinMatches[mn] || nm;

            // Aplica em CADA clip selecionado
            for (var ci2 = 0; ci2 < stdSel.length; ci2++) {
                var clip = stdSel[ci2];
                var qeClip = sel[ci2];
                var targetComp = null;
                var _equivalentOnly = false; // true = efeito equivalente (legacy): props SO por nome

                if (isBuiltin) {
                    // Built-in: usa component existente (NÃO addVideoEffect)
                    targetComp = _findComponent(clip, mn, displayName);
                    if (!targetComp) {
                        // Tenta variations
                        var altNames = ['Motion','Opacity','Time Remapping','Velocidade','Movimento','Opacidade'];
                        for (var an = 0; an < altNames.length && !targetComp; an++) {
                            targetComp = _findComponent(clip, '', altNames[an]);
                        }
                    }
                    if (!targetComp) { _lwDbgPreset('    NO_BUILTIN_COMP for ' + mn); continue; }
                    _lwDbgPreset('    builtin component found: ' + (function(){try{return targetComp.displayName;}catch(e){return '?';}})());
                } else {
                    // Custom effect: addVideoEffect com VERIFICACAO do matchName.
                    // BUGS antigos: (1) lookup por nome trazia efeito de TERCEIRO com o mesmo
                    // displayName (ex: Impact Blur "Gaussian Blur" no lugar do nativo) e o preset
                    // escrevia parametros aleatorios nele; (2) o match por displayName pegava um
                    // component PRE-EXISTENTE do clip (ex: Transform antigo) e sobrescrevia ele.
                    // Agora: tenta identifiers em ordem (matchName primeiro — o QE resolve
                    // matchName no addVideoEffect), acha o component NOVO por diff da lista,
                    // e SO usa se o matchName bater. Errado = nao escreve nada nele.
                    var _compMatchNames = function() {
                        var arr = [];
                        try {
                            for (var s = 0; s < clip.components.numItems; s++) {
                                var t = '';
                                try { t = String(clip.components[s].matchName || ''); } catch(eT) {}
                                arr.push(t);
                            }
                        } catch(eS) {}
                        return arr;
                    };
                    var humanNm = String(mn).replace(/^AE\.ADBE\s*/, '').replace(/^AE\./, '').replace(/^ADBE\s*/, '');
                    // Candidatos em ordem esperta:
                    // - Se o matchName termina em digito (versao moderna, ex "Gaussian Blur 2")
                    //   E existe "<nome> (Legacy)" na lista, tenta o LEGACY antes do nome puro —
                    //   evita adicionar o efeito de terceiro que roubou o nome (Impact "Gaussian Blur").
                    //   O legacy e' aceito por EQUIVALENCIA (mesma familia, mesmos parametros).
                    var tries = [mn];
                    var legacyNm = nm ? (nm + ' (Legacy)') : '';
                    var mnModern = /\d\s*$/.test(String(mn));
                    if (nm) {
                        if (mnModern && legacyNm && _fxListHas(legacyNm)) tries.push(legacyNm);
                        tries.push(nm);
                        if (!mnModern && legacyNm && _fxListHas(legacyNm)) tries.push(legacyNm);
                    }
                    if (humanNm && humanNm !== nm) {
                        tries.push(humanNm);
                        if (_fxListHas(humanNm + ' (Legacy)')) tries.push(humanNm + ' (Legacy)');
                    }
                    // Base da familia: tira "(Legacy)" e digitos finais → "AE.ADBE Gaussian Blur 2"
                    // e "AE.ADBE Gaussian Blur" tem a MESMA base (parametros identicos)
                    var _fxBase = function(s) {
                        return String(s || '').replace(/\s*\(legacy\)\s*$/i, '').replace(/\s*\d+$/, '').replace(/\s+$/, '');
                    };
                    targetComp = null;
                    for (var lk = 0; lk < tries.length && !targetComp; lk++) {
                        var ident = tries[lk];
                        if (!ident) continue;
                        var fxObj = null;
                        try { fxObj = qe.project.getVideoEffectByName(ident); } catch(eLk) {}
                        if (!fxObj) continue;
                        var beforeList = _compMatchNames();
                        try {
                            if (selKind === 'audio') qeClip.addAudioEffect(fxObj);
                            else qeClip.addVideoEffect(fxObj);
                        } catch(eAdd) { _lwDbgPreset('    addEffect("' + ident + '") ERR: ' + eAdd); continue; }
                        // Espera o Premiere registrar o novo component
                        var afterN = beforeList.length;
                        for (var poll = 0; poll < 10; poll++) {
                            try { afterN = clip.components.numItems; } catch(ePN) {}
                            if (afterN > beforeList.length) break;
                            $.sleep(30);
                        }
                        if (afterN <= beforeList.length) { _lwDbgPreset('    "' + ident + '": add nao registrou component (identifier invalido?)'); continue; }
                        // Acha o component NOVO: primeiro indice onde a lista diverge da antiga
                        var newComp = null;
                        try {
                            for (var ck = 0; ck < clip.components.numItems; ck++) {
                                var cmn = '';
                                try { cmn = String(clip.components[ck].matchName || ''); } catch(eMn) {}
                                if (ck >= beforeList.length || cmn !== beforeList[ck]) { newComp = clip.components[ck]; break; }
                            }
                        } catch(eNw) {}
                        if (!newComp) { try { newComp = clip.components[clip.components.numItems - 1]; } catch(eLast) {} }
                        var newMn = '';
                        try { newMn = String(newComp.matchName || ''); } catch(eNm2) {}
                        if (newMn === mn) {
                            targetComp = newComp;
                            _equivalentOnly = false;
                            _lwDbgPreset('    ✓ efeito correto via "' + ident + '" → ' + newMn);
                        } else if (_fxBase(newMn) && _fxBase(newMn) === _fxBase(mn)) {
                            // Mesma FAMILIA (ex: Transform obsoleto vs Geometry2, Gaussian legacy vs 2):
                            // parametros identicos → aceita, mas aplica props SO por nome (sem indice)
                            targetComp = newComp;
                            _equivalentOnly = true;
                            _lwDbgPreset('    ✓ efeito EQUIVALENTE via "' + ident + '" → ' + newMn + ' (pedido: ' + mn + ')');
                        } else {
                            // Veio efeito ERRADO (colisao de nome). NAO escreve params nele —
                            // era isso que gerava "blur fantasma" com valores aleatorios.
                            _lwDbgPreset('    ✗ colisao: pedi ' + mn + ', veio ' + newMn + ' via "' + ident + '" — ignorando esse component');
                        }
                    }
                    if (!targetComp) { _lwDbgPreset('    NO_FX: nenhum identifier resultou em ' + mn); errors.push('no_fx:' + mn); continue; }
                    _lwDbgPreset('    target component: matchName=' + (function(){try{return targetComp.matchName;}catch(e){return '?';}})() + ' displayName=' + (function(){try{return targetComp.displayName;}catch(e){return '?';}})());

                    // Dump das properties do target component pra ver nomes reais
                    // CRITICAL: NÃO usar 'var props' aqui — ES3 hoisting sobrescreveria
                    // a variável effect.props da função outer e quebraria o loop abaixo!
                    try {
                        var compProps = targetComp.properties;
                        if (compProps && compProps.numItems) {
                            _lwDbgPreset('    component has ' + compProps.numItems + ' properties:');
                            for (var pi3 = 0; pi3 < Math.min(15, compProps.numItems); pi3++) {
                                var pName = '';
                                try { pName = String(compProps[pi3].displayName || compProps[pi3].name || '?'); } catch(eN) {}
                                _lwDbgPreset('      prop[' + pi3 + ']: ' + pName);
                            }
                        } else {
                            _lwDbgPreset('    component has NO properties or compProps.numItems=0');
                        }
                    } catch(ePropsDump) { _lwDbgPreset('    properties dump err: ' + ePropsDump); }
                }

                // ── SETA cada propriedade ──
                // Tempos no preset sao ABSOLUTOS (do clip-fonte do autor). AnchorIn/OutPoint
                // delimitam esse clip-fonte. O TYPE do preset define como reancorar aqui:
                //   0 = Scale to Clip Length (estica os tempos pro tamanho DESTE clip)
                //   1 = Anchor to In Point   (offset a partir do inicio do clip)
                //   2 = Anchor to Out Point  (offset a partir do FIM do clip)
                var anchorInTicks = 0;
                try {
                    if (effect._anchor && effect._anchor.AnchorInPoint) {
                        anchorInTicks = parseFloat(effect._anchor.AnchorInPoint);
                    }
                } catch(eAnc) {}
                var anchorOutTicks = anchorInTicks;
                try {
                    if (effect._anchor && effect._anchor.AnchorOutPoint) {
                        anchorOutTicks = parseFloat(effect._anchor.AnchorOutPoint) || anchorInTicks;
                    }
                } catch(eAo) {}
                var presetType = 1;
                try {
                    if (effect._anchor && effect._anchor.Type != null && effect._anchor.Type !== '') {
                        presetType = parseInt(effect._anchor.Type, 10);
                        if (isNaN(presetType)) presetType = 1;
                    }
                } catch(eTp) {}
                var clipInPointTicks = 0, clipOutTicks = 0;
                try { if (clip.inPoint && clip.inPoint.ticks != null) clipInPointTicks = parseFloat(clip.inPoint.ticks); } catch(eIp0) {}
                try { if (clip.outPoint && clip.outPoint.ticks != null) clipOutTicks = parseFloat(clip.outPoint.ticks); } catch(eOp0) {}
                if (clipOutTicks <= clipInPointTicks) clipOutTicks = clipInPointTicks;
                // Dimensoes da sequence (pra converter pontos normalizados → pixel quando preciso)
                var seqW = 1920, seqH = 1080;
                try { seqW = parseInt(seq.frameSizeHorizontal, 10) || 1920; } catch(eSw0) {}
                try { seqH = parseInt(seq.frameSizeVertical, 10) || 1080; } catch(eSh0) {}
                // Converte tempo ABSOLUTO do preset → tempo no clip atual conforme o Type
                var _targetTicks = function(rawKfTicks) {
                    if (presetType === 0 && anchorOutTicks > anchorInTicks && clipOutTicks > clipInPointTicks) {
                        var frac = (rawKfTicks - anchorInTicks) / (anchorOutTicks - anchorInTicks);
                        return clipInPointTicks + frac * (clipOutTicks - clipInPointTicks);
                    }
                    if (presetType === 2) {
                        return clipOutTicks - (anchorOutTicks - rawKfTicks);
                    }
                    var off = rawKfTicks - anchorInTicks;
                    if (off < 0) off = 0;
                    return clipInPointTicks + off;
                };
                _lwDbgPreset('    anchor in=' + _ticksToSec(anchorInTicks).toFixed(3) + 's out=' + _ticksToSec(anchorOutTicks).toFixed(3) + 's type=' + presetType + ' | clip in=' + _ticksToSec(clipInPointTicks).toFixed(3) + 's out=' + _ticksToSec(clipOutTicks).toFixed(3) + 's | seq=' + seqW + 'x' + seqH);

                // ORDEM: props ESTATICAS primeiro (Uniform Scale precisa estar setado
                // ANTES dos keyframes de Scale — senao o zoom aplica com o uniform
                // errado e estica torto), depois as animadas.
                var propOrder = [];
                for (var oi1 = 0; oi1 < props.length; oi1++) {
                    if (props[oi1] && !(props[oi1].IsTimeVarying && props[oi1].Keyframes)) propOrder.push(oi1);
                }
                for (var oi2 = 0; oi2 < props.length; oi2++) {
                    if (props[oi2] && props[oi2].IsTimeVarying && props[oi2].Keyframes) propOrder.push(oi2);
                }
                for (var oi = 0; oi < propOrder.length; oi++) {
                    var pi = propOrder[oi];
                    var prop = props[pi];
                    if (!prop) continue;
                    // Excalibur às vezes omite Name (ex: Uniform Scale). Fallback: usa o
                    // mesmo índice no targetComp.properties (a ordem do JSON costuma
                    // bater com a ordem das props do effect).
                    var ppObj = null;
                    var propLabel = prop.Name || '(unnamed@' + pi + ')';
                    if (prop.Name) {
                        ppObj = _findProp(targetComp, prop.Name);
                    }
                    if (!ppObj && !_equivalentOnly) {
                        // Fallback por índice — NUNCA em efeito equivalente (a ordem das props
                        // pode divergir entre versoes; por nome e' seguro, por indice nao)
                        try {
                            if (targetComp.properties && pi < targetComp.properties.numItems) {
                                ppObj = targetComp.properties[pi];
                                try { propLabel = String(ppObj.displayName || propLabel); } catch(eDn) {}
                                _lwDbgPreset('      using index fallback for ' + propLabel);
                            }
                        } catch(eIdx) {}
                    }
                    if (!ppObj) { _lwDbgPreset('      NO_PROP: ' + propLabel); continue; }

                    try {
                        if (prop.IsTimeVarying && prop.Keyframes) {
                            // KEYFRAMES animados — replica fiel do preset:
                            // 1. Keys originais SNAPADOS no grid de frames da sequence
                            //    (keys fora do grid deixavam shakes com amplitude errada)
                            // 2. Interpolacao REAL via setInterpolationTypeAtKey (linear/hold/bezier)
                            // 3. Simula keyframes por frame (igual easing system) SO quando o ease
                            //    do preset e' forte (influence > 35%) e o trecho e' longo — com a
                            //    curva reconstruida do speed/influence REAL do preset.
                            //    (antes: 5 samples de curva generica em TODO segmento = distorcia tudo)
                            try { ppObj.setTimeVarying(true); } catch(eTv) {}
                            var seqFrameTicks = 0;
                            try { seqFrameTicks = parseFloat(seq.timebase) || 0; } catch(eTb) {}
                            // PONTOS (Position/Anchor): preset guarda NORMALIZADO (0..1), mas a
                            // property pode ser pixel-space (ex: Transform). Auto-detect pelo valor
                            // atual (igual o motion tracker) — sem isso o shake virava sub-pixel
                            // e o clip ia parar no canto da tela.
                            var pxScale = null;
                            if (prop.ParameterControlType === 6) {
                                try {
                                    var curV = ppObj.getValue();
                                    if (curV && curV.length >= 2 && (Math.abs(curV[0]) > 2 || Math.abs(curV[1]) > 2)) pxScale = [seqW, seqH];
                                } catch(ePd) {}
                                _lwDbgPreset('      point-space: ' + (pxScale ? 'PIXEL (converte ×' + seqW + 'x' + seqH + ')' : 'normalizado'));
                            }

                            // PASS 1: parse (t, valor, interp, ease in/out) — formato do keyframe:
                            // [ticks, valor, interpMode, ?, inSpeed, inInfluence, outSpeed, outInfluence]
                            var parsedKfs = [];
                            var kfStrs = String(prop.Keyframes).split(';');
                            for (var ki = 0; ki < kfStrs.length; ki++) {
                                var kfS = kfStrs[ki];
                                if (!kfS) continue;
                                var fields = kfS.split(',');
                                if (fields.length < 2) continue;
                                var rawTicks = parseFloat(fields[0]);
                                // Reancora conforme o TYPE do preset (scale / anchor in / anchor out)
                                var absTicks = _targetTicks(rawTicks);
                                if (absTicks < clipInPointTicks) absTicks = clipInPointTicks;
                                // Snap no grid de frames da sequence (relativo ao in-point)
                                if (seqFrameTicks > 0) absTicks = clipInPointTicks + Math.round((absTicks - clipInPointTicks) / seqFrameTicks) * seqFrameTicks;
                                var val = _parseVal(fields[1], prop.ParameterControlType);
                                if (pxScale && val instanceof Array && val.length >= 2) val = [val[0] * pxScale[0], val[1] * pxScale[1]];
                                parsedKfs.push({
                                    ticks: absTicks, val: val,
                                    interp: fields.length > 2 ? (parseInt(fields[2], 10) || 0) : 0,
                                    inSpd:  fields.length > 4 ? (parseFloat(fields[4]) || 0) : 0,
                                    inInf:  fields.length > 5 ? (parseFloat(fields[5]) || 0) : 0,
                                    outSpd: fields.length > 6 ? (parseFloat(fields[6]) || 0) : 0,
                                    outInf: fields.length > 7 ? (parseFloat(fields[7]) || 0) : 0
                                });
                            }
                            parsedKfs.sort(function(a, b) { return a.ticks - b.ticks; });
                            // Dedupe: o snap pode juntar 2 keys no mesmo frame — mantem o ultimo
                            var dedupKfs = [];
                            for (var dk = 0; dk < parsedKfs.length; dk++) {
                                if (dedupKfs.length && Math.abs(dedupKfs[dedupKfs.length - 1].ticks - parsedKfs[dk].ticks) < 1) {
                                    dedupKfs[dedupKfs.length - 1] = parsedKfs[dk];
                                } else {
                                    dedupKfs.push(parsedKfs[dk]);
                                }
                            }
                            parsedKfs = dedupKfs;

                            // PASS 2: adiciona keys originais + interpolacao nativa
                            var canSetInterp = false;
                            try { canSetInterp = (typeof ppObj.setInterpolationTypeAtKey === 'function'); } catch(eCk) {}
                            var addedCount = 0;
                            for (var pi2 = 0; pi2 < parsedKfs.length; pi2++) {
                                var kf = parsedKfs[pi2];
                                try {
                                    var kfTime = new Time(); kfTime.ticks = String(Math.round(kf.ticks));
                                    ppObj.addKey(kfTime);
                                    var setTime = new Time(); setTime.ticks = String(Math.round(kf.ticks));
                                    ppObj.setValueAtKey(setTime, kf.val);
                                    if (canSetInterp) {
                                        // kfInterpMode: 0=Linear, 4=Hold, 5=Bezier
                                        var itp = (kf.interp === 4) ? 4 : ((kf.interp === 5) ? 5 : 0);
                                        try {
                                            var itT = new Time(); itT.ticks = String(Math.round(kf.ticks));
                                            ppObj.setInterpolationTypeAtKey(itT, itp, true);
                                        } catch(eIt) {}
                                    }
                                    addedCount++;
                                    _lwDbgPreset('        kf@' + _ticksToSec(kf.ticks).toFixed(3) + 's = ' + kf.val + ' (interp=' + kf.interp + ')');
                                } catch(eKf) { _lwDbgPreset('      kf err @' + _ticksToSec(kf.ticks).toFixed(3) + 's: ' + eKf); }
                            }

                            // PASS 3: ease FORTE → simula a curva com 1 key por frame,
                            // reconstruida do speed/influence do preset (estilo AE):
                            // influence = fracao do tempo do trecho; speed = valor/segundo.
                            var samplesAdded = 0;
                            if (seqFrameTicks > 0) {
                                for (var pp1 = 0; pp1 < parsedKfs.length - 1; pp1++) {
                                    var kA = parsedKfs[pp1];
                                    var kB = parsedKfs[pp1 + 1];
                                    if (kA.interp === 4) continue; // hold: sem interpolacao no trecho
                                    var durTicks = kB.ticks - kA.ticks;
                                    if (durTicks <= 0) continue;
                                    var segFrames = durTicks / seqFrameTicks;
                                    if (segFrames < 3) continue; // trecho curto (shake): linear/bezier nativo e' exato
                                    var maxInf = Math.max(kA.outInf || 0, kB.inInf || 0);
                                    if (maxInf <= 0.35) continue; // ease fraco: bezier nativo resolve
                                    var isNum = (typeof kA.val === 'number' && typeof kB.val === 'number');
                                    var isVec = (kA.val instanceof Array && kB.val instanceof Array);
                                    if (!isNum && !isVec) continue;
                                    var dv;
                                    if (isNum) { dv = kB.val - kA.val; }
                                    else {
                                        dv = 0;
                                        for (var vd = 0; vd < kA.val.length; vd++) {
                                            var dd = (kB.val[vd] || 0) - (kA.val[vd] || 0);
                                            dv += dd * dd;
                                        }
                                        dv = Math.sqrt(dv);
                                    }
                                    if (Math.abs(dv) < 0.000001) continue;
                                    var durSec = durTicks / TICKS_PER_SEC;
                                    var iA = Math.min(1, Math.max(0.01, kA.outInf || 0.1667));
                                    var iB = Math.min(1, Math.max(0.01, kB.inInf || 0.1667));
                                    var cp1x = iA, cp2x = 1 - iB;
                                    var cp1y = ((kA.outSpd || 0) * iA * durSec) / dv;
                                    var cp2y = 1 - (((kB.inSpd || 0) * iB * durSec) / dv);
                                    if (!isFinite(cp1y)) cp1y = cp1x;
                                    if (!isFinite(cp2y)) cp2y = cp2x;
                                    cp1y = Math.max(-2, Math.min(3, cp1y));
                                    cp2y = Math.max(-2, Math.min(3, cp2y));
                                    // 1 sample por frame (se o trecho for muito longo, pula frames — cap ~90)
                                    var stepFrames = Math.max(1, Math.ceil(segFrames / 90));
                                    for (var sf = stepFrames; sf < segFrames - 0.5; sf += stepFrames) {
                                        var frac = sf / segFrames;
                                        var eased = evalBezierCurve(frac, cp1x, cp1y, cp2x, cp2y);
                                        var midTicks = kA.ticks + Math.round(sf) * seqFrameTicks;
                                        var midVal;
                                        if (isNum) { midVal = kA.val + (kB.val - kA.val) * eased; }
                                        else {
                                            midVal = [];
                                            for (var d2 = 0; d2 < kA.val.length; d2++) {
                                                midVal.push(kA.val[d2] + ((kB.val[d2] || 0) - kA.val[d2]) * eased);
                                            }
                                        }
                                        try {
                                            var mt = new Time(); mt.ticks = String(Math.round(midTicks));
                                            ppObj.addKey(mt);
                                            var st = new Time(); st.ticks = String(Math.round(midTicks));
                                            ppObj.setValueAtKey(st, midVal);
                                            if (canSetInterp) {
                                                try {
                                                    var st2 = new Time(); st2.ticks = String(Math.round(midTicks));
                                                    ppObj.setInterpolationTypeAtKey(st2, 0, true); // samples = linear
                                                } catch(eIl) {}
                                            }
                                            samplesAdded++;
                                        } catch(eSm) {}
                                    }
                                }
                            }
                            _lwDbgPreset('      ✓ ' + propLabel + ' = TimeVarying (' + addedCount + ' kfs + ' + samplesAdded + ' samples ease-forte, interp ' + (canSetInterp ? 'nativa' : 'INDISPONIVEL') + ')');
                        } else if (prop.StartKeyframe) {
                            // StartKeyframe = valor estatico REAL da prop (CurrentValue e' so o
                            // ultimo valor de UI e pode estar stale).
                            var skFields = String(prop.StartKeyframe).split(',');
                            if (skFields.length >= 2) {
                                var skVal = _parseVal(skFields[1], prop.ParameterControlType);
                                // Pontos: mesma conversao normalizado→pixel dos keyframes
                                if (prop.ParameterControlType === 6 && skVal instanceof Array && skVal.length >= 2) {
                                    try {
                                        var curS = ppObj.getValue();
                                        if (curS && curS.length >= 2 && (Math.abs(curS[0]) > 2 || Math.abs(curS[1]) > 2)) {
                                            skVal = [skVal[0] * seqW, skVal[1] * seqH];
                                        }
                                    } catch(ePs) {}
                                }
                                try { ppObj.setValue(skVal, true); _lwDbgPreset('      ✓ ' + propLabel + ' = (start) ' + skFields[1] + (skVal instanceof Array ? ' → [' + skVal.join(',') + ']' : '')); }
                                catch(eSk) { _lwDbgPreset('      setVal (start) err: ' + eSk); }
                            }
                        } else if (prop.CurrentValue !== null && prop.CurrentValue !== undefined && String(prop.CurrentValue) !== '0' && String(prop.CurrentValue) !== '') {
                            // STATIC VALUE — só se não for "0" placeholder do Excalibur
                            // (Excalibur exporta "0" pra props que não foram modificadas pelo preset;
                            //  setar 0 em Scale/Opacity quebra o clip)
                            var val = _parseVal(prop.CurrentValue, prop.ParameterControlType);
                            try { ppObj.setValue(val, true); } catch(eSv) {
                                try { ppObj.setValue(val); } catch(eSv2) { _lwDbgPreset('      setVal err: ' + eSv2); }
                            }
                            _lwDbgPreset('      ✓ ' + prop.Name + ' = ' + prop.CurrentValue);
                        } else {
                            _lwDbgPreset('      ⊘ ' + prop.Name + ' = SKIP (no real data; CurrentValue=' + prop.CurrentValue + ')');
                        }
                    } catch(eOuter) { _lwDbgPreset('      prop err: ' + eOuter); }
                }
                totalApplied++;
            }
        }
        _lwDbgPreset('═══ DONE: applied ' + totalApplied + ' effects to ' + stdSel.length + ' clips ═══');
        return 'OK:applied:' + totalApplied;
    } catch(eAll) {
        _lwDbgPreset('OUTER ERR: ' + eAll);
        return 'ERR:apply_preset_data:' + eAll;
    }
}

function lwSearchApply(kind, matchName) {
    // Debug top-level: registra TODO call que chega no host (independente do kind)
    try {
        var _fApply = new File(Folder.desktop.fsName + '/lion-apply-trace.txt');
        if (_fApply.open('a')) {
            _fApply.encoding = 'UTF-8';
            _fApply.write('[' + (new Date()).toString() + '] lwSearchApply kind="' + kind + '" matchName="' + String(matchName).slice(0,200) + '"\n');
            _fApply.close();
        }
    } catch(eDbg) {}
    try {
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        var seq = app.project.activeSequence;
        try { app.enableQE(); } catch(e0) {}
        if (typeof qe === 'undefined' || !qe.project) return 'ERR:qe_not_available';
        var qeSeq = qe.project.getActiveSequence();
        if (!qeSeq) return 'ERR:no_qe_sequence';

        if (kind === 'video-fx') {
            // Acha QE clips selecionados (video) — com recovery via snapshot
            var selVideoClips = _lwFindSelectedQEClips(qeSeq, true, seq);
            if (selVideoClips.length === 0) {
                var rcV = _lwRecoverFromSnapshot(seq, qeSeq);
                if (rcV.qeClips.length > 0 && rcV.kind === 'video') selVideoClips = rcV.qeClips;
            }
            if (selVideoClips.length === 0) return 'NO_CLIP_SELECTED';
            // Pega o effect object pelo nome
            var fxObj = null;
            try { fxObj = qe.project.getVideoEffectByName(matchName); } catch(e1) {}
            if (!fxObj) return 'ERR:effect_not_found:' + matchName;
            // Aplica em todos os clips selecionados
            var appliedV = 0;
            for (var iv = 0; iv < selVideoClips.length; iv++) {
                try { selVideoClips[iv].addVideoEffect(fxObj); appliedV++; } catch(eAv) {}
            }
            if (appliedV === 0) return 'ERR:could_not_apply';
            return 'OK:applied:' + appliedV;
        } else if (kind === 'audio-fx') {
            var selAudioClips = _lwFindSelectedQEClips(qeSeq, false, seq);
            if (selAudioClips.length === 0) {
                var rcA = _lwRecoverFromSnapshot(seq, qeSeq);
                if (rcA.qeClips.length > 0 && rcA.kind === 'audio') selAudioClips = rcA.qeClips;
            }
            if (selAudioClips.length === 0) return 'NO_AUDIO_CLIP_SELECTED';
            var afxObj = null;
            try { afxObj = qe.project.getAudioEffectByName(matchName); } catch(e2) {}
            if (!afxObj) return 'ERR:audio_effect_not_found:' + matchName;
            var appliedA = 0;
            for (var ia = 0; ia < selAudioClips.length; ia++) {
                try { selAudioClips[ia].addAudioEffect(afxObj); appliedA++; } catch(eAa) {}
            }
            if (appliedA === 0) return 'ERR:could_not_apply';
            return 'OK:applied:' + appliedA;
        } else if (kind === 'video-tx' || kind === 'audio-tx') {
            return 'ERR:transitions_not_implemented_yet';
        } else if (kind === 'preset') {
            return _lwApplyPreset(matchName); // matchName aqui é o full path do .prfpset
        } else if (kind === 'audio-source') {
            return _lwInsertAudioSource(matchName); // matchName é o nodeId/name do project item
        } else {
            return 'ERR:unknown_kind:' + kind;
        }
    } catch (eAll) {
        return 'ERR:' + eAll.toString().slice(0, 200);
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — GET CONTEXT (tipo do clip selecionado, playhead, etc)
// ═══════════════════════════════════════════════════════════════════
// GLOBAL: snapshot da última selection capturada (sobrevive perda de foco do Premiere)
// Quando LION SEARCH abre, foco vai pra ela → Premiere às vezes "perde" a selection.
// Snapshot permite reaplicar o efeito mesmo se getSelection() retornar vazio.
var _lwSelectionSnapshot = { clips: [], capturedAt: 0, seqRef: null };

// Recupera clips do snapshot se selection atual está vazia (snapshot < 60s).
// Retorna { stdClips: [...], qeClips: [...], kind: 'video'|'audio'|null }
function _lwRecoverFromSnapshot(seq, qeSeq) {
    if (!_lwSelectionSnapshot || !_lwSelectionSnapshot.clips || _lwSelectionSnapshot.clips.length === 0) {
        return { stdClips: [], qeClips: [], kind: null };
    }
    // Snapshot precisa ser RECENTE (60s) e MESMA sequência
    var ageSec = ((new Date()).getTime() - _lwSelectionSnapshot.capturedAt) / 1000;
    if (ageSec > 60) return { stdClips: [], qeClips: [], kind: null };
    // Determina kind do snapshot
    var kindCount = { video: 0, audio: 0 };
    for (var sci = 0; sci < _lwSelectionSnapshot.clips.length; sci++) kindCount[_lwSelectionSnapshot.clips[sci].kind]++;
    var kind = kindCount.video >= kindCount.audio ? 'video' : 'audio';
    var trackList = (kind === 'video') ? seq.videoTracks : seq.audioTracks;
    var stdClips = [];
    var qeClips = [];
    // Match por nodeId + start time (segurança)
    for (var ti = 0; ti < trackList.numTracks; ti++) {
        var tk = trackList[ti];
        for (var ci = 0; ci < tk.clips.numItems; ci++) {
            var c = tk.clips[ci];
            try {
                var cNodeId = '';
                try { cNodeId = String(c.nodeId || ''); } catch(eN) {}
                var cStart = 0;
                try { cStart = parseFloat(c.start.seconds); } catch(eS) {}
                for (var sn = 0; sn < _lwSelectionSnapshot.clips.length; sn++) {
                    var snap = _lwSelectionSnapshot.clips[sn];
                    if (snap.kind !== kind) continue;
                    // Match por nodeId (mais confiável) ou start+name (fallback)
                    if ((cNodeId && cNodeId === snap.nodeId) ||
                        (Math.abs(cStart - snap.startSec) < 0.05 && ti === snap.trackIdx)) {
                        stdClips.push(c);
                        // Acha QE clip correspondente (mesmo index)
                        try {
                            var qeTrackList = (kind === 'video') ? qeSeq : qeSeq;
                            var qeT = (kind === 'video') ? qeSeq.getVideoTrackAt(ti) : qeSeq.getAudioTrackAt(ti);
                            if (qeT) {
                                var qeC = qeT.getItemAt(ci);
                                if (qeC) qeClips.push(qeC);
                            }
                        } catch(eQE) {}
                        break;
                    }
                }
            } catch(eC) {}
        }
    }
    return { stdClips: stdClips, qeClips: qeClips, kind: kind };
}

function lwSearchGetContext() {
    try {
        var ctx = {
            selectionType: 'none', // 'video' | 'audio' | 'mixed' | 'none'
            videoSelected: 0,
            audioSelected: 0,
            playheadSec: 0,
            playheadTicks: '0',
            hasSequence: false,
        };
        if (!app.project || !app.project.activeSequence) return _lwCtxJson(ctx);
        var seq = app.project.activeSequence;
        ctx.hasSequence = true;

        // Snapshot pra reaplicar caso foco do LION SEARCH derrube selection
        var snapClips = [];

        // Conta clips selecionados E captura snapshot
        var i, j, tk;
        for (i = 0; i < seq.videoTracks.numTracks; i++) {
            tk = seq.videoTracks[i];
            for (j = 0; j < tk.clips.numItems; j++) {
                try {
                    if (tk.clips[j].isSelected()) {
                        ctx.videoSelected++;
                        try {
                            snapClips.push({
                                kind: 'video',
                                trackIdx: i,
                                nodeId: tk.clips[j].nodeId || '',
                                startSec: parseFloat(tk.clips[j].start.seconds),
                                endSec: parseFloat(tk.clips[j].end.seconds),
                                name: String(tk.clips[j].name || '')
                            });
                        } catch(eSn) {}
                    }
                } catch(e) {}
            }
        }
        for (i = 0; i < seq.audioTracks.numTracks; i++) {
            tk = seq.audioTracks[i];
            for (j = 0; j < tk.clips.numItems; j++) {
                try {
                    if (tk.clips[j].isSelected()) {
                        ctx.audioSelected++;
                        try {
                            snapClips.push({
                                kind: 'audio',
                                trackIdx: i,
                                nodeId: tk.clips[j].nodeId || '',
                                startSec: parseFloat(tk.clips[j].start.seconds),
                                endSec: parseFloat(tk.clips[j].end.seconds),
                                name: String(tk.clips[j].name || '')
                            });
                        } catch(eSn) {}
                    }
                } catch(e) {}
            }
        }
        // Salva snapshot só se tem selection (senão preserva o anterior)
        if (snapClips.length > 0) {
            _lwSelectionSnapshot = { clips: snapClips, capturedAt: (new Date()).getTime(), seqRef: seq };
        }
        if (ctx.videoSelected > 0 && ctx.audioSelected > 0) ctx.selectionType = 'mixed';
        else if (ctx.videoSelected > 0) ctx.selectionType = 'video';
        else if (ctx.audioSelected > 0) ctx.selectionType = 'audio';
        else ctx.selectionType = 'none';

        // Playhead
        try {
            var pos = seq.getPlayerPosition();
            if (pos) {
                ctx.playheadSec = pos.seconds;
                ctx.playheadTicks = String(pos.ticks);
            }
        } catch(eP) {}

        return _lwCtxJson(ctx);
    } catch (e) {
        return '{"error":"' + _lwEscapeJson(e.toString()) + '"}';
    }
}

function _lwCtxJson(ctx) {
    var out = '{';
    out += '"selectionType":"' + ctx.selectionType + '"';
    out += ',"videoSelected":' + ctx.videoSelected;
    out += ',"audioSelected":' + ctx.audioSelected;
    out += ',"playheadSec":' + ctx.playheadSec;
    out += ',"playheadTicks":"' + ctx.playheadTicks + '"';
    out += ',"hasSequence":' + (ctx.hasSequence ? 'true' : 'false');
    out += '}';
    return out;
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — LIST AUDIO SOURCES
// Scaneia: 1) projetos abertos (active + outros se Premiere expor)
//          2) pasta de SFX configurável (Lion Workspace settings)
// ═══════════════════════════════════════════════════════════════════
function lwSearchListAudioSources(sfxFolderPath) {
    var items = [];
    var seen = {};
    var debug = [];
    var dbgLog = [];
    function dlog(m) { dbgLog.push(m); }

    try {
        // ─── PARTE 1: Projetos abertos no Premiere ───
        var projects = [];
        var seenProjects = {};
        function addProj(p, src) {
            if (!p) return;
            try {
                var key = '';
                try { key = p.path || p.name || ''; } catch(eK) {}
                if (key && seenProjects[key]) return;
                seenProjects[key] = 1;
                projects.push({ proj: p, src: src });
                dlog('  [+] proj via ' + src + ': ' + (p.name || '?'));
            } catch(e) {}
        }

        // 1a) Inspeciona app — list ALL methods/properties pra debug
        try {
            var appKeys = '';
            for (var k in app) appKeys += k + ',';
            dlog('app keys (parcial): ' + appKeys.substr(0, 600));
        } catch(eK0) { dlog('app-keys-err: ' + eK0); }

        // 1b) Habilita QE pra ter acesso multi-projeto
        try { app.enableQE(); } catch(eQE) {}

        // 1c) QE DOM: qe.numProjects / qe.getProject(i)
        try {
            if (typeof qe !== 'undefined' && qe) {
                var qeKeys = '';
                for (var qk in qe) qeKeys += qk + ',';
                dlog('qe keys (parcial): ' + qeKeys.substr(0, 400));
                var qeNumProj = 0;
                try { qeNumProj = qe.numProjects || 0; } catch(eQN) {}
                dlog('qe.numProjects=' + qeNumProj);
                for (var qpi = 0; qpi < qeNumProj; qpi++) {
                    var qp = null;
                    try { qp = qe.getProject(qpi); } catch(eGP) {}
                    if (qp) {
                        var qpName = '';
                        try { qpName = qp.name || ''; } catch(eN1) {}
                        dlog('  qe.getProject(' + qpi + ').name=' + qpName);
                    }
                }
                debug.push('qe-projects=' + qeNumProj);
            }
        } catch(eQ) { debug.push('qe-err:' + eQ.toString().slice(0, 40)); dlog('qe err: ' + eQ); }

        // 1d) app.openDocuments
        try {
            if (typeof app.openDocuments !== 'undefined') {
                var nDocs = 0;
                try { nDocs = app.openDocuments.numItems || app.openDocuments.length || 0; } catch(eN) { nDocs = 0; }
                dlog('app.openDocuments exists; numItems/length=' + nDocs);
                for (var d = 0; d < nDocs; d++) {
                    var doc = null;
                    try { doc = app.openDocuments[d]; } catch(eD) {}
                    if (!doc) try { doc = app.openDocuments.getItemAt(d); } catch(eD2) {}
                    if (doc) addProj(doc, 'openDocuments[' + d + ']');
                }
                debug.push('openDocs=' + nDocs);
            } else { dlog('app.openDocuments = undefined'); debug.push('openDocs=undef'); }
        } catch(e1) { debug.push('openDocs-err:' + e1.toString().slice(0, 40)); dlog('openDocs err: ' + e1); }

        // 1e) app.projects
        try {
            if (typeof app.projects !== 'undefined') {
                var nProj = 0;
                try { nProj = app.projects.numProjects || app.projects.numItems || app.projects.length || 0; } catch(eP) { nProj = 0; }
                dlog('app.projects exists; numProjects/length=' + nProj);
                for (var ip = 0; ip < nProj; ip++) {
                    var pr = null;
                    try { pr = app.projects[ip]; } catch(ePr) {}
                    if (pr) addProj(pr, 'projects[' + ip + ']');
                }
                debug.push('projects=' + nProj);
            } else { dlog('app.projects = undefined'); debug.push('projects=undef'); }
        } catch(e2) { debug.push('projects-err:' + e2.toString().slice(0, 40)); dlog('projects err: ' + e2); }

        // 1f) Fallback: active project (sempre adicionado)
        try {
            if (app.project) {
                addProj(app.project, 'active');
            }
        } catch(e3) {}

        dlog('Total projetos detectados: ' + projects.length);

        // 1e) Itera cada projeto
        for (var pi = 0; pi < projects.length; pi++) {
            var proj = projects[pi].proj;
            var projName = '';
            try { projName = proj.name || ''; } catch(eN) {}
            try { projName = projName.replace(/\.prproj$/i, ''); } catch(eR) {}

            var rootItem = null;
            try { rootItem = proj.rootItem; } catch(eRoot) {}
            if (!rootItem) { debug.push('p[' + pi + ']:noroot'); continue; }

            var beforeCount = items.length;
            _lwScanForAudio(rootItem, '', items, seen, 0, projName);
            dlog('  proj[' + projName + '] -> +' + (items.length - beforeCount) + ' áudios');
        }

        // ─── PARTE 2: Pasta SFX configurável ───
        if (sfxFolderPath) {
            dlog('Scanning SFX folder: ' + sfxFolderPath);
            var sfxF = new Folder(sfxFolderPath);
            if (sfxF.exists) {
                var beforeFs = items.length;
                _lwScanFsForAudio(sfxF, '', items, seen, 0);
                dlog('  SFX folder -> +' + (items.length - beforeFs) + ' áudios');
                debug.push('sfx-folder=' + (items.length - beforeFs));
            } else {
                dlog('  SFX folder NÃO existe: ' + sfxFolderPath);
                debug.push('sfx-folder=missing');
            }
        }
    } catch(e) {
        return _lwSearchJson(items, 'audio-scan-err:' + e.toString().slice(0, 80), []);
    }

    // Escreve debug detalhado
    try {
        var fdbg = new File(Folder.temp.fsName + '/lion-search-audio-debug.txt');
        if (fdbg.open('w')) {
            fdbg.encoding = 'UTF-8';
            for (var dli = 0; dli < dbgLog.length; dli++) fdbg.writeln(dbgLog[dli]);
            fdbg.writeln('');
            fdbg.writeln('--- Total: ' + items.length + ' ---');
            for (var fi = 0; fi < Math.min(20, items.length); fi++) {
                fdbg.writeln(items[fi].name + ' | ' + items[fi].category);
            }
            fdbg.close();
        }
    } catch(eFw) {}

    return _lwSearchJson(items, 'audio:' + debug.join(',') + ',total=' + items.length, []);
}

// Scaneia pasta do filesystem por arquivos de áudio
function _lwScanFsForAudio(folder, parentPath, items, seen, depth) {
    if (!folder || depth > 8) return;
    var files = null;
    try { files = folder.getFiles(); } catch(e) { return; }
    if (!files) return;
    for (var i = 0; i < files.length; i++) {
        var f = files[i];
        if (f instanceof Folder) {
            var sub = parentPath ? parentPath + ' / ' + f.name : f.name;
            _lwScanFsForAudio(f, sub, items, seen, depth + 1);
        } else {
            var fname = String(f.name);
            var lname = fname.toLowerCase();
            if (lname.match(/\.(mp3|wav|aac|m4a|flac|ogg|aif|aiff|wma|opus)$/)) {
                var fullPath = f.fsName;
                if (seen[fullPath]) continue;
                seen[fullPath] = 1;
                var displayName = fname;
                items.push({
                    kind: 'audio-source',
                    name: displayName,
                    matchName: 'FS:' + fullPath, // prefix FS: pra apply saber que é do filesystem
                    category: '🗂️ ' + (parentPath || 'SFX'),
                    sfxPath: fullPath,
                });
            }
        }
    }
}

function _lwScanForAudio(item, parentPath, items, seen, depth, projName) {
    if (!item || depth > 8) return;
    try {
        // type 1=BIN, 2=CLIP, 3=ROOT, 4=FILE — vamos recursar em todos
        var name = '';
        try { name = item.name || ''; } catch(eN) {}
        var children = null;
        try { children = item.children; } catch(eC) {}

        if (children && children.numItems > 0) {
            var subPath = (parentPath || depth === 0) ? (parentPath ? parentPath + ' / ' + name : name) : '';
            for (var i = 0; i < children.numItems; i++) {
                _lwScanForAudio(children[i], subPath, items, seen, depth + 1, projName);
            }
            return;
        }

        // Não tem filhos — é uma folha (clip de mídia)
        var nodeId = '';
        try { nodeId = item.nodeId || name; } catch(eId) { nodeId = name; }

        // Detecta se tem áudio
        var hasAudio = false;
        try {
            if (item.hasAudio && typeof item.hasAudio === 'function') hasAudio = item.hasAudio();
            else if (item.hasAudio === true) hasAudio = true;
        } catch(eH) {}

        // Fallback por extensão
        if (!hasAudio) {
            var lname = String(name).toLowerCase();
            if (lname.match(/\.(mp3|wav|aac|m4a|flac|ogg|aif|aiff|wma|opus)$/)) hasAudio = true;
        }
        // Detecta se tem video tb (pra rotular categoria)
        var hasVideo = false;
        try {
            if (item.hasVideo && typeof item.hasVideo === 'function') hasVideo = item.hasVideo();
        } catch(eV) {}

        // Dedup por file path (mesmo som em múltiplos projetos = 1 entrada só)
        // Tenta MÚLTIPLAS APIs pra obter o path — algumas falham silenciosamente
        var mediaPath = '';
        try { if (item.getMediaPath) mediaPath = item.getMediaPath() || ''; } catch(eMp1) {}
        if (!mediaPath) { try { mediaPath = item.canonicalUri || ''; } catch(eMp2) {} }
        if (!mediaPath) { try { mediaPath = item.path || ''; } catch(eMp3) {} }
        // canonicalUri vem como "file:///C:/..." — normaliza
        if (mediaPath && mediaPath.indexOf('file:') === 0) {
            try { mediaPath = decodeURIComponent(mediaPath.replace(/^file:\/+/, '')); } catch(eD) {}
        }
        var dedupKey = mediaPath || ((projName || '') + '|' + nodeId + '|' + name);
        if (hasAudio && !seen[dedupKey]) {
            // FILTRO OFFLINE: pula items cujo arquivo não existe no disco.
            // Items sem mediaPath OU com mediaPath inválido → pula.
            if (!mediaPath) {
                try { _lwDbgAudio('  scan skip (no path): ' + name); } catch(eDl) {}
                return;
            }
            var fileExists = false;
            var triedPaths = [];
            // Tenta o path direto
            try {
                triedPaths.push(mediaPath);
                var ff = new File(mediaPath);
                fileExists = ff.exists;
            } catch(eFE) {}
            // URL-decoded
            if (!fileExists) {
                try {
                    var decoded = decodeURIComponent(mediaPath);
                    if (decoded !== mediaPath) {
                        triedPaths.push(decoded);
                        var ff2 = new File(decoded);
                        fileExists = ff2.exists;
                    }
                } catch(eD) {}
            }
            // Sem prefixo file://
            if (!fileExists && mediaPath.indexOf('file:') === 0) {
                try {
                    var noPrefix = mediaPath.replace(/^file:\/+/, '');
                    triedPaths.push(noPrefix);
                    var ff3 = new File(noPrefix);
                    fileExists = ff3.exists;
                } catch(eF) {}
            }
            // Mac: tenta com slashes invertidos (path windows com mediaPath linux?)
            if (!fileExists && mediaPath.indexOf('/') >= 0) {
                try {
                    var swapped = mediaPath.replace(/\//g, '\\');
                    if (swapped !== mediaPath) {
                        triedPaths.push(swapped);
                        var ff4 = new File(swapped);
                        fileExists = ff4.exists;
                    }
                } catch(eS) {}
            }
            if (!fileExists) {
                try { _lwDbgAudio('  scan SKIP OFFLINE: ' + name + ' | path=' + mediaPath + ' | tried=' + triedPaths.length); } catch(eDl2) {}
                return;
            }

            // FILTRO DURAÇÃO: precisa ter in/out points válidos pra ser inserível
            // Items sem duração (clipes vazios, marcadores) não fazem sentido no LION SEARCH
            var hasValidDuration = false;
            try {
                var outP = null;
                // Tenta audio (4), video (1), both (0) — qualquer um com out > 0 vale
                for (var mtt = 0; mtt < 3 && !hasValidDuration; mtt++) {
                    try {
                        var op = item.getOutPoint([4,1,0][mtt]);
                        if (op && parseFloat(op.seconds) > 0) hasValidDuration = true;
                    } catch(eDD) {}
                }
            } catch(eDur) {}
            if (!hasValidDuration) { return; } // pula items sem duração inserível

            seen[dedupKey] = 1;
            var categoryParts = [];
            if (projName) categoryParts.push(projName);
            if (parentPath) categoryParts.push(parentPath);
            else categoryParts.push(hasVideo ? 'Audio+Video' : 'Audio');
            // matchName: prefere PATH (global unique, funciona em qualquer projeto).
            // Prefixo "FS:" pro código de insert saber que é caminho de arquivo
            var matchName = 'FS:' + mediaPath;
            items.push({
                kind: 'audio-source',
                name: name || '(sem nome)',
                matchName: matchName,
                category: categoryParts.join(' / '),
                projName: projName || '',
            });
        }
    } catch(e) {}
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — LIST PRESETS (.prfpset no filesystem)
// Estratégia: scan agressivo em múltiplos paths, escreve debug detalhado.
//   Win: %USERPROFILE%\Documents\Adobe\Premiere Pro\<ver>\Profile-<user>\...
//        %APPDATA%\Adobe\Premiere Pro\<ver>\Profile-<user>\...
//        + OneDrive Documents (se sync ativo)
//   Mac: ~/Documents/Adobe/Premiere Pro/<ver>/Profile-<user>/...
//        ~/Library/Application Support/Adobe/Premiere Pro/<ver>/Profile-<user>/...
// ═══════════════════════════════════════════════════════════════════
function lwSearchListPresets() {
    var items = [];
    var debug = '';
    var pathsTried = [];
    var pathsExisted = [];
    var dbgLog = []; // Log detalhado pra arquivo

    function dlog(msg) { try { dbgLog.push(msg); } catch(e) {} }

    // ─── ESTRATÉGIA 1: API QE (Excalibur usa essa via) ───
    try {
        if (typeof qe === 'undefined') { try { app.enableQE(); } catch(eE) {} }
        if (typeof qe !== 'undefined' && qe && qe.project) {
            // Lista TODAS as propriedades/métodos de qe.project (debug)
            var qePjKeys = '';
            try { for (var k in qe.project) qePjKeys += k + ','; } catch(eK) {}
            dlog('qe.project keys: ' + qePjKeys.substr(0, 800));

            // Tenta cada possível API de preset
            var presetApiNames = [
                'getEffectPresetList',
                'getEffectPresets',
                'getPresetList',
                'getPresets',
                'getVideoEffectPresetList',
                'getAudioEffectPresetList',
                'effectPresets',
                'presets',
            ];
            for (var pa = 0; pa < presetApiNames.length; pa++) {
                var apiName = presetApiNames[pa];
                try {
                    var fn = qe.project[apiName];
                    if (!fn) continue;
                    var result = null;
                    if (typeof fn === 'function') {
                        try { result = fn.call(qe.project); } catch(eC) { dlog('  ' + apiName + '() throw: ' + eC); continue; }
                    } else {
                        result = fn;
                    }
                    if (result) {
                        var rType = typeof result;
                        var rLen = 0;
                        try { rLen = result.numItems || result.length || 0; } catch(eL) {}
                        dlog('  qe.project.' + apiName + ' = ' + rType + ' (len=' + rLen + ')');
                        // Se for lista, processa
                        if (rLen > 0) {
                            for (var ri = 0; ri < rLen; ri++) {
                                var entry = null;
                                try { entry = result[ri]; } catch(eR) {}
                                if (!entry) try { entry = result.getItemAt(ri); } catch(eR2) {}
                                if (!entry) continue;
                                var name = '', match = '', cat = '';
                                if (typeof entry === 'string') {
                                    name = entry; match = entry;
                                } else {
                                    try { name = entry.name || entry.title || ''; } catch(eN) {}
                                    try { match = entry.matchName || entry.name || ''; } catch(eM) {}
                                    try { cat = entry.category || ''; } catch(eCC) {}
                                }
                                if (name) {
                                    items.push({
                                        kind: 'preset',
                                        name: name,
                                        matchName: match || name,
                                        category: cat || ('QE: ' + apiName),
                                    });
                                }
                            }
                        }
                    }
                } catch(eApi) { dlog('  ' + apiName + ' err: ' + eApi); }
            }
        }
    } catch(eQE) { dlog('QE preset err: ' + eQE); }

    if (items.length > 0) {
        dlog('Total presets via QE: ' + items.length);
    }

    try {
        var rootCandidates = [];

        // 0) APIs do app (mais confiáveis — sabem o caminho exato da Adobe)
        try {
            if (typeof app.getPProPrefPath === 'function') {
                var prefPath = app.getPProPrefPath();
                if (prefPath) {
                    dlog('app.getPProPrefPath()=' + prefPath);
                    // pref path é tipo .../Profile-<user>/, então adiciona seu pai
                    rootCandidates.push(prefPath);
                    // Também adiciona o pai (volta níveis pra Adobe Premiere Pro/)
                    var p2 = String(prefPath).replace(/\\/g, '/');
                    var parts = p2.split('/');
                    while (parts.length > 0) {
                        var newPath = parts.join('/');
                        if (newPath) rootCandidates.push(newPath);
                        // Para quando achar "Adobe" no path
                        if (parts[parts.length - 1] === 'Adobe' || parts.length <= 2) break;
                        parts.pop();
                    }
                }
            }
        } catch(eAPP) { dlog('getPProPrefPath err: ' + eAPP); }

        try {
            if (typeof app.getPProSystemPrefPath === 'function') {
                var sysPrefPath = app.getPProSystemPrefPath();
                if (sysPrefPath) {
                    dlog('app.getPProSystemPrefPath()=' + sysPrefPath);
                    rootCandidates.push(sysPrefPath);
                }
            }
        } catch(eAPS) { dlog('getPProSystemPrefPath err: ' + eAPS); }

        // app.path = caminho do executável da Adobe Premiere Pro
        try {
            if (app.path) {
                dlog('app.path=' + app.path);
                // app.path é o .exe file, get parent dir
                var ap = String(app.path).replace(/\\/g, '/');
                var apParent = ap.substring(0, ap.lastIndexOf('/'));
                if (apParent) {
                    rootCandidates.push(apParent);
                    rootCandidates.push(apParent + '/Effect Presets');
                    // Mac .app bundle
                    rootCandidates.push(apParent + '/../Resources');
                }
            }
        } catch(eAp) { dlog('app.path err: ' + eAp); }

        // 1) Folder.myDocuments
        try {
            if (Folder.myDocuments) {
                rootCandidates.push(Folder.myDocuments.fsName + '/Adobe/Premiere Pro');
            }
        } catch(e1) { dlog('myDocuments err: ' + e1); }

        // 2) Folder.appData (Win: %APPDATA% | Mac: ~/Library/Application Support)
        try {
            if (Folder.appData) {
                rootCandidates.push(Folder.appData.fsName + '/Adobe/Premiere Pro');
            }
        } catch(e2) { dlog('appData err: ' + e2); }

        // 3) Folder.userData
        try {
            if (Folder.userData && Folder.userData.fsName !== Folder.appData.fsName) {
                rootCandidates.push(Folder.userData.fsName + '/Adobe/Premiere Pro');
            }
        } catch(e3) {}

        // 4) Variáveis env
        if (File.fs === 'Windows') {
            try {
                var userProfile = $.getenv('USERPROFILE');
                if (userProfile) {
                    rootCandidates.push(userProfile + '/Documents/Adobe/Premiere Pro');
                    rootCandidates.push(userProfile + '/OneDrive/Documents/Adobe/Premiere Pro');
                    rootCandidates.push(userProfile + '/OneDrive/Documentos/Adobe/Premiere Pro');
                    rootCandidates.push(userProfile + '/Documentos/Adobe/Premiere Pro');
                }
                var appData = $.getenv('APPDATA');
                if (appData) rootCandidates.push(appData + '/Adobe/Premiere Pro');
            } catch(eEnv) { dlog('env err: ' + eEnv); }
        } else if (File.fs === 'Macintosh') {
            try {
                var home = $.getenv('HOME');
                if (home) {
                    rootCandidates.push(home + '/Documents/Adobe/Premiere Pro');
                    rootCandidates.push(home + '/Library/Application Support/Adobe/Premiere Pro');
                }
                rootCandidates.push('/Library/Application Support/Adobe/Premiere Pro');
            } catch(eMac) { dlog('mac env err: ' + eMac); }
        }

        // Dedup paths
        var uniqRoots = [];
        var seenPath = {};
        for (var dr = 0; dr < rootCandidates.length; dr++) {
            var pp = rootCandidates[dr];
            if (!pp || seenPath[pp]) continue;
            seenPath[pp] = 1;
            uniqRoots.push(pp);
        }

        dlog('--- Paths tentados ---');
        for (var lp = 0; lp < uniqRoots.length; lp++) dlog('  ' + uniqRoots[lp]);

        // Encontra pastas de presets em cada root
        var presetRoots = [];
        var seenRoots = {};
        for (var rc = 0; rc < uniqRoots.length; rc++) {
            var rootPath = uniqRoots[rc];
            pathsTried.push(rootPath);
            var rootF = new Folder(rootPath);
            if (!rootF.exists) { dlog('  [NAO EXISTE] ' + rootPath); continue; }
            pathsExisted.push(rootPath);
            dlog('  [OK] ' + rootPath);
            _lwFindPresetRoots(rootF, presetRoots, seenRoots, 0, dlog);
        }

        dlog('--- Pastas de preset encontradas ---');
        for (var lpr = 0; lpr < presetRoots.length; lpr++) dlog('  ' + presetRoots[lpr].fsName);

        // Scan cada root de preset
        var seen = {};
        for (var r = 0; r < presetRoots.length; r++) {
            _lwScanPresetFolder(presetRoots[r], '', items, seen, 0);
        }
        dlog('Total .prfpset achados: ' + items.length);

        // Dedup por NAME (mesmo preset em múltiplas versões = 1 entrada só)
        var dedupMap = {};
        var dedupedItems = [];
        for (var di = 0; di < items.length; di++) {
            var nKey = items[di].name + '|' + (items[di].category || '');
            if (!dedupMap[nKey]) {
                dedupMap[nKey] = 1;
                dedupedItems.push(items[di]);
            }
        }
        items = dedupedItems;
        dlog('Após dedup por nome: ' + items.length);

        debug = 'tried=' + pathsTried.length + ',existed=' + pathsExisted.length + ',roots=' + presetRoots.length + ',presets=' + items.length;
    } catch(e) {
        debug = 'preset-scan-err:' + e.toString().slice(0, 80);
        dlog('OUTER ERR: ' + e);
    }

    // Escreve log detalhado em arquivo temp pra debug
    try {
        var fdbg = new File(Folder.temp.fsName + '/lion-search-presets-debug.txt');
        if (fdbg.open('w')) {
            fdbg.encoding = 'UTF-8';
            fdbg.writeln('Platform: ' + File.fs);
            fdbg.writeln('Folder.myDocuments: ' + (Folder.myDocuments ? Folder.myDocuments.fsName : 'null'));
            fdbg.writeln('Folder.appData: ' + (Folder.appData ? Folder.appData.fsName : 'null'));
            fdbg.writeln('Folder.userData: ' + (Folder.userData ? Folder.userData.fsName : 'null'));
            try { fdbg.writeln('USERPROFILE: ' + $.getenv('USERPROFILE')); } catch(e) {}
            try { fdbg.writeln('HOME: ' + $.getenv('HOME')); } catch(e) {}
            try { fdbg.writeln('APPDATA: ' + $.getenv('APPDATA')); } catch(e) {}
            fdbg.writeln('');
            for (var dl = 0; dl < dbgLog.length; dl++) fdbg.writeln(dbgLog[dl]);
            fdbg.writeln('');
            fdbg.writeln('--- First 10 presets ---');
            for (var fi = 0; fi < Math.min(10, items.length); fi++) {
                fdbg.writeln(items[fi].name + ' @ ' + items[fi].matchName);
            }
            fdbg.close();
        }
    } catch(eFw) {}

    return _lwSearchJson(items, debug, []);
}

// Encontra recursivamente pastas com .prfpset OU pastas com nome conhecido
// Estratégia: marca uma pasta como "preset root" se contém pelo menos 1 .prfpset
// OU se o nome bate com "Effect Presets" / similar
function _lwFindPresetRoots(folder, out, seen, depth, dlog) {
    if (!folder || depth > 10) return;
    var files = null;
    try { files = folder.getFiles(); } catch(e) {
        if (dlog) dlog('  getFiles err @ ' + folder.fsName + ': ' + e);
        return;
    }
    if (!files) return;
    var hasPrfpsetFile = false;
    for (var fi = 0; fi < files.length; fi++) {
        var ff = files[fi];
        if (ff instanceof File) {
            var nm = String(ff.name).toLowerCase();
            if (nm.length > 8 && nm.substr(nm.length - 8) === '.prfpset') {
                hasPrfpsetFile = true;
                break;
            }
        }
    }
    var fname = String(folder.name);
    var lname = fname.toLowerCase();
    var isKnownPresetFolder = (
        fname === 'Effect Presets and Custom Items'
        || fname === 'Effect Presets'
        || lname === 'effect presets and custom items'
        || lname === 'effect presets'
        || (lname.indexOf('preset') >= 0 && (lname.indexOf('effect') >= 0 || lname.indexOf('custom') >= 0))
    );
    if (hasPrfpsetFile || isKnownPresetFolder) {
        if (!seen[folder.fsName]) {
            seen[folder.fsName] = 1;
            out.push(folder);
            if (dlog) dlog('  [PRESET ROOT] ' + folder.fsName + (hasPrfpsetFile ? ' (has .prfpset)' : ' (named)'));
        }
    }
    // Continua descendo em todas as subpastas (limitado por depth)
    for (var i = 0; i < files.length; i++) {
        var f = files[i];
        if (!(f instanceof Folder)) continue;
        _lwFindPresetRoots(f, out, seen, depth + 1, dlog);
    }
}

// Helper: ExtendScript File.name retorna URL-encoded ("Cross%20Dissolve.prfpset")
// — converte pra string legível
function _lwDecodeName(s) {
    if (!s) return '';
    try { return decodeURIComponent(s); } catch(e) { return s; }
}

function _lwScanPresetFolder(folder, parentPath, items, seen, depth) {
    if (!folder || depth > 8) return;
    var files;
    try { files = folder.getFiles(); } catch(e) { return; }
    if (!files) return;
    for (var i = 0; i < files.length; i++) {
        var f = files[i];
        if (f instanceof Folder) {
            var sub = parentPath ? parentPath + ' / ' + _lwDecodeName(f.name) : _lwDecodeName(f.name);
            _lwScanPresetFolder(f, sub, items, seen, depth + 1);
        } else {
            var fname = _lwDecodeName(String(f.name));
            var lower = fname.toLowerCase();
            if (lower.length > 8 && lower.substr(lower.length - 8) === '.prfpset') {
                var displayName = fname.substr(0, fname.length - 8);
                var displayLower = displayName.toLowerCase();
                if (displayLower === '.ds_store') continue;
                // Container files com várias presets dentro: marca pra parser parsear
                var isContainer = false;
                var containerNames = ['factory presets', 'lumetri presets', 'lumetri color presets', 'maskpresets', 'effect presets and custom items'];
                for (var cn = 0; cn < containerNames.length; cn++) {
                    if (displayLower === containerNames[cn]) { isContainer = true; break; }
                }
                var fullPath = f.fsName;
                if (seen[fullPath]) continue;
                seen[fullPath] = 1;
                var pathLower = String(parentPath + '/' + displayName).toLowerCase();
                var audioOnly = (pathLower.indexOf('audio') >= 0);
                var videoOnly = (pathLower.indexOf('video') >= 0 && !audioOnly);
                items.push({
                    kind: 'preset',
                    name: displayName,
                    matchName: fullPath,
                    category: parentPath || 'Presets',
                    audioOnly: audioOnly,
                    videoOnly: videoOnly,
                    isContainer: isContainer, // plugin client vai parsear se for container
                });
            }
        }
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — APPLY PRESET (drag .prfpset no clip selecionado)
// ═══════════════════════════════════════════════════════════════════
// Helper: escreve debug pra arquivo (apply preset)
function _lwDbgPreset(msg) {
    try {
        var f = new File(Folder.temp.fsName + '/lion-search-apply-preset.txt');
        if (f.open('a')) {
            f.encoding = 'UTF-8';
            f.write('[' + (new Date()).toString() + '] ' + msg + '\n');
            f.close();
        }
    } catch(e) {}
}

// Helper: resolve path tentando múltiplas variações (URL-encoded, slashes diferentes)
// Retorna o primeiro path que existe, ou null se nenhum existe
function _lwResolveExistingPath(rawPath) {
    if (!rawPath) return null;
    var p = String(rawPath);
    // Remove file:// prefix se houver
    p = p.replace(/^file:\/+/, '');
    // Lista de variações pra tentar
    var tries = [p];
    try { tries.push(decodeURIComponent(p)); } catch(e1) {}
    tries.push(p.replace(/\//g, '\\'));
    tries.push(p.replace(/\\/g, '/'));
    try { tries.push(decodeURIComponent(p.replace(/\//g, '\\'))); } catch(e2) {}
    try { tries.push(decodeURIComponent(p.replace(/\\/g, '/'))); } catch(e3) {}

    var seen = {};
    for (var i = 0; i < tries.length; i++) {
        var t = tries[i];
        if (!t || seen[t]) continue;
        seen[t] = 1;
        try {
            var f = new File(t);
            if (f.exists) return t;
        } catch(eE) {}
    }
    return null;
}

// Helper: encontra um project item por nome no rootItem (busca recursiva)
function _lwFindProjectItemByName(bin, targetName, depth) {
    if (!bin || depth > 12) return null;
    try {
        if (!bin.children || bin.children.numItems === 0) return null;
        for (var i = 0; i < bin.children.numItems; i++) {
            var it = bin.children[i];
            var nm = '';
            try { nm = it.name || ''; } catch(eN) {}
            if (nm === targetName) return it;
            // Recurse em bins
            var ch = null;
            try { ch = it.children; } catch(eC) {}
            if (ch && ch.numItems > 0) {
                var found = _lwFindProjectItemByName(it, targetName, depth + 1);
                if (found) return found;
            }
        }
    } catch(e) {}
    return null;
}

// Helper: tenta chamar um método em obj com args, retorna {ok, err}
function _lwTryCall(obj, methodName, args, label) {
    try {
        var fn = obj[methodName];
        if (typeof fn !== 'function') return { ok: false, err: methodName + ' not a function (typeof=' + typeof fn + ')' };
        var result;
        if (args.length === 0) result = fn.call(obj);
        else if (args.length === 1) result = fn.call(obj, args[0]);
        else result = fn.apply(obj, args);
        return { ok: true, label: label, result: result };
    } catch (e) {
        return { ok: false, err: methodName + ': ' + e.toString().slice(0, 80), label: label };
    }
}

function _lwApplyPreset(presetIdentifier) {
    try {
        _lwDbgPreset('═══ Apply Preset: ' + presetIdentifier + ' ═══');
        if (!presetIdentifier) return 'ERR:preset_id_empty';
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        try { app.enableQE(); } catch(e0) {}
        if (typeof qe === 'undefined' || !qe.project) return 'ERR:qe_not_available';

        // Parse identifier (containerPath#presetName ou só path/nome)
        var presetName = presetIdentifier;
        var presetPath = presetIdentifier;
        var hashIdx = presetIdentifier.indexOf('#');
        if (hashIdx > 0) {
            presetPath = presetIdentifier.substring(0, hashIdx);
            presetName = presetIdentifier.substring(hashIdx + 1);
        } else if (!/\.prfpset$/i.test(presetIdentifier)) {
            presetPath = '';
        }
        _lwDbgPreset('  presetName="' + presetName + '"');
        _lwDbgPreset('  presetPath="' + presetPath + '"');

        var qeSeq = qe.project.getActiveSequence();
        if (!qeSeq) return 'ERR:no_qe_sequence';
        var seq = app.project.activeSequence;
        var sel = _lwFindSelectedQEClips(qeSeq, true, seq);
        var selKind = 'video';
        if (sel.length === 0) { sel = _lwFindSelectedQEClips(qeSeq, false, seq); selKind = 'audio'; }

        // RECOVERY: se selection volta vazia (Premiere perdeu por causa do foco do LION SEARCH),
        // usa snapshot capturado em lwSearchGetContext (que rodou ANTES da janela tirar foco)
        var stdSel = [];
        if (sel.length === 0) {
            _lwDbgPreset('  selection vazia — tentando recovery via snapshot');
            var rec = _lwRecoverFromSnapshot(seq, qeSeq);
            if (rec.qeClips.length > 0) {
                sel = rec.qeClips;
                stdSel = rec.stdClips;
                selKind = rec.kind;
                _lwDbgPreset('  recovery: ' + sel.length + ' qe + ' + stdSel.length + ' std (kind=' + selKind + ', snapshot age=' + (((new Date()).getTime() - _lwSelectionSnapshot.capturedAt)/1000).toFixed(1) + 's)');
            } else {
                _lwDbgPreset('  recovery falhou — snapshot vazio ou stale');
                return 'NO_CLIP_SELECTED';
            }
        }
        _lwDbgPreset('  ' + sel.length + ' QE clips (' + selKind + ')');

        // Pega standard API clips (alguns métodos só funcionam aqui) — só se recovery não populou
        if (stdSel.length === 0) {
            try {
                var trackList = (selKind === 'video') ? seq.videoTracks : seq.audioTracks;
                for (var ti = 0; ti < trackList.numTracks; ti++) {
                    var tk = trackList[ti];
                    for (var ci = 0; ci < tk.clips.numItems; ci++) {
                        if (tk.clips[ci].isSelected()) stdSel.push(tk.clips[ci]);
                    }
                }
            } catch(eStd) {}
        }
        _lwDbgPreset('  ' + stdSel.length + ' std clips');

        // Inspeciona métodos do clip pra debug (apesar de for-in não mostrar nativos)
        if (stdSel.length > 0) {
            try {
                var keys = '';
                for (var sk in stdSel[0]) keys += sk + ',';
                _lwDbgPreset('  std clip keys: ' + keys.substr(0, 400));
            } catch(eKS) {}
        }

        var applied = 0;
        var allErrs = [];

        // ─── Prepare File object ───
        var pf = null;
        if (presetPath) {
            try {
                pf = new File(presetPath);
                if (!pf.exists) {
                    _lwDbgPreset('  ERROR: preset file does not exist: ' + presetPath);
                    pf = null;
                }
            } catch(eF) { _lwDbgPreset('  File obj err: ' + eF); }
        }

        // ═══════════════════════════════════════════════════════════════════
        // ESTRATÉGIA: REGEX scan (1000x mais rápido que XML() em arquivo 22MB)
        // ExtendScript XML class demora 46s+ pra parsear arquivo grande — inútil.
        // Em vez disso: indexOf da presetName, extrai janela ao redor, regex.
        // ═══════════════════════════════════════════════════════════════════
        if (pf && presetName) {
            try {
                var t0 = (new Date()).getTime();
                _lwDbgPreset('  → REGEX scan strategy');
                pf.encoding = 'UTF-8';
                if (pf.open('r')) {
                    var xmlRaw = pf.read();
                    pf.close();
                    _lwDbgPreset('  read ' + xmlRaw.length + ' chars in ' + ((new Date()).getTime() - t0) + 'ms');

                    // Gzip check: 0x1f 0x8b
                    var firstByte = xmlRaw.length > 0 ? xmlRaw.charCodeAt(0) : 0;
                    if (firstByte === 0x1f) {
                        _lwDbgPreset('  ✗ file is gzipped — JSX cannot decompress; needs plugin Node.js side');
                    } else {
                        // ─── ESTRATÉGIA NOVA: regex direto, sem XML() ───
                        var tFind = (new Date()).getTime();
                        // Escapa caracteres especiais no nome do preset pro regex
                        var escName = String(presetName).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                        // Procura <Name>presetName</Name> (com whitespace tolerante)
                        var nameTagRe = new RegExp('<Name[^>]*>\\s*' + escName + '\\s*</Name>', 'i');
                        var nameMatch = nameTagRe.exec(xmlRaw);
                        if (!nameMatch) {
                            // Tenta encoding alternativo: pode ter & encoded
                            var escNameXml = String(presetName).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
                            var escNameXml2 = escNameXml.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                            var nameTagRe2 = new RegExp('<Name[^>]*>\\s*' + escNameXml2 + '\\s*</Name>', 'i');
                            nameMatch = nameTagRe2.exec(xmlRaw);
                        }
                        _lwDbgPreset('  regex find name: ' + ((new Date()).getTime() - tFind) + 'ms, match=' + (nameMatch ? 'YES @ pos=' + nameMatch.index : 'NO'));

                        if (!nameMatch) {
                            // Debug: lista alguns Names próximos ao primeiro <Name> pra ver formato
                            _lwDbgPreset('  presetName "' + presetName + '" NOT in XML — listando primeiros 20 Names:');
                            var allNamesRe = /<Name[^>]*>([^<]+)<\/Name>/g;
                            var anm; var nameCount = 0;
                            while ((anm = allNamesRe.exec(xmlRaw)) !== null && nameCount < 20) {
                                _lwDbgPreset('    Name[' + nameCount + ']: ' + anm[1].substr(0, 80));
                                nameCount++;
                            }
                        } else {
                            // DUMP: 800 chars logo após o <Name> match pra ver estrutura
                            var afterDump = xmlRaw.substring(nameMatch.index, Math.min(xmlRaw.length, nameMatch.index + 800));
                            _lwDbgPreset('  XML dump após match (800 chars):');
                            // Quebra em pedaços de 200 chars
                            for (var dd = 0; dd < afterDump.length; dd += 200) {
                                _lwDbgPreset('    >> ' + afterDump.substr(dd, 200).replace(/\r?\n/g, ' '));
                            }

                            // Procura ObjectRef próximo do Name pra saber onde está o EffectPreset
                            // Padrão Adobe: BinTreeItem com Name → Ref pra EffectPreset com EffectAppliedList
                            var refRe = /<Ref\s+ObjectRef="(\d+)"/g;
                            var refSearch = xmlRaw.substring(nameMatch.index, Math.min(xmlRaw.length, nameMatch.index + 2000));
                            var refs = [];
                            var refM;
                            while ((refM = refRe.exec(refSearch)) !== null) refs.push(refM[1]);
                            _lwDbgPreset('  refs found near preset: ' + refs.slice(0, 5).join(','));

                            // Acha o ObjectID logo antes do Name (este é o BinTreeItem do preset)
                            var beforeName = xmlRaw.substring(Math.max(0, nameMatch.index - 500), nameMatch.index);
                            var ownIdMatch = /<\w+\s+ObjectID="(\d+)"[^>]*>(?:(?!<\w+\s+ObjectID).)*$/m.exec(beforeName);
                            var ownId = ownIdMatch ? ownIdMatch[1] : '';
                            _lwDbgPreset('  preset BinTreeItem ObjectID=' + ownId);

                            // Estratégia 1: janela GRANDE (100KB cada lado = 200KB total)
                            var startPos = Math.max(0, nameMatch.index - 100000);
                            var endPos = Math.min(xmlRaw.length, nameMatch.index + 100000);
                            var window = xmlRaw.substring(startPos, endPos);
                            _lwDbgPreset('  extracted ' + window.length + ' chars window (200KB)');

                            // Coleta TODOS os MatchName na janela
                            var matchNames = [];
                            var mnSeen = {};
                            var mnRe = /<MatchName[^>]*>([^<]+)<\/MatchName>/g;
                            var mnm;
                            while ((mnm = mnRe.exec(window)) !== null) {
                                var mnv = String(mnm[1]).replace(/^\s+|\s+$/g, '');
                                if (mnv && !mnSeen[mnv]) { mnSeen[mnv] = 1; matchNames.push(mnv); }
                            }
                            _lwDbgPreset('  found ' + matchNames.length + ' MatchNames in window');

                            // Estratégia 2: Se janela não acha, busca via ObjectRef
                            // Se achamos refs perto do Name, cada ref aponta pra um EffectApplied
                            // que contém ObjectRef → Effect com MatchName
                            if (matchNames.length === 0 && refs.length > 0) {
                                _lwDbgPreset('  fallback: seguindo ObjectRefs');
                                for (var ri = 0; ri < refs.length; ri++) {
                                    // Procura <... ObjectID="REF"> no arquivo
                                    var refIdRe = new RegExp('<\\w+\\s+ObjectID="' + refs[ri] + '"[^>]*>', 'g');
                                    var refIdMatch = refIdRe.exec(xmlRaw);
                                    if (refIdMatch) {
                                        // Pega 5KB após esse ObjectID — deve conter MatchNames
                                        var refSection = xmlRaw.substring(refIdMatch.index, Math.min(xmlRaw.length, refIdMatch.index + 5000));
                                        _lwDbgPreset('    ref ' + refs[ri] + ' found @ pos=' + refIdMatch.index);
                                        var refMnRe = /<MatchName[^>]*>([^<]+)<\/MatchName>/g;
                                        var refMnm;
                                        while ((refMnm = refMnRe.exec(refSection)) !== null) {
                                            var refMnv = String(refMnm[1]).replace(/^\s+|\s+$/g, '');
                                            if (refMnv && !mnSeen[refMnv]) { mnSeen[refMnv] = 1; matchNames.push(refMnv); }
                                        }
                                    }
                                }
                                _lwDbgPreset('  after ObjectRef chase: ' + matchNames.length + ' MatchNames');
                            }

                            for (var mi = 0; mi < Math.min(20, matchNames.length); mi++) {
                                _lwDbgPreset('    fx[' + mi + ']: ' + matchNames[mi]);
                            }

                            // Aplica cada efeito via qe.project.getVideoEffectByName/MatchName
                            // Mudança: NÃO pula Motion/Opacity — tenta aplicar valores via clip.components
                            var addedCount = 0;
                            var motionApplied = false;
                            for (var mi2 = 0; mi2 < matchNames.length; mi2++) {
                                var mn2 = matchNames[mi2];
                                var fxObj = null;
                                // Tenta múltiplos métodos
                                var lookups = ['getVideoEffectByMatchName', 'getVideoEffectByName'];
                                for (var lkp = 0; lkp < lookups.length && !fxObj; lkp++) {
                                    try {
                                        if (typeof qe.project[lookups[lkp]] === 'function') {
                                            fxObj = qe.project[lookups[lkp]](mn2);
                                        }
                                    } catch(eLk) {}
                                }
                                if (!fxObj) {
                                    // Tenta nome humanizado (remove "AE.ADBE " prefix)
                                    var human = mn2.replace(/^AE\.ADBE\s*/, '').replace(/^ADBE\s*/, '');
                                    try { fxObj = qe.project.getVideoEffectByName(human); } catch(eH) {}
                                    if (fxObj) _lwDbgPreset('    found via humanized: ' + human);
                                }
                                if (!fxObj) { _lwDbgPreset('    NO_FX: ' + mn2); continue; }
                                for (var sci = 0; sci < sel.length; sci++) {
                                    try {
                                        sel[sci].addVideoEffect(fxObj);
                                        addedCount++;
                                        if (mn2 === 'AE.ADBE Motion') motionApplied = true;
                                    } catch(eAdd) {
                                        _lwDbgPreset('    addEffect err for ' + mn2 + ': ' + eAdd);
                                    }
                                }
                                _lwDbgPreset('    ✓ added: ' + mn2);
                            }

                            // ─── PARSE VALORES: pra cada componente, tenta extrair e setar valores ───
                            // Premiere preset tem <Properties> com valores serializados.
                            // Pra clip.components[0] (Motion built-in), aplica params do preset XML.
                            if (stdSel.length > 0) {
                                _lwDbgPreset('  → tentando aplicar VALORES do preset (Motion/builtin)');
                                try {
                                    // Procura por <ScaleH>, <Position>, etc no preset window
                                    var paramPatterns = {
                                        'Scale': /<Scale[^>]*>([\d.\-eE]+)<\/Scale>/,
                                        'ScaleX': /<ScaleX[^>]*>([\d.\-eE]+)<\/ScaleX>/,
                                        'ScaleY': /<ScaleY[^>]*>([\d.\-eE]+)<\/ScaleY>/,
                                        'Rotation': /<Rotation[^>]*>([\d.\-eE]+)<\/Rotation>/,
                                        'Anchor': /<Anchor[^>]*>([\d.,\-eE\s]+)<\/Anchor>/,
                                        'Position': /<Position[^>]*>([\d.,\-eE\s]+)<\/Position>/,
                                        'Opacity': /<Opacity[^>]*>([\d.\-eE]+)<\/Opacity>/,
                                    };
                                    var foundParams = {};
                                    for (var pn in paramPatterns) {
                                        var pm = paramPatterns[pn].exec(window);
                                        if (pm) {
                                            foundParams[pn] = pm[1];
                                            _lwDbgPreset('    param ' + pn + ' = ' + pm[1]);
                                        }
                                    }

                                    // Aplica em cada clip selecionado
                                    if (Object.keys ? Object.keys(foundParams).length > 0 : true) {
                                        for (var sti = 0; sti < stdSel.length; sti++) {
                                            try {
                                                var c0 = stdSel[sti].components;
                                                if (!c0 || !c0.numItems) continue;
                                                _lwDbgPreset('    clip[' + sti + '] has ' + c0.numItems + ' components');
                                                // Lista components
                                                for (var ci3 = 0; ci3 < c0.numItems; ci3++) {
                                                    var cmp = c0[ci3];
                                                    var cmpName = '';
                                                    try { cmpName = cmp.displayName || cmp.matchName || '?'; } catch(eCn) {}
                                                    _lwDbgPreset('      cmp[' + ci3 + ']: ' + cmpName);
                                                    // Lista properties do component
                                                    try {
                                                        var props = cmp.properties;
                                                        if (props && props.numItems) {
                                                            for (var pi = 0; pi < Math.min(8, props.numItems); pi++) {
                                                                var pp = props[pi];
                                                                var ppn = '';
                                                                try { ppn = pp.displayName || pp.name || '?'; } catch(ePpn) {}
                                                                _lwDbgPreset('        prop[' + pi + ']: ' + ppn);
                                                            }
                                                        }
                                                    } catch(ePps) {}
                                                }
                                            } catch(eCC) { _lwDbgPreset('    components inspect err: ' + eCC); }
                                        }
                                    }
                                } catch(eParam) { _lwDbgPreset('  param parse err: ' + eParam); }
                            }

                            if (addedCount > 0) {
                                applied = sel.length;
                                _lwDbgPreset('  ✓ REGEX strategy DONE: addedCount=' + addedCount + ' on ' + sel.length + ' clips, total=' + ((new Date()).getTime() - t0) + 'ms');
                            } else {
                                _lwDbgPreset('  ✗ Nenhum efeito custom — preset só altera builtin (Motion). Valores precisam ser aplicados manualmente por enquanto.');
                            }
                        }
                        // SKIP o XML() parse antigo (era muito lento)
                        var SKIP_OLD_XML = true;
                        if (SKIP_OLD_XML) { /* old code abaixo dentro de else inalcançável */ } else {
                        // Remove a declaração <?xml ...?> e BOM se houver
                        var xmlClean = xmlRaw.replace(/^﻿/, '').replace(/^\s*<\?xml[^?]+\?>\s*/, '');
                        var xmlObj = null;
                        try { xmlObj = new XML(xmlClean); } catch(eX) {
                            _lwDbgPreset('  XML parse err: ' + (eX.toString ? eX.toString() : eX));
                        }

                        if (xmlObj) {
                            // Procura nó com <Name>presetName</Name> dentro do XML.
                            // Estrutura varia (EffectPresetItem, BinTreeItem com tipo, etc).
                            // Vamos walk recursivo e achar TODOS os MatchName dentro do nó com Name match.
                            var foundPresetNode = null;
                            function _walkFindPreset(node, depth) {
                                if (foundPresetNode || depth > 20) return;
                                try {
                                    // Check Name child do node
                                    var nameProp = node.Name;
                                    if (nameProp && nameProp.length() > 0) {
                                        var nameStr = String(nameProp);
                                        if (nameStr === presetName) {
                                            foundPresetNode = node;
                                            return;
                                        }
                                    }
                                    // Tenta @ Name attribute também
                                    try {
                                        var nameAttr = node.attribute('Name');
                                        if (nameAttr && nameAttr.length() > 0 && String(nameAttr) === presetName) {
                                            foundPresetNode = node;
                                            return;
                                        }
                                    } catch(eNA) {}
                                    var children = node.children();
                                    var nch = children.length();
                                    for (var i = 0; i < nch && !foundPresetNode; i++) {
                                        _walkFindPreset(children[i], depth + 1);
                                    }
                                } catch(eW) {}
                            }
                            _walkFindPreset(xmlObj, 0);
                            _lwDbgPreset('  presetNode found in XML: ' + (foundPresetNode ? 'YES' : 'NO'));

                            // Se não achou pelo Name child, tenta full-text search por presetName
                            // e usa o container pai como nó do preset
                            if (!foundPresetNode) {
                                _lwDbgPreset('  fallback: scanning XML for preset name text...');
                                // Walk tudo, registra todos os <Name> que tem valor parecido (debug)
                                var allNames = [];
                                function _collectNames(node, depth) {
                                    if (depth > 20) return;
                                    try {
                                        var nameProp = node.Name;
                                        if (nameProp && nameProp.length() > 0) {
                                            var ns = String(nameProp);
                                            if (ns) allNames.push(ns);
                                        }
                                        var children = node.children();
                                        var nch = children.length();
                                        for (var i = 0; i < nch; i++) _collectNames(children[i], depth + 1);
                                    } catch(eN) {}
                                }
                                _collectNames(xmlObj, 0);
                                _lwDbgPreset('  total <Name> values in XML: ' + allNames.length);
                                // Mostra os primeiros 10
                                for (var nai = 0; nai < Math.min(10, allNames.length); nai++) {
                                    _lwDbgPreset('    Name[' + nai + ']: ' + allNames[nai]);
                                }
                                // Tenta match insensitive ou parcial
                                var presetNameLower = String(presetName).toLowerCase();
                                function _walkFindPresetLoose(node, depth) {
                                    if (foundPresetNode || depth > 20) return;
                                    try {
                                        var nameProp = node.Name;
                                        if (nameProp && nameProp.length() > 0) {
                                            var ns = String(nameProp).toLowerCase();
                                            if (ns === presetNameLower) {
                                                foundPresetNode = node;
                                                return;
                                            }
                                        }
                                        var children = node.children();
                                        var nch = children.length();
                                        for (var i = 0; i < nch && !foundPresetNode; i++) {
                                            _walkFindPresetLoose(children[i], depth + 1);
                                        }
                                    } catch(eL) {}
                                }
                                _walkFindPresetLoose(xmlObj, 0);
                                _lwDbgPreset('  loose match: ' + (foundPresetNode ? 'YES' : 'NO'));
                            }

                            if (foundPresetNode) {
                                // DUMP: estrutura do preset node pra entender XML schema do .prfpset
                                try {
                                    var dumpXml = String(foundPresetNode.toXMLString ? foundPresetNode.toXMLString() : foundPresetNode);
                                    _lwDbgPreset('  presetNode XML (1500 chars):');
                                    // Quebra em pedaços de 200 chars pra log
                                    for (var dx = 0; dx < Math.min(1500, dumpXml.length); dx += 200) {
                                        _lwDbgPreset('    >> ' + dumpXml.substr(dx, 200));
                                    }
                                } catch(eDmp) { _lwDbgPreset('  dump err: ' + eDmp); }

                                // Coleta TODOS os valores de texto que parecem identificadores de efeito.
                                // .prfpset XML usa schemas variados — não dá pra confiar em <MatchName> só.
                                // Estratégia: coleta TODO leaf de texto, e detecta padrões de matchName.
                                var effectMatchNames = [];
                                var effectDisplayNames = [];
                                var allTexts = []; // pra debug
                                var elementNameCounts = {}; // pra ver schema

                                function _collectFxNames(node, depth) {
                                    if (depth > 30) return;
                                    try {
                                        var elName = '';
                                        try { elName = node.name() ? String(node.name()) : ''; } catch(eNm) {}
                                        // Conta elementos pra entender schema
                                        if (elName) elementNameCounts[elName] = (elementNameCounts[elName] || 0) + 1;

                                        // Se for um text node ou element folha com texto
                                        var children = node.children();
                                        var nch = children.length();
                                        if (nch === 0 || (nch === 1 && String(children[0].name()) === '')) {
                                            // Leaf — pega valor
                                            var val = String(node);
                                            if (val && val.length < 200) {
                                                allTexts.push({el: elName, val: val});
                                                // Detecta padrão de matchName: "AE.ADBE..." ou "ADBE..." ou contém "AE."
                                                if (/^AE\.|^ADBE |^AE_|ADBE/.test(val)) {
                                                    effectMatchNames.push(val);
                                                }
                                                // Heurística: nome de efeito é tipicamente curto sem "AE." prefix
                                                // Vamos guardar nomes potenciais — qualquer string que pareça nome de efeito
                                                if (/^(EffectName|Name|Title|DisplayName)$/i.test(elName)) {
                                                    if (val && val !== presetName) effectDisplayNames.push(val);
                                                }
                                            }
                                        }
                                        for (var i = 0; i < nch; i++) _collectFxNames(children[i], depth + 1);
                                    } catch(eC) {}
                                }
                                _collectFxNames(foundPresetNode, 0);
                                // Log do schema descoberto
                                var schemaStr = '';
                                for (var ek in elementNameCounts) schemaStr += ek + '=' + elementNameCounts[ek] + ' ';
                                _lwDbgPreset('  preset elements schema: ' + schemaStr.substr(0, 400));
                                _lwDbgPreset('  preset has matchNames=' + effectMatchNames.length + ' displayNames=' + effectDisplayNames.length + ' totalTexts=' + allTexts.length);
                                for (var fmi = 0; fmi < Math.min(20, effectMatchNames.length); fmi++) {
                                    _lwDbgPreset('    fx mn[' + fmi + ']: ' + effectMatchNames[fmi]);
                                }
                                for (var fdi = 0; fdi < Math.min(20, effectDisplayNames.length); fdi++) {
                                    _lwDbgPreset('    fx dn[' + fdi + ']: ' + effectDisplayNames[fdi]);
                                }
                                // Mostra primeiros 10 texts pra ver outros campos
                                for (var ti10 = 0; ti10 < Math.min(15, allTexts.length); ti10++) {
                                    _lwDbgPreset('    text[' + allTexts[ti10].el + ']: ' + allTexts[ti10].val.substr(0, 80));
                                }

                                // ═══ Mapeia matchName → display name conhecidos ═══
                                // Premiere expõe efeitos por nome local (português, inglês...) — getVideoEffectByName usa display
                                var fxNames = effectDisplayNames.concat(effectMatchNames);
                                // Dedup
                                var seen = {};
                                var fxUnique = [];
                                for (var fu = 0; fu < fxNames.length; fu++) {
                                    if (!seen[fxNames[fu]]) { seen[fxNames[fu]] = 1; fxUnique.push(fxNames[fu]); }
                                }

                                // Efeitos "builtin" do clip (Motion, Opacity, Time Remapping) NÃO precisam ser adicionados — já existem
                                var builtinSkip = {'Motion':1, 'Opacity':1, 'Time Remapping':1, 'AE.ADBE Motion':1, 'AE.ADBE Opacity':1, 'AE.ADBE Timecode':1, 'AE.ADBE MarkerParam':1};

                                var addedCount = 0;
                                for (var fxi = 0; fxi < fxUnique.length; fxi++) {
                                    var fxName = fxUnique[fxi];
                                    if (builtinSkip[fxName]) { _lwDbgPreset('    skip builtin: ' + fxName); continue; }
                                    var fxObj = null;
                                    // Tenta múltiplos métodos de lookup (alguns aceitam matchName, outros displayName)
                                    var lookupMethods = ['getVideoEffectByName', 'getVideoEffectByMatchName', 'getVideoEffectByDisplayName'];
                                    for (var lmi = 0; lmi < lookupMethods.length && !fxObj; lmi++) {
                                        try {
                                            if (typeof qe.project[lookupMethods[lmi]] === 'function') {
                                                fxObj = qe.project[lookupMethods[lmi]](fxName);
                                                if (fxObj) _lwDbgPreset('    fx found via qe.project.' + lookupMethods[lmi] + '("' + fxName + '")');
                                            }
                                        } catch(eLM) {}
                                    }
                                    if (!fxObj && selKind === 'audio') {
                                        try { fxObj = qe.project.getAudioEffectByName(fxName); } catch(eGA) {}
                                    }
                                    if (!fxObj) { _lwDbgPreset('    NO_FX_OBJ for: ' + fxName); continue; }
                                    // Aplica em todos os clips selecionados
                                    for (var sci = 0; sci < sel.length; sci++) {
                                        try {
                                            if (selKind === 'audio') sel[sci].addAudioEffect(fxObj);
                                            else sel[sci].addVideoEffect(fxObj);
                                            addedCount++;
                                        } catch(eAdd) { _lwDbgPreset('    addEffect err: ' + eAdd); }
                                    }
                                    _lwDbgPreset('    ✓ added: ' + fxName);
                                }
                                if (addedCount > 0) {
                                    applied = sel.length;
                                    _lwDbgPreset('  ✓ XML strategy: addedCount=' + addedCount + ' on ' + sel.length + ' clip(s)');
                                } else {
                                    _lwDbgPreset('  ✗ XML strategy: no effects could be added');
                                }
                            }
                        }
                        } // close else do SKIP_OLD_XML
                    }
                } else {
                    _lwDbgPreset('  could not open preset file for read');
                }
            } catch(eXmlAll) {
                _lwDbgPreset('  XML strategy outer err: ' + eXmlAll);
            }
        }

        // Se XML strategy aplicou, retorna sucesso
        if (applied > 0) {
            _lwDbgPreset('═══ DONE: applied=' + applied + ' (via XML parse) ═══');
            return 'OK:applied:' + applied;
        }

        // ─── Procura item já existente no projeto (NÃO importa, evita dialog "File format not supported") ───
        // Containers .prfpset NÃO podem ser importados via importFiles() — sempre dão erro.
        // Em vez disso: procura entre items já no projeto, OU usa QE pra buscar o preset internamente.
        var importedProjectItem = null;
        if (presetName) {
            try {
                importedProjectItem = _lwFindProjectItemByName(app.project.rootItem, presetName, 0);
                _lwDbgPreset('  projItem search by name "' + presetName + '" = ' + (importedProjectItem ? 'FOUND' : 'not in project'));
            } catch(eF) { _lwDbgPreset('  find err: ' + eF); }
        }

        // ─── Inspeção QE: dump métodos pra descobrir API de preset ───
        var qePresetObj = null;
        try {
            var qeKeys = '';
            for (var qk in qe.project) qeKeys += qk + ',';
            _lwDbgPreset('  qe.project keys: ' + qeKeys.substr(0, 400));
        } catch(eQK) { _lwDbgPreset('  qe keys err: ' + eQK); }

        // Tenta APIs conhecidas de QE pra obter o preset object
        var qePresetMethods = [
            'getEffectPresetByName',
            'getPresetByName',
            'getVideoEffectPresetByName',
            'getAudioEffectPresetByName',
            'getMotionPresetByName',
        ];
        for (var qpm = 0; qpm < qePresetMethods.length; qpm++) {
            try {
                if (typeof qe.project[qePresetMethods[qpm]] === 'function') {
                    var qpRes = qe.project[qePresetMethods[qpm]](presetName);
                    _lwDbgPreset('  qe.project.' + qePresetMethods[qpm] + '("' + presetName + '") = ' + (qpRes ? 'OBJECT' : 'null'));
                    if (qpRes && !qePresetObj) qePresetObj = qpRes;
                }
            } catch(eQpm) { _lwDbgPreset('  qe.' + qePresetMethods[qpm] + ' err: ' + eQpm); }
        }

        // Dump métodos do primeiro QE clip pra ver opções de preset
        if (sel.length > 0) {
            try {
                var qcKeys = '';
                for (var qck in sel[0]) qcKeys += qck + ',';
                _lwDbgPreset('  qe clip[0] keys: ' + qcKeys.substr(0, 400));
            } catch(eQCK) {}
        }

        // ─── Tenta TODAS as combinações de método × tipo de argumento × tipo de clip ───
        var clipTargets = [
            { obj: stdSel, label: 'std' },
            { obj: sel, label: 'qe' },
        ];

        var methodNames = [
            'applyEffectPreset',
            'applyPreset',
            'applyPresetByName',
            'loadEffectPreset',
            'addEffectPreset',
            'setEffectPreset',
            'applyEffectPresetByName',
        ];

        var argSets = [];
        if (qePresetObj) argSets.push({ args: [qePresetObj], label: 'qePresetObj' });
        if (importedProjectItem) argSets.push({ args: [importedProjectItem], label: 'projectItem' });
        if (pf) argSets.push({ args: [pf], label: 'File' });
        if (presetPath) argSets.push({ args: [presetPath], label: 'pathStr' });
        if (presetName) argSets.push({ args: [presetName], label: 'nameStr' });
        // Adiciona métodos QE-específicos pra cobrir addEffectPreset + similares
        var methodNamesExtended = [
            'applyEffectPreset',
            'applyPreset',
            'applyPresetByName',
            'loadEffectPreset',
            'addEffectPreset',
            'setEffectPreset',
            'applyEffectPresetByName',
            'addVideoEffectPreset',
            'addAudioEffectPreset',
            'addPreset',
        ];
        methodNames = methodNamesExtended;

        for (var ti2 = 0; ti2 < clipTargets.length && applied === 0; ti2++) {
            var target = clipTargets[ti2];
            for (var mi = 0; mi < methodNames.length && applied === 0; mi++) {
                for (var ai = 0; ai < argSets.length && applied === 0; ai++) {
                    var argSet = argSets[ai];
                    var label = target.label + '.' + methodNames[mi] + '(' + argSet.label + ')';
                    var anyOk = false;
                    for (var ci2 = 0; ci2 < target.obj.length; ci2++) {
                        var r = _lwTryCall(target.obj[ci2], methodNames[mi], argSet.args, label);
                        if (r.ok) { anyOk = true; }
                        else { allErrs.push(r.err); }
                    }
                    if (anyOk) {
                        applied = target.obj.length;
                        _lwDbgPreset('  ✓ ' + label + ' SUCCESS — applied ' + applied);
                        break;
                    }
                }
            }
        }

        // Se nada funcionou, dump TODOS os erros pra ver pattern
        if (applied === 0) {
            _lwDbgPreset('  ✗ ALL FAILED');
            for (var ei = 0; ei < Math.min(20, allErrs.length); ei++) {
                _lwDbgPreset('    err[' + ei + ']: ' + allErrs[ei]);
            }
        }

        _lwDbgPreset('═══ DONE: applied=' + applied + ' total errs=' + allErrs.length + ' ═══');
        if (applied === 0) {
            // Retorna primeiros 2 erros únicos pra usuário ver
            var uniqErrs = {};
            var firstErrs = [];
            for (var ue = 0; ue < allErrs.length && firstErrs.length < 2; ue++) {
                var ek = allErrs[ue].split(':')[0];
                if (!uniqErrs[ek]) { uniqErrs[ek] = 1; firstErrs.push(allErrs[ue]); }
            }
            return 'ERR:preset_apply_unsupported:' + firstErrs.join(' | ');
        }
        return 'OK:applied:' + applied;
    } catch (e) {
        _lwDbgPreset('OUTER ERR: ' + e);
        return 'ERR:' + e.toString().slice(0, 200);
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — INSERT AUDIO SOURCE (gap finder + nova track se preciso)
// ═══════════════════════════════════════════════════════════════════
// Helper: BUFFERED debug — acumula em array, flush 1x no final (era 20+ I/Os = ~500ms perdidos)
var _lwAudioDbgBuf = [];
function _lwDbgAudio(msg) {
    _lwAudioDbgBuf.push('[' + (new Date()).toString() + '] ' + msg);
}
function _lwFlushAudioDbg() {
    if (_lwAudioDbgBuf.length === 0) return;
    var blob = _lwAudioDbgBuf.join('\n') + '\n';
    _lwAudioDbgBuf = [];
    try {
        var f = new File(Folder.temp.fsName + '/lion-search-audio-insert.txt');
        if (f.open('a')) { f.encoding = 'UTF-8'; f.write(blob); f.close(); }
    } catch(e) {}
    try {
        var fd = new File(Folder.desktop.fsName + '/lion-audio-debug.txt');
        if (fd.open('a')) { fd.encoding = 'UTF-8'; fd.write(blob); fd.close(); }
    } catch(eD) {}
}

// Nova função: insere áudio da Audio Library com trim + volume dB opcionais.
// inSec/outSec = -1 → usa duração total. volumeDb = 0 → sem ajuste.
function lwInsertAudioWithTrim(filePath, inSec, outSec, volumeDb) {
    try {
        var nodeIdOrName = 'FS:' + filePath;
        _lwLibTrimIn = (inSec != null && inSec >= 0) ? parseFloat(inSec) : null;
        _lwLibTrimOut = (outSec != null && outSec >= 0) ? parseFloat(outSec) : null;
        _lwLibVolumeDb = (volumeDb != null && !isNaN(parseFloat(volumeDb)) && Math.abs(parseFloat(volumeDb)) >= 0.01) ? parseFloat(volumeDb) : null;
        var result = _lwInsertAudioSource(nodeIdOrName);
        _lwLibTrimIn = null;
        _lwLibTrimOut = null;
        _lwLibVolumeDb = null;
        return result;
    } catch(eW) {
        _lwLibTrimIn = null; _lwLibTrimOut = null; _lwLibVolumeDb = null;
        return 'ERR:lib_insert:' + eW.toString().slice(0, 200);
    }
}
// Globals pra passar trim/volume do entrypoint pra _lwInsertAudioSourceImpl sem mudar assinatura
var _lwLibTrimIn = null;
var _lwLibTrimOut = null;
var _lwLibVolumeDb = null;

// Wrapper que SEMPRE flusha debug no final (sucesso ou erro) — 1 I/O em vez de ~20
function _lwInsertAudioSource(nodeIdOrName) {
    var result;
    try {
        result = _lwInsertAudioSourceImpl(nodeIdOrName);
    } catch(eW) {
        result = 'ERR:wrapper:' + eW.toString().slice(0, 200);
    }
    try { _lwFlushAudioDbg(); } catch(eF) {}
    return result;
}

function _lwInsertAudioSourceImpl(nodeIdOrName) {
    try {
        _lwDbgAudio('═══ Insert Audio: ' + nodeIdOrName + ' ═══');
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        var seq = app.project.activeSequence;

        var foundItem = null;
        var preImportedFromFs = false;

        // Se for FS path (path do arquivo), importa direto ou busca no projeto ativo
        if (nodeIdOrName.indexOf('FS:') === 0) {
            var fsPath = nodeIdOrName.substr(3);
            _lwDbgAudio('  FS path: ' + fsPath);

            // Normaliza path pra comparação tolerante (URL-encoding, slashes)
            function _normP(p) {
                if (!p) return '';
                var s = String(p).replace(/^file:\/+/, '').replace(/\\/g, '/').toLowerCase();
                try { s = decodeURIComponent(s); } catch(e) {}
                return s;
            }
            var fsPathNorm = _normP(fsPath);
            _lwDbgAudio('  fsPathNorm: ' + fsPathNorm);

            // Busca no projeto ativo PRIMEIRO (mais provável, já importado)
            function _findFsByPath(bin, holder, depth) {
                if (holder.found || depth > 12) return;
                try {
                    if (!bin.children || bin.children.numItems === 0) return;
                    for (var ii = 0; ii < bin.children.numItems; ii++) {
                        if (holder.found) return;
                        var it = bin.children[ii];
                        var ch = null;
                        try { ch = it.children; } catch(eC) {}
                        if (ch && ch.numItems > 0) { _findFsByPath(it, holder, depth + 1); }
                        else {
                            var p = '';
                            try { p = it.getMediaPath(); } catch(eM) {}
                            if (_normP(p) === fsPathNorm) { holder.found = it; return; }
                        }
                    }
                } catch(e) {}
            }
            // SEMPRE verifica que arquivo existe no disco ANTES de tentar inserir
            // (mesmo que item esteja no projeto, file pode estar offline → insere offline)
            var fsResolved = _lwResolveExistingPath(fsPath);
            if (!fsResolved) {
                _lwDbgAudio('  ✗ FILE OFFLINE — arquivo não existe no disco: ' + fsPath);
                return 'ERR:sfx_file_not_on_disk:' + fsPath;
            }
            fsPath = fsResolved;
            fsPathNorm = _normP(fsPath);

            var holder = { found: null };
            _findFsByPath(app.project.rootItem, holder, 0);
            if (holder.found) {
                _lwDbgAudio('  found in active project (no import needed)');
                foundItem = holder.found;
                preImportedFromFs = true; // já está no projeto ativo, pula walk de parent
            } else {
                _lwDbgAudio('  not in active project — importing');
                // Check tolerante (URL-encoding etc)
                var resolvedFsPath = _lwResolveExistingPath(fsPath);
                if (!resolvedFsPath) {
                    _lwDbgAudio('  RESOLVED PATH = null (arquivo não existe)');
                    return 'ERR:sfx_file_not_on_disk:' + fsPath;
                }
                _lwDbgAudio('  resolvedPath: ' + resolvedFsPath);
                fsPath = resolvedFsPath;
                fsPathNorm = _normP(fsPath);
                // Importa com suppressUI=true
                var importOk = false;
                try { importOk = app.project.importFiles([fsPath], true); } catch(eI) { _lwDbgAudio('  import err: ' + eI); }
                _lwDbgAudio('  importFiles=' + importOk);
                if (!importOk) return 'ERR:sfx_import_failed:' + fsPath;
                // Polling progressivo — fast inicial, slow se demorar (até 1.5s total)
                // Premiere às vezes leva 300-800ms pra processar import de arquivo grande.
                holder = { found: null };
                var pollT0 = (new Date()).getTime();
                // 5 fast tries (10ms) + 15 medium (40ms) + 10 slow (60ms) = total ~1.25s
                var pollSchedule = [10,10,10,10,10, 40,40,40,40,40,40,40,40,40,40, 60,60,60,60,60,60,60,60,60,60];
                for (var pollN = 0; pollN < pollSchedule.length; pollN++) {
                    _findFsByPath(app.project.rootItem, holder, 0);
                    if (holder.found) break;
                    $.sleep(pollSchedule[pollN]);
                }
                _lwDbgAudio('  poll: ' + ((new Date()).getTime() - pollT0) + 'ms, tries=' + (pollN+1) + ' found=' + (holder.found ? 'YES' : 'NO'));
                if (!holder.found) {
                    // Última tentativa com fallback: pega o último leaf adicionado (newest)
                    _lwDbgAudio('  fallback: looking for newest leaf in project');
                    var newest = null;
                    function _getLastLeafFs(bin, depth) {
                        if (!bin || depth > 12 || newest) return;
                        try {
                            if (!bin.children || bin.children.numItems === 0) return;
                            for (var li = bin.children.numItems - 1; li >= 0 && !newest; li--) {
                                var it5 = bin.children[li];
                                var ch5 = null;
                                try { ch5 = it5.children; } catch(eCh5) {}
                                if (ch5 && ch5.numItems > 0) _getLastLeafFs(it5, depth + 1);
                                else { newest = it5; return; }
                            }
                        } catch(e) {}
                    }
                    _getLastLeafFs(app.project.rootItem, 0);
                    if (newest) {
                        _lwDbgAudio('  fallback found newest leaf: ' + (function(){try{return newest.name;}catch(e){return '?';}})());
                        holder.found = newest;
                    } else {
                        _lwDbgAudio('  imported but not found by path AND no leaves');
                        return 'ERR:sfx_imported_not_found:' + fsPath;
                    }
                }
                _lwDbgAudio('  imported + found');
                foundItem = holder.found;
                preImportedFromFs = true;
            }
        }

        // ─── nodeId/name path: identifier NÃO é FS path ───
        if (!foundItem) {
            _lwDbgAudio('  → nodeId search (não-FS): "' + nodeIdOrName + '"');
        }

        // Se não veio de FS, acha o project item — procura em TODOS os projetos abertos
        function _findInBin(bin) {
            if (foundItem) return;
            try {
                if (!bin.children || bin.children.numItems === 0) return;
                for (var i = 0; i < bin.children.numItems; i++) {
                    if (foundItem) return;
                    var it = bin.children[i];
                    var ch = null;
                    try { ch = it.children; } catch(eC) {}
                    if (ch && ch.numItems > 0) { _findInBin(it); }
                    else {
                        var nid = '';
                        try { nid = it.nodeId || it.name || ''; } catch(eN) {}
                        var nm = '';
                        try { nm = it.name || ''; } catch(eM) {}
                        if (nid === nodeIdOrName || nm === nodeIdOrName) { foundItem = it; return; }
                    }
                }
            } catch(e) {}
        }

        // Coleta todos os projetos abertos
        var projects = [];
        try {
            if (app.openDocuments) {
                var nDocs = app.openDocuments.numItems || app.openDocuments.length || 0;
                for (var d = 0; d < nDocs; d++) {
                    var doc = null;
                    try { doc = app.openDocuments[d]; } catch(eDD) {}
                    if (!doc) try { doc = app.openDocuments.getItemAt(d); } catch(eDD2) {}
                    if (doc) projects.push(doc);
                }
            }
        } catch(eOd) {}
        try {
            if (app.projects && projects.length === 0) {
                var nProj = app.projects.numProjects || app.projects.length || 0;
                for (var ip = 0; ip < nProj; ip++) {
                    var pr = null;
                    try { pr = app.projects[ip]; } catch(ePj) {}
                    if (pr) projects.push(pr);
                }
            }
        } catch(eP) {}
        if (projects.length === 0) projects.push(app.project); // fallback

        _lwDbgAudio('  found ' + projects.length + ' open projects');

        // Procura primeiro no projeto ativo (mais provável + insertClip funciona)
        try { _findInBin(app.project.rootItem); } catch(eA) {}
        _lwDbgAudio('  search in active project: ' + (foundItem ? 'FOUND' : 'not found'));

        // Se não achou, tenta nos outros projetos
        if (!foundItem) {
            for (var pp = 0; pp < projects.length && !foundItem; pp++) {
                var pr2 = projects[pp];
                if (pr2 === app.project) continue; // já tentou
                try { _findInBin(pr2.rootItem); } catch(eB) {}
                _lwDbgAudio('  search in proj[' + pp + ']: ' + (foundItem ? 'FOUND' : 'not found'));
            }
        }

        if (!foundItem) {
            _lwDbgAudio('  ✗ audio_source_not_found: nodeId="' + nodeIdOrName + '" — catalog stale? try reabrir LION SEARCH');
            return 'ERR:audio_source_not_found:' + nodeIdOrName;
        }
        _lwDbgAudio('  itemName=' + (function(){try{return foundItem.name;}catch(e){return '?';}})());

        // Se o item já foi achado via FS path no projeto ativo (preImportedFromFs)
        // OU veio do projeto ativo via nodeId search, PULA o walk de parent (que trava em offline).
        // Pra item de OUTRO projeto (preImportedFromFs=false e veio de outro project), faz import.
        var itemBelongsToActive = preImportedFromFs; // FS no projeto ativo já confirma pertence
        if (!itemBelongsToActive) {
            try {
                // Walk up parents — se chegar no rootItem do app.project, pertence
                var p = foundItem;
                for (var pw = 0; pw < 20 && p; pw++) {
                    if (p === app.project.rootItem) { itemBelongsToActive = true; break; }
                    try { p = p.parent; } catch(eP) { break; }
                    if (!p) break;
                }
            } catch(eBel) {}
        }
        _lwDbgAudio('  itemBelongsToActive=' + itemBelongsToActive + ' (preImportedFromFs=' + preImportedFromFs + ')');

        if (!itemBelongsToActive) {
            var mediaPath = '';
            try { mediaPath = foundItem.getMediaPath(); } catch(eMp) {}
            if (!mediaPath) {
                try { mediaPath = foundItem.canonicalUri || ''; } catch(eC) {}
            }
            if (!mediaPath) return 'ERR:could_not_get_path_other_project';

            // Check tolerante: getMediaPath() pode retornar URL-encoded (%20 pra espaços),
            // forward slashes, etc — File.exists falha. Tenta múltiplas variações.
            var resolvedPath = _lwResolveExistingPath(mediaPath);
            if (!resolvedPath) {
                return 'ERR:file_not_on_disk:' + mediaPath;
            }
            mediaPath = resolvedPath; // usa o path que funcionou

            function _normPath(p) {
                if (!p) return '';
                var s = String(p);
                try { s = decodeURIComponent(s); } catch(eD) {}
                s = s.replace(/\\/g, '/').toLowerCase();
                s = s.replace(/^file:\/\/+/, '');
                return s;
            }
            var targetNorm = _normPath(mediaPath);
            var targetFilename = '';
            try {
                var slashIdx = Math.max(targetNorm.lastIndexOf('/'), targetNorm.lastIndexOf('\\'));
                targetFilename = slashIdx >= 0 ? targetNorm.substr(slashIdx + 1) : targetNorm;
            } catch(eFn) {}

            function _countLeaves(bin, depth) {
                if (!bin || depth > 10) return 0;
                var n = 0;
                try {
                    if (!bin.children || bin.children.numItems === 0) return 0;
                    for (var ic = 0; ic < bin.children.numItems; ic++) {
                        var ck = bin.children[ic];
                        var ck2 = null;
                        try { ck2 = ck.children; } catch(eCh) {}
                        if (ck2 && ck2.numItems > 0) n += _countLeaves(ck, depth + 1);
                        else n++;
                    }
                } catch(e) {}
                return n;
            }
            var leavesBefore = _countLeaves(app.project.rootItem, 0);

            // Importa com suppressUI=true SEMPRE pra não mostrar dialog Premiere
            var importedOk = false;
            try {
                importedOk = app.project.importFiles([mediaPath], true, app.project.getInsertionBin(), false);
            } catch(eImp) {}
            if (!importedOk) {
                // Segunda tentativa também com suppressUI=true
                try {
                    importedOk = app.project.importFiles([mediaPath], true);
                } catch(eImp2) { return 'ERR:import_failed:' + eImp2.toString().slice(0, 100); }
            }

            // Poll rápido em vez de sleep fixo — quebra assim que Premiere processou
            var preLeaves = leavesBefore;
            for (var pN = 0; pN < 10; pN++) {
                try {
                    var nowLeaves = _countLeaves(app.project.rootItem, 0);
                    if (nowLeaves > preLeaves) break;
                } catch(eL) {}
                $.sleep(20);
            }

            // Busca o item — 3 estratégias
            var newItem = null;

            // Estratégia 1: match por path normalizado
            function _findByPathNorm(bin, depth) {
                if (newItem || depth > 10) return;
                try {
                    if (!bin.children || bin.children.numItems === 0) return;
                    for (var ii = 0; ii < bin.children.numItems; ii++) {
                        if (newItem) return;
                        var it2 = bin.children[ii];
                        var ch2 = null;
                        try { ch2 = it2.children; } catch(eC2) {}
                        if (ch2 && ch2.numItems > 0) { _findByPathNorm(it2, depth + 1); }
                        else {
                            var pp2 = '';
                            try { pp2 = it2.getMediaPath(); } catch(eMp2) {}
                            if (_normPath(pp2) === targetNorm) { newItem = it2; return; }
                        }
                    }
                } catch(eF) {}
            }
            _findByPathNorm(app.project.rootItem, 0);

            // Estratégia 2: match por filename + path contém targetFilename
            if (!newItem && targetFilename) {
                function _findByFilename(bin, depth) {
                    if (newItem || depth > 10) return;
                    try {
                        if (!bin.children || bin.children.numItems === 0) return;
                        for (var ii2 = 0; ii2 < bin.children.numItems; ii2++) {
                            if (newItem) return;
                            var it3 = bin.children[ii2];
                            var ch3 = null;
                            try { ch3 = it3.children; } catch(eCh3) {}
                            if (ch3 && ch3.numItems > 0) { _findByFilename(it3, depth + 1); }
                            else {
                                var nm3 = '';
                                try { nm3 = String(it3.name).toLowerCase(); } catch(eNm) {}
                                var pp3 = '';
                                try { pp3 = _normPath(it3.getMediaPath()); } catch(eMp3) {}
                                // Match se nome bate ou path contém filename
                                if (nm3 === targetFilename || nm3 + '.' + (it3.type || '') === targetFilename
                                    || (pp3 && pp3.indexOf(targetFilename) >= 0)) {
                                    newItem = it3;
                                    return;
                                }
                            }
                        }
                    } catch(eF2) {}
                }
                _findByFilename(app.project.rootItem, 0);
            }

            // Estratégia 3: pega o último leaf adicionado (newest item)
            if (!newItem) {
                var leavesAfter = _countLeaves(app.project.rootItem, 0);
                if (leavesAfter > leavesBefore) {
                    // Pega o último leaf — provavelmente é o que importamos
                    function _getLastLeaf(bin, depth) {
                        if (!bin || depth > 10) return null;
                        try {
                            if (!bin.children || bin.children.numItems === 0) return null;
                            // Itera de trás pra frente — último item costuma ser o mais novo
                            for (var li = bin.children.numItems - 1; li >= 0; li--) {
                                var it4 = bin.children[li];
                                var ch4 = null;
                                try { ch4 = it4.children; } catch(eCh4) {}
                                if (ch4 && ch4.numItems > 0) {
                                    var deep = _getLastLeaf(it4, depth + 1);
                                    if (deep) return deep;
                                } else {
                                    return it4;
                                }
                            }
                        } catch(e) {}
                        return null;
                    }
                    newItem = _getLastLeaf(app.project.rootItem, 0);
                }
            }

            if (!newItem) return 'ERR:imported_but_not_found:' + mediaPath + ' (target=' + targetFilename + ')';
            foundItem = newItem;
        }

        _lwDbgAudio('  → foundItem ready, name=' + (function(){try{return foundItem.name;}catch(e){return '?';}})() + ' type=' + (function(){try{return foundItem.type;}catch(e){return '?';}})());

        // Calcula duração — respeita in/out se setados
        var insertDurationSec = 0;
        var hasInOut = false;
        var inSec = 0, outSec = 0;
        try {
            // Tenta múltiplos mediaType pra in/out
            var inPt = null, outPt = null;
            var mediaTypes = [4, 1, 0]; // audio, video, both
            for (var mt = 0; mt < mediaTypes.length; mt++) {
                try {
                    var ip = foundItem.getInPoint(mediaTypes[mt]);
                    var op = foundItem.getOutPoint(mediaTypes[mt]);
                    if (ip && op) { inPt = ip; outPt = op; _lwDbgAudio('  inOut found via mediaType=' + mediaTypes[mt]); break; }
                } catch(eMT) {}
            }
            if (inPt && outPt) {
                inSec = parseFloat(inPt.seconds);
                outSec = parseFloat(outPt.seconds);
                _lwDbgAudio('  in/out seconds: ' + inSec + ' / ' + outSec);
                if (outSec > inSec) {
                    insertDurationSec = outSec - inSec;
                    hasInOut = true;
                }
            } else {
                _lwDbgAudio('  no in/out points found');
            }
        } catch(eIO) { _lwDbgAudio('  in/out err: ' + eIO); }

        // Se não tem in/out válido, usa duração total
        if (insertDurationSec <= 0) {
            try {
                var dur = foundItem.getOutPoint(0); // tenta full duration
                if (dur && dur.seconds > 0) insertDurationSec = parseFloat(dur.seconds);
                _lwDbgAudio('  fallback duration via getOutPoint(0): ' + insertDurationSec);
            } catch(eD) { _lwDbgAudio('  fallback duration err: ' + eD); }
        }
        if (insertDurationSec <= 0) {
            // Fallback: tenta via clip projecting.. assume 5s mínimo
            insertDurationSec = 5;
            _lwDbgAudio('  using 5s default duration');
        }
        _lwDbgAudio('  insertDurationSec=' + insertDurationSec + ' hasInOut=' + hasInOut);

        // OVERRIDE da Audio Library: se user definiu trim in/out (start/end seconds),
        // aplica setInPoint/setOutPoint no item DO PROJECT (não no clip) pra cortar.
        if (_lwLibTrimIn != null || _lwLibTrimOut != null) {
            var libIn = (_lwLibTrimIn != null && _lwLibTrimIn >= 0) ? _lwLibTrimIn : 0;
            var libOut = (_lwLibTrimOut != null && _lwLibTrimOut > 0) ? _lwLibTrimOut : insertDurationSec;
            _lwDbgAudio('  AudioLib trim: in=' + libIn + 's out=' + libOut + 's');
            try {
                // Cria Time objects pra set in/out — 1 sec = 254016000000 ticks
                var TICKS_PER_SEC = 254016000000;
                var inTime = new Time();
                inTime.ticks = String(Math.round(libIn * TICKS_PER_SEC));
                var outTime = new Time();
                outTime.ticks = String(Math.round(libOut * TICKS_PER_SEC));
                // Aplica em audio (4)
                try { foundItem.setInPoint(inTime, 4); _lwDbgAudio('    setInPoint OK'); } catch(eSI) { _lwDbgAudio('    setInPoint err: ' + eSI); }
                try { foundItem.setOutPoint(outTime, 4); _lwDbgAudio('    setOutPoint OK'); } catch(eSO) { _lwDbgAudio('    setOutPoint err: ' + eSO); }
                insertDurationSec = libOut - libIn;
                hasInOut = true;
                _lwDbgAudio('  trim applied: insertDurationSec=' + insertDurationSec);
            } catch(eLibTrim) { _lwDbgAudio('  AudioLib trim outer err: ' + eLibTrim); }
        }

        // Posição playhead
        var insertStartSec = 0;
        try {
            var pp = seq.getPlayerPosition();
            if (pp) insertStartSec = parseFloat(pp.seconds);
        } catch(ePh) { _lwDbgAudio('  playhead err: ' + ePh); }
        var insertEndSec = insertStartSec + insertDurationSec;
        _lwDbgAudio('  playhead=' + insertStartSec + 's | window=[' + insertStartSec + ',' + insertEndSec + ']');

        // Sequence info
        _lwDbgAudio('  seq.audioTracks.numTracks=' + seq.audioTracks.numTracks);

        // Encontra audio track sem overlap
        var targetTrackIdx = -1;
        var EPS = 0.001;
        for (var ti = 0; ti < seq.audioTracks.numTracks; ti++) {
            var atk = seq.audioTracks[ti];
            var hasOverlap = false;
            var clipCount = 0;
            try { clipCount = atk.clips.numItems; } catch(eCN) {}
            for (var ci = 0; ci < clipCount; ci++) {
                var c = atk.clips[ci];
                var cStart = 0, cEnd = 0;
                try { cStart = parseFloat(c.start.seconds); cEnd = parseFloat(c.end.seconds); } catch(eC) {}
                // Overlap se [cStart, cEnd) intersecta [insertStartSec, insertEndSec)
                if (cStart < insertEndSec - EPS && cEnd > insertStartSec + EPS) { hasOverlap = true; break; }
            }
            _lwDbgAudio('    track A' + (ti+1) + ': clips=' + clipCount + ' overlap=' + hasOverlap);
            if (!hasOverlap) { targetTrackIdx = ti; break; }
        }
        _lwDbgAudio('  targetTrackIdx (after scan) = ' + targetTrackIdx);

        // Se nenhuma track livre, cria nova audio track
        if (targetTrackIdx < 0) {
            _lwDbgAudio('  no free track — creating new');
            var beforeNumTracks = 0;
            try { beforeNumTracks = seq.audioTracks.numTracks; } catch(eBN) {}
            // Tenta múltiplas signatures (Premiere varia muito entre versões)
            var audioAttempts = [
                { name: 'QE.addTracks(0,0,1,1,numA)', fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 0, 1, 1, seq.audioTracks.numTracks); } },
                { name: 'QE.addTracks(0,1,numA)',     fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 1, seq.audioTracks.numTracks); } },
                { name: 'QE.addAudioTrack(numA)',     fn: function(){ app.enableQE(); qe.project.getActiveSequence().addAudioTrack(seq.audioTracks.numTracks); } },
                { name: 'QE.addAudioTrack()',         fn: function(){ app.enableQE(); qe.project.getActiveSequence().addAudioTrack(); } },
                { name: 'QE.addTracks(0,1)',          fn: function(){ app.enableQE(); qe.project.getActiveSequence().addTracks(0, 1); } },
                { name: 'addTracks(0,0,1,1,-1)',      fn: function(){ seq.addTracks(0, 0, 1, 1, -1); } },
                { name: 'audioTracks.addTrack(1)',    fn: function(){ seq.audioTracks.addTrack(1); } },
                { name: 'audioTracks.addTrack()',     fn: function(){ seq.audioTracks.addTrack(); } }
            ];
            for (var aa = 0; aa < audioAttempts.length; aa++) {
                var aAtt = audioAttempts[aa];
                try {
                    aAtt.fn();
                    for (var pt = 0; pt < 8; pt++) {
                        if (seq.audioTracks.numTracks > beforeNumTracks) break;
                        $.sleep(20);
                    }
                    _lwDbgAudio('  attempt[' + aAtt.name + ']: numTracks ' + beforeNumTracks + ' -> ' + seq.audioTracks.numTracks);
                    if (seq.audioTracks.numTracks > beforeNumTracks) {
                        // Acha a track EMPTY (sem clips) — funciona se QE inserir no topo OU no fim
                        var emptyAudioIdxs = [];
                        for (var ei = 0; ei < seq.audioTracks.numTracks; ei++) {
                            try { if (seq.audioTracks[ei].clips.numItems === 0) emptyAudioIdxs.push(ei); } catch(eEI) {}
                        }
                        if (emptyAudioIdxs.length > 0) {
                            targetTrackIdx = emptyAudioIdxs[emptyAudioIdxs.length - 1]; // empty mais ao topo
                            _lwDbgAudio('  -> using empty audio track at idx ' + targetTrackIdx);
                            break;
                        }
                        // Fallback: assume última
                        targetTrackIdx = seq.audioTracks.numTracks - 1;
                        _lwDbgAudio('  -> falling back to last audio track at idx ' + targetTrackIdx);
                        break;
                    }
                } catch(eAtt) { _lwDbgAudio('  attempt[' + aAtt.name + '] threw: ' + eAtt); }
            }
        }
        if (targetTrackIdx < 0) { _lwDbgAudio('  → no_track_available (all attempts failed)'); return 'ERR:no_track_available'; }
        _lwDbgAudio('  final targetTrackIdx=' + targetTrackIdx);

        // Insere o clip no playhead
        try {
            var targetTrack = seq.audioTracks[targetTrackIdx];
            var beforeCount = 0;
            try { beforeCount = targetTrack.clips.numItems; } catch(eBC) {}
            _lwDbgAudio('  targetTrack got, clips before=' + beforeCount);

            // Sanity check: is targetTrack locked?
            try {
                var isLocked = targetTrack.isLocked ? targetTrack.isLocked() : '?';
                _lwDbgAudio('  targetTrack.isLocked=' + isLocked);
            } catch(eLk) {}

            // Snapshot do total de clips em TODAS as audio tracks (pra detectar insert em outra track)
            function _totalAudioClips() {
                var total = 0;
                try {
                    for (var ti3 = 0; ti3 < seq.audioTracks.numTracks; ti3++) {
                        try { total += seq.audioTracks[ti3].clips.numItems; } catch(eT) {}
                    }
                } catch(eTT) {}
                return total;
            }
            var totalBefore = _totalAudioClips();
            _lwDbgAudio('  total audio clips before=' + totalBefore);

            // Tenta insertClip primeiro
            var insertOk = false;
            var lastInsertErr = '';
            try {
                _lwDbgAudio('  trying insertClip(foundItem, ' + insertStartSec + ')...');
                targetTrack.insertClip(foundItem, insertStartSec);
                insertOk = true;
                _lwDbgAudio('  insertClip OK (no throw)');
            } catch (eIns) {
                lastInsertErr = eIns.toString();
                _lwDbgAudio('  insertClip THREW: ' + lastInsertErr);
            }

            // Premiere pode demorar 50-200ms pra atualizar clips.numItems — poll
            var afterCount = beforeCount;
            var totalAfter = totalBefore;
            for (var rc = 0; rc < 10; rc++) {
                try { afterCount = targetTrack.clips.numItems; } catch(eAC) {}
                totalAfter = _totalAudioClips();
                if (afterCount > beforeCount || totalAfter > totalBefore) break;
                $.sleep(20);
            }
            _lwDbgAudio('  clips after insertClip target=' + afterCount + '/' + beforeCount + ' total=' + totalAfter + '/' + totalBefore);

            // Trust: se TOTAL aumentou (mesmo que target track count não), insert deu certo em algum lugar
            if (insertOk && afterCount === beforeCount && totalAfter === totalBefore) {
                _lwDbgAudio('  insertClip claimed OK but NEITHER target NOR total increased — treating as fail');
                insertOk = false;
                lastInsertErr = 'silent_fail (no clip count increased anywhere)';
            } else if (insertOk && afterCount === beforeCount && totalAfter > totalBefore) {
                _lwDbgAudio('  insertClip went to DIFFERENT track but worked (total increased)');
            }

            // Fallback 1: overwriteClip (não conflita com IDs em alguns casos)
            if (!insertOk) {
                try {
                    _lwDbgAudio('  trying overwriteClip(foundItem, ' + insertStartSec + ')...');
                    targetTrack.overwriteClip(foundItem, insertStartSec);
                    insertOk = true;
                    _lwDbgAudio('  overwriteClip OK');
                } catch (eOw) {
                    lastInsertErr += ' | overwrite: ' + eOw.toString();
                    _lwDbgAudio('  overwriteClip THREW: ' + eOw);
                }
                var afterCount2 = 0;
                try { afterCount2 = targetTrack.clips.numItems; } catch(eAC2) {}
                _lwDbgAudio('  clips after overwriteClip=' + afterCount2);
                if (insertOk && afterCount2 === beforeCount) {
                    _lwDbgAudio('  overwriteClip claimed OK but clip count unchanged');
                    insertOk = false;
                    lastInsertErr += ' | overwrite_silent_fail';
                }
            }

            // Fallback 2: QE-based insertion
            if (!insertOk) {
                _lwDbgAudio('  trying QE-based insertion...');
                try {
                    app.enableQE();
                    var qSeq2 = qe.project.getActiveSequence();
                    var qTrack = qSeq2.getAudioTrackAt(targetTrackIdx);
                    if (qTrack && qTrack.insert) {
                        // QE insert needs timecode string
                        var tc = '00:00:00:00';
                        try {
                            var fps = seq.timebase ? (254016000000 / parseFloat(seq.timebase)) : 30;
                            var totalFrames = Math.floor(insertStartSec * fps);
                            var hh = Math.floor(totalFrames / (fps*3600));
                            var mm = Math.floor((totalFrames % (fps*3600)) / (fps*60));
                            var ss = Math.floor((totalFrames % (fps*60)) / fps);
                            var ff = Math.floor(totalFrames % fps);
                            tc = (hh<10?'0':'')+hh+':'+(mm<10?'0':'')+mm+':'+(ss<10?'0':'')+ss+':'+(ff<10?'0':'')+ff;
                        } catch(eTc) {}
                        _lwDbgAudio('  QE.insert tc=' + tc);
                        qTrack.insert(foundItem.getProjectItemViewPreset ? foundItem.getProjectItemViewPreset() : foundItem, tc);
                        var afterCount3 = 0;
                        try { afterCount3 = targetTrack.clips.numItems; } catch(eAC3) {}
                        _lwDbgAudio('  clips after QE insert=' + afterCount3);
                        if (afterCount3 > beforeCount) insertOk = true;
                    } else {
                        _lwDbgAudio('  QE track or insert not available');
                    }
                } catch(eQI) { _lwDbgAudio('  QE insert THREW: ' + eQI); lastInsertErr += ' | qe: ' + eQI; }
            }

            if (!insertOk) {
                _lwDbgAudio('  ALL methods failed: ' + lastInsertErr);
                // Detecta erro de "multiple projects with the same ID" e dá mensagem clara
                if (/same id|multiple projects/i.test(lastInsertErr)) {
                    return 'ERR:duplicate_project_id';
                }
                return 'ERR:insert_failed:' + lastInsertErr.slice(0, 200);
            }

            // Se tinha in/out e o clip inserido pegou tudo, ajusta o end
            // Aproveita pra capturar o newClip pra aplicar volume dB também
            var newClipRef = null;
            if (hasInOut || _lwLibVolumeDb != null) {
                for (var nci = targetTrack.clips.numItems - 1; nci >= 0; nci--) {
                    var nc = targetTrack.clips[nci];
                    var ncStart = 0;
                    try { ncStart = parseFloat(nc.start.seconds); } catch(eNS) {}
                    if (Math.abs(ncStart - insertStartSec) < 0.5) { newClipRef = nc; break; }
                }
                if (newClipRef && hasInOut) {
                    try { newClipRef.end = insertStartSec + insertDurationSec; _lwDbgAudio('  trimmed clip end to ' + (insertStartSec + insertDurationSec)); } catch(eEd) { _lwDbgAudio('  trim err: ' + eEd); }
                }
            }

            // VOLUME dB: aplica no Audio Volume component do clip recém-inserido
            // Premiere usa unidade interna ~15 (0dB) - vez disso usamos conversão dB→gain.
            // Audio clip "Volume" component tem property "Level" cujo valor é 0-1 (gain linear, 1.0 = 0dB).
            // Pra setValue em dB → converte: gain = 10^(dB/20), depois multiplicado por escala interna.
            if (_lwLibVolumeDb != null && newClipRef) {
                _lwDbgAudio('  applying volumeDb=' + _lwLibVolumeDb);
                try {
                    var comps = newClipRef.components;
                    var compsN = 0;
                    try { compsN = comps.numItems; } catch(eN) {}
                    _lwDbgAudio('  newClip has ' + compsN + ' components');
                    // Dump nomes pra debug
                    for (var dI = 0; dI < compsN; dI++) {
                        try {
                            var dc = comps[dI];
                            var dn = '';
                            try { dn = String(dc.displayName || ''); } catch(eD1) {}
                            var dmn = '';
                            try { dmn = String(dc.matchName || ''); } catch(eD2) {}
                            _lwDbgAudio('    comp[' + dI + '] displayName="' + dn + '" matchName="' + dmn + '"');
                        } catch(eD) {}
                    }
                    // Convert dB pra gain LINEAR (0-1 onde 1 = 0dB)
                    // Mas Premiere Volume usa float onde 0dB = 1.0, +6dB ≈ 2.0, -6dB ≈ 0.5
                    var gainLinear = Math.pow(10, _lwLibVolumeDb / 20);
                    _lwDbgAudio('  gainLinear = ' + gainLinear + ' (from ' + _lwLibVolumeDb + 'dB)');

                    // Procura component "Volume" (matchName "AE.ADBE Fixed Volume" ou "ADBE Volume")
                    var volComp = null;
                    for (var cI = 0; cI < compsN && !volComp; cI++) {
                        try {
                            var cmp = comps[cI];
                            var cmpName = '';
                            var cmpMatch = '';
                            try { cmpName = String(cmp.displayName || ''); } catch(eDn) {}
                            try { cmpMatch = String(cmp.matchName || ''); } catch(eMn) {}
                            if (/^volume$|Audio Levels|Channel Volume|ADBE Volume|ADBE Audio Levels/i.test(cmpName) ||
                                /Volume|AudioLevels/i.test(cmpMatch)) {
                                volComp = cmp;
                                _lwDbgAudio('    matched volume component: ' + cmpName + ' / ' + cmpMatch);
                            }
                        } catch(eCm) {}
                    }
                    // Fallback: primeiro component que tenha prop chamada "Level"
                    if (!volComp) {
                        for (var cI2 = 0; cI2 < compsN && !volComp; cI2++) {
                            try {
                                var cmp2 = comps[cI2];
                                var pps = cmp2.properties;
                                if (pps && pps.numItems > 0) {
                                    for (var pj = 0; pj < pps.numItems; pj++) {
                                        var pn = '';
                                        try { pn = String(pps[pj].displayName || ''); } catch(ePj) {}
                                        if (/^Level$|Volume Level|Channel Volume/i.test(pn)) {
                                            volComp = cmp2;
                                            _lwDbgAudio('    fallback match via prop name: ' + pn);
                                            break;
                                        }
                                    }
                                }
                            } catch(eFb) {}
                        }
                    }

                    if (volComp) {
                        var props = volComp.properties;
                        var propsN = 0;
                        try { propsN = props.numItems; } catch(ePN) {}
                        _lwDbgAudio('    vol component has ' + propsN + ' properties');
                        var levelProp = null;
                        for (var pI = 0; pI < propsN; pI++) {
                            try {
                                var pp = props[pI];
                                var pName = '';
                                try { pName = String(pp.displayName || ''); } catch(ePnE) {}
                                _lwDbgAudio('      prop[' + pI + '] name="' + pName + '"');
                                if (/^Level$|Volume Level|Audio Level/i.test(pName)) {
                                    levelProp = pp;
                                    break;
                                }
                            } catch(ePI) {}
                        }
                        // Se não achou pelo nome, usa o primeiro (Premiere "Volume" comp tem Level no idx 1 ou 0)
                        if (!levelProp && propsN > 0) {
                            try { levelProp = props[propsN >= 2 ? 1 : 0]; _lwDbgAudio('    using prop idx fallback'); } catch(eIdx) {}
                        }
                        if (levelProp) {
                            // FIX: setValue(-10) não lança erro mas Premiere clampa pro fundo do range
                            // (essencialmente silencia). Solução: lê o valor ATUAL (que representa 0dB
                            // por default — escala varia: 1.0 em Premiere moderno, 15.0 em legacy)
                            // e aplica ganho LINEAR proporcional. Funciona em qualquer escala.
                            var currentVal = null;
                            try { currentVal = levelProp.getValue(); } catch(eGv) { _lwDbgAudio('    getValue err: ' + eGv); }
                            _lwDbgAudio('    currentVal (0dB ref) = ' + currentVal);
                            if (currentVal == null || isNaN(currentVal) || currentVal <= 0) {
                                // Sem ref válida — usa default 1.0 (gain linear puro)
                                currentVal = 1.0;
                                _lwDbgAudio('    using fallback currentVal=1.0');
                            }
                            // newValue = currentVal * 10^(dB/20)
                            var newVal = currentVal * gainLinear;
                            _lwDbgAudio('    target newVal = ' + currentVal + ' * ' + gainLinear + ' = ' + newVal);
                            try {
                                levelProp.setValue(newVal, true);
                                _lwDbgAudio('    ✓ setValue OK with proportional value');
                            } catch(eSv) {
                                _lwDbgAudio('    setValue err: ' + eSv);
                            }
                        } else {
                            _lwDbgAudio('    ✗ no level property found');
                        }
                    } else {
                        _lwDbgAudio('  ✗ no volume component found in clip');
                    }
                } catch(eVol) { _lwDbgAudio('  volume apply outer err: ' + eVol); }
            }

            _lwDbgAudio('  ═══ OK:inserted:track' + (targetTrackIdx + 1) + ':' + insertDurationSec.toFixed(2) + 's ═══');
            return 'OK:inserted:track' + (targetTrackIdx + 1) + ':' + insertDurationSec.toFixed(2) + 's';
        } catch (eOuter) {
            _lwDbgAudio('  OUTER catch: ' + eOuter);
            return 'ERR:insert_failed:' + eOuter.toString().slice(0, 200);
        }
    } catch (eAll) {
        return 'ERR:' + eAll.toString().slice(0, 200);
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — LIST ALL (effects + presets + audio sources merged)
// ═══════════════════════════════════════════════════════════════════
function lwSearchListAllExtended(sfxFolderPath, skipAudioScan) {
    var items = [];
    var debug = [];
    var errors = [];
    var masterDebug = [];
    var _skipAudio = (skipAudioScan === true);

    function mlog(m) { try { masterDebug.push(m); } catch(e) {} }

    function mergeFrom(jsonStr, label) {
        try {
            if (typeof JSON !== 'undefined') {
                var parsed = JSON.parse(jsonStr);
                if (parsed.items) {
                    for (var i = 0; i < parsed.items.length; i++) items.push(parsed.items[i]);
                }
                if (parsed.debug) debug.push(label + ':' + parsed.debug);
                if (parsed.error) errors.push(label + ':' + parsed.error);
                mlog('  ' + label + ': items=' + (parsed.items ? parsed.items.length : 0) + ' debug=' + (parsed.debug || '') + (parsed.error ? ' err=' + parsed.error : ''));
            }
        } catch(e) {
            errors.push(label + ':' + e.toString().slice(0, 60));
            mlog('  ' + label + ': PARSE-ERR ' + e);
        }
    }

    mlog('=== LION SEARCH MASTER DEBUG ===');
    mlog('time: ' + (new Date()).toString());

    // Inspeção inicial do app/qe
    try {
        mlog('app.version=' + app.version);
        mlog('app.project.name=' + (app.project ? app.project.name : 'null'));
    } catch(eA) { mlog('app inspect err: ' + eA); }
    try { app.enableQE(); } catch(eQE) {}
    try {
        mlog('typeof qe=' + (typeof qe));
        if (typeof qe !== 'undefined' && qe) {
            try { mlog('qe.numProjects=' + (qe.numProjects || 'undefined')); } catch(eQNP) {}
            // Lista TODAS as keys/methods de qe
            try {
                var qeKeys = '';
                for (var qk in qe) qeKeys += qk + ',';
                mlog('qe keys: ' + qeKeys);
            } catch(eQK) { mlog('qe keys err: ' + eQK); }
            // Lista TODAS as keys/methods de qe.project
            try {
                if (qe.project) {
                    var qePjKeys = '';
                    for (var qpk in qe.project) qePjKeys += qpk + ',';
                    mlog('qe.project keys: ' + qePjKeys);
                }
            } catch(eQPK) { mlog('qe.project keys err: ' + eQPK); }
        }
    } catch(eQ) { mlog('qe inspect err: ' + eQ); }
    try {
        mlog('typeof app.openDocuments=' + (typeof app.openDocuments));
        if (typeof app.openDocuments !== 'undefined') {
            try { mlog('app.openDocuments.numItems=' + (app.openDocuments.numItems || 'undef')); } catch(eND) {}
            try { mlog('app.openDocuments.length=' + (app.openDocuments.length || 'undef')); } catch(eLD) {}
        }
    } catch(eOD) { mlog('openDocs inspect err: ' + eOD); }
    try {
        mlog('typeof app.projects=' + (typeof app.projects));
        if (typeof app.projects !== 'undefined') {
            try { mlog('app.projects.numProjects=' + (app.projects.numProjects || 'undef')); } catch(eNP) {}
        }
    } catch(eAP) { mlog('app.projects inspect err: ' + eAP); }
    // Keys do app
    try {
        var appKeys = '';
        for (var ak in app) appKeys += ak + ',';
        mlog('app keys: ' + appKeys);
    } catch(eAK) { mlog('app keys err: ' + eAK); }

    // ─── EFFECTS ───
    mlog('--- listing effects ---');
    var fxJson = '';
    try { fxJson = lwSearchListAll(); } catch(e1) { errors.push('fx-call:' + e1); mlog('fx call exception: ' + e1); }
    mergeFrom(fxJson, 'fx');

    // ─── PRESETS ───
    mlog('--- listing presets ---');
    var presetsJson = '';
    try { presetsJson = lwSearchListPresets(); } catch(e2) { errors.push('preset-call:' + e2); mlog('preset call exception: ' + e2); }
    mergeFrom(presetsJson, 'pre');

    // ─── AUDIO SOURCES ───
    // Se skipAudioScan=true (useLibraryOnly), pula completamente o scan
    // de áudio dos projetos. Áudios da biblioteca são injetados depois.
    if (_skipAudio) {
        mlog('--- audio scan SKIPPED (useLibraryOnly=true) ---');
    } else {
        mlog('--- listing audio sources ---');
        var audioJson = '';
        try { audioJson = lwSearchListAudioSources(sfxFolderPath || ''); } catch(e3) { errors.push('audio-call:' + e3); mlog('audio call exception: ' + e3); }
        mergeFrom(audioJson, 'aud');
    }

    mlog('=== TOTAL items: ' + items.length + ' ===');

    // Escreve master debug — tenta Desktop primeiro (mais fácil de achar), TEMP como fallback
    var debugPaths = [];
    try { debugPaths.push(Folder.desktop.fsName + '/lion-search-debug.txt'); } catch(eDk) {}
    try { debugPaths.push(Folder.temp.fsName + '/lion-search-master-debug.txt'); } catch(eTmp) {}
    for (var dp = 0; dp < debugPaths.length; dp++) {
        try {
            var fdbg = new File(debugPaths[dp]);
            if (fdbg.open('w')) {
                fdbg.encoding = 'UTF-8';
                for (var di = 0; di < masterDebug.length; di++) fdbg.writeln(masterDebug[di]);
                fdbg.writeln('');
                fdbg.writeln('--- Sample of first items per kind ---');
                var samples = { 'video-fx': [], 'audio-fx': [], 'preset': [], 'audio-source': [] };
                for (var si = 0; si < items.length; si++) {
                    var k = items[si].kind;
                    if (samples[k] && samples[k].length < 5) samples[k].push(items[si]);
                }
                for (var sk in samples) {
                    var cnt = 0;
                    for (var x = 0; x < items.length; x++) if (items[x].kind === sk) cnt++;
                    fdbg.writeln('  ' + sk + ' (count: ' + cnt + '):');
                    for (var ssi = 0; ssi < samples[sk].length; ssi++) {
                        fdbg.writeln('    - ' + samples[sk][ssi].name + ' [' + samples[sk][ssi].category + ']');
                    }
                }
                fdbg.close();
            }
        } catch(eDw) {}
    }
    return _lwSearchJsonWithFlags(items, debug.join(','), errors);
}

// Versão estendida que preserva audioOnly/videoOnly nos items (usados pra filtrar)
function _lwSearchJsonWithFlags(items, debug, errors) {
    var out = '{"items":[';
    for (var i = 0; i < items.length; i++) {
        if (i > 0) out += ',';
        var it = items[i];
        out += '{"kind":"' + _lwEscapeJson(it.kind) + '"';
        out += ',"name":"' + _lwEscapeJson(it.name) + '"';
        out += ',"matchName":"' + _lwEscapeJson(it.matchName) + '"';
        out += ',"category":"' + _lwEscapeJson(it.category) + '"';
        if (it.audioOnly) out += ',"audioOnly":true';
        if (it.videoOnly) out += ',"videoOnly":true';
        if (it.isContainer) out += ',"isContainer":true';
        out += '}';
    }
    out += '],"debug":"' + _lwEscapeJson(debug || '') + '"';
    out += ',"error":"' + _lwEscapeJson((errors && errors.length) ? errors.join(' | ') : '') + '"}';
    return out;
}
