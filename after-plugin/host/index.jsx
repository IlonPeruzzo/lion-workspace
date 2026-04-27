// ============================================================================
// Lion Workspace — After Effects ExtendScript host
// ============================================================================
// Todas as funções abaixo retornam strings: "OK:..." ou "ERROR:..."
// O plugin (client/index.html) chama via csInterface.evalScript() e parseia.

#target aftereffects

// ────────────────────────────────────────────────────────────────────────────
// HELPERS GENÉRICOS
// ────────────────────────────────────────────────────────────────────────────

function lwAEPing() {
    try { return 'OK:AE' + (parseFloat(app.version) || 0); }
    catch (e) { return 'ERROR:' + e.toString(); }
}

function lwAEGetProjectName() {
    try {
        if (!app.project || !app.project.file) return '';
        return decodeURIComponent(app.project.file.name.replace(/\.aepx?$/i, ''));
    } catch (e) { return ''; }
}

function lwAEGetActiveCompName() {
    try {
        var item = app.project && app.project.activeItem;
        return (item && item instanceof CompItem) ? item.name : '';
    } catch (e) { return ''; }
}

function lwAEGetSessionInfo() {
    try {
        var info = {
            projectName: lwAEGetProjectName(),
            activeComp: lwAEGetActiveCompName(),
            aeVersion: parseFloat(app.version) || 0,
            numItems: app.project ? app.project.numItems : 0
        };
        return 'OK:' + JSON.stringify(info);
    } catch (e) { return 'ERROR:' + e.toString(); }
}

// Returns active CompItem or null
function _activeComp() {
    var item = app.project && app.project.activeItem;
    return (item && item instanceof CompItem) ? item : null;
}

// Returns array of selected layers (or all layers if none selected)
function _selectedLayers(comp) {
    if (!comp) return [];
    var sel = comp.selectedLayers;
    if (sel && sel.length > 0) return sel;
    return [];
}

// Get layer source size (width x height) at current time
function _layerSize(layer, time) {
    try {
        if (typeof time === 'undefined') time = (layer.containingComp || _activeComp()).time;
        var rect = layer.sourceRectAtTime(time, false);
        return [rect.width, rect.height];
    } catch (e) {
        return [layer.source ? layer.source.width : 100, layer.source ? layer.source.height : 100];
    }
}

// ────────────────────────────────────────────────────────────────────────────
// ANCHOR POINT PRO
// Recebe pos: 'tl','tc','tr','cl','cc','cr','bl','bc','br' (9 main)
// + sub-positions tipo 'tlc','tlcl','q1' etc se quiser. Por simplicidade
// só tratamos as 9 principais aqui (suficiente pro AE).
// ────────────────────────────────────────────────────────────────────────────
function lwAESetAnchor(pos, compensate) {
    app.beginUndoGroup('Lion Workspace — Set Anchor');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }

        var fx = 0.5, fy = 0.5;
        if (pos.indexOf('l') === 0 || pos === 'tl' || pos === 'cl' || pos === 'bl') fx = 0;
        if (pos.indexOf('r') === pos.length - 1 || pos === 'tr' || pos === 'cr' || pos === 'br') fx = 1;
        if (pos === 'tc' || pos === 'tl' || pos === 'tr') fy = 0;
        if (pos === 'bc' || pos === 'bl' || pos === 'br') fy = 1;
        // 5x5 grid sub-positions
        var subMap = {
            tlc:[0.25,0],   tcr:[0.75,0],
            cl:[0,0.5],     cc:[0.5,0.5],   cr:[1,0.5],
            tlcl:[0,0.25],  trcr:[1,0.25],
            blcl:[0,0.75],  brcr:[1,0.75],
            clcc:[0.25,0.5], cccr:[0.75,0.5],
            tccc:[0.5,0.25], bccc:[0.5,0.75],
            q1:[0.25,0.25], q2:[0.75,0.25], q3:[0.25,0.75], q4:[0.75,0.75]
        };
        if (subMap[pos]) { fx = subMap[pos][0]; fy = subMap[pos][1]; }

        var n = 0;
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            var size = _layerSize(L, comp.time);
            var w = size[0], h = size[1];
            var ap = L.property('Transform').property('Anchor Point');
            var posProp = L.property('Transform').property('Position');
            var oldAp = ap.value;
            var newAp = [w * fx, h * fy];
            if (compensate) {
                // Move position by the same delta so the layer doesn't jump
                var oldPos = posProp.value;
                // delta in layer space → world space (no rotation/scale handling for simplicity)
                var dx = (newAp[0] - oldAp[0]);
                var dy = (newAp[1] - oldAp[1]);
                // Scale into account (uniform)
                var sc = L.property('Transform').property('Scale').value;
                var sx = sc[0] / 100, sy = sc[1] / 100;
                var newPos;
                if (oldPos.length === 3) {
                    newPos = [oldPos[0] + dx * sx, oldPos[1] + dy * sy, oldPos[2]];
                } else {
                    newPos = [oldPos[0] + dx * sx, oldPos[1] + dy * sy];
                }
                posProp.setValue(newPos);
            }
            // Preserve Z if 3D
            if (oldAp.length === 3) ap.setValue([newAp[0], newAp[1], oldAp[2]]);
            else ap.setValue(newAp);
            n++;
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// ────────────────────────────────────────────────────────────────────────────
// EASING / CURVAS
// Aplica bezier (cp1x, cp1y, cp2x, cp2y) nos keyframes selecionados
// usando KeyframeEase (influence + speed). Conversão aproximada:
//   influence_out = cp1x * 100, influence_in = (1 - cp2x) * 100
// (Não é matematicamente exato vs cubic-bezier do CSS mas é fiel o suficiente
//  pra os presets típicos.)
// ────────────────────────────────────────────────────────────────────────────
function lwAEApplyEasing(cp1x, cp1y, cp2x, cp2y, applyMode) {
    app.beginUndoGroup('Lion Workspace — Apply Easing');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }

        var influenceOut = Math.max(0.1, Math.min(100, cp1x * 100));
        var influenceIn  = Math.max(0.1, Math.min(100, (1 - cp2x) * 100));
        // Speed: AE uses units/sec — use 0 for ease curves (let influence shape it)
        var speedOut = 0, speedIn = 0;
        var totalKeys = 0, totalProps = 0;

        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            var props = _collectAnimatedProps(L);
            for (var p = 0; p < props.length; p++) {
                var prop = props[p];
                var keys = prop.selectedKeys;
                if (!keys || keys.length < 2) continue;
                if (applyMode === 'all') {
                    keys = [];
                    for (var k = 1; k <= prop.numKeys; k++) keys.push(k);
                }
                totalProps++;
                for (var ki = 0; ki < keys.length; ki++) {
                    var idx = keys[ki];
                    try {
                        var dim = (typeof prop.value === 'number') ? 1 : prop.value.length;
                        var easeIn = []; var easeOut = [];
                        for (var d = 0; d < dim; d++) {
                            easeIn.push(new KeyframeEase(speedIn, influenceIn));
                            easeOut.push(new KeyframeEase(speedOut, influenceOut));
                        }
                        prop.setTemporalEaseAtKey(idx, easeIn, easeOut);
                        // Also set interpolation type to BEZIER
                        prop.setInterpolationTypeAtKey(idx, KeyframeInterpolationType.BEZIER, KeyframeInterpolationType.BEZIER);
                        totalKeys++;
                    } catch (ek) {}
                }
            }
        }
        app.endUndoGroup();
        return 'OK:' + totalKeys + ':' + totalProps;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// Recursively collect all keyframe-able properties on a layer
function _collectAnimatedProps(layer) {
    var out = [];
    function walk(group) {
        for (var i = 1; i <= group.numProperties; i++) {
            var p = group.property(i);
            if (p.propertyType === PropertyType.PROPERTY) {
                if (p.canVaryOverTime && p.numKeys > 0) out.push(p);
            } else {
                walk(p);
            }
        }
    }
    walk(layer);
    return out;
}

// ────────────────────────────────────────────────────────────────────────────
// EXPRESSIONS (Wiggle, Shake, Loop, etc)
// ────────────────────────────────────────────────────────────────────────────
function lwAEApplyExpression(targetProp, kind, params) {
    app.beginUndoGroup('Lion Workspace — Expression: ' + kind);
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }
        var p = params || {};

        // Build the expression string
        var expr = '';
        if (kind === 'wiggle') {
            var freq = p.freq || 2, amp = p.amp || 30;
            expr = 'wiggle(' + freq + ', ' + amp + ');';
        } else if (kind === 'shake') {
            var sf = p.freq || 8, sa = p.amp || 50;
            expr = 'wiggle(' + sf + ', ' + sa + ');';
        } else if (kind === 'loopout-cycle') {
            expr = 'loopOut("cycle");';
        } else if (kind === 'loopout-pingpong') {
            expr = 'loopOut("pingpong");';
        } else if (kind === 'loopout-continue') {
            expr = 'loopOut("continue");';
        } else if (kind === 'loopin-cycle') {
            expr = 'loopIn("cycle");';
        } else if (kind === 'bounce') {
            // Classic bounce overshoot expression
            expr = 'amp = ' + (p.amp || 0.1) + ';\n' +
                'freq = ' + (p.freq || 2) + ';\n' +
                'decay = ' + (p.decay || 8) + ';\n' +
                'n = 0;\n' +
                'if (numKeys > 0) {\n' +
                '  n = nearestKey(time).index;\n' +
                '  if (key(n).time > time) n--;\n' +
                '}\n' +
                'if (n == 0) value;\n' +
                'else {\n' +
                '  t = time - key(n).time;\n' +
                '  v = velocityAtTime(key(n).time - 0.001);\n' +
                '  value + v * amp * Math.sin(freq * t * 2 * Math.PI) / Math.exp(decay * t);\n' +
                '}';
        } else if (kind === 'inertia') {
            expr = 'amp = ' + (p.amp || 0.05) + ';\n' +
                'freq = ' + (p.freq || 4.0) + ';\n' +
                'decay = ' + (p.decay || 5.0) + ';\n' +
                'n = 0;\n' +
                'if (numKeys > 0) {\n' +
                '  n = nearestKey(time).index;\n' +
                '  if (key(n).time > time) n--;\n' +
                '}\n' +
                'if (n == 0) { t = 0; }\n' +
                'else { t = time - key(n).time; }\n' +
                'if (n > 0) {\n' +
                '  v = velocityAtTime(key(n).time - 0.001) * amp;\n' +
                '  value + v * Math.sin(freq * t * 2 * Math.PI) / Math.exp(decay * t);\n' +
                '} else value;';
        } else if (kind === 'clear') {
            expr = '';
        } else {
            app.endUndoGroup();
            return 'UNKNOWN_KIND';
        }

        var n = 0;
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            var prop = _resolveTargetProp(L, targetProp);
            if (!prop) continue;
            try {
                if (prop.canSetExpression) {
                    if (kind === 'clear') {
                        prop.expression = '';
                        prop.expressionEnabled = false;
                    } else {
                        prop.expression = expr;
                        prop.expressionEnabled = true;
                    }
                    n++;
                }
            } catch (e) {}
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// targetProp: 'position' | 'rotation' | 'scale' | 'opacity' | 'all-transform'
function _resolveTargetProp(layer, target) {
    var t = layer.property('Transform');
    if (target === 'position') return t.property('Position');
    if (target === 'rotation') return t.property('Rotation') || t.property('Z Rotation');
    if (target === 'scale')    return t.property('Scale');
    if (target === 'opacity')  return t.property('Opacity');
    if (target === 'anchor')   return t.property('Anchor Point');
    return t.property('Position'); // default
}

// ────────────────────────────────────────────────────────────────────────────
// QUICK RIGS
// ────────────────────────────────────────────────────────────────────────────

// Cria um Null Object e parenteia as layers selecionadas a ele
function lwAECreateNullParent() {
    app.beginUndoGroup('Lion Workspace — Null Parent');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }

        var nullL = comp.layers.addNull();
        nullL.name = 'CONTROL';
        nullL.label = 9; // green
        // Posiciona o null no centro da bbox das layers selecionadas
        var minX=1e9, minY=1e9, maxX=-1e9, maxY=-1e9;
        for (var i = 0; i < layers.length; i++) {
            var pos = layers[i].property('Transform').property('Position').value;
            if (pos[0] < minX) minX = pos[0];
            if (pos[1] < minY) minY = pos[1];
            if (pos[0] > maxX) maxX = pos[0];
            if (pos[1] > maxY) maxY = pos[1];
        }
        nullL.property('Transform').property('Position').setValue([(minX+maxX)/2, (minY+maxY)/2]);
        // Parent
        for (var j = 0; j < layers.length; j++) {
            try { layers[j].parent = nullL; } catch (e) {}
        }
        app.endUndoGroup();
        return 'OK:' + layers.length;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// Auto-orient layers
// mode: 'off' | 'path' | 'camera'
function lwAEAutoOrient(mode) {
    app.beginUndoGroup('Lion Workspace — Auto-Orient');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }
        var aoType;
        if (mode === 'path') aoType = AutoOrientType.ALONG_SUBPATH;
        else if (mode === 'camera') aoType = AutoOrientType.CAMERA_OR_POINT_OF_INTEREST;
        else aoType = AutoOrientType.NO_AUTO_ORIENT;
        var n = 0;
        for (var i = 0; i < layers.length; i++) {
            try { layers[i].autoOrient = aoType; n++; } catch (e) {}
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// Cria uma câmera 2-node
function lwAECreateCamera() {
    app.beginUndoGroup('Lion Workspace — Camera Rig');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var cam = comp.layers.addCamera('Camera', [comp.width/2, comp.height/2]);
        cam.label = 13; // blue
        // Adiciona um null com nome 'CAM_CTRL' parenteado pra controlar
        var nullL = comp.layers.addNull();
        nullL.name = 'CAM_CTRL';
        nullL.label = 9;
        nullL.property('Transform').property('Position').setValue([comp.width/2, comp.height/2, 0]);
        try { cam.parent = nullL; } catch (e) {}
        app.endUndoGroup();
        return 'OK:1';
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// ────────────────────────────────────────────────────────────────────────────
// ALIGN / DISTRIBUTE
// where: 'top','bottom','left','right','centerH','centerV','center'
// reference: 'comp' | 'selection'
// ────────────────────────────────────────────────────────────────────────────
function lwAEAlign(where, reference) {
    app.beginUndoGroup('Lion Workspace — Align');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }

        var refLeft, refTop, refRight, refBottom, refCenterX, refCenterY;
        if (reference === 'selection' && layers.length > 1) {
            refLeft = 1e9; refTop = 1e9; refRight = -1e9; refBottom = -1e9;
            for (var i = 0; i < layers.length; i++) {
                var b = _layerBounds(layers[i], comp.time);
                if (b.left < refLeft) refLeft = b.left;
                if (b.top < refTop) refTop = b.top;
                if (b.right > refRight) refRight = b.right;
                if (b.bottom > refBottom) refBottom = b.bottom;
            }
        } else {
            refLeft = 0; refTop = 0; refRight = comp.width; refBottom = comp.height;
        }
        refCenterX = (refLeft + refRight) / 2;
        refCenterY = (refTop + refBottom) / 2;

        var n = 0;
        for (var j = 0; j < layers.length; j++) {
            var L = layers[j];
            var b = _layerBounds(L, comp.time);
            var pos = L.property('Transform').property('Position').value;
            var newX = pos[0], newY = pos[1];
            // pos is the layer's anchor point in comp space; bbox is built around it
            var halfW = (b.right - b.left) / 2;
            var halfH = (b.bottom - b.top) / 2;
            var apX = pos[0] - b.left;       // distance from bbox.left to anchor
            var apY = pos[1] - b.top;        // distance from bbox.top to anchor
            if (where === 'left')      newX = refLeft + apX;
            else if (where === 'right')newX = refRight - (b.right - pos[0]);
            else if (where === 'top')  newY = refTop + apY;
            else if (where === 'bottom') newY = refBottom - (b.bottom - pos[1]);
            else if (where === 'centerH') newX = refCenterX - halfW + apX;
            else if (where === 'centerV') newY = refCenterY - halfH + apY;
            else if (where === 'center') {
                newX = refCenterX - halfW + apX;
                newY = refCenterY - halfH + apY;
            }
            var newPos = pos.length === 3 ? [newX, newY, pos[2]] : [newX, newY];
            L.property('Transform').property('Position').setValue(newPos);
            n++;
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

function _layerBounds(layer, time) {
    try {
        var rect = layer.sourceRectAtTime(time, false);
        var pos = layer.property('Transform').property('Position').value;
        var ap = layer.property('Transform').property('Anchor Point').value;
        var sc = layer.property('Transform').property('Scale').value;
        var sx = sc[0]/100, sy = sc[1]/100;
        var w = rect.width * sx, h = rect.height * sy;
        var left = pos[0] - ap[0] * sx;
        var top  = pos[1] - ap[1] * sy;
        return { left: left, top: top, right: left + w, bottom: top + h, width: w, height: h };
    } catch (e) {
        return { left: 0, top: 0, right: 0, bottom: 0, width: 0, height: 0 };
    }
}

// Distribui horizontal ou verticalmente entre as bordas
function lwAEDistribute(axis) {
    app.beginUndoGroup('Lion Workspace — Distribute');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (layers.length < 3) { app.endUndoGroup(); return 'NEED_3'; }

        // Ordena pelo centro do eixo
        var arr = [];
        for (var i = 0; i < layers.length; i++) {
            var b = _layerBounds(layers[i], comp.time);
            var c = axis === 'h' ? (b.left+b.right)/2 : (b.top+b.bottom)/2;
            arr.push({ layer: layers[i], bounds: b, center: c });
        }
        arr.sort(function(a,b){ return a.center - b.center; });

        var first = arr[0], last = arr[arr.length-1];
        var step = (last.center - first.center) / (arr.length - 1);

        for (var k = 1; k < arr.length - 1; k++) {
            var L = arr[k].layer;
            var pos = L.property('Transform').property('Position').value;
            var targetCenter = first.center + step * k;
            var delta = targetCenter - arr[k].center;
            var newPos;
            if (axis === 'h') {
                newPos = pos.length === 3 ? [pos[0]+delta, pos[1], pos[2]] : [pos[0]+delta, pos[1]];
            } else {
                newPos = pos.length === 3 ? [pos[0], pos[1]+delta, pos[2]] : [pos[0], pos[1]+delta];
            }
            L.property('Transform').property('Position').setValue(newPos);
        }
        app.endUndoGroup();
        return 'OK:' + (arr.length - 2);
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// ────────────────────────────────────────────────────────────────────────────
// COLOR / LUT
// ────────────────────────────────────────────────────────────────────────────

// Aplica um Apply Color LUT effect com o arquivo .cube selecionado
function lwAEApplyLUT(filePath) {
    app.beginUndoGroup('Lion Workspace — Apply LUT');
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }
        var f = new File(filePath);
        if (!f.exists) { app.endUndoGroup(); return 'FILE_NOT_FOUND'; }

        var n = 0;
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            try {
                var fx = L.property('Effects').addProperty('ADBE Apply Color LUT2');
                if (!fx) fx = L.property('Effects').addProperty('Apply Color LUT');
                if (fx) {
                    // Find the file param
                    var fileProp = fx.property('ADBE Apply Color LUT2-0001') || fx.property(1);
                    if (fileProp && fileProp.setValue) fileProp.setValue(f);
                    n++;
                }
            } catch (e) {}
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}

// Adiciona Curves preset rápido (lift, contrast, fade)
function lwAEColorPreset(preset) {
    app.beginUndoGroup('Lion Workspace — Color: ' + preset);
    try {
        var comp = _activeComp();
        if (!comp) { app.endUndoGroup(); return 'NO_COMP'; }
        var layers = _selectedLayers(comp);
        if (!layers.length) { app.endUndoGroup(); return 'NO_SELECTION'; }

        var n = 0;
        for (var i = 0; i < layers.length; i++) {
            var L = layers[i];
            try {
                if (preset === 'cinematic-fade') {
                    var lev = L.property('Effects').addProperty('ADBE Pro Levels2');
                    if (lev) {
                        // Lift the blacks slightly for a faded look
                        try { lev.property('Output Black').setValue(15); } catch(e) {}
                        try { lev.property('Gamma').setValue(0.95); } catch(e) {}
                    }
                    n++;
                } else if (preset === 'punch-contrast') {
                    var c = L.property('Effects').addProperty('ADBE Brightness & Contrast 2');
                    if (c) {
                        try { c.property('Contrast').setValue(20); } catch(e) {}
                    }
                    n++;
                } else if (preset === 'warm') {
                    var cb = L.property('Effects').addProperty('ADBE Color Balance (HLS)');
                    // Fall back to Photo Filter
                    if (!cb) cb = L.property('Effects').addProperty('ADBE Photo Filter');
                    if (cb) n++;
                } else if (preset === 'cool') {
                    var pf = L.property('Effects').addProperty('ADBE Photo Filter');
                    if (pf) {
                        try { pf.property('Filter').setValue(2); } catch(e) {} // Cooling Filter
                        n++;
                    }
                }
            } catch (e) {}
        }
        app.endUndoGroup();
        return 'OK:' + n;
    } catch (e) {
        app.endUndoGroup();
        return 'ERROR:' + e.toString();
    }
}
