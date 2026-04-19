// ExtendScript for Adobe Premiere Pro — Lion Workspace Plugin

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
        var ticks = Number(pos.ticks);
        var nudge = createTimeAt(ticks + 1);
        seq.setPlayerPosition(nudge.ticks);
        seq.setPlayerPosition(pos.ticks);
    } catch (e) {}
}

function applyBezierEasing(cp1x, cp1y, cp2x, cp2y, mode) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var totalApplied = 0;
        var totalKeys = 0;
        var errors = [];
        var STEPS = 10;

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
                                    var easedFraction = evalBezierCurve(fraction, cp1x, cp1y, cp2x, cp2y);
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

    // AUTO-DETECT: normalized (0-1) vs pixel mode
    // If both anchor values are < 2.0, Premiere is using normalized coordinates
    var isNormalized = (Math.abs(oldAnchor[0]) < 2.0 && Math.abs(oldAnchor[1]) < 2.0);

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
        var initial = null;
        try {
            if (app.project && app.project.path) {
                var pf = new File(app.project.path);
                if (pf.parent && pf.parent.exists) initial = pf.parent;
            }
        } catch (e) {}
        var folder = (initial || Folder.desktop).selectDlg('Escolher pasta para salvar prints');
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
        var clips = getSelectedClips();
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

                var okImport = false;
                try {
                    if (targetBin) {
                        okImport = app.project.importFiles([fp], false, targetBin, false);
                    } else {
                        okImport = app.project.importFiles([fp]);
                    }
                } catch (e1) {}
                // Fallback: import to root
                if (!okImport) {
                    try { app.project.importFiles([fp]); okImport = true; } catch (e2) {}
                }

                if (!okImport) { failed++; continue; }
                imported++;
                $.sleep(500);

                // Insere no playhead (smart track: encontra track vazia ou cria nova)
                if (insertAtPlayhead && seq) {
                    var clipItem = _cpFindItemByPath(fp);
                    if (clipItem) {
                        try {
                            var insertTime = seq.getPlayerPosition();
                            var mType = _getMediaType(fp);
                            var smartTrack = findEmptyTrackAtPlayhead(seq, mType);
                            if (smartTrack) {
                                smartTrack.insertClip(clipItem, insertTime);
                                inserted++;
                            }
                        } catch (e) {}
                    }
                }
            } catch (e) { failed++; }
        }

        return 'OK:' + imported + ':' + inserted + ':' + failed;
    } catch (e) { return 'ERR:' + e.toString(); }
}

// ============================================// ============================================

// Smart track finder: finds an empty video or audio track at playhead, or creates a new one
// mediaType: 'video' or 'audio'
// Returns a track object (never null — creates one if needed)
function findEmptyTrackAtPlayhead(seq, mediaType) {
    var playhead = seq.getPlayerPosition();
    var playTicks = Number(playhead.ticks);
    var tracks = (mediaType === 'audio') ? seq.audioTracks : seq.videoTracks;

    // Check each existing track for a gap at the playhead
    for (var t = 0; t < tracks.numTracks; t++) {
        var track = tracks[t];
        var hasClipAtPlayhead = false;
        try {
            for (var c = 0; c < track.clips.numItems; c++) {
                var clip = track.clips[c];
                var clipStart = Number(clip.start.ticks);
                var clipEnd = Number(clip.end.ticks);
                if (playTicks >= clipStart && playTicks < clipEnd) {
                    hasClipAtPlayhead = true;
                    break;
                }
            }
        } catch (e) {}
        if (!hasClipAtPlayhead) return track;
    }

    // All tracks occupied — add a new track
    try {
        if (mediaType === 'audio') {
            seq.audioTracks.addTrack();
            return seq.audioTracks[seq.audioTracks.numTracks - 1];
        } else {
            seq.videoTracks.addTrack();
            return seq.videoTracks[seq.videoTracks.numTracks - 1];
        }
    } catch (e) {
        // addTrack not available (older Premiere) — fallback to first track
        return tracks[0];
    }
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
        var mType = _getMediaType(filePath);
        var track = findEmptyTrackAtPlayhead(seq, mType);
        if (!track) return 'NO_TRACK';

        track.insertClip(clipItem, insertTime);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// ============================================
// AUTOCUT — Silence-based auto-cutting
// Receives silence timestamps from FFmpeg analysis
// and performs razor cuts + removes silent segments
// ============================================

// Get the media file path of the first selected clip
function autoCutGetClipInfo() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var clip = clips[0];
        var mediaPath = '';
        try {
            var pi = clip.projectItem;
            if (pi) mediaPath = pi.getMediaPath() || '';
        } catch (e) {}

        if (!mediaPath) return 'NO_MEDIA';

        // Get clip timing info
        var clipStart = Number(clip.start.ticks);
        var clipEnd = Number(clip.end.ticks);
        var clipInPoint = 0;
        try { clipInPoint = Number(clip.inPoint.ticks); } catch (e) {}

        // Get ticks per second for conversion
        var tps = 254016000000; // Premiere default ticks per second
        try {
            var t = new Time();
            t.seconds = 1;
            tps = Number(t.ticks);
        } catch (e) {}

        // Find which track and index the clip is on
        var trackIndex = -1;
        var clipIndex = -1;
        var trackType = 'video';
        // Search video tracks
        for (var vt = 0; vt < seq.videoTracks.numTracks; vt++) {
            var track = seq.videoTracks[vt];
            for (var c = 0; c < track.clips.numItems; c++) {
                if (track.clips[c] === clip || Number(track.clips[c].start.ticks) === clipStart) {
                    trackIndex = vt;
                    clipIndex = c;
                    break;
                }
            }
            if (trackIndex >= 0) break;
        }
        // Search audio tracks if not found in video
        if (trackIndex < 0) {
            for (var at = 0; at < seq.audioTracks.numTracks; at++) {
                var atrack = seq.audioTracks[at];
                for (var c = 0; c < atrack.clips.numItems; c++) {
                    if (atrack.clips[c] === clip || Number(atrack.clips[c].start.ticks) === clipStart) {
                        trackIndex = at;
                        clipIndex = c;
                        trackType = 'audio';
                        break;
                    }
                }
                if (trackIndex >= 0) break;
            }
        }

        var info = {
            mediaPath: mediaPath,
            clipStart: clipStart,
            clipEnd: clipEnd,
            clipInPoint: clipInPoint,
            ticksPerSecond: tps,
            trackIndex: trackIndex,
            clipIndex: clipIndex,
            trackType: trackType,
            seqName: seq.name
        };
        return 'OK:' + JSON.stringify(info);
    } catch (e) {
        return 'ERROR:' + e.toString();
    }
}

// ─── Backup sequence before destructive operations ───
// Tries seq.clone() (Premiere 2024+); falls back to project save
function autoCutBackupSequence() {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';
        var origName = seq.name;
        var stamp = new Date();
        var tag = stamp.getFullYear() + '-' + (stamp.getMonth()+1) + '-' + stamp.getDate() +
                  '_' + stamp.getHours() + 'h' + stamp.getMinutes();
        var backupName = origName + ' [BACKUP ' + tag + ']';

        // Try Sequence.clone() (newer Premiere versions)
        try {
            if (typeof seq.clone === 'function') {
                var cloned = seq.clone();
                if (cloned) {
                    try { cloned.name = backupName; } catch(e) {}
                    // Return focus to original sequence
                    try { app.project.activeSequence = seq; } catch(e) {}
                    return 'CLONED:' + backupName;
                }
            }
        } catch(eClone) {}

        // Fallback 1: save project to disk (if it has been saved before)
        try {
            if (app.project.path && app.project.path.length > 0) {
                app.project.save();
                return 'SAVED';
            }
        } catch(eSave) {}

        // Fallback 2: at least add a sequence marker as a visual anchor
        try {
            var markers = seq.markers;
            if (markers && markers.createMarker) {
                var m = markers.createMarker(0);
                try { m.name = 'Pre-AutoCut: ' + tag; } catch(e) {}
                try { m.comments = 'AutoCut iniciou aqui. Use Ctrl+Z pra desfazer.'; } catch(e) {}
                return 'MARKER';
            }
        } catch(eMark) {}

        return 'NONE';
    } catch (e) {
        return 'ERR:' + e.toString();
    }
}

// ─── Extract helper: use QE sequence.extract() to remove range + ripple ───
// Returns true if clip count changed (silence removed)
function _doExtractRange(startTick, endTick, qeSeq, stdTrack) {
    if (!qeSeq) return false;
    var before = stdTrack.clips.numItems;
    var startStr = String(Math.round(startTick));
    var endStr = String(Math.round(endTick));
    try {
        qeSeq.setInPoint(startStr);
        $.sleep(50);
        qeSeq.setOutPoint(endStr);
        $.sleep(50);
        if (typeof qeSeq.extract === 'function') {
            qeSeq.extract();
            $.sleep(300);
            return stdTrack.clips.numItems !== before;
        }
    } catch(e) {}
    return false;
}

// ─── Razor helper: tries multiple methods, returns true if clip count increased ───
function _doRazorAt(tickVal, targetTrack, qeTrack, qeSeq, seq) {
    var nb = targetTrack.clips.numItems;
    var tickStr = String(Math.round(tickVal));

    // Method A: Set playhead → get native ticks → QE track razor
    if (qeTrack) {
        try {
            seq.setPlayerPosition(tickStr);
            $.sleep(50);
            var nativeTicks = seq.getPlayerPosition().ticks;
            qeTrack.razor(nativeTicks);
            $.sleep(200);
            if (targetTrack.clips.numItems > nb) return 'A';
        } catch (e) {}
    }

    // Method B: QE track razor with our ticks directly
    if (qeTrack) {
        try {
            qeTrack.razor(tickStr);
            $.sleep(200);
            if (targetTrack.clips.numItems > nb) return 'B';
        } catch (e) {}
    }

    // Method C: Standard DOM track.razor with Time object
    try {
        targetTrack.razor(createTimeAt(tickVal));
        $.sleep(200);
        if (targetTrack.clips.numItems > nb) return 'C';
    } catch (e) {}

    // Method D: QE sequence-level razor (cuts all targeted tracks)
    if (qeSeq) {
        try {
            seq.setPlayerPosition(tickStr);
            $.sleep(50);
            qeSeq.razor(seq.getPlayerPosition().ticks);
            $.sleep(200);
            if (targetTrack.clips.numItems > nb) return 'D';
        } catch (e) {}
    }

    return '';
}

// Perform the actual cuts based on silence data
// silencesJson: JSON string of [{start, end}] in seconds (relative to media start)
// padding: seconds of padding to keep around speech
// mode: 'remove' (delete silences) or 'cut' (just razor, don't delete)
// trackIdx: video/audio track index (from autoCutGetClipInfo)
// trackTyp: 'video' or 'audio'
function autoCutExecute(silencesJson, padding, mode, trackIdx, trackTyp) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var silences;
        try { silences = eval('(' + silencesJson + ')'); } catch (ep) { return 'PARSE_ERR'; }
        if (!silences || silences.length === 0) return 'NO_SILENCES';

        padding = padding || 0;
        mode = mode || 'remove';
        trackIdx = parseInt(String(trackIdx), 10);
        trackTyp = String(trackTyp || 'video');

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var clip = clips[0];
        var clipStartTicks = Number(clip.start.ticks);
        var clipEndTicks = Number(clip.end.ticks);
        var clipInTicks = 0;
        try { clipInTicks = Number(clip.inPoint.ticks); } catch (e) {}

        var tps = 254016000000;
        try { var t = new Time(); t.seconds = 1; tps = Number(t.ticks); } catch (e) {}

        // ─── Find track ───
        var targetTrack = null;
        var fIdx = trackIdx, fTyp = trackTyp;
        if (!isNaN(trackIdx) && trackIdx >= 0) {
            try {
                if (trackTyp === 'audio') targetTrack = seq.audioTracks[trackIdx];
                else targetTrack = seq.videoTracks[trackIdx];
            } catch (e) {}
        }
        if (!targetTrack) {
            for (var vt = 0; vt < seq.videoTracks.numTracks; vt++) {
                var vtrack = seq.videoTracks[vt];
                for (var vc = 0; vc < vtrack.clips.numItems; vc++) {
                    if (Math.abs(Number(vtrack.clips[vc].start.ticks) - clipStartTicks) < tps * 0.01) {
                        targetTrack = vtrack; fIdx = vt; fTyp = 'video'; break;
                    }
                }
                if (targetTrack) break;
            }
        }
        if (!targetTrack) {
            for (var at = 0; at < seq.audioTracks.numTracks; at++) {
                var atrack = seq.audioTracks[at];
                for (var ac = 0; ac < atrack.clips.numItems; ac++) {
                    if (Math.abs(Number(atrack.clips[ac].start.ticks) - clipStartTicks) < tps * 0.01) {
                        targetTrack = atrack; fIdx = at; fTyp = 'audio'; break;
                    }
                }
                if (targetTrack) break;
            }
        }
        if (!targetTrack) return 'NO_TRACK';
        if (targetTrack.clips.numItems === 0) return 'TRACK_EMPTY';

        // ─── Target ONLY this track (required for razor) ───
        try {
            for (var ti = 0; ti < seq.videoTracks.numTracks; ti++) {
                seq.videoTracks[ti].setTargeted(ti === fIdx && fTyp === 'video', true);
            }
            for (var tai = 0; tai < seq.audioTracks.numTracks; tai++) {
                seq.audioTracks[tai].setTargeted(tai === fIdx && fTyp === 'audio', true);
            }
        } catch (eTgt) {}

        // ─── Enable QE DOM ───
        var qeSeq = null, qeTrack = null;
        try {
            app.enableQE();
            qeSeq = qe.project.getActiveSequence();
            if (fTyp === 'audio') qeTrack = qeSeq.getAudioTrackAt(fIdx);
            else qeTrack = qeSeq.getVideoTrackAt(fIdx);
        } catch (eqe) {}

        // ─── Convert silence times to timeline ticks ───
        var cutPoints = [];
        for (var i = 0; i < silences.length; i++) {
            var s = silences[i];
            var silStartSec = Math.max(0, s.start + padding);
            var silEndSec = Math.max(silStartSec, s.end - padding);
            if (silEndSec <= silStartSec) continue;

            var tlStart = clipStartTicks + Math.round((silStartSec * tps) - clipInTicks);
            var tlEnd = clipStartTicks + Math.round((silEndSec * tps) - clipInTicks);

            if (tlStart < clipStartTicks) tlStart = clipStartTicks;
            if (tlEnd > clipEndTicks) tlEnd = clipEndTicks;
            if (tlEnd <= tlStart) continue;

            cutPoints.push({ start: tlStart, end: tlEnd });
        }

        if (cutPoints.length === 0) return 'NO_CUTS';
        cutPoints.sort(function(a, b) { return b.start - a.start; });

        var cutsApplied = 0, removedDuration = 0;
        var dbg = [];
        var tol = Math.round(tps * 0.05);
        var origPlayhead = seq.getPlayerPosition().ticks;

        // Unlock target track so razor/extract can operate on it
        try { if (typeof targetTrack.setLocked === 'function') targetTrack.setLocked(false); } catch(e) {}

        // ── Save current in/out points so we can restore them after extract ──
        var origInTicks = null, origOutTicks = null;
        try { origInTicks = seq.getInPoint ? seq.getInPoint().ticks : null; } catch(e) {}
        try { origOutTicks = seq.getOutPoint ? seq.getOutPoint().ticks : null; } catch(e) {}

        // Detect if extract works on first cut; if yes, use it for all
        var useExtract = null; // null = undetermined, true/false after first test

        for (var ci = 0; ci < cutPoints.length; ci++) {
            var cp = cutPoints[ci];
            try {
                var numBefore = targetTrack.clips.numItems;
                var method = '';

                // ── TRY 1: qeSeq.extract() — cleanest, removes range + ripples in one shot ──
                if (mode === 'remove' && useExtract !== false && qeSeq) {
                    if (_doExtractRange(cp.start, cp.end, qeSeq, targetTrack)) {
                        method = 'extract';
                        useExtract = true;
                        removedDuration += (cp.end - cp.start) / tps;
                        cutsApplied++;
                        if (ci === 0) dbg.push('m=extract');
                        continue;
                    } else if (ci === 0) {
                        useExtract = false; // first try failed, don't attempt again
                        dbg.push('extract:fail');
                    }
                }

                // ── TRY 2: Razor at both points + remove middle clip (fallback) ──
                var rEnd = _doRazorAt(cp.end, targetTrack, qeTrack, qeSeq, seq);
                var rStart = _doRazorAt(cp.start, targetTrack, qeTrack, qeSeq, seq);
                var numAfter = targetTrack.clips.numItems;

                if (ci === 0) {
                    dbg.push('razor=' + rEnd + '/' + rStart);
                    dbg.push('n=' + numBefore + '>' + numAfter);
                }

                if (numAfter <= numBefore) {
                    // Razor didn't create new clips — dump diagnostics for first failure
                    if (ci === 0) {
                        dbg.push('tps=' + tps);
                        dbg.push('cp=' + (Math.round(cp.start / tps * 100) / 100) + '-' + (Math.round(cp.end / tps * 100) / 100) + 's');
                        dbg.push('clip=' + (Math.round(clipStartTicks / tps * 100) / 100) + '-' + (Math.round(clipEndTicks / tps * 100) / 100) + 's');
                        dbg.push('qeT=' + (qeTrack ? typeof qeTrack.razor : 'null'));
                        dbg.push('qeS=' + (qeSeq ? typeof qeSeq.razor : 'null'));
                        dbg.push('qeX=' + (qeSeq ? typeof qeSeq.extract : 'null'));
                    }
                    continue;
                }

                // ── Razor worked! Find and remove silence clip ──
                if (mode === 'remove') {
                    var found = false;
                    for (var fc = targetTrack.clips.numItems - 1; fc >= 0; fc--) {
                        var tc = targetTrack.clips[fc];
                        var ts = Number(tc.start.ticks);
                        var te = Number(tc.end.ticks);
                        if (Math.abs(ts - cp.start) <= tol && Math.abs(te - cp.end) <= tol) {
                            removedDuration += (te - ts) / tps;
                            tc.remove(true, true);
                            cutsApplied++;
                            found = true;
                            break;
                        }
                    }
                    if (!found && ci < 3) {
                        dbg.push('miss' + ci + ':w=' + Math.round(cp.start) + '-' + Math.round(cp.end));
                    }
                } else {
                    cutsApplied++;
                }
                $.sleep(30);
            } catch (eCut) {
                dbg.push('e' + ci + ':' + eCut.toString().substring(0, 30));
            }
        }

        // Restore original in/out points (extract sets them)
        try { if (origInTicks && seq.setInPoint) seq.setInPoint(origInTicks); } catch(e) {}
        try { if (origOutTicks && seq.setOutPoint) seq.setOutPoint(origOutTicks); } catch(e) {}

        // Restore playhead
        try { seq.setPlayerPosition(origPlayhead); } catch (e) {}
        forceUIRefresh();

        var result = 'OK:' + cutsApplied + ':' + (Math.round(removedDuration * 10) / 10);
        if (dbg.length > 0) result += '|DBG:' + dbg.join(';');
        return result;
    } catch (e) {
        return 'ERROR:' + e.toString();
    }
}
