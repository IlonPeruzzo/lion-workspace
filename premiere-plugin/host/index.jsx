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

        return 'OK:' + imported + ':' + inserted + ':' + failed;
    } catch (e) { return 'ERR:' + e.toString(); }
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

// Acha track LIVRE pra o range [startTicks, endTicks). Se nenhuma livre,
// cria nova track. NUNCA retorna track com clip no range — protege contra
// overwrite/insert acidental em mídia existente.
function _cpFindFreeTrackForRange(seq, mediaType, startTicks, endTicks) {
    var tracks = (mediaType === 'audio') ? seq.audioTracks : seq.videoTracks;

    // Pra cada track existente, vê se o range [start, end) está LIVRE
    for (var t = 0; t < tracks.numTracks; t++) {
        var track = tracks[t];
        var occupied = false;
        try {
            // Track travada (lock) também não conta como livre
            if (track.isLocked && track.isLocked()) { occupied = true; }
            for (var c = 0; !occupied && c < track.clips.numItems; c++) {
                var clip = track.clips[c];
                var clipStart = Number(clip.start.ticks);
                var clipEnd = Number(clip.end.ticks);
                // Overlap se: clipStart < endTicks E clipEnd > startTicks
                if (clipStart < endTicks && clipEnd > startTicks) {
                    occupied = true;
                }
            }
        } catch (e) { occupied = true; }
        if (!occupied) return track;
    }

    // Nenhuma track livre — cria uma nova track ACIMA da última
    try {
        if (mediaType === 'audio') {
            // addTrack(numAudioTracks, audioTrackType, position)
            // Posição -1 = no fim (acima de todas)
            try { seq.audioTracks.addTrack(1); } catch(e) {
                try { seq.audioTracks.addTrack(); } catch(e2) {}
            }
            return seq.audioTracks[seq.audioTracks.numTracks - 1];
        } else {
            try { seq.videoTracks.addTrack(1); } catch(e) {
                try { seq.videoTracks.addTrack(); } catch(e2) {}
            }
            return seq.videoTracks[seq.videoTracks.numTracks - 1];
        }
    } catch (e) {
        // addTrack indisponível — última cartada: a primeira track
        // (vai usar overwriteClip de qualquer jeito, então só sobrescreve essa região)
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
                var path = "";
                try { path = foundClip.projectItem.getMediaPath(); } catch(e) {}
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
// numa track ACIMA (cria nova track se preciso). Se origem nao e timeline,
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

// Helper: itera todas as QE tracks (video ou audio) e retorna a lista de QE clips selecionados.
// Tenta 2 estratégias: 1) QE isSelected, 2) match por start time com clips selecionados do API normal.
function _lwFindSelectedQEClips(qeSeq, isVideo, seq) {
    var found = [];
    var numTracks = isVideo ? qeSeq.numVideoTracks : qeSeq.numAudioTracks;

    // Estratégia 1: QE isSelected()
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
    if (found.length > 0) return found;

    // Estratégia 2: pega clips selecionados do API normal e mapeia pra QE por start time
    if (!seq) return found;
    var trackList = isVideo ? seq.videoTracks : seq.audioTracks;
    var selStarts = [];
    for (var ti = 0; ti < trackList.numTracks; ti++) {
        var tk = trackList[ti];
        for (var ci = 0; ci < tk.clips.numItems; ci++) {
            try {
                if (tk.clips[ci].isSelected()) {
                    selStarts.push({
                        trackIdx: ti,
                        startSec: tk.clips[ci].start.seconds,
                        startTicks: String(tk.clips[ci].start.ticks),
                        name: tk.clips[ci].name,
                    });
                }
            } catch(eIs) {}
        }
    }
    for (var si = 0; si < selStarts.length; si++) {
        var sel = selStarts[si];
        var qeT = null;
        try { qeT = isVideo ? qeSeq.getVideoTrackAt(sel.trackIdx) : qeSeq.getAudioTrackAt(sel.trackIdx); } catch(eGT) {}
        if (!qeT) continue;
        var matched = null;
        var qN = 0;
        try { qN = qeT.numItems; } catch(eQN) { qN = 0; }
        for (var qj = 0; qj < qN; qj++) {
            var qit = null;
            try { qit = qeT.getItemAt(qj); } catch(eGI) {}
            if (!qit) continue;
            try { if (qit.type === 1) continue; } catch(eQT) {}
            // Tenta match por ticks (string), seconds (float), ou name
            try {
                var qts = String(qit.start.ticks);
                if (qts === sel.startTicks) { matched = qit; break; }
            } catch(eQTk) {}
            try {
                if (Math.abs(qit.start.seconds - sel.startSec) < 0.5) {
                    if (!matched) matched = qit;
                    try { if (qit.name === sel.name) { matched = qit; break; } } catch(eN2) {}
                }
            } catch(eQS) {}
        }
        if (matched) found.push(matched);
    }
    return found;
}

// Aplica um effect/transition num clip selecionado.
// kind = "video-fx" | "audio-fx" | "video-tx" | "audio-tx"
// matchName = nome do effect (em PR 2024+, é o display name retornado por getVideoEffectList)
function lwSearchApply(kind, matchName) {
    try {
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        var seq = app.project.activeSequence;
        try { app.enableQE(); } catch(e0) {}
        if (typeof qe === 'undefined' || !qe.project) return 'ERR:qe_not_available';
        var qeSeq = qe.project.getActiveSequence();
        if (!qeSeq) return 'ERR:no_qe_sequence';

        if (kind === 'video-fx') {
            // Acha QE clips selecionados (video)
            var selVideoClips = _lwFindSelectedQEClips(qeSeq, true, seq);
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

        // Conta clips selecionados
        var i, j, tk;
        for (i = 0; i < seq.videoTracks.numTracks; i++) {
            tk = seq.videoTracks[i];
            for (j = 0; j < tk.clips.numItems; j++) {
                try { if (tk.clips[j].isSelected()) ctx.videoSelected++; } catch(e) {}
            }
        }
        for (i = 0; i < seq.audioTracks.numTracks; i++) {
            tk = seq.audioTracks[i];
            for (j = 0; j < tk.clips.numItems; j++) {
                try { if (tk.clips[j].isSelected()) ctx.audioSelected++; } catch(e) {}
            }
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
        var mediaPath = '';
        try { mediaPath = item.getMediaPath ? item.getMediaPath() : ''; } catch(eMp) {}
        // Fallback dedup key: name+nodeId+proj se não tiver path
        var dedupKey = mediaPath || ((projName || '') + '|' + nodeId + '|' + name);
        if (hasAudio && !seen[dedupKey]) {
            seen[dedupKey] = 1;
            // Categoria: [Projeto] / Bin / Subbin
            var categoryParts = [];
            if (projName) categoryParts.push(projName);
            if (parentPath) categoryParts.push(parentPath);
            else categoryParts.push(hasVideo ? 'Audio+Video' : 'Audio');
            items.push({
                kind: 'audio-source',
                name: name || '(sem nome)',
                matchName: nodeId,
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

function _lwApplyPreset(presetIdentifier) {
    try {
        _lwDbgPreset('=== Apply Preset: ' + presetIdentifier + ' ===');
        if (!presetIdentifier) return 'ERR:preset_id_empty';
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        try { app.enableQE(); } catch(e0) {}
        if (typeof qe === 'undefined' || !qe.project) return 'ERR:qe_not_available';

        // Parse identifier
        var presetName = presetIdentifier;
        var presetPath = presetIdentifier;
        var hashIdx = presetIdentifier.indexOf('#');
        if (hashIdx > 0) {
            presetPath = presetIdentifier.substring(0, hashIdx);
            presetName = presetIdentifier.substring(hashIdx + 1);
        } else if (!/\.prfpset$/i.test(presetIdentifier)) {
            presetPath = '';
        }
        _lwDbgPreset('  presetName=' + presetName + ' presetPath=' + presetPath);

        var qeSeq = qe.project.getActiveSequence();
        if (!qeSeq) return 'ERR:no_qe_sequence';
        var seq = app.project.activeSequence;
        var sel = _lwFindSelectedQEClips(qeSeq, true, seq);
        var selKind = 'video';
        if (sel.length === 0) { sel = _lwFindSelectedQEClips(qeSeq, false, seq); selKind = 'audio'; }
        if (sel.length === 0) return 'NO_CLIP_SELECTED';
        _lwDbgPreset('  ' + sel.length + ' QE clips selected (' + selKind + ')');

        // Pega a versão "standard API" do clip selecionado (necessário pra applyPreset que aceita ProjectItem)
        var stdSel = [];
        try {
            var trackList = (selKind === 'video') ? seq.videoTracks : seq.audioTracks;
            for (var ti = 0; ti < trackList.numTracks; ti++) {
                var tk = trackList[ti];
                for (var ci = 0; ci < tk.clips.numItems; ci++) {
                    if (tk.clips[ci].isSelected()) stdSel.push(tk.clips[ci]);
                }
            }
        } catch(eStd) { _lwDbgPreset('  err getting std clips: ' + eStd); }
        _lwDbgPreset('  ' + stdSel.length + ' standard clips selected');

        var applied = 0;
        var allErrs = [];

        // ─── ESTRATÉGIA 1: importa container, acha preset por nome no bin, aplica via stdClip ───
        if (presetPath && presetName && /\.prfpset$/i.test(presetPath)) {
            _lwDbgPreset('  STRAT 1: import container + find preset');
            try {
                var pf = new File(presetPath);
                if (pf.exists) {
                    // Importa container — Premiere cria bin com presets dentro
                    var beforeCount = app.project.rootItem.children.numItems;
                    var importOk = false;
                    try { importOk = app.project.importFiles([presetPath]); } catch(eI) { _lwDbgPreset('    importFiles err: ' + eI); }
                    _lwDbgPreset('    importFiles=' + importOk + ' beforeCount=' + beforeCount + ' afterCount=' + app.project.rootItem.children.numItems);
                    $.sleep(150);
                    // Acha o preset por nome
                    var presetItem = _lwFindProjectItemByName(app.project.rootItem, presetName, 0);
                    _lwDbgPreset('    findByName(' + presetName + ')=' + (presetItem ? 'FOUND' : 'NOT FOUND'));
                    if (presetItem) {
                        for (var si = 0; si < stdSel.length; si++) {
                            try {
                                stdSel[si].applyPreset(presetItem);
                                applied++;
                                _lwDbgPreset('    stdClip[' + si + '].applyPreset(item) OK');
                            } catch(eAp) {
                                allErrs.push('s1.stdApplyItem:' + eAp.toString().slice(0, 60));
                                _lwDbgPreset('    stdClip[' + si + '].applyPreset err: ' + eAp);
                                // Tenta via QE clip também
                                try { sel[si].applyEffectPreset(presetItem); applied++; _lwDbgPreset('    qeClip applyEffectPreset(item) OK'); }
                                catch(eAq) { allErrs.push('s1.qeApplyItem:' + eAq.toString().slice(0, 60)); }
                            }
                        }
                    }
                }
            } catch(eS1) { allErrs.push('s1:' + eS1.toString().slice(0, 60)); _lwDbgPreset('    s1 outer err: ' + eS1); }
        }

        // ─── ESTRATÉGIA 2: applyEffectPreset(File) direto via QE — funciona pra preset individual ───
        if (applied === 0 && presetPath && /\.prfpset$/i.test(presetPath)) {
            _lwDbgPreset('  STRAT 2: qeClip.applyEffectPreset(File)');
            try {
                var pf2 = new File(presetPath);
                if (pf2.exists) {
                    for (var sj = 0; sj < sel.length; sj++) {
                        try {
                            sel[sj].applyEffectPreset(pf2);
                            applied++;
                            _lwDbgPreset('    qeClip[' + sj + '].applyEffectPreset(File) OK');
                        } catch(eA2) { allErrs.push('s2.qeApplyFile:' + eA2.toString().slice(0, 60)); _lwDbgPreset('    qeClip err: ' + eA2); }
                    }
                }
            } catch(eS2) { allErrs.push('s2:' + eS2.toString().slice(0, 60)); }
        }

        // ─── ESTRATÉGIA 3: stdClip.applyEffectPreset(filePath string) ───
        if (applied === 0 && presetPath) {
            _lwDbgPreset('  STRAT 3: stdClip.applyEffectPreset(path)');
            for (var sk = 0; sk < stdSel.length; sk++) {
                try { stdSel[sk].applyEffectPreset(presetPath); applied++; _lwDbgPreset('    OK'); }
                catch(eA3) { allErrs.push('s3.stdApplyPath:' + eA3.toString().slice(0, 60)); }
            }
        }

        // ─── ESTRATÉGIA 4: stdClip.applyPreset(filePath string) ───
        if (applied === 0 && presetPath) {
            _lwDbgPreset('  STRAT 4: stdClip.applyPreset(path)');
            for (var sl = 0; sl < stdSel.length; sl++) {
                try { stdSel[sl].applyPreset(presetPath); applied++; _lwDbgPreset('    OK'); }
                catch(eA4) { allErrs.push('s4.stdApply:' + eA4.toString().slice(0, 60)); }
            }
        }

        _lwDbgPreset('  RESULT: applied=' + applied + ' errs=' + allErrs.join(' | '));
        if (applied === 0) {
            return 'ERR:preset_apply_unsupported:' + (allErrs.length ? allErrs.slice(0, 2).join(';') : 'all methods failed');
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
function _lwInsertAudioSource(nodeIdOrName) {
    try {
        if (!app.project || !app.project.activeSequence) return 'NO_SEQUENCE';
        var seq = app.project.activeSequence;

        var foundItem = null;
        var preImportedFromFs = false;

        // Se for FS path (do SFX folder), importa direto
        if (nodeIdOrName.indexOf('FS:') === 0) {
            var fsPath = nodeIdOrName.substr(3);
            var fsFile = new File(fsPath);
            if (!fsFile.exists) return 'ERR:sfx_file_not_found:' + fsPath;
            // Verifica se já tá importado (busca por mediaPath)
            function _findFsByPath(bin, target, holder) {
                if (holder.found) return;
                try {
                    if (!bin.children || bin.children.numItems === 0) return;
                    for (var ii = 0; ii < bin.children.numItems; ii++) {
                        if (holder.found) return;
                        var it = bin.children[ii];
                        var ch = null;
                        try { ch = it.children; } catch(eC) {}
                        if (ch && ch.numItems > 0) { _findFsByPath(it, target, holder); }
                        else {
                            var p = '';
                            try { p = it.getMediaPath(); } catch(eM) {}
                            if (p === target) { holder.found = it; return; }
                        }
                    }
                } catch(e) {}
            }
            var holder = { found: null };
            _findFsByPath(app.project.rootItem, fsPath, holder);
            if (holder.found) {
                foundItem = holder.found;
            } else {
                // Importa pro projeto ativo
                var importOk = false;
                try { importOk = app.project.importFiles([fsPath]); } catch(eI) {}
                if (!importOk) return 'ERR:sfx_import_failed:' + fsPath;
                holder = { found: null };
                _findFsByPath(app.project.rootItem, fsPath, holder);
                if (!holder.found) return 'ERR:sfx_imported_not_found:' + fsPath;
                foundItem = holder.found;
                preImportedFromFs = true;
            }
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

        // Procura primeiro no projeto ativo (mais provável + insertClip funciona)
        try { _findInBin(app.project.rootItem); } catch(eA) {}

        // Se não achou, tenta nos outros projetos
        if (!foundItem) {
            for (var pp = 0; pp < projects.length && !foundItem; pp++) {
                var pr2 = projects[pp];
                if (pr2 === app.project) continue; // já tentou
                try { _findInBin(pr2.rootItem); } catch(eB) {}
            }
        }

        if (!foundItem) return 'ERR:audio_source_not_found:' + nodeIdOrName;

        // Se o item não pertence ao projeto ativo, IMPORTA o arquivo pro projeto ativo
        // pra poder fazer track.insertClip (que requer item do mesmo projeto).
        var itemBelongsToActive = false;
        try {
            // Walk up parents — se chegar no rootItem do app.project, pertence
            var p = foundItem;
            for (var pw = 0; pw < 20 && p; pw++) {
                if (p === app.project.rootItem) { itemBelongsToActive = true; break; }
                try { p = p.parent; } catch(eP) { break; }
                if (!p) break;
            }
        } catch(eBel) {}

        if (!itemBelongsToActive) {
            // Pega o file path e importa
            var mediaPath = '';
            try { mediaPath = foundItem.getMediaPath(); } catch(eMp) {}
            if (!mediaPath) {
                try { mediaPath = foundItem.canonicalUri || ''; } catch(eC) {}
            }
            if (!mediaPath) return 'ERR:could_not_get_path_other_project';

            // Helper: normaliza path pra comparar (lowercase, forward slashes, sem URL-encoding)
            function _normPath(p) {
                if (!p) return '';
                var s = String(p);
                try { s = decodeURIComponent(s); } catch(eD) {}
                s = s.replace(/\\/g, '/').toLowerCase();
                // Remove file:// prefix se houver
                s = s.replace(/^file:\/\/+/, '');
                return s;
            }
            var targetNorm = _normPath(mediaPath);
            // Extrai filename pra fallback
            var targetFilename = '';
            try {
                var slashIdx = Math.max(targetNorm.lastIndexOf('/'), targetNorm.lastIndexOf('\\'));
                targetFilename = slashIdx >= 0 ? targetNorm.substr(slashIdx + 1) : targetNorm;
            } catch(eFn) {}

            // Conta itens antes do import (pra detectar o novo)
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

            // Importa pro projeto ativo
            var importedOk = false;
            try {
                importedOk = app.project.importFiles([mediaPath], 1, app.project.getInsertionBin(), 0);
            } catch(eImp) {}
            if (!importedOk) {
                try {
                    importedOk = app.project.importFiles([mediaPath]);
                } catch(eImp2) { return 'ERR:import_failed:' + eImp2.toString().slice(0, 100); }
            }

            // Aguarda Premiere processar o import
            try { $.sleep(80); } catch(eSl) {}

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
                    if (ip && op) { inPt = ip; outPt = op; break; }
                } catch(eMT) {}
            }
            if (inPt && outPt) {
                inSec = parseFloat(inPt.seconds);
                outSec = parseFloat(outPt.seconds);
                if (outSec > inSec) {
                    insertDurationSec = outSec - inSec;
                    hasInOut = true;
                }
            }
        } catch(eIO) {}

        // Se não tem in/out válido, usa duração total
        if (insertDurationSec <= 0) {
            try {
                var dur = foundItem.getOutPoint(0); // tenta full duration
                if (dur && dur.seconds > 0) insertDurationSec = parseFloat(dur.seconds);
            } catch(eD) {}
        }
        if (insertDurationSec <= 0) {
            // Fallback: tenta via clip projecting.. assume 5s mínimo
            insertDurationSec = 5;
        }

        // Posição playhead
        var insertStartSec = 0;
        try {
            var pp = seq.getPlayerPosition();
            if (pp) insertStartSec = parseFloat(pp.seconds);
        } catch(ePh) {}
        var insertEndSec = insertStartSec + insertDurationSec;

        // Encontra audio track sem overlap
        var targetTrackIdx = -1;
        var EPS = 0.001;
        for (var ti = 0; ti < seq.audioTracks.numTracks; ti++) {
            var atk = seq.audioTracks[ti];
            var hasOverlap = false;
            for (var ci = 0; ci < atk.clips.numItems; ci++) {
                var c = atk.clips[ci];
                var cStart = 0, cEnd = 0;
                try { cStart = parseFloat(c.start.seconds); cEnd = parseFloat(c.end.seconds); } catch(eC) {}
                // Overlap se [cStart, cEnd) intersecta [insertStartSec, insertEndSec)
                if (cStart < insertEndSec - EPS && cEnd > insertStartSec + EPS) { hasOverlap = true; break; }
            }
            if (!hasOverlap) { targetTrackIdx = ti; break; }
        }

        // Se nenhuma track livre, cria nova audio track
        if (targetTrackIdx < 0) {
            try {
                // addTracks(numVid, vidIdx, numAudio, audioChType, audioIdx)
                // audioChType: 0=Mono, 1=Stereo, 2=5.1, 3=Adaptive
                seq.addTracks(0, 0, 1, 1, -1); // 1 stereo audio no final
                targetTrackIdx = seq.audioTracks.numTracks - 1;
            } catch(eAdd) {
                // Fallback via QE
                try {
                    app.enableQE();
                    var qSeq = qe.project.getActiveSequence();
                    if (qSeq && qSeq.addAudioTrack) {
                        qSeq.addAudioTrack();
                        targetTrackIdx = seq.audioTracks.numTracks - 1;
                    }
                } catch(eQA) { return 'ERR:could_not_create_audio_track:' + eAdd.toString().slice(0, 80); }
            }
        }
        if (targetTrackIdx < 0) return 'ERR:no_track_available';

        // Insere o clip no playhead
        try {
            var targetTrack = seq.audioTracks[targetTrackIdx];
            var beforeCount = targetTrack.clips.numItems;

            // Tenta insertClip primeiro
            var insertOk = false;
            var lastInsertErr = '';
            try {
                targetTrack.insertClip(foundItem, insertStartSec);
                insertOk = true;
            } catch (eIns) {
                lastInsertErr = eIns.toString();
            }

            // Fallback 1: overwriteClip (não conflita com IDs em alguns casos)
            if (!insertOk) {
                try {
                    targetTrack.overwriteClip(foundItem, insertStartSec);
                    insertOk = true;
                } catch (eOw) {
                    lastInsertErr += ' | overwrite: ' + eOw.toString();
                }
            }

            if (!insertOk) {
                // Detecta erro de "multiple projects with the same ID" e dá mensagem clara
                if (/same id|multiple projects/i.test(lastInsertErr)) {
                    return 'ERR:duplicate_project_id';
                }
                return 'ERR:insert_failed:' + lastInsertErr.slice(0, 200);
            }

            // Se tinha in/out e o clip inserido pegou tudo, ajusta o end
            if (hasInOut) {
                var newClip = null;
                for (var nci = targetTrack.clips.numItems - 1; nci >= 0; nci--) {
                    var nc = targetTrack.clips[nci];
                    var ncStart = 0;
                    try { ncStart = parseFloat(nc.start.seconds); } catch(eNS) {}
                    if (Math.abs(ncStart - insertStartSec) < 0.5) { newClip = nc; break; }
                }
                if (newClip) {
                    try { newClip.end = insertStartSec + insertDurationSec; } catch(eEd) {}
                }
            }

            return 'OK:inserted:track' + (targetTrackIdx + 1) + ':' + insertDurationSec.toFixed(2) + 's';
        } catch (eOuter) {
            return 'ERR:insert_failed:' + eOuter.toString().slice(0, 200);
        }
    } catch (eAll) {
        return 'ERR:' + eAll.toString().slice(0, 200);
    }
}

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — LIST ALL (effects + presets + audio sources merged)
// ═══════════════════════════════════════════════════════════════════
function lwSearchListAllExtended(sfxFolderPath) {
    var items = [];
    var debug = [];
    var errors = [];
    var masterDebug = [];

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
    mlog('--- listing audio sources ---');
    var audioJson = '';
    try { audioJson = lwSearchListAudioSources(sfxFolderPath || ''); } catch(e3) { errors.push('audio-call:' + e3); mlog('audio call exception: ' + e3); }
    mergeFrom(audioJson, 'aud');

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
