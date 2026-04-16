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
// ANCHOR POINT PRO
// Moves clip anchor point to a named position (tl, tc, tr, cl, cc, cr, bl, bc, br)
// optionally compensating position so the clip stays visually in place.
// ============================================

// Find the Motion component of a clip
function findMotionComponent(clip) {
    try {
        var components = clip.components;
        // Search by localized name
        for (var i = 0; i < components.numItems; i++) {
            var comp = components[i];
            var name = comp.displayName;
            if (name === 'Motion' || name === 'Movimento' || name === 'Movement' ||
                name === 'Bewegung' || name === 'Mouvement' || name === 'Movimiento') {
                return comp;
            }
        }
        // Fallback: check if component has Position-like property (index 1 is typically Motion)
        // Index 0 is the clip media itself, NOT Motion
        for (var i = 1; i < components.numItems; i++) {
            try {
                var props = components[i].properties;
                if (props && props.numItems >= 4) {
                    // Motion has Position, Scale, Rotation, Anchor Point — at least 4 props
                    return components[i];
                }
            } catch (e2) {}
        }
    } catch (e) {}
    return null;
}

// Find a specific property in a component by matching name (case-insensitive partial match)
function findMotionProperty(component, names) {
    try {
        // Exact match first
        for (var p = 0; p < component.properties.numItems; p++) {
            var prop = component.properties[p];
            var pname = prop.displayName;
            for (var n = 0; n < names.length; n++) {
                if (pname === names[n]) return prop;
            }
        }
        // Case-insensitive partial match fallback
        for (var p = 0; p < component.properties.numItems; p++) {
            var prop = component.properties[p];
            var pnameLow = prop.displayName.toLowerCase();
            for (var n = 0; n < names.length; n++) {
                if (pnameLow.indexOf(names[n].toLowerCase()) >= 0) return prop;
            }
        }
    } catch (e) {}
    return null;
}

// Get source clip dimensions using multiple methods
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
            if (wi >= 0) {
                var ws = md.indexOf('>', wi) + 1;
                var we = md.indexOf('<', ws);
                w = parseInt(md.substring(ws, we), 10);
            }
            var hi = md.indexOf('Column.Intrinsic.VideoHeight');
            if (hi < 0) hi = md.indexOf('VideoHeight');
            if (hi >= 0) {
                var hs = md.indexOf('>', hi) + 1;
                var he = md.indexOf('<', hs);
                h = parseInt(md.substring(hs, he), 10);
            }
            if (w > 0 && h > 0) return { width: w, height: h };
        }
    } catch (e) {}

    // Method 3: XMP metadata
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

// Apply a named anchor position to selected clips
// position: 'tl','tc','tr','cl','cc','cr','bl','bc','br'
// compensate: boolean — if true, adjust Position to keep clip visually in place
function setClipAnchorPoint(position, compensate) {
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

        for (var ci = 0; ci < clips.length; ci++) {
            var clip = clips[ci];
            try {
                var motion = findMotionComponent(clip);
                if (!motion) continue;

                var anchorProp = findMotionProperty(motion, [
                    'Anchor Point', 'Ponto de ancoragem', 'Ponto de Ancoragem',
                    'Anchor', 'Ankerpunkt', 'Point d\'ancrage', 'Punto de anclaje'
                ]);
                var positionProp = findMotionProperty(motion, [
                    'Position', 'Posição', 'Posicao', 'Posición'
                ]);
                var scaleProp = findMotionProperty(motion, [
                    'Scale', 'Escala', 'Skalierung', 'Échelle'
                ]);

                if (!anchorProp || !positionProp) continue;

                var oldAnchor = anchorProp.getValue();
                var oldPosition = positionProp.getValue();
                if (!oldAnchor || !oldPosition) continue;

                // Get REAL source dimensions (not sequence dimensions)
                var srcDims = getSourceDimensions(clip);
                var srcW, srcH;
                if (srcDims) {
                    srcW = srcDims.width;
                    srcH = srcDims.height;
                } else {
                    // Last resort: if anchor looks like it's at center, derive from it
                    // Otherwise use sequence dimensions
                    if (oldAnchor[0] > 10 && oldAnchor[1] > 10) {
                        srcW = Math.round(oldAnchor[0] * 2);
                        srcH = Math.round(oldAnchor[1] * 2);
                    } else {
                        srcW = seqW;
                        srcH = seqH;
                    }
                }

                // Calculate target anchor in source pixels
                var newAx, newAy;
                switch (position) {
                    case 'tl': newAx = 0;       newAy = 0;       break;
                    case 'tc': newAx = srcW / 2; newAy = 0;       break;
                    case 'tr': newAx = srcW;     newAy = 0;       break;
                    case 'cl': newAx = 0;        newAy = srcH / 2; break;
                    case 'cc': newAx = srcW / 2; newAy = srcH / 2; break;
                    case 'cr': newAx = srcW;     newAy = srcH / 2; break;
                    case 'bl': newAx = 0;        newAy = srcH;     break;
                    case 'bc': newAx = srcW / 2; newAy = srcH;     break;
                    case 'br': newAx = srcW;     newAy = srcH;     break;
                    default:   newAx = srcW / 2; newAy = srcH / 2;
                }

                // Get scale
                var scale = 1;
                try {
                    var sv = scaleProp ? scaleProp.getValue() : 100;
                    if (typeof sv === 'number') scale = sv / 100;
                    else if (sv instanceof Array) scale = sv[0] / 100;
                } catch (es) {}

                // Set new anchor
                anchorProp.setValue([newAx, newAy]);

                // Compensate position so clip stays visually in place
                if (compensate) {
                    var dAx = newAx - oldAnchor[0];
                    var dAy = newAy - oldAnchor[1];

                    // Position is normalized (0-1). Default center = [0.5, 0.5]
                    var newPx = oldPosition[0] + (dAx * scale) / seqW;
                    var newPy = oldPosition[1] + (dAy * scale) / seqH;

                    // Sanity check: position should stay in reasonable range (-2 to 3)
                    if (newPx > -2 && newPx < 3 && newPy > -2 && newPy < 3) {
                        positionProp.setValue([newPx, newPy]);
                    }
                }

                debug = 'src=' + srcW + 'x' + srcH + ' anchor=[' + Math.round(newAx) + ',' + Math.round(newAy) + '] pos=[' + (oldPosition[0]).toFixed(3) + ',' + (oldPosition[1]).toFixed(3) + '] scale=' + scale.toFixed(2);
                totalApplied++;
            } catch (ec) {
                debug = 'err:' + ec.toString();
            }
        }

        if (totalApplied === 0) return 'NO_APPLIED|' + debug;
        forceUIRefresh();
        return 'OK:' + totalApplied + '|' + debug;
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// Add a keyframe at the playhead for Position and Anchor Point (for animation work)
function addAnchorKeyframe() {
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
                var motion = findMotionComponent(clip);
                if (!motion) continue;

                var anchorProp = findMotionProperty(motion, [
                    'Anchor Point', 'Ponto de ancoragem', 'Ponto de Ancoragem',
                    'Anchor', 'Ankerpunkt', 'Point d\'ancrage', 'Punto de anclaje'
                ]);
                var positionProp = findMotionProperty(motion, [
                    'Position', 'Posição', 'Posicao', 'Posición'
                ]);

                // Fallback: scan by value shape if name match failed
                if (!positionProp || !anchorProp) {
                    for (var fp = 0; fp < motion.properties.numItems; fp++) {
                        try {
                            var fprop = motion.properties[fp];
                            var fname = fprop.displayName.toLowerCase();
                            if (!anchorProp && (fname.indexOf('anchor') >= 0 || fname.indexOf('ancor') >= 0)) {
                                anchorProp = fprop;
                            } else if (!positionProp && fname.indexOf('posi') >= 0) {
                                positionProp = fprop;
                            }
                        } catch (efp) {}
                    }
                }

                // Calculate clip-relative time using seconds (avoids ticks precision issues)
                var clipTime = new Time();
                try {
                    var relativeSecs = playhead.seconds - clip.start.seconds + clip.inPoint.seconds;
                    clipTime.seconds = relativeSecs;
                } catch (et) {
                    clipTime = playhead;
                }

                // Enable time-varying if not already
                try {
                    if (anchorProp && !anchorProp.isTimeVarying()) anchorProp.setTimeVarying(true);
                } catch (e) {}
                try {
                    if (positionProp && !positionProp.isTimeVarying()) positionProp.setTimeVarying(true);
                } catch (e) {}

                // Add keyframes at current time
                try { if (anchorProp) anchorProp.addKey(clipTime); } catch (e) {}
                try { if (positionProp) positionProp.addKey(clipTime); } catch (e) {}

                totalApplied++;
            } catch (ec) {}
        }

        if (totalApplied === 0) return 'NO_APPLIED';
        forceUIRefresh();
        return 'OK:' + totalApplied;
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}

// ============================================
// COPY / PASTE (Lion Copy-Pasta)
// Paste: importa arquivo(s)/imagem do clipboard do SO pra timeline
// Copy:  pega arquivo do clip selecionado e poe no clipboard do SO
// ============================================

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

                // Insere no playhead
                if (insertAtPlayhead && seq) {
                    var clipItem = _cpFindItemByPath(fp);
                    if (clipItem) {
                        try {
                            var insertTime = seq.getPlayerPosition();
                            var videoTrack = seq.videoTracks[0];
                            if (videoTrack) {
                                videoTrack.insertClip(clipItem, insertTime);
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
        var videoTrack = seq.videoTracks[0];
        if (!videoTrack) return 'NO_VIDEO_TRACK';

        videoTrack.insertClip(clipItem, insertTime);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}
