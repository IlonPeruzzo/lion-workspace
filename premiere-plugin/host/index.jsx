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

// Get clip source dimensions (best-effort)
function getClipDimensions(clip) {
    var seq = app.project.activeSequence;
    var fallbackW = 1920, fallbackH = 1080;
    try {
        fallbackW = parseInt(seq.frameSizeHorizontal, 10) || 1920;
        fallbackH = parseInt(seq.frameSizeVertical, 10) || 1080;
    } catch (e) {}

    // Try to get actual source dimensions from projectItem metadata
    try {
        var pi = clip.projectItem;
        if (pi) {
            // Premiere 2020+: column metadata
            try {
                var vw = parseInt(pi.getMetadataValue('Column.Intrinsic.VideoWidth'), 10);
                var vh = parseInt(pi.getMetadataValue('Column.Intrinsic.VideoHeight'), 10);
                if (vw > 0 && vh > 0) return { width: vw, height: vh };
            } catch (e1) {}
            // Try footage interpretation
            try {
                var interp = pi.getFootageInterpretation();
                if (interp && interp.frameSize) {
                    var fw = interp.frameSize.width || interp.frameSize[0];
                    var fh = interp.frameSize.height || interp.frameSize[1];
                    if (fw > 0 && fh > 0) return { width: fw, height: fh };
                }
            } catch (e2) {}
        }
    } catch (epi) {}

    return { width: fallbackW, height: fallbackH };
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

        var seqW = 1920, seqH = 1080;
        try {
            seqW = seq.frameSizeHorizontal || 1920;
            seqH = seq.frameSizeVertical || 1080;
        } catch (e) {}

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
                var scaleProp = findMotionProperty(motion, [
                    'Scale', 'Escala', 'Skalierung', 'Échelle'
                ]);

                // Fallbacks: scan properties by value shape
                if (!positionProp || !anchorProp) {
                    for (var fp = 0; fp < motion.properties.numItems; fp++) {
                        try {
                            var fprop = motion.properties[fp];
                            var fval = fprop.getValue();
                            if (fval instanceof Array && fval.length === 2) {
                                var fname = fprop.displayName.toLowerCase();
                                if (!anchorProp && (fname.indexOf('anchor') >= 0 || fname.indexOf('ancor') >= 0)) {
                                    anchorProp = fprop;
                                } else if (!positionProp && fname.indexOf('posi') >= 0) {
                                    positionProp = fprop;
                                }
                            }
                            if (!scaleProp) {
                                var sname = fprop.displayName.toLowerCase();
                                if (sname.indexOf('scale') >= 0 || sname.indexOf('escala') >= 0) {
                                    scaleProp = fprop;
                                }
                            }
                        } catch (efp) {}
                    }
                }

                if (!anchorProp || !positionProp) continue;

                // Get source dimensions — try deriving from current anchor if at default center
                var dims = getClipDimensions(clip);
                var srcW = dims.width;
                var srcH = dims.height;
                try {
                    var curAnchor = anchorProp.getValue();
                    // If anchor is at default center, we can derive true source dims
                    if (curAnchor && curAnchor.length >= 2 && curAnchor[0] > 1 && curAnchor[1] > 1) {
                        srcW = Math.round(curAnchor[0] * 2);
                        srcH = Math.round(curAnchor[1] * 2);
                    }
                } catch (ead) {}

                // Calculate target anchor in clip-source pixels
                var newAx, newAy;
                switch (position) {
                    case 'tl': newAx = 0;      newAy = 0;      break;
                    case 'tc': newAx = srcW/2; newAy = 0;      break;
                    case 'tr': newAx = srcW;   newAy = 0;      break;
                    case 'cl': newAx = 0;      newAy = srcH/2; break;
                    case 'cc': newAx = srcW/2; newAy = srcH/2; break;
                    case 'cr': newAx = srcW;   newAy = srcH/2; break;
                    case 'bl': newAx = 0;      newAy = srcH;   break;
                    case 'bc': newAx = srcW/2; newAy = srcH;   break;
                    case 'br': newAx = srcW;   newAy = srcH;   break;
                    default:   newAx = srcW/2; newAy = srcH/2;
                }

                // Read current values
                var oldAnchor = null, oldPosition = null, scale = 1;
                try { oldAnchor = anchorProp.getValue(); } catch (e) {}
                try { oldPosition = positionProp.getValue(); } catch (e) {}
                try {
                    var sv = scaleProp ? scaleProp.getValue() : 100;
                    scale = (typeof sv === 'number') ? sv / 100 : 1;
                } catch (e) {}

                // Set new anchor point
                try {
                    anchorProp.setValue([newAx, newAy]);
                } catch (ea) { continue; }

                // Compensate position if requested
                if (compensate && oldAnchor && oldPosition && oldAnchor.length >= 2 && oldPosition.length >= 2) {
                    var deltaX = (newAx - oldAnchor[0]) * scale;
                    var deltaY = (newAy - oldAnchor[1]) * scale;

                    // Detect if Position is normalized (0-1 range) or pixel-based
                    // In Premiere Pro ExtendScript, Position is normalized (default [0.5, 0.5])
                    var px = oldPosition[0];
                    var py = oldPosition[1];
                    var isNormalized = (Math.abs(px) <= 5 && Math.abs(py) <= 5);

                    var newPx, newPy;
                    if (isNormalized) {
                        newPx = px + deltaX / seqW;
                        newPy = py + deltaY / seqH;
                    } else {
                        newPx = px + deltaX;
                        newPy = py + deltaY;
                    }

                    try {
                        positionProp.setValue([newPx, newPy]);
                    } catch (ep) {}
                }

                totalApplied++;
            } catch (ec) {}
        }

        if (totalApplied === 0) return 'NO_APPLIED';
        return 'OK:' + totalApplied;
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
        return 'OK:' + totalApplied;
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
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
        var videoTrack = seq.videoTracks[0];
        if (!videoTrack) return 'NO_VIDEO_TRACK';

        videoTrack.insertClip(clipItem, insertTime);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}
