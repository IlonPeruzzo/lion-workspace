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
function applyBezierEasing(cp1x, cp1y, cp2x, cp2y, mode) {
    try {
        if (!app.project) return 'NO_PROJECT';
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        var clips = getSelectedClips();
        if (clips.length === 0) return 'NO_SELECTION';

        var totalApplied = 0;
        var STEPS = 10; // Number of intermediate keyframes to create

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
                            if (mode === 'between') {
                                // Process consecutive pairs
                                for (var k = 0; k < keys.length - 1; k++) {
                                    pairs.push([k, k + 1]);
                                }
                            } else {
                                // Process all consecutive pairs
                                for (var k = 0; k < keys.length - 1; k++) {
                                    pairs.push([k, k + 1]);
                                }
                            }

                            // Process each pair (in reverse to not mess up indices)
                            for (var pi = pairs.length - 1; pi >= 0; pi--) {
                                var kA = pairs[pi][0];
                                var kB = pairs[pi][1];
                                var timeA = keys[kA];
                                var timeB = keys[kB];
                                var valA = prop.getValueAtKey(timeA);
                                var valB = prop.getValueAtKey(timeB);

                                var ticksA = timeA.ticks;
                                var ticksB = timeB.ticks;
                                var duration = ticksB - ticksA;
                                if (duration <= 0) continue;

                                // Create intermediate keyframes along the bezier curve
                                for (var s = 1; s < STEPS; s++) {
                                    var fraction = s / STEPS; // 0.1, 0.2, ... 0.9
                                    var easedFraction = evalBezierCurve(fraction, cp1x, cp1y, cp2x, cp2y);

                                    // Calculate time for this intermediate point
                                    var midTicks = ticksA + Math.round(duration * fraction);
                                    var midTime = new Time();
                                    midTime.ticks = String(midTicks);

                                    // Calculate interpolated value
                                    var midVal;
                                    if (typeof valA === 'number') {
                                        midVal = valA + (valB - valA) * easedFraction;
                                    } else if (valA instanceof Array) {
                                        midVal = [];
                                        for (var d = 0; d < valA.length; d++) {
                                            midVal.push(valA[d] + (valB[d] - valA[d]) * easedFraction);
                                        }
                                    } else {
                                        continue; // Skip unsupported value types
                                    }

                                    // Add keyframe at intermediate point
                                    try {
                                        prop.addKey(midTime);
                                        prop.setValueAtKey(midTime, midVal);
                                    } catch (ek) {}
                                }
                                totalApplied++;
                            }
                        } catch (ep) {}
                    }
                }
            } catch (ec) {}
        }

        if (totalApplied === 0) return 'NO_KEYS';
        return 'OK:' + totalApplied;
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
        for (var i = 0; i < components.numItems; i++) {
            var comp = components[i];
            var name = comp.displayName;
            if (name === 'Motion' || name === 'Movimento' || name === 'Movement' ||
                name === 'Bewegung' || name === 'Mouvement' || name === 'Movimiento') {
                return comp;
            }
        }
        // Fallback: first component is typically Motion
        if (components.numItems > 0) return components[0];
    } catch (e) {}
    return null;
}

// Find a specific property in a component by matching name
function findMotionProperty(component, names) {
    try {
        for (var p = 0; p < component.properties.numItems; p++) {
            var prop = component.properties[p];
            var pname = prop.displayName;
            for (var n = 0; n < names.length; n++) {
                if (pname === names[n]) return prop;
            }
        }
    } catch (e) {}
    return null;
}

// Get clip source dimensions (best-effort)
function getClipDimensions(clip) {
    var seq = app.project.activeSequence;
    var w = 1920, h = 1080;
    try {
        w = seq.frameSizeHorizontal || 1920;
        h = seq.frameSizeVertical || 1080;
    } catch (e) {}

    // Try to get actual source dimensions from the project item
    try {
        var pi = clip.projectItem;
        if (pi) {
            // Try QE DOM
            try {
                app.enableQE();
                var qep = qe.project;
                // Search for this clip in active sequence tracks
                // This is unreliable but try it
            } catch (eq) {}
        }
    } catch (epi) {}

    return { width: w, height: h };
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
                    'Ankerpunkt', 'Point d\'ancrage', 'Punto de anclaje'
                ]);
                var positionProp = findMotionProperty(motion, [
                    'Position', 'Posição', 'Posicao', 'Posición'
                ]);
                var scaleProp = findMotionProperty(motion, [
                    'Scale', 'Escala', 'Skalierung', 'Échelle'
                ]);

                // Fallbacks by typical index order (Position=0, Scale=1, ..., Anchor=4)
                if (!positionProp) { try { positionProp = motion.properties[0]; } catch (e) {} }
                if (!scaleProp)    { try { scaleProp = motion.properties[1]; } catch (e) {} }
                if (!anchorProp)   { try { anchorProp = motion.properties[4]; } catch (e) {} }

                if (!anchorProp || !positionProp) continue;

                // Get source dimensions
                var dims = getClipDimensions(clip);
                var srcW = dims.width;
                var srcH = dims.height;

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
                    anchorProp.setValue([newAx, newAy], true);
                } catch (ea) {
                    // Try without second parameter
                    try { anchorProp.setValue([newAx, newAy]); } catch (ea2) { continue; }
                }

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
                        positionProp.setValue([newPx, newPy], true);
                    } catch (ep) {
                        try { positionProp.setValue([newPx, newPy]); } catch (ep2) {}
                    }
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
                    'Anchor Point', 'Ponto de ancoragem', 'Ponto de Ancoragem'
                ]);
                var positionProp = findMotionProperty(motion, [
                    'Position', 'Posição', 'Posicao'
                ]);

                if (!positionProp) { try { positionProp = motion.properties[0]; } catch (e) {} }
                if (!anchorProp)   { try { anchorProp = motion.properties[4]; } catch (e) {} }

                // Calculate clip-relative time
                var clipTime = new Time();
                try {
                    var clipStart = clip.start.ticks;
                    var ph = playhead.ticks;
                    var inPoint = clip.inPoint.ticks;
                    var relativeTicks = String(Number(ph) - Number(clipStart) + Number(inPoint));
                    clipTime.ticks = relativeTicks;
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

        // Get active sequence
        var seq = app.project.activeSequence;
        if (!seq) return 'NO_SEQUENCE';

        // Find the clip in the project (search all bins recursively)
        var f = new File(filePath);
        var clipItem = null;
        var fileName = f.displayName.replace(/\.[^.]+$/, '');

        function searchBin(bin) {
            for (var i = 0; i < bin.children.numItems; i++) {
                var child = bin.children[i];
                if (child.type === 2) {
                    // It's a bin, recurse
                    var found = searchBin(child);
                    if (found) return found;
                } else {
                    // Match by name
                    if (child.name === f.displayName || child.name === fileName) {
                        return child;
                    }
                }
            }
            return null;
        }

        clipItem = searchBin(app.project.rootItem);
        if (!clipItem) return 'CLIP_NOT_FOUND';

        // Get playhead position (current time indicator)
        var insertTime = seq.getPlayerPosition();

        // Insert into first video track (V1)
        var videoTrack = seq.videoTracks[0];
        if (!videoTrack) return 'NO_VIDEO_TRACK';

        videoTrack.insertClip(clipItem, insertTime);
        return 'OK';
    } catch (e) {
        return 'ERROR: ' + e.toString();
    }
}
