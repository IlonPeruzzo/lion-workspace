// ════════════════════════════════════════════════════════════════════
// MOTION TRACKER — Lion Workspace
//
// Trackeia ponto/bbox em vídeos usando OpenCV.js (WASM).
//
// Algoritmos:
//   - KLT (Lucas-Kanade) com pirâmide — point tracker
//   - Template matching com atualização periódica — bbox tracker (substitui CSRT)
//
// Smoothing:
//   - Whittaker — penaliza segunda derivada (suave + segue dados)
//   - None
//
// Dependências:
//   - @techstark/opencv-js (WASM, ~7MB)
//   - ffmpeg/ffprobe (já bundlado no app pra YouTube/bg-remove)
// ════════════════════════════════════════════════════════════════════
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn } = require('child_process');

let _cv = null;
let _cvLoading = null;

// CRÍTICO: o módulo opencv-js é um THENABLE (tem .then) mesmo quando não é
// Promise. Async functions seguem a cadeia .then em return/await → engasga.
// SOLUÇÃO: funções async NUNCA retornam _cv direto. Retornam `true` (primitivo).
// Callers acessam _cv via getCv() (sync getter).
async function _ensureCv() {
    if (_cv && _cv.Mat) return true;
    if (_cvLoading) return _cvLoading;
    _cvLoading = (async () => {
        console.log('[motion-tracker] _ensureCv: requiring @techstark/opencv-js...');
        const cvModule = require('@techstark/opencv-js');
        console.log('[motion-tracker] _ensureCv: required, is Promise=' + (cvModule instanceof Promise) + ' has then=' + (typeof cvModule.then === 'function'));
        if (cvModule instanceof Promise) {
            const c = await cvModule;
            _cv = c;
            console.log('[motion-tracker] _ensureCv: resolved from Promise');
        } else {
            console.log('[motion-tracker] _ensureCv: awaiting onRuntimeInitialized callback');
            await new Promise((resolve) => {
                cvModule.onRuntimeInitialized = () => {
                    console.log('[motion-tracker] _ensureCv: onRuntimeInitialized fired, resolving');
                    resolve();
                };
            });
            _cv = cvModule;
        }
        console.log('[motion-tracker] _ensureCv: stored in _cv (sync), returning true');
        return true;  // primitivo — não tem .then, não trava o async
    })();
    return _cvLoading;
}

function getCv() {
    if (!_cv || !_cv.Mat) throw new Error('OpenCV not ready — call _ensureCv first');
    return _cv;
}

// ─── Frame extraction via ffmpeg ────────────────────────────────────
// Retorna iterator: yield { frameIdx, mat (cv.Mat RGBA) } por frame
// Usa rawvideo pipe pra evitar I/O em disco (rápido)
async function* _streamFrames(videoPath, ffmpegPath, startFrame, endFrame, width, height, fps) {
    await _ensureCv();
    const cv = getCv();
    const filters = [];
    if (startFrame > 0 || (endFrame > 0 && endFrame > startFrame)) {
        const end = (endFrame > 0) ? endFrame : '9999999';
        filters.push(`select=between(n\\,${startFrame}\\,${end})`);
    }
    const args = [
        '-i', videoPath,
        ...(filters.length ? ['-vf', filters.join(',')] : []),
        '-vsync', '0',
        '-f', 'rawvideo',
        '-pix_fmt', 'rgba',
        '-'
    ];
    const ff = spawn(ffmpegPath, args, { stdio: ['ignore', 'pipe', 'pipe'] });
    const frameSize = width * height * 4;
    let buf = Buffer.alloc(0);
    let frameIdx = startFrame;

    const chunks = [];
    let resolvePending = null;
    let endedWithError = null;
    let done = false;

    ff.stdout.on('data', chunk => {
        chunks.push(chunk);
        if (resolvePending) { const r = resolvePending; resolvePending = null; r(); }
    });
    ff.on('close', code => {
        done = true;
        if (code !== 0 && code !== null) endedWithError = new Error('ffmpeg exit ' + code);
        if (resolvePending) { const r = resolvePending; resolvePending = null; r(); }
    });
    ff.on('error', e => { endedWithError = e; done = true; if (resolvePending) { const r = resolvePending; resolvePending = null; r(); } });

    try {
        while (true) {
            // wait for enough bytes or end
            while (buf.length < frameSize && !done) {
                if (chunks.length > 0) {
                    buf = Buffer.concat([buf, ...chunks.splice(0)]);
                } else {
                    await new Promise(res => { resolvePending = res; });
                }
            }
            if (buf.length < frameSize) {
                if (endedWithError) throw endedWithError;
                break;
            }
            const frameBytes = buf.subarray(0, frameSize);
            buf = buf.subarray(frameSize);
            const mat = cv.matFromArray(height, width, cv.CV_8UC4, frameBytes);
            yield { frameIdx, mat };
            frameIdx++;
        }
    } finally {
        try { ff.kill('SIGKILL'); } catch (e) {}
    }
}

// ─── Probe metadata ──────────────────────────────────────────────────
async function probeVideo(videoPath, ffprobePath) {
    console.log('[motion-tracker] probeVideo: spawning ffprobe at ' + ffprobePath + ' for ' + videoPath);
    return new Promise((resolve, reject) => {
        const args = ['-v', 'error', '-select_streams', 'v:0',
            '-show_entries', 'stream=width,height,r_frame_rate,nb_frames,duration',
            '-show_entries', 'format=duration',
            '-of', 'json', videoPath];
        let fp;
        try { fp = spawn(ffprobePath, args); }
        catch (eS) { console.error('[motion-tracker] probeVideo spawn err:', eS); reject(eS); return; }
        let out = '';
        let errOut = '';
        fp.stdout.on('data', d => out += d);
        fp.stderr.on('data', d => errOut += d);
        fp.on('close', code => {
            console.log('[motion-tracker] probeVideo: ffprobe exit code=' + code + ' stdout.len=' + out.length + ' stderr.len=' + errOut.length);
            if (errOut) console.log('[motion-tracker] probeVideo stderr:', errOut.slice(0, 300));
            if (code !== 0) return reject(new Error('ffprobe exit ' + code + ': ' + errOut.slice(0, 200)));
            try {
                const j = JSON.parse(out);
                const s = j.streams[0];
                const [n, d] = (s.r_frame_rate || '30/1').split('/').map(Number);
                const meta = {
                    width: s.width, height: s.height,
                    fps: n / (d || 1),
                    frameCount: parseInt(s.nb_frames, 10) || 0,
                    duration: parseFloat(s.duration || j.format?.duration || 0),
                };
                console.log('[motion-tracker] probeVideo: ' + JSON.stringify(meta));
                resolve(meta);
            } catch (e) { reject(e); }
        });
        fp.on('error', e => { console.error('[motion-tracker] probeVideo runtime err:', e); reject(e); });
    });
}

// ─── Whittaker smoothing — pure JS ──────────────────────────────────
// Resolve (I + λ DᵀD) y_smooth = y_raw onde D é segunda derivada.
// É um sistema linear pentadiagonal — Thomas algorithm modificado seria
// O(n), mas pra n<10000 (vídeo curto) Gauss eliminação simples já basta.
function _whittakerSmooth(y, lambda = 100) {
    const n = y.length;
    if (n < 3) return y.slice();
    // Build A = I + λ DᵀD (banded matrix). Para D = 2nd diff,
    // DᵀD é simétrica banda-2 com diagonal [1,5,6,...,6,5,1] e off-diags.
    // Simplificação: usa eliminação de Gauss densa em buffer Float64.
    // Não é o mais eficiente mas robusto pra n ≤ ~5000 frames.
    const A = new Float64Array(n * n);
    for (let i = 0; i < n; i++) A[i * n + i] = 1; // I
    // Add λ * DᵀD onde D[i, i]=1, D[i, i+1]=-2, D[i, i+2]=1 (2nd diff)
    for (let i = 0; i < n - 2; i++) {
        const r = [1, -2, 1];
        for (let a = 0; a < 3; a++) for (let b = 0; b < 3; b++) {
            A[(i + a) * n + (i + b)] += lambda * r[a] * r[b];
        }
    }
    // Solve A x = y via Gauss elimination com pivot parcial
    const b = Array.from(y);
    for (let i = 0; i < n; i++) {
        // partial pivot
        let max = Math.abs(A[i * n + i]), maxRow = i;
        for (let k = i + 1; k < n; k++) {
            if (Math.abs(A[k * n + i]) > max) { max = Math.abs(A[k * n + i]); maxRow = k; }
        }
        if (maxRow !== i) {
            for (let c = i; c < n; c++) {
                const t = A[i * n + c]; A[i * n + c] = A[maxRow * n + c]; A[maxRow * n + c] = t;
            }
            const t = b[i]; b[i] = b[maxRow]; b[maxRow] = t;
        }
        // eliminate below
        for (let k = i + 1; k < n; k++) {
            const factor = A[k * n + i] / A[i * n + i];
            if (factor === 0) continue;
            for (let c = i; c < n; c++) A[k * n + c] -= factor * A[i * n + c];
            b[k] -= factor * b[i];
        }
    }
    // back substitution
    const x = new Array(n);
    for (let i = n - 1; i >= 0; i--) {
        let sum = b[i];
        for (let c = i + 1; c < n; c++) sum -= A[i * n + c] * x[c];
        x[i] = sum / A[i * n + i];
    }
    return x;
}

// ─── Point tracker (Lucas-Kanade) ───────────────────────────────────
// points: [{ x, y }] — pontos iniciais (1 ou mais)
// Retorna: [{ frame, time, points: [{x, y, confidence, lost}] }]
async function trackPointsKLT(videoPath, ffmpegPath, ffprobePath, points, options = {}) {
    console.log('[motion-tracker] trackPointsKLT entered: videoPath=' + videoPath);
    console.log('[motion-tracker] trackPointsKLT: ffmpegPath=' + ffmpegPath + ' ffprobePath=' + ffprobePath);
    console.log('[motion-tracker] trackPointsKLT: points=' + points.length + ' options=' + JSON.stringify(options));
    const opts = Object.assign({
        startFrame: 0,
        endFrame: -1,
        windowSize: 21,
        maxLevel: 5,
        criteriaEps: 0.03,
        maxIterations: 30,
        confidenceThreshold: 0.3,
        maxLostFrames: 5,
        smoothing: 'whittaker',  // 'whittaker' | 'none'
        whittakerLambda: 100,
        onProgress: null,  // (currentFrame, totalFrames) => void
    }, options);

    await _ensureCv();
    const cv = getCv();
    console.log('[motion-tracker] trackPointsKLT: cv obtained, calling probeVideo');
    const meta = await probeVideo(videoPath, ffprobePath);
    const endFrame = opts.endFrame > 0 ? Math.min(opts.endFrame, meta.frameCount - 1) : meta.frameCount - 1;
    const totalToProcess = endFrame - opts.startFrame + 1;

    const winSize = new cv.Size(opts.windowSize, opts.windowSize);
    const criteria = new cv.TermCriteria(
        cv.TERM_CRITERIA_EPS | cv.TERM_CRITERIA_COUNT,
        opts.maxIterations, opts.criteriaEps
    );

    // Per-point state
    const numPoints = points.length;
    const tracks = points.map((p, i) => ({
        id: i,
        history: [],         // [{ frame, time, x, y, confidence, lost }]
        currentXY: [p.x, p.y],
        consecutiveLost: 0,
        active: true,
    }));

    // Seed first frame
    tracks.forEach((tr, i) => {
        tr.history.push({
            frame: opts.startFrame,
            time: opts.startFrame / meta.fps,
            x: tr.currentXY[0], y: tr.currentXY[1],
            confidence: 1.0, lost: false,
        });
    });

    let prevGray = null;
    let processed = 0;

    for await (const { frameIdx, mat } of _streamFrames(
        videoPath, ffmpegPath, opts.startFrame, endFrame,
        meta.width, meta.height, meta.fps
    )) {
        const gray = new cv.Mat();
        cv.cvtColor(mat, gray, cv.COLOR_RGBA2GRAY);
        mat.delete();

        if (prevGray === null) {
            // first frame — initial point already recorded
            prevGray = gray;
            processed++;
            if (opts.onProgress) opts.onProgress(processed, totalToProcess);
            continue;
        }

        // Build prev points array (only active tracks)
        const activeTracks = tracks.filter(t => t.active);
        if (activeTracks.length === 0) {
            gray.delete();
            break; // all lost
        }
        const prevPts = cv.matFromArray(activeTracks.length, 1, cv.CV_32FC2,
            activeTracks.flatMap(t => t.currentXY));
        const nextPts = new cv.Mat();
        const status = new cv.Mat();
        const err = new cv.Mat();

        try {
            cv.calcOpticalFlowPyrLK(prevGray, gray, prevPts, nextPts, status, err,
                winSize, opts.maxLevel, criteria);
        } catch (e) {
            console.error('[motion-tracker] KLT err:', e);
            prevPts.delete(); nextPts.delete(); status.delete(); err.delete(); gray.delete();
            break;
        }

        // Validate via backward flow
        const backPts = new cv.Mat();
        const backStatus = new cv.Mat();
        const backErr = new cv.Mat();
        try {
            cv.calcOpticalFlowPyrLK(gray, prevGray, nextPts, backPts, backStatus, backErr,
                winSize, opts.maxLevel, criteria);
        } catch (e) { /* tolerate, just no back check */ }

        // Update each active track
        const time = frameIdx / meta.fps;
        for (let i = 0; i < activeTracks.length; i++) {
            const tr = activeTracks[i];
            const nx = nextPts.data32F[i * 2];
            const ny = nextPts.data32F[i * 2 + 1];
            const ok = status.data[i] === 1;

            let confidence = 0;
            let lost = !ok;

            if (ok) {
                // Backward error
                let backError = 0;
                if (backPts.data32F && backStatus.data[i] === 1) {
                    const bx = backPts.data32F[i * 2];
                    const by = backPts.data32F[i * 2 + 1];
                    backError = Math.hypot(bx - tr.currentXY[0], by - tr.currentXY[1]);
                }
                // KLT err
                const klErr = err.data32F[i] || 0;
                const errorConf = Math.max(0, 1 - klErr / 100);
                const backConf = Math.max(0, 1 - backError / 5);
                // Bounds
                const inBounds = nx >= 0 && nx < meta.width && ny >= 0 && ny < meta.height;
                confidence = inBounds ? (errorConf * 0.6 + backConf * 0.4) : 0;
                if (confidence < opts.confidenceThreshold) lost = true;
            }

            if (lost) {
                tr.consecutiveLost++;
                tr.history.push({ frame: frameIdx, time, x: null, y: null, confidence: 0, lost: true });
                if (tr.consecutiveLost >= opts.maxLostFrames) tr.active = false;
            } else {
                tr.consecutiveLost = 0;
                tr.currentXY = [nx, ny];
                tr.history.push({ frame: frameIdx, time, x: nx, y: ny, confidence, lost: false });
            }
        }

        prevPts.delete(); nextPts.delete(); status.delete(); err.delete();
        backPts.delete(); backStatus.delete(); backErr.delete();
        prevGray.delete();
        prevGray = gray;

        processed++;
        if (opts.onProgress && processed % 5 === 0) opts.onProgress(processed, totalToProcess);
    }
    if (prevGray) prevGray.delete();

    // Smoothing per track
    if (opts.smoothing === 'whittaker') {
        for (const tr of tracks) {
            const validIdx = [];
            const xs = [], ys = [];
            tr.history.forEach((p, i) => {
                if (!p.lost && p.x != null) { validIdx.push(i); xs.push(p.x); ys.push(p.y); }
            });
            if (xs.length < 3) continue;
            const xS = _whittakerSmooth(xs, opts.whittakerLambda);
            const yS = _whittakerSmooth(ys, opts.whittakerLambda);
            validIdx.forEach((origIdx, k) => {
                tr.history[origIdx].x = xS[k];
                tr.history[origIdx].y = yS[k];
            });
        }
    }

    return {
        success: true,
        algorithm: 'KLT',
        videoPath,
        metadata: meta,
        tracks: tracks.map(t => ({ id: t.id, history: t.history })),
    };
}

module.exports = {
    probeVideo,
    trackPointsKLT,
    _whittakerSmooth, // exposed pra testes
    _ensureCv,
};
