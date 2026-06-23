// ════════════════════════════════════════════════════════════════════
// Motion Tracker Worker — roda OpenCV.js (WASM) num Worker Thread
// isolado, pra NÃO bloquear o main process do Electron.
//
// Sem isso, o HTTP server para de responder durante o tracking →
// plugin Premiere marca "Desconectado" e Electron mostra "Não respondendo".
//
// Mensagem in:  { reqId, filePath, points, options }
// Mensagem out: { reqId, ok, result?, error? }
// ════════════════════════════════════════════════════════════════════
// Script rodando como child_process FORK — processo Node completamente
// isolado. Comunicação via process.on('message') / process.send().
// Sem libuv compartilhado, WASM do OpenCV.js inicializa sem problema
// (worker_threads tem incompatibilidade com Emscripten WASM init).

// CRÍTICO: stdout/stderr do Node são BUFFERIZADOS quando vão pra pipe
// (não TTY). Isso fazia parecer que o fork tava congelado — na verdade
// os logs ficavam na fila e só flushavam quando o buffer enchia ou o
// processo morria. Força blocking I/O pra logs aparecerem em real-time.
try { if (process.stdout._handle && process.stdout._handle.setBlocking) process.stdout._handle.setBlocking(true); } catch (e) {}
try { if (process.stderr._handle && process.stderr._handle.setBlocking) process.stderr._handle.setBlocking(true); } catch (e) {}

let tracker = null;
let loadError = null;
try {
    console.log('[mt-worker:fork] requiring motion-tracker.js...');
    tracker = require('./motion-tracker.js');
    console.log('[mt-worker:fork] module loaded OK');
} catch (e) {
    loadError = e;
    console.error('[mt-worker:fork] LOAD FAIL:', e && (e.stack || e.message));
}

function send(msg) {
    try { process.send(msg); } catch (e) { console.error('[mt-worker:fork] send err:', e); }
}

// Wrapper que escreve direto no stdout (bypassa qualquer buffering de console.log)
function flog(msg) {
    try { process.stdout.write('[FLOG] ' + msg + '\n'); } catch (e) {}
}

// Heartbeat — confirma que o event loop tá vivo
setInterval(() => flog('heartbeat ' + new Date().toISOString()), 2000);

process.on('uncaughtException', (e) => { flog('UNCAUGHT EX: ' + (e.stack || e.message)); });
process.on('unhandledRejection', (e) => { flog('UNHANDLED REJ: ' + (e && (e.stack || e.message) || e)); });

process.on('message', async (msg) => {
    if (!msg || msg.reqId == null) return;
    flog('msg in: reqId=' + msg.reqId);
    if (loadError) {
        send({ reqId: msg.reqId, ok: false, error: 'load: ' + (loadError.message || String(loadError)) });
        return;
    }
    const { reqId, filePath, ffmpegPath, ffprobePath, points, options } = msg;
    try {
        flog('STEP 1: about to await _ensureCv');
        const cv = await tracker._ensureCv();
        flog('STEP 2: _ensureCv resolved, cv has Mat=' + (cv && !!cv.Mat));
        flog('STEP 3: about to call trackPointsKLT');
        const result = await tracker.trackPointsKLT(
            filePath, ffmpegPath, ffprobePath, points || [], options || {}
        );
        flog('STEP 4: trackPointsKLT done, points=' + (result && result.tracks ? result.tracks.length : 0));
        send({ reqId, ok: true, result });
    } catch (e) {
        flog('CATCH: ' + (e && (e.stack || e.message) || String(e)));
        send({ reqId, ok: false, error: (e && (e.stack || e.message)) || String(e) });
    }
});

console.log('[mt-worker:fork] ready, sending init msg');
send({ ready: true });
