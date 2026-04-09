const { app, BrowserWindow, ipcMain, screen } = require('electron');
const path = require('path');
const http = require('http');
const fs = require('fs');
const os = require('os');
const { exec, spawn } = require('child_process');

const isWin = process.platform === 'win32';
const isMac = process.platform === 'darwin';

/* ═══════════════════════════════════════════════════════════════════
   FOREGROUND PROCESS DETECTION (platform-aware)
   ─────────────────────────────────────────────────────────────────
   Windows: A persistent PowerShell daemon (single process) compiles
   Add-Type ONCE and then loops, writing the current foreground
   process name to a temp file every ~1.2 s.  Node reads this file
   instantly — no more spawning a new shell on every check.

   macOS: Uses osascript on-demand (fast, no compilation needed).
   ═══════════════════════════════════════════════════════════════════ */

const fgFile = path.join(os.tmpdir(), `lion-fg-${process.pid}.txt`);
let fgDaemon = null;
let daemonScriptPath = null;
let minimizeScriptPath = null;

if (isWin) {
    // Daemon script — runs in a loop, writes foreground process name to file
    daemonScriptPath = path.join(os.tmpdir(), `lion-fgd-${process.pid}.ps1`);
    fs.writeFileSync(daemonScriptPath, [
        "try{Add-Type -Name WFG -Namespace NFG -MemberDefinition '[DllImport(\"user32.dll\")]public static extern IntPtr GetForegroundWindow();[DllImport(\"user32.dll\")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);'}catch{}",
        "$f='" + fgFile.replace(/'/g, "''") + "'",
        "while($true){",
        "  try{",
        "    $h=[NFG.WFG]::GetForegroundWindow();$p=[uint32]0",
        "    [NFG.WFG]::GetWindowThreadProcessId($h,[ref]$p)|Out-Null",
        "    $n=(Get-Process -Id $p -EA 0).ProcessName",
        "    if($n){[IO.File]::WriteAllText($f,$n)}",
        "  }catch{}",
        "  Start-Sleep -Milliseconds 1200",
        "}"
    ].join("\n"));

    // Minimize script (used by focus mode to hide non-allowed windows)
    minimizeScriptPath = path.join(os.tmpdir(), `lion-minimize-${process.pid}.ps1`);
    fs.writeFileSync(minimizeScriptPath, [
        "try{Add-Type -Name WM -Namespace NM -MemberDefinition '[DllImport(\"user32.dll\")]public static extern bool ShowWindow(IntPtr h,int c);[DllImport(\"user32.dll\")]public static extern IntPtr GetForegroundWindow();'}catch{}",
        "[NM.WM]::ShowWindow([NM.WM]::GetForegroundWindow(), 6)"
    ].join("\n"));
}

function startFgDaemon() {
    if (!isWin || fgDaemon) return;
    fgDaemon = spawn('powershell', [
        '-NoProfile', '-WindowStyle', 'Hidden',
        '-ExecutionPolicy', 'Bypass', '-File', daemonScriptPath
    ], { windowsHide: true, stdio: 'ignore' });

    fgDaemon.on('exit', () => {
        fgDaemon = null;
        if (!app.isQuitting) setTimeout(startFgDaemon, 3000);
    });
    fgDaemon.on('error', () => {});
}

function stopFgDaemon() {
    if (fgDaemon) { fgDaemon.kill(); fgDaemon = null; }
    try { fs.unlinkSync(fgFile); } catch {}
    try { if (daemonScriptPath) fs.unlinkSync(daemonScriptPath); } catch {}
    try { if (minimizeScriptPath) fs.unlinkSync(minimizeScriptPath); } catch {}
}

/** Get current foreground process name (async callback API for cross-platform) */
function getFgProc(cb) {
    if (isWin) {
        try { cb(fs.readFileSync(fgFile, 'utf8').trim()); }
        catch { cb(''); }
    } else if (isMac) {
        exec("osascript -e 'tell application \"System Events\" to get name of first application process whose frontmost is true'",
            { timeout: 3000 },
            (err, stdout) => cb(err ? '' : (stdout || '').trim())
        );
    } else {
        cb('');
    }
}

// Start daemon right away (before app ready — it doesn't use Electron APIs)
if (isWin) startFgDaemon();

/* ═══════ State ═══════ */
let mainWin = null;
let overlays = [];
let focusActive = false;
let focusExternal = false;
let focusInterval = null;
let originalBounds = null;

/* ═══════ Allowed / Work apps ═══════ */
const ALLOWED_APPS = [
    'lion workspace', 'electron',
    // Adobe Suite
    'adobe premiere pro', 'premiere pro', 'premiere', 'afterfx', 'after effects',
    'photoshop', 'illustrator', 'indesign', 'lightroom', 'media encoder',
    'adobe audition', 'audition', 'animate', 'adobe animate', 'character animator',
    'adobe xd', 'substance', 'adobe bridge', 'bridge',
    // 3D / VFX
    'cinema 4d', 'cinema4d', 'c4d', 'blender', 'davinci resolve', 'resolve',
    'nuke', 'fusion', 'houdini', 'maya', '3ds max', 'zbrush', 'unreal', 'unity',
    // Windows system
    'explorer', 'files', 'searchui', 'shellexperiencehost',
    'startmenuexperiencehost', 'applicationframehost',
    'textinputhost', 'runtimebroker', 'taskmgr',
    // macOS system
    'finder', 'spotlight', 'system preferences', 'system settings',
    'activity monitor'
];

const WORK_APPS_TIMER = [
    'adobe', 'premiere', 'afterfx', 'after effects', 'photoshop',
    'cinema 4d', 'cinema4d', 'c4d', 'blender', 'resolve', 'davinci',
    'nuke', 'houdini', 'maya', 'zbrush', 'unreal', 'unity',
    'illustrator', 'indesign', 'lightroom', 'media encoder', 'audition',
    'animate', 'substance', 'bridge', 'fusion', '3ds max', 'cephtmlengine',
    // macOS-only apps
    'final cut', 'motion', 'compressor', 'logic pro'
];

function isAllowedProcess(name) {
    const lower = name.toLowerCase();
    return ALLOWED_APPS.some(a => lower.includes(a));
}

function isWorkApp(name) {
    const lower = name.toLowerCase();
    return WORK_APPS_TIMER.some(a => lower.includes(a));
}

/* ═══════ Window ═══════ */
function createWindow() {
    const opts = {
        width: 1100, height: 800,
        minWidth: 900, minHeight: 600,
        title: 'Lion Workspace',
        autoHideMenuBar: true,
        backgroundColor: '#0a0a0a',
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            backgroundThrottling: false,
            preload: path.join(__dirname, 'preload.js')
        }
    };

    if (isWin) {
        opts.titleBarStyle = 'hidden';
        opts.titleBarOverlay = { color: '#0a0a0a', symbolColor: '#bef264', height: 40 };
    } else if (isMac) {
        opts.titleBarStyle = 'hiddenInset';
        opts.trafficLightPosition = { x: 15, y: 12 };
    }

    mainWin = new BrowserWindow(opts);
    mainWin.loadFile('index.html');
    mainWin.webContents.setWindowOpenHandler(() => ({ action: 'deny' }));
}

/* ═══════ Focus-mode helpers ═══════ */
function checkForegroundWindow() {
    if (!focusActive || !mainWin || mainWin.isDestroyed()) return;

    getFgProc(procName => {
        if (!procName || !focusActive) return;

        if (!isAllowedProcess(procName)) {
            if (focusExternal) {
                // External mode: minimize the non-allowed window
                if (isWin && minimizeScriptPath) {
                    exec(`powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "${minimizeScriptPath}"`, { windowsHide: true, timeout: 5000 });
                } else if (isMac) {
                    exec(`osascript -e 'tell application "System Events" to set visible of first application process whose frontmost is true to false'`, { timeout: 3000 });
                }
            } else {
                if (mainWin && !mainWin.isDestroyed()) {
                    mainWin.moveTop();
                    mainWin.focus();
                }
            }
        }
    });
}

function createOverlays() {
    const displays = screen.getAllDisplays();
    const primary = screen.getPrimaryDisplay();

    displays.forEach(display => {
        if (display.id === primary.id) return;
        const overlay = new BrowserWindow({
            x: display.bounds.x, y: display.bounds.y,
            width: display.bounds.width, height: display.bounds.height,
            frame: false, transparent: true, alwaysOnTop: true,
            skipTaskbar: true, focusable: false, resizable: false,
            webPreferences: { nodeIntegration: false, contextIsolation: true }
        });
        overlay.loadURL(`data:text/html,<html><body style="margin:0;background:rgba(0,0,0,0.92);display:flex;align-items:center;justify-content:center;height:100vh;font-family:system-ui;user-select:none"><div style="text-align:center;color:rgba(190,242,100,.5)"><div style="font-size:4rem;font-weight:800;letter-spacing:4px">FOCO</div><div style="font-size:1.2rem;margin-top:12px;opacity:.5">Pomodoro em andamento</div><div style="margin-top:20px;font-size:.85rem;opacity:.3">Apenas apps de trabalho permitidos</div></div></body></html>`);
        overlay.setAlwaysOnTop(true, 'screen-saver');
        overlays.push(overlay);
    });
}

function destroyOverlays() {
    overlays.forEach(o => { if (!o.isDestroyed()) o.close(); });
    overlays = [];
}

/* ═══════ IPC — Focus mode ═══════ */
ipcMain.handle('focus-start', () => {
    if (focusActive) return true;
    focusActive = true;

    if (mainWin && !mainWin.isDestroyed()) {
        originalBounds = mainWin.getBounds();
        mainWin.setAlwaysOnTop(true, 'screen-saver');
        mainWin.setFullScreen(true);
    }
    createOverlays();
    focusInterval = setInterval(checkForegroundWindow, 1200);
    return true;
});

ipcMain.handle('focus-stop', () => {
    focusActive = false;
    const wasExternal = focusExternal;
    focusExternal = false;

    if (focusInterval) { clearInterval(focusInterval); focusInterval = null; }
    destroyOverlays();

    if (mainWin && !mainWin.isDestroyed()) {
        if (!wasExternal) mainWin.setFullScreen(false);
        mainWin.setAlwaysOnTop(false);
        if (originalBounds) { mainWin.setBounds(originalBounds); originalBounds = null; }
    }
    return true;
});

ipcMain.handle('focus-is-active', () => focusActive);

ipcMain.handle('focus-start-external', () => {
    if (focusActive) return true;
    focusActive = true;
    focusExternal = true;

    if (mainWin && !mainWin.isDestroyed()) {
        originalBounds = mainWin.getBounds();
    }
    createOverlays();
    focusInterval = setInterval(checkForegroundWindow, 1200);
    return true;
});

/* ═══════ IPC — Timer foreground auto-pause ═══════ */
let timerFgInterval = null;
let timerFgPaused = false;

function checkTimerForeground() {
    if (!mainWin || mainWin.isDestroyed()) return;

    getFgProc(procName => {
        if (!procName || !mainWin || mainWin.isDestroyed()) return;

        if (!isWorkApp(procName) && !timerFgPaused) {
            timerFgPaused = true;
            mainWin.webContents.send('timer-fg-pause');
        } else if (isWorkApp(procName) && timerFgPaused) {
            timerFgPaused = false;
            mainWin.webContents.send('timer-fg-resume');
        }
    });
}

ipcMain.handle('check-foreground', () => {
    return new Promise(resolve => {
        getFgProc(proc => resolve(proc ? isWorkApp(proc) : false));
    });
});

ipcMain.handle('timer-fg-start', () => {
    timerFgPaused = false;
    if (!timerFgInterval) {
        timerFgInterval = setInterval(checkTimerForeground, 2500);
    }
    return true;
});

ipcMain.handle('timer-fg-stop', () => {
    if (timerFgInterval) { clearInterval(timerFgInterval); timerFgInterval = null; }
    timerFgPaused = false;
    return true;
});

/* ═══════ IPC — Launch app ═══════ */
ipcMain.handle('launch-app', (event, appName) => {
    if (isWin) {
        const shellCmds = {
            'premiere': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Premiere*' -Recurse -Filter '*.exe' -ErrorAction SilentlyContinue | Where-Object {$_.Name -like '*Premiere Pro*'} | Select-Object -First 1; if($p){Start-Process $p.FullName}else{Start-Process 'shell:AppsFolder' -ArgumentList 'Adobe Premiere Pro'}"`,
            'aftereffects': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*After*' -Recurse -Filter 'AfterFX.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'photoshop': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Photoshop*' -Recurse -Filter 'Photoshop.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'cinema4d': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Maxon*' -Recurse -Filter 'Cinema 4D.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'explorer': 'explorer.exe'
        };
        const cmd = shellCmds[appName];
        if (cmd) exec(cmd, { windowsHide: true });
    } else if (isMac) {
        const macCmds = {
            'premiere': 'open -a "Adobe Premiere Pro" 2>/dev/null',
            'aftereffects': 'open -a "Adobe After Effects" 2>/dev/null',
            'photoshop': 'open -a "Adobe Photoshop" 2>/dev/null',
            'cinema4d': 'open -a "Cinema 4D" 2>/dev/null',
            'explorer': 'open ~'
        };
        const cmd = macCmds[appName];
        if (cmd) exec(cmd);
    }

    // Temporarily allow focus to leave for 5 seconds
    if (focusActive) {
        const wasActive = focusActive;
        focusActive = false;
        setTimeout(() => { focusActive = wasActive; }, 5000);
    }
    return true;
});

ipcMain.handle('open-folder', () => {
    if (focusActive) {
        const wasActive = focusActive;
        focusActive = false;
        setTimeout(() => { focusActive = wasActive; }, 8000);
    }
    exec(isWin ? 'explorer.exe' : 'open ~', { windowsHide: true });
    return true;
});

/* ═══════ HTTP Sync Server (Premiere Pro plugin) ═══════ */
const SYNC_PORT = 9847;
let cachedState = { timer: { running: false }, pomodoro: { running: false }, projects: [] };

ipcMain.handle('sync-push-state', (event, state) => {
    cachedState = state;
    return true;
});

function startSyncServer() {
    const server = http.createServer((req, res) => {
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

        if (req.method === 'OPTIONS') {
            res.writeHead(200);
            res.end();
            return;
        }

        if (req.method === 'GET' && req.url === '/state') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(cachedState));
            return;
        }

        if (req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk);
            req.on('end', () => {
                let data = {};
                try { data = JSON.parse(body || '{}'); } catch (e) {}
                const command = req.url.replace(/^\//, '');

                // Auto-manage timer foreground check from HTTP commands
                if (command === 'timer/start') {
                    timerFgPaused = false;
                    if (!timerFgInterval) {
                        timerFgInterval = setInterval(checkTimerForeground, 2500);
                    }
                } else if (command === 'timer/stop') {
                    if (timerFgInterval) { clearInterval(timerFgInterval); timerFgInterval = null; }
                    timerFgPaused = false;
                }

                if (mainWin && !mainWin.isDestroyed()) {
                    mainWin.webContents.send('sync-command', { command, data });
                }
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ ok: true }));
            });
            return;
        }

        res.writeHead(404);
        res.end('Not found');
    });

    server.listen(SYNC_PORT, '127.0.0.1', () => {
        console.log('Lion Workspace sync: http://127.0.0.1:' + SYNC_PORT);
    });
    server.on('error', (err) => {
        console.error('Sync server error:', err.message);
    });
}

/* ═══════ App lifecycle ═══════ */
app.on('before-quit', () => {
    app.isQuitting = true;
    stopFgDaemon();
});

app.whenReady().then(() => {
    createWindow();
    startSyncServer();
});

app.on('window-all-closed', () => {
    if (!isMac) app.quit();
});

// macOS: re-create window when clicking dock icon
if (isMac) {
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
}
