const { app, BrowserWindow, ipcMain, screen, powerMonitor, Menu, globalShortcut } = require('electron');
const path = require('path');
const http = require('http');
const fs = require('fs');
const os = require('os');
const crypto = require('crypto');
const { exec, spawn } = require('child_process');
const { autoUpdater } = require('electron-updater');
const { createClient } = require('@supabase/supabase-js');

const isWin = process.platform === 'win32';
const isMac = process.platform === 'darwin';

// Global crash guards — prevent uncaught errors (antivirus locks, network, etc.)
// from showing the ugly "A JavaScript error occurred" dialog and killing the app.
process.on('uncaughtException', (err) => {
    console.error('Uncaught exception:', err && (err.stack || err.message || err));
});
process.on('unhandledRejection', (reason) => {
    console.error('Unhandled rejection:', reason);
});

/* ═══════════════════════════════════════════════════════════════════
   DATA MIGRATION — merge data from old "controle-videos" folder
   into current "lion-workspace" if it has no client data yet.
   This must run BEFORE app is ready.
   ═══════════════════════════════════════════════════════════════════ */
const oldName = 'controle-videos';
const currentUD = app.getPath('userData');
const oldUD = path.join(path.dirname(currentUD), oldName);
// Only migrate once: copy old leveldb files if current folder is empty/new
if (fs.existsSync(oldUD) && currentUD !== oldUD) {
    const curLS = path.join(currentUD, 'Local Storage', 'leveldb');
    const oldLS = path.join(oldUD, 'Local Storage', 'leveldb');
    if (fs.existsSync(oldLS) && !fs.existsSync(curLS)) {
        try {
            fs.mkdirSync(path.join(currentUD, 'Local Storage'), { recursive: true });
            fs.cpSync(oldLS, curLS, { recursive: true });
            console.log('Migrated localStorage from controle-videos to lion-workspace');
        } catch (e) { console.warn('Migration copy failed:', e); }
    }
}

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
        "  Start-Sleep -Milliseconds 600",
        "}"
    ].join("\n"));

    // Minimize script (used by focus mode to hide non-allowed windows)
    minimizeScriptPath = path.join(os.tmpdir(), `lion-minimize-${process.pid}.ps1`);
    fs.writeFileSync(minimizeScriptPath, [
        "try{Add-Type -Name WM -Namespace NM -MemberDefinition '[DllImport(\"user32.dll\")]public static extern bool ShowWindow(IntPtr h,int c);[DllImport(\"user32.dll\")]public static extern IntPtr GetForegroundWindow();'}catch{}",
        "[NM.WM]::ShowWindow([NM.WM]::GetForegroundWindow(), 6)"
    ].join("\n"));
}

let _fgRestartDelay = 3000; // exponential backoff on repeated crashes
let _fgLastStart = 0;
function startFgDaemon() {
    if (!isWin || fgDaemon) return;
    const now = Date.now();
    // If daemon lasted >60s, consider it stable and reset delay
    if (now - _fgLastStart > 60000) _fgRestartDelay = 3000;
    _fgLastStart = now;

    fgDaemon = spawn('powershell', [
        '-NoProfile', '-WindowStyle', 'Hidden',
        '-ExecutionPolicy', 'Bypass', '-File', daemonScriptPath
    ], { windowsHide: true, stdio: 'ignore' });

    fgDaemon.on('exit', () => {
        fgDaemon = null;
        if (!app.isQuitting) {
            setTimeout(startFgDaemon, _fgRestartDelay);
            // Exponential backoff: 3s → 6s → 12s → 24s → 48s → cap 60s
            _fgRestartDelay = Math.min(_fgRestartDelay * 2, 60000);
        }
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
const BASE_ALLOWED = [
    'lion workspace', 'electron',
    // Adobe Suite
    'adobe premiere pro', 'premiere pro', 'premiere', 'afterfx', 'after effects',
    'photoshop', 'illustrator', 'indesign', 'lightroom', 'media encoder',
    'adobe audition', 'audition', 'animate', 'adobe animate', 'character animator',
    'adobe xd', 'substance', 'adobe bridge', 'bridge', 'acrobat',
    // 3D / VFX / Motion
    'cinema 4d', 'cinema4d', 'c4d', 'blender', 'davinci resolve', 'resolve',
    'nuke', 'fusion', 'houdini', 'maya', '3ds max', 'zbrush', 'unreal', 'unity',
    'modo', 'katana', 'mari', 'natron', 'hitfilm', 'vegas pro', 'vegas',
    // Audio
    'fl studio', 'ableton', 'reaper', 'audacity', 'logic pro', 'pro tools', 'cubase',
    // Design / UI
    'figma', 'sketch', 'canva', 'coreldraw', 'affinity', 'gimp', 'inkscape', 'krita',
    // Code / Dev (optional for some editors)
    'visual studio code', 'code', 'cursor', 'sublime', 'notepad++', 'terminal', 'powershell', 'cmd',
    // File / System (always allowed)
    'explorer', 'files', 'searchui', 'shellexperiencehost',
    'startmenuexperiencehost', 'applicationframehost',
    'textinputhost', 'runtimebroker', 'taskmgr', 'settings',
    // macOS system
    'finder', 'spotlight', 'system preferences', 'system settings',
    'activity monitor'
];

const WORK_APPS_TIMER = [
    // Adobe Suite
    'adobe', 'premiere', 'afterfx', 'after effects', 'photoshop',
    'illustrator', 'indesign', 'lightroom', 'media encoder', 'audition',
    'animate', 'substance', 'bridge', 'cephtmlengine',
    // 3D / VFX
    'cinema 4d', 'cinema4d', 'c4d', 'blender', 'resolve', 'davinci',
    'nuke', 'houdini', 'maya', 'zbrush', 'unreal', 'unity',
    'fusion', '3ds max', 'modo', 'katana', 'mari', 'natron',
    // Video
    'vegas', 'hitfilm', 'capcut', 'final cut', 'motion', 'compressor',
    // Design
    'figma', 'sketch', 'canva', 'affinity', 'gimp', 'inkscape', 'krita',
    // Audio
    'fl studio', 'ableton', 'reaper', 'pro tools', 'cubase', 'logic pro', 'audacity'
];

// Dynamic whitelist from renderer (user-added apps)
let userWhitelist = [];

function isAllowedProcess(name) {
    const lower = name.toLowerCase();
    if (BASE_ALLOWED.some(a => lower.includes(a))) return true;
    if (userWhitelist.some(a => lower.includes(a.toLowerCase()))) return true;
    return false;
}

function isWorkApp(name) {
    const lower = name.toLowerCase();
    return WORK_APPS_TIMER.some(a => lower.includes(a));
}

// Detecta apenas Premiere Pro (não pega outros apps Adobe)
function isPremiereApp(name) {
    const lower = (name || '').toLowerCase();
    return /(adobe\s+)?premiere(\s+pro)?/i.test(lower);
}

// Detecta apenas After Effects
function isAfterApp(name) {
    const lower = (name || '').toLowerCase();
    return /afterfx|after\s+effects/i.test(lower);
}

// Modo de tracking ativo: 'work' (qualquer work app), 'premiere', 'after'
let timerFgMode = 'work';
function appMatchesMode(name, mode) {
    if (mode === 'premiere') return isPremiereApp(name);
    if (mode === 'after')    return isAfterApp(name);
    return isWorkApp(name);
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

    // When main window closes, kill focus mode só overlays don't keep the app alive
    mainWin.on('close', () => {
        if (focusActive) {
            focusActive = false;
            focusExternal = false;
            if (focusInterval) { clearInterval(focusInterval); focusInterval = null; }
            destroyOverlays();
        }
    });
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
        if (isMac) {
            // Mac: use simpleFullScreen to avoid Spaces animation conflict
            mainWin.setSimpleFullScreen(true);
            mainWin.setAlwaysOnTop(true, 'floating');
        } else {
            mainWin.setAlwaysOnTop(true, 'screen-saver');
            mainWin.setFullScreen(true);
        }
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
        if (!wasExternal) {
            if (isMac) {
                mainWin.setSimpleFullScreen(false);
            } else {
                mainWin.setFullScreen(false);
            }
        }
        mainWin.setAlwaysOnTop(false);
        // Wait for transition before restoring bounds (Mac animation)
        if (originalBounds) {
            setTimeout(() => {
                if (mainWin && !mainWin.isDestroyed() && originalBounds) {
                    mainWin.setBounds(originalBounds);
                    originalBounds = null;
                }
            }, isMac ? 400 : 100);
        }
    }
    return true;
});

ipcMain.handle('focus-is-active', () => focusActive);

ipcMain.handle('focus-set-whitelist', (event, list) => {
    userWhitelist = Array.isArray(list) ? list : [];
    return true;
});

// List running processes for app linking
ipcMain.handle('list-running-processes', () => {
    return new Promise(resolve => {
        if (isWin) {
            exec('powershell -NoProfile -WindowStyle Hidden -Command "Get-Process | Where-Object {$_.MainWindowTitle -ne \'\'} | Select-Object ProcessName, MainWindowTitle, Path | ConvertTo-Json"',
                { windowsHide: true, timeout: 8000 },
                (err, stdout) => {
                    if (err) { resolve([]); return; }
                    try {
                        let procs = JSON.parse(stdout || '[]');
                        if (!Array.isArray(procs)) procs = [procs];
                        resolve(procs.map(p => ({
                            name: p.ProcessName || '',
                            title: p.MainWindowTitle || '',
                            path: p.Path || ''
                        })).filter(p => p.name && p.name !== 'Lion Workspace'));
                    } catch { resolve([]); }
                });
        } else if (isMac) {
            exec("osascript -e 'tell application \"System Events\" to get name of every application process whose background only is false'",
                { timeout: 5000 },
                (err, stdout) => {
                    if (err) { resolve([]); return; }
                    try {
                        const raw = (stdout || '').trim();
                        // Output: "Finder, Safari, Chrome, ..." or "{Finder, Safari, ...}"
                        const cleaned = raw.replace(/^\{|\}$/g, '');
                        const names = cleaned.split(', ')
                            .map(n => n.trim())
                            .filter(n => n && n !== 'Lion Workspace' && n !== 'Electron');
                        resolve([...new Set(names)].map(n => ({ name: n, title: n, path: `/Applications/${n}.app` })));
                    } catch { resolve([]); }
                });
        } else { resolve([]); }
    });
});

// Launch app by executable path
ipcMain.handle('launch-by-path', (event, exePath) => {
    if (!exePath) return false;
    // Sanitize path
    const safePath = exePath.replace(/[;|`$()&]/g, '');
    if (isWin) {
        exec(`start "" "${safePath}"`, { windowsHide: true });
    } else if (isMac) {
        // On Mac, detect .app bundles vs executables
        if (safePath.endsWith('.app')) {
            exec(`open -a "${safePath}"`);
        } else {
            exec(`open "${safePath}"`);
        }
    }
    if (focusActive) {
        const wasActive = focusActive;
        focusActive = false;
        setTimeout(() => { focusActive = wasActive; }, 5000);
    }
    return true;
});

// Find installed path for an app (Windows: searches Program Files; macOS: /Applications)
ipcMain.handle('find-app-path', (event, appName) => {
    // Sanitize input
    const safe = (appName || '').replace(/[;|`$()&'"\\]/g, '');
    if (!safe) return Promise.resolve('');
    return new Promise(resolve => {
        if (isWin) {
            exec(`powershell -NoProfile -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files','C:\\Program Files (x86)' -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -like '*${safe}*'} | Select-Object -First 1; if($p){$e=Get-ChildItem $p.FullName -Recurse -Filter '*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($e){$e.FullName}}"`,
                { windowsHide: true, timeout: 10000 },
                (err, stdout) => resolve((stdout || '').trim() || ''));
        } else if (isMac) {
            exec(`find /Applications -maxdepth 2 -name "*${safe}*.app" -print -quit 2>/dev/null`,
                { timeout: 3000 },
                (err, stdout) => resolve((stdout || '').trim() || ''));
        } else { resolve(''); }
    });
});

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
let fgMissCount = 0;          // Grace period: must miss consecutive checks before pausing
const FG_MISS_THRESHOLD = 2;  // Need 2 consecutive misses (~2.4s) before pausing

const IDLE_THRESHOLD_SECS = 45; // Pause timer if no mouse/keyboard for 45 seconds (system-wide)

function checkTimerForeground() {
    if (!mainWin || mainWin.isDestroyed()) return;

    // Check system-wide idle time (works even when user is in Premiere)
    let systemIdle = 0;
    try { systemIdle = powerMonitor.getSystemIdleTime(); } catch(e) {}

    if (systemIdle >= IDLE_THRESHOLD_SECS) {
        // User hasn't touched mouse/keyboard for 45s — pause even if in work app
        if (!timerFgPaused) {
            timerFgPaused = true;
            mainWin.webContents.send('timer-fg-pause');
        }
        return;
    }

    getFgProc(procName => {
        if (!procName || !mainWin || mainWin.isDestroyed()) return;

        if (!appMatchesMode(procName, timerFgMode)) {
            fgMissCount++;
            if (fgMissCount >= FG_MISS_THRESHOLD && !timerFgPaused) {
                timerFgPaused = true;
                mainWin.webContents.send('timer-fg-pause');
            }
        } else {
            fgMissCount = 0;
            if (timerFgPaused) {
                timerFgPaused = false;
                mainWin.webContents.send('timer-fg-resume');
            }
        }
    });
}

ipcMain.handle('check-foreground', (_, mode) => {
    return new Promise(resolve => {
        const m = mode || 'work';
        getFgProc(proc => resolve(proc ? appMatchesMode(proc, m) : false));
    });
});

ipcMain.handle('timer-fg-start', (_, mode) => {
    timerFgPaused = false;
    fgMissCount = 0;
    timerFgMode = mode || 'work';
    if (!timerFgInterval) {
        timerFgInterval = setInterval(checkTimerForeground, 1200);
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
            'illustrator': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Illustrator*' -Recurse -Filter 'Illustrator.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'lightroom': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Lightroom*' -Recurse -Filter '*.exe' -ErrorAction SilentlyContinue | Where-Object {$_.Name -like '*Lightroom*'} | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'mediaencoder': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Media Encoder*' -Recurse -Filter '*.exe' -ErrorAction SilentlyContinue | Where-Object {$_.Name -like '*Media Encoder*'} | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'audition': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Adobe\\*Audition*' -Recurse -Filter 'Audition.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'blender': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Blender*' -Recurse -Filter 'blender.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}else{Start-Process 'shell:AppsFolder' -ArgumentList 'Blender'}"`,
            'resolve': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\Blackmagic*' -Recurse -Filter 'Resolve.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'figma': `powershell -WindowStyle Hidden -Command "Start-Process 'shell:AppsFolder' -ArgumentList 'Figma'"`,
            'obs': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files\\obs-studio*' -Recurse -Filter 'obs64.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'audacity': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files*\\Audacity*' -Recurse -Filter 'Audacity.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'vegas': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files*\\VEGAS*' -Recurse -Filter 'vegas*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'gimp': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files*\\GIMP*' -Recurse -Filter 'gimp*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'inkscape': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files*\\Inkscape*' -Recurse -Filter 'inkscape.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'krita': `powershell -WindowStyle Hidden -Command "$p=Get-ChildItem 'C:\\Program Files*\\Krita*' -Recurse -Filter 'krita.exe' -ErrorAction SilentlyContinue | Select-Object -First 1; if($p){Start-Process $p.FullName}"`,
            'capcut': `powershell -WindowStyle Hidden -Command "Start-Process 'shell:AppsFolder' -ArgumentList 'CapCut'"`,
            'canva': `powershell -WindowStyle Hidden -Command "Start-Process 'https://www.canva.com'"`,
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
            'illustrator': 'open -a "Adobe Illustrator" 2>/dev/null',
            'lightroom': 'open -a "Adobe Lightroom Classic" 2>/dev/null || open -a "Adobe Lightroom" 2>/dev/null',
            'mediaencoder': 'open -a "Adobe Media Encoder" 2>/dev/null',
            'audition': 'open -a "Adobe Audition" 2>/dev/null',
            'blender': 'open -a "Blender" 2>/dev/null',
            'resolve': 'open -a "DaVinci Resolve" 2>/dev/null',
            'figma': 'open -a "Figma" 2>/dev/null',
            'obs': 'open -a "OBS" 2>/dev/null || open -a "OBS Studio" 2>/dev/null',
            'audacity': 'open -a "Audacity" 2>/dev/null',
            'vegas': 'echo "Vegas Pro não disponível no macOS"',
            'gimp': 'open -a "GIMP-2.10" 2>/dev/null || open -a "GIMP" 2>/dev/null',
            'inkscape': 'open -a "Inkscape" 2>/dev/null',
            'krita': 'open -a "Krita" 2>/dev/null',
            'capcut': 'open -a "CapCut" 2>/dev/null',
            'canva': 'open "https://www.canva.com"',
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

ipcMain.handle('open-external', (event, url) => {
    if (url && (url.startsWith('https://') || url.startsWith('http://'))) {
        require('electron').shell.openExternal(url);
    }
    return true;
});

// Copy image to OS clipboard (works on Win/Mac/Linux uniformly).
// Used as fallback when navigator.clipboard.write is blocked (sandbox no Mac).
ipcMain.handle('copy-image-to-clipboard', (event, arrayBuf) => {
    try {
        const { clipboard, nativeImage } = require('electron');
        const buf = Buffer.from(arrayBuf);
        const img = nativeImage.createFromBuffer(buf);
        if (img.isEmpty()) return false;
        clipboard.writeImage(img);
        return true;
    } catch (e) {
        console.error('copy-image-to-clipboard:', e.message);
        return false;
    }
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

/* ═══════ YouTube Downloader (yt-dlp) ═══════ */
const ytDlpDir = path.join(app.getPath('userData'), 'yt-dlp');
const ytDlpBin = isWin ? path.join(ytDlpDir, 'yt-dlp.exe') : path.join(ytDlpDir, 'yt-dlp');
const ffmpegBin = isWin ? path.join(ytDlpDir, 'ffmpeg.exe') : path.join(ytDlpDir, 'ffmpeg');
const ffprobeBin = isWin ? path.join(ytDlpDir, 'ffprobe.exe') : path.join(ytDlpDir, 'ffprobe');
let ytDownloadProc = null;
let ytProgress = { active: false, percent: 0, speed: '', eta: '', title: '', status: 'idle', error: '' };

// ═══════ Mask Editor — Pincel mágico pra ajustar bg-remove ═══════
// Sessões ativas de edição. Cada bg-remove pode abrir um editor; plugin
// polla /bg/edit-status?token=X pra saber quando o user salvou/cancelou.
const maskEditorSessions = new Map();
// token -> { status: 'editing'|'saved'|'cancelled', processedPath, win }

// Singleton: só uma janela de mask editor aberta por vez.
// Se ja tiver uma, foca nela e retorna o token existente.
let _activeMaskEditorWin = null;
let _activeMaskEditorToken = null;

function openMaskEditor(origPath, processedPath) {
    // Se ja tem editor aberto, foca nele e retorna o token existente
    if (_activeMaskEditorWin && !_activeMaskEditorWin.isDestroyed()) {
        try {
            if (_activeMaskEditorWin.isMinimized()) _activeMaskEditorWin.restore();
            _activeMaskEditorWin.show();
            _activeMaskEditorWin.focus();
            _activeMaskEditorWin.moveTop();
        } catch(e) {}
        return _activeMaskEditorToken || '';
    }

    const token = crypto.randomBytes(8).toString('hex');
    const sess = { status: 'editing', processedPath, win: null };
    maskEditorSessions.set(token, sess);
    _activeMaskEditorToken = token;

    // Limpa sessões antigas (>1h) pra não vazar
    const HOUR = 60 * 60 * 1000;
    const now = Date.now();
    for (const [k, v] of maskEditorSessions) {
        if (v._createdAt && (now - v._createdAt) > HOUR) maskEditorSessions.delete(k);
    }
    sess._createdAt = now;

    try {
        const win = new BrowserWindow({
            width: 1280,
            height: 820,
            minWidth: 920,
            minHeight: 600,
            title: 'Editor de Máscara — Lion Workspace',
            backgroundColor: '#050505',
            show: true,
            // Frameless: o titlebar nativo do Windows era redundante com a
            // topbar custom do editor (.topbar). Agora a topbar é drag region.
            frame: false,
            titleBarStyle: 'hidden',
            webPreferences: {
                nodeIntegration: true,
                contextIsolation: false,
            },
        });
        sess.win = win;
        _activeMaskEditorWin = win;

        const editorHtml = path.join(__dirname, 'mask-editor.html');
        const qs = '?orig=' + encodeURIComponent(origPath)
                 + '&processed=' + encodeURIComponent(processedPath)
                 + '&token=' + token;
        win.loadFile(editorHtml, { search: qs.replace(/^\?/, '') });
        // Remove menu nativo (pra ficar mais clean)
        try { win.setMenuBarVisibility(false); } catch(e) {}

        win.on('closed', () => {
            const s = maskEditorSessions.get(token);
            if (s && s.status === 'editing') {
                s.status = 'cancelled';
            }
            // Libera o singleton
            if (_activeMaskEditorToken === token) {
                _activeMaskEditorWin = null;
                _activeMaskEditorToken = null;
            }
            // Mantém a entry por 5min pra plugin poder pollar status final
            setTimeout(() => maskEditorSessions.delete(token), 5 * 60 * 1000);
        });

        return token;
    } catch (e) {
        sess.status = 'error';
        sess.error = e.message;
        console.error('[mask-editor] failed to open:', e);
        return token;
    }
}

// IPC: editor envia o PNG final pra salvar
ipcMain.on('mask-editor:save', (event, payload) => {
    try {
        const { token, dataUrl, processedPath } = payload || {};
        const sess = maskEditorSessions.get(token);
        if (!sess) {
            event.sender.send('mask-editor:save-error', 'Sessão expirada');
            return;
        }
        const targetPath = processedPath || sess.processedPath;
        const base64 = String(dataUrl || '').replace(/^data:image\/png;base64,/, '');
        if (!base64) throw new Error('dataUrl vazio');
        fs.writeFileSync(targetPath, Buffer.from(base64, 'base64'));
        sess.status = 'saved';
        sess.processedPath = targetPath;
        event.sender.send('mask-editor:saved');
        // Fecha janela após pequeno delay (pro user ver "✓ Salvo")
        setTimeout(() => {
            try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
        }, 500);
    } catch (e) {
        event.sender.send('mask-editor:save-error', e.message || String(e));
    }
});

ipcMain.on('mask-editor:cancel', (event, payload) => {
    try {
        const { token } = payload || {};
        const sess = maskEditorSessions.get(token);
        if (sess) {
            sess.status = 'cancelled';
            try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
        }
    } catch(e) { console.error('[mask-editor:cancel]', e); }
});

// ════════════════════════════════════════════════════════════════════
// ROTOSCOPE EDITOR — SAM 2 video segmentation
// Fase 1 (atual): setup, UI base, abre janela, IPC infra
// Fase 2+: integração do modelo SAM 2 ONNX no worker
// ════════════════════════════════════════════════════════════════════
const rotoscopeSessions = new Map();
// token -> { status, win, srcPath, outPath, masks }

async function openRotoscopeEditor(origSrcPath, outPath) {
    const token = crypto.randomBytes(8).toString('hex');
    const sess = { status: 'preparing', srcPath: origSrcPath, outPath, masks: null, _createdAt: Date.now() };
    rotoscopeSessions.set(token, sess);

    const now = Date.now();
    for (const [k, v] of rotoscopeSessions) {
        if (v._createdAt && (now - v._createdAt) > 2 * 60 * 60 * 1000) rotoscopeSessions.delete(k);
    }

    // PRÉ-TRANSCODE: garante H.264 pra HTML5 <video> decodar.
    // Sem isso, ProRes/HEVC/DNxHD davam canvas vazio — vid.onerror silencioso.
    let workingSrc = origSrcPath;
    if (ffmpegReady() || (await ensureFfmpeg())) {
        try {
            const tmp = path.join(currentUD, 'roto-trim-' + Date.now() + '.mp4');
            console.log('[rotoscope] pré-transcode H.264:', origSrcPath, '→', tmp);
            await new Promise((resolve, reject) => {
                const proc = bgvSpawn(ffmpegBin, [
                    '-y', '-i', origSrcPath,
                    '-c:v', 'libx264', '-preset', 'ultrafast', '-crf', '20',
                    '-pix_fmt', 'yuv420p',
                    '-g', '1', '-keyint_min', '1', '-sc_threshold', '0',
                    '-c:a', 'aac', '-b:a', '128k',
                    '-movflags', '+faststart',
                    tmp,
                ], { windowsHide: true });
                let err = '';
                proc.stderr.on('data', d => { err += d.toString(); if (err.length > 4096) err = err.slice(-4096); });
                proc.on('close', code => {
                    if (code === 0 && fs.existsSync(tmp)) resolve();
                    else reject(new Error('ffmpeg trim falhou: ' + err.slice(-300)));
                });
                proc.on('error', e => reject(e));
            });
            workingSrc = tmp;
            sess._tempTranscoded = tmp;
        } catch (e) {
            console.error('[rotoscope] pre-transcode failed, usando source original:', e.message);
        }
    }

    sess.workingSrc = workingSrc;
    sess.status = 'editing';

    try {
        const win = new BrowserWindow({
            width: 1400, height: 900,
            minWidth: 1000, minHeight: 640,
            title: 'Rotoscope — Lion Workspace',
            backgroundColor: '#050505',
            show: true,
            frame: false,
            titleBarStyle: 'hidden',
            webPreferences: {
                nodeIntegration: true,
                contextIsolation: false,
                webgl: true,
                webSecurity: false,    // permite file:// pro <video> e fetch local
            },
        });
        sess.win = win;

        const editorHtml = path.join(__dirname, 'rotoscope-editor.html');
        const qs = '?token=' + encodeURIComponent(token)
                 + '&src=' + encodeURIComponent(workingSrc)
                 + '&origSrc=' + encodeURIComponent(origSrcPath)
                 + '&out=' + encodeURIComponent(outPath || '');
        win.loadFile(editorHtml, { search: qs.replace(/^\?/, '') });
        try { win.setMenuBarVisibility(false); } catch(e) {}

        win.on('closed', () => {
            const s = rotoscopeSessions.get(token);
            if (s && s.status === 'editing') s.status = 'cancelled';
            // Cleanup do temp transcoded
            if (s && s._tempTranscoded) {
                try { if (fs.existsSync(s._tempTranscoded)) fs.unlinkSync(s._tempTranscoded); } catch(e) {}
            }
            setTimeout(() => rotoscopeSessions.delete(token), 5 * 60 * 1000);
        });

        return token;
    } catch (e) {
        sess.status = 'error';
        sess.error = e.message;
        return token;
    }
}

// Probe via ffprobe — reusa lógica do bg-video
ipcMain.handle('rotoscope:probe', async (event, payload) => {
    try {
        const src = payload?.src;
        if (!src || !fs.existsSync(src)) return { error: 'arquivo não encontrado' };
        if (!ffprobeReady() || !ffmpegReady()) await ensureFfmpeg();
        if (!ffprobeReady()) return { error: 'ffprobe não disponível' };

        return await new Promise((resolve) => {
            const proc = bgvSpawn(ffprobeBin, [
                '-v', 'error',
                '-select_streams', 'v:0',
                '-show_entries', 'stream=width,height,r_frame_rate,duration,nb_frames',
                '-show_entries', 'format=duration',
                '-of', 'json',
                src
            ], { windowsHide: true });
            let out = '';
            proc.stdout.on('data', d => out += d.toString());
            proc.on('close', () => {
                try {
                    const j = JSON.parse(out);
                    const stream = j.streams?.[0] || {};
                    const fmt = j.format || {};
                    const fpsStr = stream.r_frame_rate || '30/1';
                    const [num, den] = fpsStr.split('/').map(Number);
                    const fps = (num && den) ? num / den : 30;
                    resolve({
                        width: parseInt(stream.width) || 0,
                        height: parseInt(stream.height) || 0,
                        fps: Math.round(fps * 1000) / 1000,
                        duration: parseFloat(stream.duration || fmt.duration || '0') || 0,
                    });
                } catch(e) { resolve({ error: 'parse falhou: ' + e.message }); }
            });
            proc.on('error', e => resolve({ error: e.message }));
        });
    } catch (e) { return { error: e.message }; }
});

// Window controls da janela frameless do rotoscope
ipcMain.on('rotoscope:win-action', (event, payload) => {
    try {
        const { token, action } = payload || {};
        const sess = rotoscopeSessions.get(token);
        if (!sess || !sess.win || sess.win.isDestroyed()) return;
        const w = sess.win;
        if (action === 'minimize') w.minimize();
        else if (action === 'maximize') {
            if (w.isMaximized()) w.unmaximize();
            else w.maximize();
        } else if (action === 'close') {
            sess.status = sess.status === 'editing' ? 'cancelled' : sess.status;
            w.close();
        }
    } catch(e) {}
});

ipcMain.on('rotoscope:ready', (event, payload) => {
    console.log('[rotoscope] worker ready:', payload?.token);
});

ipcMain.on('rotoscope:cancel', (event, payload) => {
    const sess = rotoscopeSessions.get(payload?.token);
    if (sess) {
        sess.status = 'cancelled';
        try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
    }
});

ipcMain.on('rotoscope:save', (event, payload) => {
    const sess = rotoscopeSessions.get(payload?.token);
    if (sess) {
        sess.masks = payload?.masks || {};
        sess.status = 'saved';
    }
});

ipcMain.on('rotoscope:save-done', (event, payload) => {
    const sess = rotoscopeSessions.get(payload?.token);
    if (sess) {
        sess.status = 'saved';
        sess.outPath = payload?.outPath || sess.outPath;
        // Fecha janela após save
        setTimeout(() => {
            try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
        }, 800);
    }
});

// ════════════════════════════════════════════════════════════════════
// SAM (Segment Anything Model) — RODA NO MAIN PROCESS via Node.js
// ────────────────────────────────────────────────────────────────────
// Movido pra cá pq Electron renderer com transformers.js tinha problemas
// recorrentes de loading (ESM/asar/onnxruntime-web). Node.js Main rodando
// onnxruntime-node binding nativo é MUITO mais robusto.
// ────────────────────────────────────────────────────────────────────
// Estado global (1 modelo carregado, compartilhado entre sessões):
//   samState.tx = referência pro módulo @huggingface/transformers
//   samState.model = SamModel carregado (SlimSAM-77 quantized)
//   samState.processor = AutoProcessor casado
//   samState.loading = Promise (se em curso)
//   samState.loaded = boolean
//   samState.embeddings = Map<token, { embeddings, originalSizes, reshapedSizes }>
// ════════════════════════════════════════════════════════════════════
const samState = {
    tx: null,
    model: null,
    processor: null,
    loading: null,
    loaded: false,
    embeddings: new Map(), // token -> embedding cache
};

function broadcastSamProgress(eventSender, payload) {
    try { if (eventSender && !eventSender.isDestroyed()) eventSender.send('roto:sam-progress', payload); } catch(e) {}
}

async function loadSamModel(eventSender) {
    if (samState.loaded) return { ok: true };
    if (samState.loading) return await samState.loading;

    samState.loading = (async () => {
        try {
            broadcastSamProgress(eventSender, { phase: 'init', msg: 'Inicializando transformers.js...', pct: 5 });

            // Lazy require — só carrega quando o user pede SAM
            const tx = require('@huggingface/transformers');
            samState.tx = tx;

            // Configura cacheDir pra um lugar gravável (não dentro do asar!)
            const cacheRoot = path.join(currentUD, 'sam-cache');
            try { fs.mkdirSync(cacheRoot, { recursive: true }); } catch(e) {}
            tx.env.cacheDir = cacheRoot;
            tx.env.useFSCache = true;
            tx.env.allowRemoteModels = true;
            tx.env.allowLocalModels = true;
            console.log('[sam] cacheDir =', cacheRoot);

            // SAM-Base (Meta original) — qualidade muito superior à SlimSAM.
            // Quantizado q8 fica ~24MB. Com thread-limit + encode lazy não trava o PC.
            const modelId = 'Xenova/sam-vit-base';
            broadcastSamProgress(eventSender, { phase: 'download', msg: 'Carregando SAM-Base q8 (~24MB primeira vez)...', pct: 10 });

            // Limita threads pra metade dos cores (deixa OS respirando, não trava PC).
            // executionMode 'sequential' (default) — 'parallel' usa todos cores em ops paralelas
            // e foi o que travou tudo. Sequential com poucas threads = bem mais responsivo.
            const cpuCount = os.cpus().length;
            const numThreads = Math.max(1, Math.min(4, Math.floor(cpuCount / 2)));
            console.log('[sam] CPU threads:', numThreads, 'de', cpuCount);
            const sessionOpts = {
                graphOptimizationLevel: 'all',
                executionMode: 'sequential',
                intraOpNumThreads: numThreads,
                interOpNumThreads: 1,
                enableCpuMemArena: true,
                enableMemPattern: true,
            };

            // Processor (configs JSON, leve)
            samState.processor = await tx.AutoProcessor.from_pretrained(modelId, {
                progress_callback: (data) => {
                    console.log('[sam] processor:', data.status, data.file, data.progress);
                    if (data.status === 'progress' && data.progress != null) {
                        broadcastSamProgress(eventSender, {
                            phase: 'download', msg: `Baixando ${data.file || 'config'}...`,
                            pct: 10 + Math.min(20, Math.floor((data.progress || 0) * 0.2))
                        });
                    }
                },
            });
            broadcastSamProgress(eventSender, { phase: 'download', msg: 'Processor OK · baixando modelo .onnx...', pct: 30 });

            // Model — q8 quantizado + thread limits (não trava PC)
            samState.model = await tx.SamModel.from_pretrained(modelId, {
                dtype: 'q8',
                session_options: sessionOpts,
                progress_callback: (data) => {
                    console.log('[sam] model:', data.status, data.file, data.progress);
                    if (data.status === 'progress' && data.progress != null) {
                        broadcastSamProgress(eventSender, {
                            phase: 'download', msg: `Baixando ${data.file || 'model.onnx'}: ${Math.round(data.progress)}%`,
                            pct: 30 + Math.min(60, Math.floor((data.progress || 0) * 0.6))
                        });
                    } else if (data.status === 'done' && data.file) {
                        broadcastSamProgress(eventSender, {
                            phase: 'download', msg: `Carregando ${data.file}...`, pct: 92
                        });
                    } else if (data.status === 'ready') {
                        broadcastSamProgress(eventSender, { phase: 'ready', msg: 'Iniciando inferência...', pct: 95 });
                    }
                },
            });

            samState.loaded = true;
            broadcastSamProgress(eventSender, { phase: 'ready', msg: 'SAM pronto', pct: 100 });
            console.log('[sam] modelo carregado ✓');
            return { ok: true };
        } catch (e) {
            console.error('[sam] load failed:', e);
            samState.loaded = false;
            samState.model = null;
            samState.processor = null;
            samState.loading = null;
            broadcastSamProgress(eventSender, { phase: 'error', msg: e.message || String(e), pct: 0 });
            return { error: e.message || String(e), stack: (e.stack || '').slice(0, 600) };
        }
    })();

    const result = await samState.loading;
    if (!result.error) samState.loading = null; // reset on success too (so re-call after error retries)
    return result;
}

// Carrega o modelo SAM
ipcMain.handle('roto:sam-load', async (event) => {
    return await loadSamModel(event.sender);
});

// Encoda um frame e cacheia embeddings na main process pelo token
ipcMain.handle('roto:sam-encode', async (event, payload) => {
    try {
        if (!samState.loaded) return { error: 'SAM não carregado — chame roto:sam-load primeiro' };
        const { token, width, height, rgba } = payload || {};
        if (!token) return { error: 'token ausente' };
        if (!width || !height || !rgba) return { error: 'dados de imagem ausentes' };

        // rgba vem como Uint8Array (estructure clone via IPC)
        const u8 = rgba instanceof Uint8Array ? rgba :
                   (rgba.buffer ? new Uint8Array(rgba.buffer, rgba.byteOffset || 0, rgba.byteLength) :
                                  new Uint8Array(rgba));
        const tx = samState.tx;
        const rawImage = new tx.RawImage(u8, width, height, 4);
        const inputs = await samState.processor(rawImage);
        const embeddings = await samState.model.get_image_embeddings(inputs);

        samState.embeddings.set(token, {
            embeddings,
            inputs,        // mantém pra original_sizes / reshaped_input_sizes
            width, height,
            t: Date.now(),
        });

        // Limpa embeddings antigos (>30min)
        const cutoff = Date.now() - 30 * 60 * 1000;
        for (const [k, v] of samState.embeddings) {
            if (v.t < cutoff) samState.embeddings.delete(k);
        }

        return { ok: true };
    } catch (e) {
        console.error('[sam-encode]', e);
        return { error: e.message || String(e) };
    }
});

// Segmenta com lista de clicks e retorna binary mask
ipcMain.handle('roto:sam-segment', async (event, payload) => {
    try {
        if (!samState.loaded) return { error: 'SAM não carregado' };
        const { token, clicks } = payload || {};
        const cache = samState.embeddings.get(token);
        if (!cache) return { error: 'embeddings não encontrado — encode o frame primeiro' };
        if (!Array.isArray(clicks) || clicks.length === 0) return { error: 'sem pontos' };

        const points = clicks.map(c => [c.x, c.y]);
        const labels = clicks.map(c => c.kind === 'positive' ? 1 : 0);

        // v4 API: usa reshape_input_points + add_input_labels diretamente
        // (em v4 o processor() exige uma imagem, mas como já temos embeddings,
        // criamos os tensors de prompt direto via image_processor)
        const imgProc = samState.processor.image_processor;
        const inputPointsTensor = imgProc.reshape_input_points(
            [[points]],
            cache.inputs.original_sizes,
            cache.inputs.reshaped_input_sizes,
        );
        const inputLabelsTensor = imgProc.add_input_labels(
            [[labels]],
            inputPointsTensor,
        );

        // Reusa embeddings cacheados — model.forward aceita image_embeddings + input_points
        const outputs = await samState.model({
            input_points: inputPointsTensor,
            input_labels: inputLabelsTensor,
            image_embeddings: cache.embeddings.image_embeddings,
            image_positional_embeddings: cache.embeddings.image_positional_embeddings,
        });

        const masks = await samState.processor.post_process_masks(
            outputs.pred_masks,
            cache.inputs.original_sizes,
            cache.inputs.reshaped_input_sizes,
        );

        const maskTensor = masks[0];
        const dims = maskTensor.dims;
        const H = dims[dims.length - 2];
        const W = dims[dims.length - 1];
        const data = maskTensor.data;

        // Pega scores pra escolher melhor predição
        const iouScores = outputs.iou_scores?.data;
        const numPreds = dims.length === 4 ? dims[1] : 1;
        let bestIdx = 0;
        if (iouScores && numPreds > 1) {
            let bestScore = -1;
            for (let i = 0; i < numPreds; i++) {
                if (iouScores[i] > bestScore) { bestScore = iouScores[i]; bestIdx = i; }
            }
        }
        const offset = bestIdx * (W * H);

        // Empacota em Uint8Array (1 byte por pixel: 0 ou 255)
        const out = new Uint8Array(W * H);
        for (let i = 0; i < W * H; i++) {
            out[i] = data[offset + i] ? 255 : 0;
        }

        return { ok: true, width: W, height: H, mask: out };
    } catch (e) {
        console.error('[sam-segment]', e);
        return { error: e.message || String(e), stack: (e.stack || '').slice(0, 400) };
    }
});

// Limpa embedding cache de um token (chamado quando renderer fecha)
ipcMain.on('roto:sam-clear', (event, payload) => {
    try { samState.embeddings.delete(payload?.token); } catch(e) {}
});

// EXPORT — aplica masks + gera .mov ProRes 4444 com alpha
ipcMain.handle('rotoscope:export', async (event, payload) => {
    try {
        const { token, masks, srcPath, outPath, fps, totalFrames, width, height } = payload || {};
        if (!srcPath || !fs.existsSync(srcPath)) return { error: 'src não existe' };
        if (!ffmpegReady() || !ffprobeReady()) await ensureFfmpeg();
        if (!ffmpegReady()) return { error: 'ffmpeg não disponível' };

        const finalOut = outPath || path.join(path.dirname(srcPath),
            path.basename(srcPath, path.extname(srcPath)) + '_rotoscope.mov');

        // Cria pasta temp pra masks PNG
        const tmpDir = path.join(currentUD, 'rotoscope-' + Date.now());
        fs.mkdirSync(tmpDir, { recursive: true });

        // Salva cada mask como PNG na pasta temp (interpola pra frames sem mask)
        const maskKeys = Object.keys(masks).map(Number).sort((a,b)=>a-b);
        if (maskKeys.length === 0) return { error: 'sem máscaras' };

        for (let i = 0; i < totalFrames; i++) {
            // Acha mask mais próxima (frame anterior ou igual)
            let useFrame = maskKeys[0];
            for (const k of maskKeys) {
                if (k <= i) useFrame = k; else break;
            }
            const m = masks[useFrame];
            if (!m) continue;
            const base64 = m.dataUrl.replace(/^data:image\/png;base64,/, '');
            const fname = path.join(tmpDir, 'mask-' + String(i).padStart(6,'0') + '.png');
            fs.writeFileSync(fname, Buffer.from(base64, 'base64'));
        }

        // FFmpeg: aplica masks como alpha frame-a-frame
        // 2 inputs: source video + masks PNG sequence
        // filter: alphamerge usa luminância da mask como alpha
        const args = [
            '-y',
            '-i', srcPath,
            '-framerate', String(fps),
            '-i', path.join(tmpDir, 'mask-%06d.png'),
            '-filter_complex',
            // [1:v] = masks · scale pra match do video · usa luma como alpha
            '[1:v]scale=' + width + ':' + height + ',format=gray[mask];[0:v][mask]alphamerge[out]',
            '-map', '[out]',
            '-map', '0:a?',
            '-c:v', 'prores_ks',
            '-profile:v', '4',
            '-pix_fmt', 'yuva444p10le',
            '-c:a', 'aac', '-b:a', '192k',
            '-shortest',
            finalOut,
        ];
        console.log('[rotoscope/export] ffmpeg', args.join(' '));

        await new Promise((resolve, reject) => {
            const proc = bgvSpawn(ffmpegBin, args, { windowsHide: true });
            let stderr = '';
            proc.stderr.on('data', d => { stderr += d.toString(); if (stderr.length > 8192) stderr = stderr.slice(-8192); });
            proc.on('close', code => {
                if (code === 0 && fs.existsSync(finalOut)) resolve();
                else reject(new Error('ffmpeg code ' + code + ': ' + stderr.slice(-400)));
            });
            proc.on('error', e => reject(e));
        });

        // Cleanup PNGs temp
        try {
            for (const f of fs.readdirSync(tmpDir)) fs.unlinkSync(path.join(tmpDir, f));
            fs.rmdirSync(tmpDir);
        } catch(e) {}

        const sess = rotoscopeSessions.get(token);
        if (sess) sess.outPath = finalOut;

        return { ok: true, outPath: finalOut };
    } catch (e) {
        return { error: e.message || String(e) };
    }
});

// Window controls da janela frameless do editor
ipcMain.on('mask-editor:win-action', (event, payload) => {
    try {
        const { token, action } = payload || {};
        const sess = maskEditorSessions.get(token);
        if (!sess || !sess.win || sess.win.isDestroyed()) return;
        const w = sess.win;
        if (action === 'minimize') w.minimize();
        else if (action === 'maximize') {
            if (w.isMaximized()) w.unmaximize();
            else w.maximize();
        } else if (action === 'close') {
            sess.status = sess.status === 'editing' ? 'cancelled' : sess.status;
            w.close();
        }
    } catch(e) { console.error('[mask-editor:win-action]', e); }
});

// ════════════════════════════════════════════════════════════════════
// BG REMOVE — VÍDEO (MediaPipe via worker hidden + FFmpeg encoder)
// Pipeline: worker abre vídeo → MediaPipe segmenta cada frame com WebGL →
//           pixels RGBA enviados via IPC → main escreve no stdin do ffmpeg
//           → ffmpeg encoda em WebM (VP9 alpha) ou MOV (ProRes 4444)
// ════════════════════════════════════════════════════════════════════
const { spawn: bgvSpawn } = require('child_process');

const bgVideoSessions = new Map();
// token -> { status, progress, encoderProc, outPath, win, srcPath, format, error, startedAt }

function makeBgVideoSession() {
    const token = crypto.randomBytes(8).toString('hex');
    const sess = {
        status: 'starting',          // starting|encoding|done|cancelled|error
        progress: { current: 0, total: 0, pct: 0, etaSec: 0, phase: 'init' },
        encoderProc: null,
        outPath: null,
        win: null,
        startedAt: Date.now(),
        error: null,
    };
    bgVideoSessions.set(token, sess);
    // Limpa sessions antigas (>2h)
    const now = Date.now();
    for (const [k, v] of bgVideoSessions) {
        if (v.startedAt && (now - v.startedAt) > 2 * 60 * 60 * 1000) {
            bgVideoSessions.delete(k);
        }
    }
    return token;
}

function startBgVideoWorker(srcPath, format, quality, delegate) {
    const token = makeBgVideoSession();
    const sess = bgVideoSessions.get(token);
    sess.srcPath = srcPath;
    sess.format = format;
    sess.quality = quality;
    sess.delegate = (delegate || 'GPU').toUpperCase();
    sess.lastPreview = null; // dataUrl da última preview pro plugin/app pollar

    try {
        const win = new BrowserWindow({
            width: 720,
            height: 520,
            show: false,                  // hidden — worker não tem UI pro user
            backgroundColor: '#050505',
            title: 'Lion BG Video Worker',
            webPreferences: {
                nodeIntegration: true,
                contextIsolation: false,
                webgl: true,
                offscreen: false,
            },
        });
        sess.win = win;

        const workerHtml = path.join(__dirname, 'bg-video-worker.html');
        const qs = '?token=' + encodeURIComponent(token)
                 + '&src=' + encodeURIComponent(srcPath)
                 + '&format=' + encodeURIComponent(format || 'webm')
                 + '&quality=' + encodeURIComponent(quality || 'medium')
                 + '&delegate=' + encodeURIComponent(sess.delegate);
        win.loadFile(workerHtml, { search: qs.replace(/^\?/, '') });

        // Pra debug — descomenta pra ver o worker
        // win.webContents.openDevTools({ mode: 'detach' });
        // win.show();

        win.on('closed', () => {
            const s = bgVideoSessions.get(token);
            if (s && (s.status === 'starting' || s.status === 'encoding')) {
                s.status = 'cancelled';
                try { if (s.encoderProc) s.encoderProc.kill('SIGKILL'); } catch(e) {}
            }
        });

        return token;
    } catch (e) {
        sess.status = 'error';
        sess.error = e.message;
        return token;
    }
}

// ─── IPC: probe via ffprobe (dimensões, fps, duração) ─────────────
ipcMain.handle('bg-video:probe', async (event, payload) => {
    try {
        const src = payload?.src;
        if (!src || !fs.existsSync(src)) return { error: 'arquivo não encontrado' };
        // Auto-download ffmpeg/ffprobe se não estiver disponível
        if (!ffprobeReady() || !ffmpegReady()) {
            console.log('[bg-video:probe] baixando ffmpeg/ffprobe (1ª vez)...');
            await ensureFfmpeg();
        }
        if (!ffprobeReady()) return { error: 'ffprobe não disponível (download falhou — verifique conexão)' };

        return await new Promise((resolve) => {
            const proc = bgvSpawn(ffprobeBin, [
                '-v', 'error',
                '-select_streams', 'v:0',
                '-show_entries', 'stream=width,height,r_frame_rate,duration,nb_frames',
                '-show_entries', 'format=duration',
                '-of', 'json',
                src
            ], { windowsHide: true });
            let out = '';
            proc.stdout.on('data', d => out += d.toString());
            proc.on('close', () => {
                try {
                    const j = JSON.parse(out);
                    const stream = j.streams?.[0] || {};
                    const fmt = j.format || {};
                    const fpsStr = stream.r_frame_rate || '30/1';
                    const [num, den] = fpsStr.split('/').map(Number);
                    const fps = (num && den) ? num / den : 30;
                    const dur = parseFloat(stream.duration || fmt.duration || '0') || 0;
                    resolve({
                        width: parseInt(stream.width) || 0,
                        height: parseInt(stream.height) || 0,
                        fps: Math.round(fps * 1000) / 1000,
                        duration: dur,
                        nbFrames: parseInt(stream.nb_frames) || 0,
                    });
                } catch(e) { resolve({ error: 'parse falhou: ' + e.message }); }
            });
            proc.on('error', e => resolve({ error: e.message }));
        });
    } catch (e) { return { error: e.message }; }
});

// ─── IPC: get userData path (worker não tem acesso direto) ──────
ipcMain.handle('bg-video:get-userdata', () => app.getPath('userData'));

// ─── IPC: inicia FFmpeg encoder com stdin pipe ─────────────────
ipcMain.handle('bg-video:start-encoder', async (event, payload) => {
    try {
        const { token, width, height, fps, format, srcPath } = payload || {};
        const sess = bgVideoSessions.get(token);
        if (!sess) return { error: 'sessão inválida' };
        if (!fs.existsSync(ffmpegBin)) return { error: 'ffmpeg não disponível' };

        // Output path: ao lado do ARQUIVO ORIGINAL (não do temp transcoded).
        // Antes salvava em currentUD/ com nome 'bg-video-trim-XXX_nobg.webm'
        // que era confuso e ficava no roaming. Agora salva ao lado do vídeo
        // original com nome bonito.
        const refPath = sess._origSrcPath || srcPath;
        const dir = path.dirname(refPath);
        const base = path.basename(refPath, path.extname(refPath));
        const ext = format === 'mov' ? '.mov' : '.webm';
        let outPath = path.join(dir, base + '_nobg' + ext);
        // Se o dir não for gravável, fallback pro currentUD
        try {
            const testFile = path.join(dir, '.lw-write-test-' + Date.now());
            fs.writeFileSync(testFile, ''); fs.unlinkSync(testFile);
        } catch (e) {
            outPath = path.join(currentUD, base + '_nobg' + ext);
        }
        // Garante unicidade
        let n = 0;
        while (fs.existsSync(outPath) && n < 100) {
            n++;
            outPath = path.join(path.dirname(outPath), path.basename(outPath, ext) + '-' + n + ext);
            // evita acumular sufixos: refaz do base
            outPath = path.join(path.dirname(outPath), base + '_nobg-' + n + ext);
        }

        // FFmpeg args:
        //   stdin: rawvideo RGBA WxH @ fps
        //   audio: pega do source original (não toca)
        //   output: WebM VP9 alpha OU MOV ProRes 4444
        const inputArgs = [
            '-y',
            '-f', 'rawvideo',
            '-pix_fmt', 'rgba',
            '-s', `${width}x${height}`,
            '-r', String(fps),
            '-i', 'pipe:0',
            // áudio do original (se houver)
            '-i', srcPath,
            '-map', '0:v:0',
            '-map', '1:a:0?',
        ];

        let outputArgs;
        if (format === 'mov') {
            // ProRes 4444 com alpha (yuva444p10le) — qualidade máxima
            outputArgs = [
                '-c:v', 'prores_ks',
                '-profile:v', '4',
                '-pix_fmt', 'yuva444p10le',
                '-c:a', 'aac', '-b:a', '192k',
                '-shortest',
                outPath
            ];
        } else {
            // WebM VP9 com alpha (yuva420p)
            outputArgs = [
                '-c:v', 'libvpx-vp9',
                '-pix_fmt', 'yuva420p',
                '-b:v', '0',
                '-crf', '30',
                '-deadline', 'realtime',
                '-cpu-used', '4',
                '-row-mt', '1',
                '-c:a', 'libopus', '-b:a', '128k',
                '-shortest',
                outPath
            ];
        }

        const args = inputArgs.concat(outputArgs);
        console.log('[bg-video] ffmpeg', args.join(' '));

        const proc = bgvSpawn(ffmpegBin, args, { windowsHide: true });
        sess.encoderProc = proc;
        sess.outPath = outPath;
        sess.status = 'encoding';

        let stderrBuf = '';
        proc.stderr.on('data', d => {
            stderrBuf += d.toString();
            // Mantém só últimos 4KB pra não estourar memória
            if (stderrBuf.length > 4096) stderrBuf = stderrBuf.slice(-4096);
        });
        proc.on('error', e => {
            sess.status = 'error';
            sess.error = 'ffmpeg error: ' + e.message;
        });
        proc.on('close', code => {
            sess._exitCode = code;
            sess._stderrTail = stderrBuf;
        });

        // Backpressure: stdin pode encher. Promisify writes.
        proc.stdin.on('error', (e) => {
            // EPIPE comum se ffmpeg morre — captura
            console.warn('[bg-video] stdin error:', e.message);
        });

        return outPath;
    } catch (e) { return { error: e.message }; }
});

// ─── IPC: escreve frame RGBA no stdin do ffmpeg ────────────────
ipcMain.handle('bg-video:write-frame', async (event, payload) => {
    try {
        const { token, buffer } = payload || {};
        const sess = bgVideoSessions.get(token);
        if (!sess || !sess.encoderProc) return { error: 'encoder não iniciado' };

        const buf = Buffer.isBuffer(buffer) ? buffer : Buffer.from(buffer);
        return await new Promise((resolve) => {
            const ok = sess.encoderProc.stdin.write(buf, (err) => {
                if (err) resolve({ error: err.message });
                else resolve({ ok: true });
            });
            if (!ok) {
                // Backpressure: aguarda drain
                sess.encoderProc.stdin.once('drain', () => resolve({ ok: true }));
            }
        });
    } catch (e) { return { error: e.message }; }
});

// ─── IPC: fecha stdin e espera ffmpeg terminar ─────────────────
ipcMain.handle('bg-video:close-encoder', async (event, payload) => {
    try {
        const { token } = payload || {};
        const sess = bgVideoSessions.get(token);
        if (!sess || !sess.encoderProc) return { error: 'encoder não iniciado' };

        return await new Promise((resolve) => {
            const proc = sess.encoderProc;
            proc.on('close', code => {
                sess.status = code === 0 ? 'done' : 'error';
                if (code !== 0) {
                    sess.error = 'ffmpeg saiu com code ' + code + ': ' + (sess._stderrTail || '').slice(-500);
                    resolve({ error: sess.error });
                } else {
                    resolve({ ok: true, outPath: sess.outPath });
                }
                // Fecha worker window depois de pequeno delay
                setTimeout(() => {
                    try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
                }, 1000);
            });
            try { proc.stdin.end(); } catch(e) {}
        });
    } catch (e) { return { error: e.message }; }
});

// ─── IPC: aborta encoder (cancel) ──────────────────────────────
ipcMain.handle('bg-video:abort-encoder', async (event, payload) => {
    try {
        const { token } = payload || {};
        const sess = bgVideoSessions.get(token);
        if (!sess) return { error: 'sessão inválida' };
        sess.status = 'cancelled';
        try { if (sess.encoderProc) sess.encoderProc.kill('SIGKILL'); } catch(e) {}
        try { if (sess.outPath && fs.existsSync(sess.outPath)) fs.unlinkSync(sess.outPath); } catch(e) {}
        try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
        return { ok: true };
    } catch (e) { return { error: e.message }; }
});

// ─── IPC: progress / phase / log do worker pra rastreamento ─────
ipcMain.on('bg-video:progress', (event, payload) => {
    const sess = bgVideoSessions.get(payload?.token);
    if (sess) {
        sess.progress.current = payload.current || 0;
        sess.progress.total = payload.total || 0;
        sess.progress.pct = payload.pct || 0;
        sess.progress.etaSec = payload.etaSec || 0;
    }
});
ipcMain.on('bg-video:phase', (event, payload) => {
    const sess = bgVideoSessions.get(payload?.token);
    if (sess) sess.progress.phase = payload.phase || '';
});
ipcMain.on('bg-video:log', (event, payload) => {
    console.log('[bg-video][' + (payload?.token || '?').slice(0,4) + '] ' + payload?.msg);
});
ipcMain.on('bg-video:done', (event, payload) => {
    const sess = bgVideoSessions.get(payload?.token);
    if (!sess) return;
    if (payload?.ok) {
        sess.status = 'done';
        sess.outPath = payload.outPath || sess.outPath;
    } else {
        sess.status = 'error';
        sess.error = payload?.error || 'erro desconhecido';
    }
    // Cleanup do arquivo temp do pré-transcode
    if (sess._tempTranscoded) {
        try { if (fs.existsSync(sess._tempTranscoded)) fs.unlinkSync(sess._tempTranscoded); } catch(e) {}
    }
});
ipcMain.on('bg-video:ready', (event, payload) => {
    console.log('[bg-video] worker ready:', payload?.token);
});

// Preview frames vindo do worker — guarda pra cliente pollar via HTTP
ipcMain.on('bg-video:preview', (event, payload) => {
    const sess = bgVideoSessions.get(payload?.token);
    if (sess) {
        sess.lastPreview = payload.dataUrl;
        sess.lastPreviewFrame = payload.frame || 0;
    }
    // Repassa pra mainWindow se existir (UI do app principal)
    try {
        if (mainWin && !mainWin.isDestroyed()) {
            mainWin.webContents.send('bg-video:preview', payload);
        }
    } catch(e) {}
});

// ─── HTTP: detecta GPUs disponíveis ───────────────────────────
async function detectGpus() {
    const result = { hasGpu: false, vendor: 'unknown', renderer: '', platform: process.platform };
    try {
        // Cria um BrowserWindow hidden temporário pra fazer query WebGL
        const tmpWin = new BrowserWindow({
            width: 100, height: 100, show: false,
            webPreferences: { nodeIntegration: true, contextIsolation: false, webgl: true }
        });
        await tmpWin.loadURL('data:text/html,<canvas id="c"></canvas><script>' +
            'const c=document.getElementById("c");' +
            'const gl=c.getContext("webgl2")||c.getContext("webgl");' +
            'if(gl){' +
            '  const ext=gl.getExtension("WEBGL_debug_renderer_info");' +
            '  const r=ext?gl.getParameter(ext.UNMASKED_RENDERER_WEBGL):"";' +
            '  const v=ext?gl.getParameter(ext.UNMASKED_VENDOR_WEBGL):"";' +
            '  document.title="GPU::"+v+"::"+r;' +
            '} else { document.title="GPU::none::"; }' +
            '</script>');
        await new Promise(r => setTimeout(r, 200));
        const title = tmpWin.getTitle();
        if (title.indexOf('GPU::') === 0) {
            const parts = title.substring(5).split('::');
            result.vendor = parts[0] || '';
            result.renderer = parts[1] || '';
            result.hasGpu = !!(result.renderer && result.vendor !== 'none');
        }
        try { tmpWin.close(); } catch(e) {}
    } catch (e) { console.error('[bg-video] gpu detect:', e); }
    return result;
}

// AI config storage (OpenAI API key)
const aiConfigPath = path.join(currentUD, 'ai-config.json');
function getAiConfig() { try { return JSON.parse(fs.readFileSync(aiConfigPath, 'utf8')); } catch(e) { return {}; } }
function saveAiConfig(cfg) { try { fs.writeFileSync(aiConfigPath, JSON.stringify(cfg), 'utf8'); } catch(e) {} }

function ytDlpReady() { return fs.existsSync(ytDlpBin); }
function ffmpegReady() { return fs.existsSync(ffmpegBin); }

// Safe-remove: try to delete a file, swallow EPERM/ENOENT
function safeRm(p) {
    try { if (fs.existsSync(p)) fs.unlinkSync(p); return true; }
    catch (e) { return false; }
}

function downloadBinary(url, dest) {
    return new Promise((resolve, reject) => {
        // Pre-cleanup: remove any stale partial file (handles EPERM from locked files)
        try { if (fs.existsSync(dest)) fs.unlinkSync(dest); }
        catch (e) { return reject(new Error('Não foi possível limpar arquivo anterior: ' + e.message + '. Feche o antivírus ou outras instâncias do app.')); }

        let file;
        try { file = fs.createWriteStream(dest); }
        catch (e) { return reject(new Error('Não foi possível criar arquivo: ' + e.message)); }

        const cleanup = () => { try { file.close(); } catch {} safeRm(dest); };

        file.on('error', (err) => { cleanup(); reject(err); });

        const request = require('https').get(url, { headers: { 'User-Agent': 'LionWorkspace/1.0' } }, (response) => {
            if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
                cleanup();
                downloadBinary(response.headers.location, dest).then(resolve, reject);
                return;
            }
            if (response.statusCode !== 200) { cleanup(); reject(new Error('HTTP ' + response.statusCode)); return; }
            response.pipe(file);
            file.on('finish', () => { try { file.close(); } catch {} resolve(); });
        });
        request.on('error', (err) => { cleanup(); reject(err); });
        request.setTimeout(120000, () => { request.destroy(); cleanup(); reject(new Error('Timeout no download')); });
    });
}

async function ensureYtDlp() {
    if (!fs.existsSync(ytDlpDir)) fs.mkdirSync(ytDlpDir, { recursive: true });
    if (ytDlpReady()) return true;
    const url = isWin
        ? 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp.exe'
        : 'https://github.com/yt-dlp/yt-dlp/releases/latest/download/yt-dlp_macos';
    try {
        const dest = isWin ? ytDlpBin : ytDlpBin;
        await downloadBinary(url, dest);
        if (!isWin) fs.chmodSync(dest, '755');
        return true;
    } catch (e) { console.error('yt-dlp download failed:', e.message); return false; }
}

function ffprobeReady() { return fs.existsSync(ffprobeBin); }

async function ensureFfmpeg() {
    // Antes só checava ffmpeg — agora exige BOTH ffmpeg E ffprobe.
    // bg-remove de vídeo precisa do ffprobe pra detectar dimensões/fps.
    if (ffmpegReady() && ffprobeReady()) return true;

    // Try to find system ffmpeg/ffprobe first
    const findBin = (name) => new Promise(resolve => {
        exec(isWin ? `where ${name}` : `which ${name}`, { timeout: 5000 }, (err, stdout) => {
            resolve(err ? '' : (stdout || '').trim().split('\n')[0].trim());
        });
    });

    // Link a system binary to our dir
    function linkBin(systemPath, localPath) {
        if (!systemPath || !fs.existsSync(systemPath)) return false;
        try {
            if (fs.existsSync(localPath)) return true;
            if (isWin) fs.copyFileSync(systemPath, localPath);
            else fs.symlinkSync(systemPath, localPath);
            return true;
        } catch (e) { return false; }
    }

    // 1) Check PATH
    const sysFF = await findBin('ffmpeg');
    const sysProbe = await findBin('ffprobe');
    if (sysFF) {
        linkBin(sysFF, ffmpegBin);
        linkBin(sysProbe, ffprobeBin);
        if (ffmpegReady() && ffprobeReady()) return true;
    }

    // 2) macOS: check common Homebrew / MacPorts paths
    if (isMac) {
        const macPaths = [
            '/opt/homebrew/bin',    // Homebrew Apple Silicon
            '/usr/local/bin',       // Homebrew Intel
            '/opt/local/bin',       // MacPorts
        ];
        for (const dir of macPaths) {
            const ff = path.join(dir, 'ffmpeg');
            const fp = path.join(dir, 'ffprobe');
            if (fs.existsSync(ff)) {
                linkBin(ff, ffmpegBin);
                linkBin(fp, ffprobeBin);
                if (ffmpegReady() && ffprobeReady()) return true;
            }
        }

        // 3) macOS: download ffmpeg + ffprobe from evermeet.cx (universal builds)
        try {
            const dlDir = path.join(ytDlpDir, 'ff-extract');
            fs.mkdirSync(dlDir, { recursive: true });

            for (const bin of ['ffmpeg', 'ffprobe']) {
                const zipUrl = `https://evermeet.cx/ffmpeg/getrelease/${bin}/zip`;
                const zipPath = path.join(dlDir, `${bin}.zip`);
                const targetPath = bin === 'ffmpeg' ? ffmpegBin : ffprobeBin;

                await downloadBinary(zipUrl, zipPath);
                // Unzip on Mac
                await new Promise((resolve, reject) => {
                    exec(`unzip -o "${zipPath}" -d "${dlDir}"`, { timeout: 60000 },
                        (err) => { if (err) reject(err); else resolve(); });
                });
                const extracted = path.join(dlDir, bin);
                if (fs.existsSync(extracted)) {
                    fs.copyFileSync(extracted, targetPath);
                    fs.chmodSync(targetPath, 0o755);
                }
            }
            // Cleanup
            try { fs.rmSync(dlDir, { recursive: true, force: true }); } catch {}
            return ffmpegReady() && ffprobeReady();
        } catch (e) { console.error('macOS ffmpeg download failed:', e.message); return false; }
    }

    // 4) Windows: download and extract
    try {
        const ffUrl = 'https://github.com/yt-dlp/FFmpeg-Builds/releases/download/latest/ffmpeg-master-latest-win64-gpl.zip';
        const zipPath = path.join(ytDlpDir, 'ffmpeg.zip');
        const extractDir = path.join(ytDlpDir, 'ffmpeg-extract');

        // Pre-clean any leftovers from previous failed attempts
        safeRm(zipPath);
        try { fs.rmSync(extractDir, { recursive: true, force: true }); } catch {}

        await downloadBinary(ffUrl, zipPath);
        await new Promise((resolve, reject) => {
            exec(`powershell -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${extractDir}' -Force; $f=Get-ChildItem '${extractDir}' -Recurse -Filter 'ffmpeg.exe' | Select-Object -First 1; Copy-Item $f.FullName '${ffmpegBin}' -Force; $p=Get-ChildItem '${extractDir}' -Recurse -Filter 'ffprobe.exe' | Select-Object -First 1; if($p){Copy-Item $p.FullName '${ffprobeBin}' -Force}"`,
                { windowsHide: true, timeout: 120000 },
                (err) => { if (err) reject(err); else resolve(); });
        });
        safeRm(zipPath);
        try { fs.rmSync(extractDir, { recursive: true, force: true }); } catch {}
        return ffmpegReady() && ffprobeReady();
    } catch (e) { console.error('ffmpeg download failed:', e.message); return false; }
}

// Strip playlist/index params from YouTube URLs — keep only the video ID
function cleanYtUrl(raw) {
    try {
        const u = new URL(raw);
        if (u.hostname.includes('youtube.com') && u.searchParams.has('v')) {
            return u.origin + u.pathname + '?v=' + u.searchParams.get('v');
        }
    } catch (e) {}
    return raw;
}

// ────────────────────────────────────────────────────────────────────
// YouTube bot-detection bypass helpers
// YouTube vem requerendo cookies/auth pra muitos vídeos desde 2024.
// Estratégia: usar `tv` player client (que dispensa PO token) e tentar
// cookies do browser do user em ordem de prioridade.
// ────────────────────────────────────────────────────────────────────

// Browsers detectados na ordem de tentativa
function detectInstalledBrowsers() {
    const browsers = [];
    if (isWin) {
        const candidates = [
            ['chrome',  path.join(process.env.LOCALAPPDATA || '', 'Google/Chrome/User Data')],
            ['edge',    path.join(process.env.LOCALAPPDATA || '', 'Microsoft/Edge/User Data')],
            ['brave',   path.join(process.env.LOCALAPPDATA || '', 'BraveSoftware/Brave-Browser/User Data')],
            ['firefox', path.join(process.env.APPDATA || '', 'Mozilla/Firefox/Profiles')],
            ['opera',   path.join(process.env.APPDATA || '', 'Opera Software/Opera Stable')],
        ];
        for (const [name, p] of candidates) {
            try { if (fs.existsSync(p)) browsers.push(name); } catch(e) {}
        }
    } else if (isMac) {
        const home = os.homedir();
        const candidates = [
            ['chrome',  path.join(home, 'Library/Application Support/Google/Chrome')],
            ['safari',  path.join(home, 'Library/Cookies/Cookies.binarycookies')],
            ['firefox', path.join(home, 'Library/Application Support/Firefox/Profiles')],
            ['brave',   path.join(home, 'Library/Application Support/BraveSoftware/Brave-Browser')],
            ['edge',    path.join(home, 'Library/Application Support/Microsoft Edge')],
        ];
        for (const [name, p] of candidates) {
            try { if (fs.existsSync(p)) browsers.push(name); } catch(e) {}
        }
    }
    return browsers;
}

// Args base pra bypass de bot — sem cookies (rápido/leve)
// Ordem dos clients importa:
// - web/android: streams não-DRM, mais limpos
// - tv/tv_simply: ajuda bot detection mas pode retornar DRM em alguns videos
// - mweb: backup
function ytBypassArgs() {
    return [
        '--extractor-args', 'youtube:player_client=default,web,android,mweb,tv_simply',
    ];
}

// Args alternativos quando primeiro falha com DRM
function ytBypassArgsAltDRM() {
    return [
        '--extractor-args', 'youtube:player_client=android,web_safari,ios,web',
    ];
}

// Detecta se o stderr de yt-dlp indica bot detection (precisa de cookies)
function ytIsBotError(stderr) {
    if (!stderr) return false;
    return /Sign in to confirm|not a bot|cookies-from-browser|Use --cookies/i.test(stderr);
}

function ytIsDrmError(stderr) {
    if (!stderr) return false;
    return /DRM protected|DRM-protected|drm protected/i.test(stderr);
}

// Roda yt-dlp com bypass + retry com cookies se cair em bot error.
// onProc(proc, attemptIdx) chamado quando spawn → opção pra hooked progress parsing
// ou para registrar em ytDownloadProc, etc.
async function runYtDlpRobust(userArgs, opts = {}) {
    const { onProc, onStdoutLine, onStderrLine, timeout } = opts;
    const browsersToTry = detectInstalledBrowsers();
    // 1ª tentativa: sem cookies, com bypass clients ideais
    // 2ª: alt clients (pra DRM)
    // 3ª+: cookies de cada browser
    const attempts = [
        { args: [...ytBypassArgs()], label: 'bypass-default' },
        { args: [...ytBypassArgsAltDRM()], label: 'bypass-alt-drm' },
        ...browsersToTry.map(b => ({ args: ['--cookies-from-browser', b, ...ytBypassArgs()], label: 'cookies-' + b })),
    ];

    let lastResult = null;
    for (let i = 0; i < attempts.length; i++) {
        const att = attempts[i];
        const fullArgs = [...att.args, ...userArgs];
        console.log('[yt-dlp]', att.label, fullArgs.join(' '));
        const result = await new Promise((resolve) => {
            const proc = spawn(ytDlpBin, fullArgs, { windowsHide: true });
            if (onProc) onProc(proc, i);
            let stdout = '', stderr = '';
            const stdoutBuf = [];
            const stderrBuf = [];
            proc.stdout.on('data', d => {
                const s = d.toString();
                stdout += s;
                if (onStdoutLine) {
                    const lines = s.split('\n');
                    if (stdoutBuf.length) { stdoutBuf[stdoutBuf.length-1] += lines.shift(); }
                    stdoutBuf.push(...lines);
                    while (stdoutBuf.length > 1) onStdoutLine(stdoutBuf.shift());
                }
            });
            proc.stderr.on('data', d => {
                const s = d.toString();
                stderr += s;
                if (onStderrLine) {
                    const lines = s.split('\n');
                    if (stderrBuf.length) { stderrBuf[stderrBuf.length-1] += lines.shift(); }
                    stderrBuf.push(...lines);
                    while (stderrBuf.length > 1) onStderrLine(stderrBuf.shift());
                }
            });
            proc.on('close', code => {
                if (onStdoutLine && stdoutBuf.length) onStdoutLine(stdoutBuf.shift());
                if (onStderrLine && stderrBuf.length) onStderrLine(stderrBuf.shift());
                resolve({ code, stdout, stderr, attempt: att.label });
            });
            proc.on('error', e => resolve({ code: -1, stdout: '', stderr: e.message, attempt: att.label }));
            if (timeout) setTimeout(() => { try { proc.kill(); } catch {} }, timeout);
        });
        lastResult = result;
        if (result.code === 0) return result;

        // Bot error: tenta próxima (com cookies)
        // DRM error: tenta próxima (alt clients)
        // Outro erro: retorna direto
        if (ytIsBotError(result.stderr)) {
            console.log('[yt-dlp] bot error detected, trying next strategy...');
        } else if (ytIsDrmError(result.stderr)) {
            console.log('[yt-dlp] DRM error detected, trying alt clients...');
        } else {
            return result; // Erro diferente — não adianta retry
        }
    }
    return lastResult;
}

ipcMain.handle('yt-ensure-deps', async () => {
    const yt = await ensureYtDlp();
    const ff = ffmpegReady() || await ensureFfmpeg();
    return { ytdlp: yt, ffmpeg: ff };
});

ipcMain.handle('yt-get-info', async (event, url) => {
    url = cleanYtUrl(url);
    if (!ytDlpReady()) { const ok = await ensureYtDlp(); if (!ok) return { error: 'yt-dlp não disponível' }; }
    const args = ['--no-download', '--print-json', '--no-warnings', '--no-playlist'];
    if (ffmpegReady()) args.push('--ffmpeg-location', path.dirname(ffmpegBin));
    args.push(url);
    const result = await runYtDlpRobust(args, { timeout: 45000 });
    if (result.code !== 0) return { error: result.stderr || 'Erro ao obter info' };
    try {
        const info = JSON.parse(result.stdout);
        const formats = (info.formats || [])
            .filter(f => f.vcodec !== 'none' || f.acodec !== 'none')
            .map(f => ({
                id: f.format_id,
                ext: f.ext,
                quality: f.format_note || f.resolution || '',
                fps: f.fps || 0,
                vcodec: f.vcodec,
                acodec: f.acodec,
                filesize: f.filesize || f.filesize_approx || 0,
                hasVideo: f.vcodec !== 'none',
                hasAudio: f.acodec !== 'none'
            }));
        return {
            title: info.title || '',
            thumbnail: info.thumbnail || '',
            duration: info.duration || 0,
            uploader: info.uploader || '',
            formats: formats
        };
    } catch (e) { return { error: 'Erro ao parsear info' }; }
});

ipcMain.handle('yt-download', async (event, { url, outputDir, format, startTime, endTime }) => {
    url = cleanYtUrl(url);
    if (!ytDlpReady()) return { error: 'yt-dlp não disponível' };
    if (ytDownloadProc) return { error: 'Download já em andamento' };

    const outputPath = outputDir || (isWin ? path.join(os.homedir(), 'Downloads') : path.join(os.homedir(), 'Downloads'));
    const args = [
        '-o', path.join(outputPath, '%(title)s.%(ext)s'),
        '--newline', '--no-warnings',
        '--no-mtime', '--no-playlist',
        '--windows-filenames'
    ];
    if (ffmpegReady()) args.push('--ffmpeg-location', path.dirname(ffmpegBin));

    // Trim info (will be applied AFTER download with ffmpeg -c copy)
    const wantTrim = ffmpegReady() && ((startTime && startTime !== '0' && startTime !== '00:00' && startTime !== '0:00') || (endTime && endTime !== 'inf'));
    const trimStart = startTime || '0';
    const trimEnd = endTime || '';

    const hasFfmpeg = ffmpegReady();
    if (format === 'mp3') {
        args.push('-x', '--audio-format', 'mp3', '--audio-quality', '0');
    } else if (hasFfmpeg) {
        // ffmpeg available — separate video+audio streams + merge to mp4.
        // IMPORTANTE: pra "best" NUNCA restringe codec. H.264 no YouTube costuma capar em 720p ou 1080p;
        // resoluções maiores (1440p/4K) só vêm em VP9/AV1. Filtro avc1 forçava cap em 720p/1080p.
        // Solução: filtro pega o melhor independente de codec; sort (-S) prefere H.264 SE existir
        // na mesma resolução, mas não exige.
        if (format === '4k') {
            args.push('-f', 'bestvideo[height<=2160]+bestaudio/best[height<=2160]/bestvideo+bestaudio/best');
        } else if (format === '1080') {
            args.push('-f', 'bestvideo[height<=1080]+bestaudio/best[height<=1080]/bestvideo+bestaudio/best');
        } else if (format === '720') {
            args.push('-f', 'bestvideo[height<=720]+bestaudio/best[height<=720]/bestvideo+bestaudio/best');
        } else {
            // best = SEM cap de resolução, SEM cap de codec — yt-dlp pega 4K se existir
            args.push('-f', 'bestvideo+bestaudio/best');
        }
        // Sort: prefere H.264/m4a quando MESMA resolução disponível, mas não filtra
        args.push('-S', 'res,codec:h264:m4a,br', '--merge-output-format', 'mp4');
    } else {
        // No ffmpeg — single stream (no merge needed). Note: 4K geralmente NÃO existe como single file.
        if (format === '4k') {
            args.push('-f', 'best[height<=2160][ext=mp4]/best[height<=2160]/best[ext=mp4]/best');
        } else if (format === '1080') {
            args.push('-f', 'best[height<=1080][ext=mp4]/best[height<=1080]/best');
        } else if (format === '720') {
            args.push('-f', 'best[height<=720][ext=mp4]/best[height<=720]/best');
        } else {
            args.push('-f', 'best[ext=mp4]/best');
        }
    }
    args.push(url);

    ytProgress = { active: true, percent: 0, speed: '', eta: '', title: '', status: 'downloading', error: '', outputDir: outputPath };

    let lastFile = '';

    // Helper pra parsear linha de stdout (compartilhado entre tentativas)
    const onStdoutLine = (line) => {
        const m = line.match(/\[download\]\s+([\d.]+)%\s+of\s+~?([\d.]+\S+)\s+at\s+([\d.]+\S+)\s+ETA\s+(\S+)/);
        if (m) {
            ytProgress.percent = parseFloat(m[1]);
            ytProgress.speed = m[3];
            ytProgress.eta = m[4];
        }
        const dm = line.match(/\[download\] Destination: (.+)/);
        if (dm) lastFile = dm[1].trim();
        const mm = line.match(/\[Merger\] Merging formats into "(.+)"/);
        if (mm) lastFile = mm[1].trim();
        const am = line.match(/\[ExtractAudio\] Destination: (.+)/);
        if (am) lastFile = am[1].trim();
        if (line.includes('[Merger]') || line.includes('[ExtractAudio]')) {
            ytProgress.status = 'merging';
            ytProgress.percent = 99;
        }
        if (line.includes('has already been downloaded')) {
            ytProgress.percent = 100;
            ytProgress.status = 'done';
        }
        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
    };

    // Roda com retry de bot bypass
    runYtDlpRobust(args, {
        onProc: (proc) => { ytDownloadProc = proc; },
        onStdoutLine,
        onStderrLine: (line) => { if (line.trim()) ytProgress.error = line.trim(); },
    }).then(result => {
        ytDownloadProc = null;
        const code = result.code;
        if (code === 0) {
            // If lastFile is empty or doesn't exist, find newest file in output dir
            if (!lastFile || !fs.existsSync(lastFile)) {
                try {
                    const files = fs.readdirSync(outputPath)
                        .map(f => ({ name: f, full: path.join(outputPath, f), stat: fs.statSync(path.join(outputPath, f)) }))
                        .filter(f => f.stat.isFile() && /\.(mp4|mkv|webm|mp3|m4a)$/i.test(f.name))
                        .sort((a, b) => b.stat.mtimeMs - a.stat.mtimeMs);
                    if (files.length > 0) lastFile = files[0].full;
                } catch (e) {}
            }

            // Trim with ffmpeg -c copy (instant, no re-encode)
            if (wantTrim && lastFile && fs.existsSync(lastFile)) {
                ytProgress.status = 'merging';
                ytProgress.percent = 99;
                if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);

                const ext = path.extname(lastFile);
                const trimmed = lastFile.replace(ext, '_cut' + ext);
                const ffArgs = ['-y', '-i', lastFile];
                if (trimStart && trimStart !== '0') ffArgs.push('-ss', trimStart);
                if (trimEnd) ffArgs.push('-to', trimEnd);
                ffArgs.push('-c', 'copy', '-avoid_negative_ts', 'make_zero', trimmed);

                const ffProc = spawn(ffmpegBin, ffArgs, { windowsHide: true });
                ffProc.on('close', fc => {
                    if (fc === 0 && fs.existsSync(trimmed)) {
                        // Replace original with trimmed
                        try { fs.unlinkSync(lastFile); } catch (e) {}
                        try { fs.renameSync(trimmed, lastFile); } catch (e) { lastFile = trimmed; }
                    }
                    ytProgress.percent = 100;
                    ytProgress.status = 'done';
                    ytProgress.active = false;
                    ytProgress.file = lastFile;
                    if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                });
                ffProc.on('error', () => {
                    // Trim failed but download OK — return original
                    ytProgress.percent = 100;
                    ytProgress.status = 'done';
                    ytProgress.active = false;
                    ytProgress.file = lastFile;
                    if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                });
                return; // Don't emit done yet, wait for ffmpeg
            }

            ytProgress.percent = 100;
            ytProgress.status = 'done';
            ytProgress.active = false;
            ytProgress.file = lastFile;
        } else {
            ytProgress.status = 'error';
            ytProgress.active = false;
            if (!ytProgress.error || ytIsBotError(result.stderr)) {
                ytProgress.error = ytIsBotError(result.stderr)
                    ? 'YouTube bloqueou o download — verifique se está logado no Chrome/Edge/Firefox e tente de novo'
                    : (result.stderr || 'Download falhou (código ' + code + ')');
            }
        }
        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
    }).catch(e => {
        ytDownloadProc = null;
        ytProgress.status = 'error';
        ytProgress.active = false;
        ytProgress.error = e.message || String(e);
        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
    });

    return { ok: true };
});

ipcMain.handle('yt-cancel', () => {
    if (ytDownloadProc) { ytDownloadProc.kill(); ytDownloadProc = null; }
    ytProgress = { active: false, percent: 0, speed: '', eta: '', title: '', status: 'cancelled', error: '' };
    return true;
});

ipcMain.handle('yt-progress', () => ytProgress);

ipcMain.handle('yt-open-folder', (event, folderPath) => {
    const dir = folderPath || (isWin ? path.join(os.homedir(), 'Downloads') : path.join(os.homedir(), 'Downloads'));
    if (isWin) exec(`explorer.exe "${dir}"`, { windowsHide: true });
    else exec(`open "${dir}"`);
    return true;
});

/* ═══════ HTTP Sync Server (Premiere Pro plugin) ═══════ */
const SYNC_PORT = 9847;
let cachedState = { timer: { running: false }, pomodoro: { running: false }, projects: [] };

// Plugin auth tokens (persisted to disk só they survive app restarts)
const pluginSessions = new Map(); // token -> { userId, email, plan, expiresAt }
const sessionsFile = path.join(currentUD, 'plugin-sessions.json');

function loadPluginSessions() {
    try {
        const data = JSON.parse(fs.readFileSync(sessionsFile, 'utf8'));
        const now = Date.now();
        for (const [token, session] of Object.entries(data)) {
            if (session.expiresAt > now) pluginSessions.set(token, session);
        }
        if (pluginSessions.size > 0) console.log('Restored ' + pluginSessions.size + ' plugin session(s)');
    } catch(e) { /* no sessions file yet */ }
}
function savePluginSessions() {
    try {
        const obj = {};
        for (const [token, session] of pluginSessions) obj[token] = session;
        fs.writeFileSync(sessionsFile, JSON.stringify(obj), 'utf8');
    } catch(e) {}
}
loadPluginSessions(); // Restore sessions on startup

// ═══════ LION SEARCH state ═══════
const pendingPluginCommands = [];
// Long-polling: requests aguardando comando chegar.
// Em vez do plugin fazer poll a cada 50ms, ele mantém uma conexão aberta
// que o servidor responde NA HORA que um comando entrar no queue.
// Latência de comando: ~0ms (versus 25ms avg do poll de 50ms).
const lionLongPollWaiters = [];
function _notifyLongPollWaiters() {
    // Se tem comando pendente E tem waiter, responde imediatamente
    while (pendingPluginCommands.length > 0 && lionLongPollWaiters.length > 0) {
        const waiter = lionLongPollWaiters.shift();
        try {
            const cmds = pendingPluginCommands.splice(0, pendingPluginCommands.length);
            clearTimeout(waiter.timeoutHandle);
            waiter.res.writeHead(200, { 'Content-Type': 'application/json' });
            waiter.res.end(JSON.stringify({ commands: cmds }));
        } catch(e) {}
    }
}
let lionSearchEffectsCache = [];
let lionSearchEffectsCachedAt = 0;
let lionSearchCatalogDebug = '';
let lionSearchCatalogError = '';
const lionSearchPendingResults = new Map();
let lionSearchWin = null;

// Persistência do catálogo em disco — disponível imediato no boot do app
// (sem precisar esperar plugin reenviar quando Premiere abre)
const LION_CATALOG_FILE = path.join(currentUD, 'lion-search-catalog.json');

function saveLionCatalogToDisk() {
    try {
        const data = {
            items: lionSearchEffectsCache,
            cachedAt: lionSearchEffectsCachedAt,
            debug: lionSearchCatalogDebug,
            error: lionSearchCatalogError,
            version: 1,
        };
        fs.writeFileSync(LION_CATALOG_FILE, JSON.stringify(data), 'utf8');
    } catch(e) { /* ignora — não-crítico */ }
}

function loadLionCatalogFromDisk() {
    try {
        const raw = fs.readFileSync(LION_CATALOG_FILE, 'utf8');
        const data = JSON.parse(raw);
        if (Array.isArray(data.items)) {
            lionSearchEffectsCache = data.items;
            lionSearchEffectsCachedAt = data.cachedAt || Date.now();
            lionSearchCatalogDebug = data.debug || '';
            lionSearchCatalogError = data.error || '';
            console.log('[lion-search] catálogo carregado do disco:', data.items.length, 'itens');
        }
    } catch(e) { /* primeira vez — sem cache ainda */ }
}
// Carrega imediatamente na boot do main.js
loadLionCatalogFromDisk();

// Enfilera um comando pro plugin executar via JSX, retorna Promise com resultado
function queuePluginCommand(type, payload, timeoutMs = 15000) {
    const commandId = crypto.randomBytes(6).toString('hex');
    return new Promise((resolve, reject) => {
        const timeout = setTimeout(() => {
            lionSearchPendingResults.delete(commandId);
            reject(new Error('Timeout aguardando plugin (' + (timeoutMs/1000) + 's) — plugin tá rodando no Premiere?'));
        }, timeoutMs);
        lionSearchPendingResults.set(commandId, { resolve, reject, timeout });
        pendingPluginCommands.push({ id: commandId, type, payload: payload || {} });
        // Acorda qualquer long-poll esperando — comando vai chegar no plugin em ~0ms
        _notifyLongPollWaiters();
    });
}

// Dedup: se um list-effects está em andamento, próximas calls await a mesma promise
let _lionListEffectsInFlight = null;

// IPC pra lion-search.html chamar: lista efeitos
ipcMain.handle('lion-search:list-effects', async (event, opts) => {
    const opts2 = opts || {};
    // Cache do disco SEMPRE serve resposta imediata (mesmo se forceRefresh):
    // - Não-force: retorna cache, ZERO espera
    // - Force: retorna cache + dispara refresh em background pro setInterval pegar
    if (!opts2.forceRefresh && lionSearchEffectsCache.length > 0) {
        return { items: lionSearchEffectsCache, fromCache: true, debug: lionSearchCatalogDebug, jsxError: lionSearchCatalogError };
    }
    // Se forceRefresh MAS tem cache, retorna cache imediatamente E dispara refresh assíncrono
    if (opts2.forceRefresh && lionSearchEffectsCache.length > 0) {
        // Background refresh — não bloqueia resposta
        setImmediate(() => {
            if (_lionListEffectsInFlight) return; // já tem outro em andamento
            _lionListEffectsInFlight = queuePluginCommand('list-effects', {}, 30000);
            _lionListEffectsInFlight.then(() => { _lionListEffectsInFlight = null; }).catch(() => { _lionListEffectsInFlight = null; });
        });
        return { items: lionSearchEffectsCache, fromCache: true, refreshing: true, debug: lionSearchCatalogDebug, jsxError: lionSearchCatalogError };
    }
    // Se já tem um scan em andamento, await ele em vez de disparar outro
    if (_lionListEffectsInFlight) {
        try {
            const result = await _lionListEffectsInFlight;
            return { items: result.items || [], fromCache: false, debug: result.debug || '', jsxError: result.error || '' };
        } catch(e) {}
    }
    try {
        // Timeout adaptativo: forceRefresh = 30s (scan completo pode ser lento),
        // não-force = 3s (feedback rápido se plugin não tá rodando).
        const timeout = opts2.forceRefresh ? 30000 : 3000;
        _lionListEffectsInFlight = queuePluginCommand('list-effects', {}, timeout);
        const result = await _lionListEffectsInFlight;
        _lionListEffectsInFlight = null;
        return {
            items: result.items || [],
            fromCache: false,
            debug: result.debug || '',
            jsxError: result.error || '',
        };
    } catch (e) {
        _lionListEffectsInFlight = null;
        if (lionSearchEffectsCache.length > 0) {
            return { items: lionSearchEffectsCache, fromCache: true, staleError: e.message };
        }
        return { error: 'Abra o Premiere com o plugin Lion Workspace (Window → Extensions) e faça login.' };
    }
});

// IPC pra lion-search.html chamar: aplica efeito
ipcMain.handle('lion-search:apply', async (event, payload) => {
    _isApplyInProgress = true;
    try {
        const result = await queuePluginCommand('apply-effect', payload, 15000);
        return result;
    } catch (e) { return { error: e.message }; }
    finally {
        // Mantém flag um pouco após apply pra cobrir notificação
        setTimeout(() => { _isApplyInProgress = false; }, 2000);
    }
});

// IPC pra "fechar" janela — usa hide() pra preservar janela pré-criada (próximo open é instantâneo)
ipcMain.on('lion-search:close', () => {
    try {
        if (lionSearchWin && !lionSearchWin.isDestroyed()) {
            lionSearchWin.hide();
            // Reset state pra próximo open (limpa input + resultados)
            try { lionSearchWin.webContents.send('lion-search:reset'); } catch(e) {}
        }
    } catch(e) {}
});

// Flag pra evitar que notificação/IPC apply traga mainWin pra frente no Mac
let _isApplyInProgress = false;
let _isNotifying = false;

// IPC: mostra notificação nativa SÓ pra erros (sucesso é silencioso —
// o plugin/UI já dá feedback visual de "inserido", não precisa toast Windows).
ipcMain.on('lion-search:notify', (event, payload) => {
    try {
        if (payload?.success === true) return; // skip sucesso
        _isNotifying = true;
        setTimeout(() => { _isNotifying = false; }, 1500);
        const { Notification } = require('electron');
        if (Notification.isSupported()) {
            const n = new Notification({
                title: payload?.title || 'LION SEARCH',
                body: payload?.body || '',
                urgency: 'normal',
            });
            n.show();
        }
    } catch(e) { console.warn('[lion-search] notif fail:', e); }
});
// Backward compat com old name
ipcMain.on('lion-search:show-error', (event, payload) => {
    try {
        const { Notification } = require('electron');
        if (Notification.isSupported()) {
            const n = new Notification({
                title: payload?.title || 'LION SEARCH',
                body: payload?.body || 'Erro ao aplicar',
            });
            n.show();
        }
    } catch(e) {}
});

// IPC: get current selection context (video/audio/none, playhead)
ipcMain.handle('lion-search:get-context', async () => {
    try {
        const result = await queuePluginCommand('get-context', {}, 4000);
        return result.context || { selectionType: 'none', hasSequence: false };
    } catch (e) {
        return { error: e.message, selectionType: 'none', hasSequence: false };
    }
});

// Periodic cleanup: prune expired plugin sessions every hour (prevents unbounded Map growth)
setInterval(() => {
    try {
        const now = Date.now();
        let pruned = 0;
        for (const [token, session] of pluginSessions) {
            if (session.expiresAt <= now) { pluginSessions.delete(token); pruned++; }
        }
        if (pruned > 0) { savePluginSessions(); console.log('[cleanup] Pruned ' + pruned + ' expired plugin session(s)'); }
    } catch(e) {}
}, 60 * 60 * 1000);

function generatePluginToken() {
    return crypto.randomBytes(32).toString('hex');
}

function authenticatePluginRequest(req) {
    const authHeader = req.headers['authorization'];
    if (!authHeader || !authHeader.startsWith('Bearer ')) return null;
    const token = authHeader.slice(7);
    const session = pluginSessions.get(token);
    if (!session) return null;
    if (Date.now() > session.expiresAt) {
        pluginSessions.delete(token);
        savePluginSessions();
        return null;
    }
    return session;
}

ipcMain.handle('sync-push-state', (event, state) => {
    cachedState = state;
    return true;
});

// ═══════════════════════════════════════════════════════════════════
// BiRefNet (Node worker) — usado pelos endpoints /bg/remove (img).
// Inferência roda em worker_thread separado pra NÃO bloquear o main process
// (plugin Premiere marcaria "Desconectado" se main bloqueasse >3s pings).
// ═══════════════════════════════════════════════════════════════════
const { Worker } = require('worker_threads');
let _bgWorker = null;
let _ytPreviewWin = null;
let _activeMtEditorWin = null; // singleton: so um motion tracker editor aberto por vez
const _bgPending = new Map();
let _bgNextReqId = 0;

function _ensureBgWorker() {
    if (_bgWorker) return _bgWorker;
    const workerPath = app.isPackaged
        ? path.join(process.resourcesPath, 'app.asar.unpacked', 'bg-img-worker.js')
        : path.join(__dirname, 'bg-img-worker.js');
    _bgWorker = new Worker(workerPath);
    _bgWorker.on('message', (msg) => {
        if (msg && msg.ready) { console.log('[bg-worker] started'); return; }
        if (msg && msg.reqId != null) {
            const cb = _bgPending.get(msg.reqId);
            if (cb) {
                _bgPending.delete(msg.reqId);
                cb(msg);
            }
        }
    });
    _bgWorker.on('error', (e) => console.error('[bg-worker] err:', e));
    _bgWorker.on('exit', (code) => {
        console.warn('[bg-worker] exit code:', code);
        // Rejeita pendentes
        for (const [id, cb] of _bgPending) cb({ ok: false, error: 'worker exited (' + code + ')' });
        _bgPending.clear();
        _bgWorker = null;
    });
    return _bgWorker;
}

// Processa imagem via BiRefNet no worker. Retorna Buffer PNG com alpha.
async function _removeBgBiRefNet(imagePath) {
    _ensureBgWorker();
    const reqId = ++_bgNextReqId;
    const outputPath = path.join(currentUD, 'bgremoved-' + Date.now() + '-' + reqId + '.png');
    const result = await new Promise((resolve) => {
        _bgPending.set(reqId, resolve);
        _bgWorker.postMessage({ reqId, imagePath, outputPath });
    });
    if (!result.ok) throw new Error(result.error || 'worker failed');
    try {
        const buf = fs.readFileSync(result.outputPath);
        try { fs.unlinkSync(result.outputPath); } catch(e) {}
        return buf;
    } catch (e) {
        throw new Error('worker output missing: ' + e.message);
    }
}

// ═══════════════════════════════════════════════════════════════════
// MOTION TRACKER — spawn Python com motion_tracker.py.
// Migracao da versao OpenCV.js (WASM, lenta) pra cv2 nativo (10-100x).
// Suporta CSRT/KCF/MOSSE/MIL/LucasKanade + Whittaker/Kalman/Supersmoother.
// ═══════════════════════════════════════════════════════════════════
let _cachedPython = null;
function _findPython() {
    if (_cachedPython) return _cachedPython;
    const { execSync } = require('child_process');
    const candidates = isWin
        ? ['py -3', 'python', 'python3', 'C:\\Python313\\python.exe', 'C:\\Python312\\python.exe', 'C:\\Python311\\python.exe', 'C:\\Python310\\python.exe']
        : ['python3', '/usr/local/bin/python3', '/opt/homebrew/bin/python3', '/usr/bin/python3', 'python'];
    for (const cmd of candidates) {
        try {
            const out = execSync(`${cmd} -c "import cv2,numpy,sys;print(sys.executable)"`, { timeout: 4000, stdio: ['ignore', 'pipe', 'ignore'] }).toString().trim();
            if (out) { _cachedPython = cmd; console.log('[motion-tracker] using python:', cmd, '=>', out); return cmd; }
        } catch (e) { /* tenta proximo */ }
    }
    return null;
}

// Spawn motion_tracker.py com pontos + opcoes. Retorna o JSON result.
async function _runMotionTracker(filePath, ffmpegPath, ffprobePath, points, options) {
    const py = _findPython();
    if (!py) {
        throw new Error('Python+OpenCV nao instalados. Abra Configuracoes > Motion Tracker pra instalar.');
    }
    const scriptPath = app.isPackaged
        ? path.join(process.resourcesPath, 'app.asar.unpacked', 'motion_tracker.py')
        : path.join(__dirname, 'motion_tracker.py');

    const opts = options || {};
    const args = [
        scriptPath,
        filePath,
        '--points', JSON.stringify(points),
        '--algorithm', opts.algorithm || 'CSRT',
        '--smoothing', opts.smoothing || 'whittaker',
        '--bbox-width', String(opts.bboxWidth || 40),
        '--bbox-height', String(opts.bboxHeight || 40),
        '--confidence-threshold', String(opts.confidenceThreshold || 0.3),
        '--max-lost-frames', String(opts.maxLostFrames || 5),
        '--whittaker-lambda', String(opts.whittakerLambda || 100.0),
    ];

    const { spawn } = require('child_process');
    // No Windows "py -3" precisa de shell. Em Mac/Linux usa array direto.
    const useShell = isWin && py.startsWith('py ');
    const proc = useShell
        ? spawn(py + ' ' + args.map(a => `"${String(a).replace(/"/g, '\\"')}"`).join(' '), { shell: true, windowsHide: true })
        : spawn(py, args, { windowsHide: true });

    return new Promise((resolve, reject) => {
        let stdout = '';
        let stderr = '';
        proc.stdout.on('data', d => { stdout += d.toString(); });
        proc.stderr.on('data', d => { stderr += d.toString(); process.stderr.write('[mt-py] ' + d); });
        proc.on('error', reject);
        proc.on('close', code => {
            if (code !== 0) {
                return reject(new Error(`Python exit ${code}: ${stderr.slice(0, 500) || stdout.slice(0, 500)}`));
            }
            try {
                resolve(JSON.parse(stdout));
            } catch (e) {
                reject(new Error('Falha parseando JSON do Python: ' + e.message + '\nSTDOUT: ' + stdout.slice(0, 500)));
            }
        });
    });
}

function startSyncServer() {
    const server = http.createServer((req, res) => {
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
        res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

        if (req.method === 'OPTIONS') {
            res.writeHead(200);
            res.end();
            return;
        }

        // === Auth endpoints (no token required) ===
        if (req.method === 'POST' && req.url === '/auth/login') {
            let body = '';
            req.on('data', chunk => body += chunk);
            req.on('end', async () => {
                try {
                    const { email, password } = JSON.parse(body);
                    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
                    if (error) {
                        res.writeHead(401, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: error.message }));
                        return;
                    }
                    // Check subscription
                    const { data: sub } = await supabase
                        .from('subscriptions')
                        .select('plan, status')
                        .eq('user_id', data.user.id)
                        .in('status', ['active', 'trialing'])
                        .limit(1)
                        .single();
                    if (!sub) {
                        res.writeHead(403, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'no_subscription', message: 'Assinatura necessária para usar o plugin.' }));
                        return;
                    }
                    // Check device
                    const fp = getDeviceFingerprint();
                    const { data: device } = await supabase
                        .from('devices')
                        .select('blocked')
                        .eq('user_id', data.user.id)
                        .eq('fingerprint', fp)
                        .single();
                    if (device?.blocked) {
                        res.writeHead(403, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'device_blocked' }));
                        return;
                    }
                    // Create plugin session token (valid 24h)
                    const token = generatePluginToken();
                    pluginSessions.set(token, {
                        userId: data.user.id,
                        email: data.user.email,
                        name: data.user.user_metadata?.full_name || '',
                        plan: sub.plan,
                        expiresAt: Date.now() + (24 * 60 * 60 * 1000)
                    });
                    savePluginSessions(); // persist to disk
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({
                        token,
                        user: { email: data.user.email, name: data.user.user_metadata?.full_name || '' },
                        plan: sub.plan,
                        app_version: app.getVersion()
                    }));
                } catch (e) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: e.message }));
                }
            });
            return;
        }

        if (req.method === 'GET' && req.url === '/auth/check') {
            const session = authenticatePluginRequest(req);
            if (!session) {
                res.writeHead(401, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ authenticated: false, app_version: app.getVersion() }));
            } else {
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ authenticated: true, user: { email: session.email, name: session.name }, plan: session.plan, app_version: app.getVersion() }));
            }
            return;
        }

        // Auto-login: plugin pega um token usando a sessão do app principal (sem senha)
        // Só funciona pra conexões localhost (127.0.0.1), o que já é garantido porque
        // o server está ouvindo apenas em 127.0.0.1.
        if (req.method === 'POST' && req.url === '/auth/local-session') {
            (async () => {
                try {
                    // First try the in-memory snapshot (set on login or app startup)
                    let user = currentAppUser;
                    // Fall back to Supabase session if snapshot is null
                    if (!user) {
                        try {
                            const { data: { session } } = await supabase.auth.getSession();
                            if (session && session.user) {
                                user = {
                                    id: session.user.id,
                                    email: session.user.email,
                                    name: session.user.user_metadata?.full_name || ''
                                };
                                currentAppUser = user; // cache for next call
                            }
                        } catch(e) {}
                    }
                    if (!user) {
                        console.log('[auto-login] No app session — plugin will show manual login');
                        res.writeHead(401, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'app_not_logged_in', message: 'Faça login no Lion Workspace primeiro.' }));
                        return;
                    }
                    console.log('[auto-login] Found app user: ' + user.email);
                    // Check subscription
                    const { data: sub } = await supabase
                        .from('subscriptions')
                        .select('plan, status')
                        .eq('user_id', user.id)
                        .in('status', ['active', 'trialing'])
                        .limit(1)
                        .single();
                    if (!sub) {
                        console.log('[auto-login] No active subscription for ' + user.email);
                        res.writeHead(403, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'no_subscription', message: 'Assinatura necessária para usar o plugin.' }));
                        return;
                    }
                    // Check device
                    const fp = getDeviceFingerprint();
                    const { data: device } = await supabase
                        .from('devices')
                        .select('blocked')
                        .eq('user_id', user.id)
                        .eq('fingerprint', fp)
                        .single();
                    if (device?.blocked) {
                        res.writeHead(403, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'device_blocked' }));
                        return;
                    }
                    // Generate plugin token
                    const token = generatePluginToken();
                    pluginSessions.set(token, {
                        userId: user.id,
                        email: user.email,
                        name: user.name || '',
                        plan: sub.plan,
                        expiresAt: Date.now() + (24 * 60 * 60 * 1000)
                    });
                    savePluginSessions();
                    console.log('[auto-login] Issued plugin token for ' + user.email);
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({
                        token,
                        user: { email: user.email, name: user.name || '' },
                        plan: sub.plan,
                        app_version: app.getVersion()
                    }));
                } catch (e) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: e.message }));
                }
            })();
            return;
        }

        if (req.method === 'POST' && req.url === '/auth/logout') {
            const authHeader = req.headers['authorization'];
            if (authHeader && authHeader.startsWith('Bearer ')) {
                pluginSessions.delete(authHeader.slice(7));
                savePluginSessions();
            }
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ ok: true }));
            return;
        }

        // Health check (no auth required — lets plugin detect app is running)
        if (req.method === 'GET' && req.url === '/ping') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ ok: true, version: app.getVersion() }));
            return;
        }

        // === YouTube preview HTML (no auth — só serve embed do YT pra origin http://) ===
        // Necessário porque CEP plugin tem origin file:// que YouTube bloqueia (Error 153).
        // Servindo daqui, a janela tem origin http://127.0.0.1 — YT aceita.
        if (req.method === 'GET' && req.url.indexOf('/yt-preview') === 0) {
            try {
                const u = new URL(req.url, 'http://127.0.0.1:9847');
                const id = (u.searchParams.get('id') || '').replace(/[^a-zA-Z0-9_-]/g, '');
                const start = Math.max(0, parseInt(u.searchParams.get('start') || '0', 10));
                const end = Math.max(0, parseInt(u.searchParams.get('end') || '0', 10));
                if (!id) {
                    res.writeHead(400, { 'Content-Type': 'text/plain' });
                    res.end('missing id');
                    return;
                }
                // HTML com identidade Lion Workspace + IFrame API que força loop no trim
                const html = `<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8">
<title>Lion Workspace · YouTube Preview</title>
<style>
:root{--bg:#0a0a0a;--bg2:#111;--accent:#bef264;--accent2:#d9f99d;--text:#f5f5f5;--text2:#a3a3a3;--text3:#636363;--border:rgba(255,255,255,.08)}
*{margin:0;padding:0;box-sizing:border-box}
html,body{height:100%;background:var(--bg);color:var(--text);font-family:-apple-system,system-ui,sans-serif;overflow:hidden;font-size:13px}
body{display:flex;flex-direction:column}
.hdr{display:flex;align-items:center;gap:9px;padding:7px 11px;background:linear-gradient(180deg,rgba(190,242,100,.06),transparent);border-bottom:1px solid rgba(190,242,100,.08);flex-shrink:0}
.logo{width:22px;height:22px;border-radius:6px;background:linear-gradient(135deg,var(--accent2),var(--accent));color:#0a0a0a;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;flex-shrink:0;box-shadow:0 2px 6px rgba(190,242,100,.25)}
.t{font-size:.74rem;font-weight:600;color:var(--text);flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.t-sub{font-size:.58rem;color:var(--text3);font-weight:500;margin-top:1px;font-family:monospace}
.t-wrap{flex:1;min-width:0}
.badge{font-family:'JetBrains Mono','SF Mono',monospace;font-size:.62rem;color:var(--accent);background:rgba(190,242,100,.1);border:1px solid rgba(190,242,100,.18);padding:3px 8px;border-radius:5px;font-weight:600;flex-shrink:0;font-variant-numeric:tabular-nums}
.player-wrap{flex:1;position:relative;background:#000;display:flex;align-items:center;justify-content:center}
#player{width:100%;height:100%}
.loading{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:var(--text3);font-size:.7rem;display:flex;align-items:center;gap:6px}
.spin{width:14px;height:14px;border:2px solid rgba(190,242,100,.2);border-top-color:var(--accent);border-radius:50%;animation:sp .7s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.ftr{display:flex;align-items:center;justify-content:space-between;padding:6px 11px;background:var(--bg2);border-top:1px solid var(--border);font-size:.62rem;flex-shrink:0}
.ts{font-family:'JetBrains Mono',monospace;color:var(--text2);font-variant-numeric:tabular-nums}
.ts strong{color:var(--accent)}
.loop{display:flex;align-items:center;gap:5px;cursor:pointer;color:var(--text3);padding:3px 7px;border-radius:5px;transition:all .12s;user-select:none}
.loop:hover{background:rgba(190,242,100,.04);color:var(--text2)}
.loop.on{color:var(--accent);background:rgba(190,242,100,.1)}
.loop svg{width:11px;height:11px}
</style>
</head><body>
<div class="hdr">
    <div class="logo">L</div>
    <div class="t-wrap">
        <div class="t">YouTube Preview</div>
        <div class="t-sub">Lion Workspace</div>
    </div>
    <div class="badge" id="trimBadge">--</div>
</div>
<div class="player-wrap">
    <div id="player"></div>
    <div class="loading" id="loading"><span class="spin"></span>Carregando player…</div>
</div>
<div class="ftr">
    <div class="ts"><strong id="curT">0:00</strong> / <span id="totT">0:00</span></div>
    <div class="loop on" id="loopBtn" onclick="toggleLoop()">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><polyline points="17 1 21 5 17 9"/><path d="M3 11V9a4 4 0 0 1 4-4h14"/><polyline points="7 23 3 19 7 15"/><path d="M21 13v2a4 4 0 0 1-4 4H3"/></svg>
        <span>Loop trim</span>
    </div>
</div>
<script>
var VIDEO_ID=${JSON.stringify(id)};
var START=${start};
var END=${end};
var player=null;
var loopEnabled=true;
function fmt(s){s=Math.max(0,Math.floor(s));var m=Math.floor(s/60);var sec=s%60;return m+':'+(sec<10?'0':'')+sec}
document.getElementById('trimBadge').textContent=(END>START)?fmt(START)+' → '+fmt(END):'sem trim';
var tag=document.createElement('script');tag.src='https://www.youtube.com/iframe_api';document.head.appendChild(tag);
window.onYouTubeIframeAPIReady=function(){
    player=new YT.Player('player',{
        videoId:VIDEO_ID,
        playerVars:{autoplay:1,controls:1,rel:0,modestbranding:1,iv_load_policy:3,fs:1,playsinline:1,start:START},
        events:{
            onReady:function(e){
                document.getElementById('loading').style.display='none';
                e.target.playVideo();
                setInterval(tick,250);
            },
            onStateChange:function(e){
                // 0=ended → loop pro start do trim
                if(e.data===0 && loopEnabled && END>START){
                    player.seekTo(START,true);player.playVideo();
                }
            }
        }
    });
};
function tick(){
    if(!player||!player.getCurrentTime)return;
    var cur=player.getCurrentTime()||0;
    var dur=player.getDuration()||0;
    document.getElementById('curT').textContent=fmt(cur);
    document.getElementById('totT').textContent=fmt(dur);
    // Trim end: param 'end' do YT é não confiável — força via JS
    if(END>START && cur>=END-0.1){
        if(loopEnabled){ player.seekTo(START,true); }
        else{ player.pauseVideo(); player.seekTo(END,true); }
    }
    // Skip antes do start (se user arrastar a barra pra trás do trim)
    if(END>START && cur<Math.max(0,START-1)){
        player.seekTo(START,true);
    }
}
function toggleLoop(){
    loopEnabled=!loopEnabled;
    document.getElementById('loopBtn').classList.toggle('on',loopEnabled);
}
</script>
</body></html>`;
                res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
                res.end(html);
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'text/plain' });
                res.end('err: ' + e.message);
            }
            return;
        }

        // === All other routes require authentication ===
        const pluginSession = authenticatePluginRequest(req);
        if (!pluginSession) {
            res.writeHead(401, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'unauthorized', message: 'Faça login no plugin.' }));
            return;
        }

        if (req.method === 'GET' && req.url === '/state') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(cachedState));
            return;
        }

        // ═══════ CLIPBOARD — leitura via Electron (cross-platform, handle PNG/WebP/DIB nativamente) ═══════
        // Plugin pode chamar isso ao invés de PowerShell/JXA pra evitar problemas de encoding/formato.
        if (req.method === 'GET' && req.url.indexOf('/clipboard/read-image') === 0) {
            try {
                const { clipboard } = require('electron');
                const u = new URL(req.url, 'http://localhost');
                let saveDir = u.searchParams.get('dir') || '';
                const tmpFallback = require('os').tmpdir();
                // Aceita só absolute path (Win drive letter ou Unix /); senão usa tmp
                const isAbs = saveDir && (path.isAbsolute(saveDir) || /^[a-zA-Z]:[\\/]/.test(saveDir));
                if (!isAbs) saveDir = tmpFallback;
                try { fs.mkdirSync(saveDir, { recursive: true }); } catch (e) { saveDir = tmpFallback; }
                // Final write-test — se dir não é writável, cai pro tmp
                try {
                    const tf = path.join(saveDir, '.lw-wt-' + Date.now());
                    fs.writeFileSync(tf, ''); fs.unlinkSync(tf);
                } catch (e) { saveDir = tmpFallback; }

                // 1) Tenta imagem direto (cobre PNG, WebP, DIB — Electron handles tudo)
                const img = clipboard.readImage();
                if (img && !img.isEmpty()) {
                    const pngBuf = img.toPNG();
                    if (pngBuf && pngBuf.length > 0) {
                        const fname = `lw-clip-${Date.now()}-${Math.random().toString(36).slice(2,10)}.png`;
                        const outPath = path.join(saveDir, fname);
                        fs.writeFileSync(outPath, pngBuf);
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ type: 'image', paths: [outPath] }));
                        return;
                    }
                }

                // 2) Fallback: clipboard.readBuffer('FileNameW') on Win or NSPasteboard via clipboard.read
                // Electron já normaliza isso via readImage, mas tentamos texto como path
                const txt = clipboard.readText();
                if (txt && txt.length < 32768) {
                    const lines = txt.split(/[\r\n]+/).map(s => s.trim().replace(/^"|"$/g, '')).filter(Boolean);
                    const valid = lines.filter(p => { try { return fs.statSync(p).isFile(); } catch (e) { return false; } });
                    if (valid.length > 0) {
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ type: 'files', paths: valid }));
                        return;
                    }
                }

                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ type: 'empty', reason: 'no-image-no-paths' }));
            } catch (e) {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ type: 'error', message: String(e.message || e) }));
            }
            return;
        }

        // ═══════ LION SEARCH — settings (hotkey config + SFX folder) ═══════
        if (req.method === 'GET' && req.url === '/lion-search/settings') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                hotkey: _lionSettings.hotkey,
                enabled: _lionSettings.enabled,
                sfxFolder: _lionSettings.sfxFolder || '',
                useLibraryOnly: _lionSettings.useLibraryOnly !== false, // default ON pra nova UX
                registered: _registeredLionHotkey,
            }));
            return;
        }

        // ═══════ Audio Library — GET (plugin lê pra alimentar LION SEARCH) ═══════
        if (req.method === 'GET' && req.url === '/lion-search/audio-library') {
            const lib = loadAudioLibrary();
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ items: lib }));
            return;
        }

        // ═══════ LION SEARCH — pending commands queue ═══════
        // Plugin polls pra pegar comandos vindos do LW (ex: "list effects",
        // "apply effect X"). Plugin executa via JSX e POSTa de volta o resultado.
        if (req.method === 'GET' && req.url.indexOf('/lion-search/pop-pending') === 0) {
            // Long-polling: ?wait=1 segura a resposta até comando chegar (max 25s)
            // Fallback rápido se não tiver ?wait — comportamento antigo
            const u = new URL(req.url, 'http://localhost');
            const wait = u.searchParams.get('wait') === '1';
            if (!wait || pendingPluginCommands.length > 0) {
                const cmds = pendingPluginCommands.splice(0, pendingPluginCommands.length);
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ commands: cmds }));
                return;
            }
            // Long-poll: aguarda até 25s por comando — server PUSH-A quando _notifyLongPollWaiters dispara
            const timeoutHandle = setTimeout(() => {
                const idx = lionLongPollWaiters.findIndex(w => w.res === res);
                if (idx >= 0) lionLongPollWaiters.splice(idx, 1);
                try {
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ commands: [] }));
                } catch(e) {}
            }, 25000);
            lionLongPollWaiters.push({ res, timeoutHandle });
            // Cleanup se cliente desconectar antes
            req.on('close', () => {
                clearTimeout(timeoutHandle);
                const idx = lionLongPollWaiters.findIndex(w => w.res === res);
                if (idx >= 0) lionLongPollWaiters.splice(idx, 1);
            });
            return;
        }
        if (req.method === 'GET' && req.url === '/lion-search/effects-cache') {
            // Cliente da lion-search.html pega catálogo cacheado (atualizado pelo plugin)
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({
                items: lionSearchEffectsCache,
                fetchedAt: lionSearchEffectsCachedAt,
                age: lionSearchEffectsCachedAt ? (Date.now() - lionSearchEffectsCachedAt) : null,
            }));
            return;
        }

        // YouTube downloader HTTP endpoints (for plugin)
        if (req.method === 'GET' && req.url === '/yt/progress') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(ytProgress));
            return;
        }

        // BG Remove Vídeo — status + progress polling (plugin/UI usa)
        if (req.method === 'GET' && req.url.indexOf('/bg/remove-video/status') === 0) {
            const u = new URL(req.url, 'http://localhost');
            const token = u.searchParams.get('token');
            const includePreview = u.searchParams.get('preview') === '1';
            const sess = token ? bgVideoSessions.get(token) : null;
            res.writeHead(200, { 'Content-Type': 'application/json' });
            if (!sess) {
                res.end(JSON.stringify({ status: 'expired' }));
            } else {
                const out = {
                    status: sess.status,
                    progress: sess.progress,
                    outPath: sess.outPath,
                    error: sess.error || null,
                    delegate: sess.delegate,
                };
                if (includePreview && sess.lastPreview) {
                    out.preview = sess.lastPreview;
                    out.previewFrame = sess.lastPreviewFrame || 0;
                }
                res.end(JSON.stringify(out));
            }
            return;
        }

        // BG Remove Vídeo — detecta GPU disponível (cliente usa pra mostrar opções)
        if (req.method === 'GET' && req.url === '/bg/remove-video/gpu-info') {
            detectGpus().then((info) => {
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify(info));
            }).catch((e) => {
                res.writeHead(500, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: e.message }));
            });
            return;
        }

        // Mask editor status — plugin polla pra saber quando user salvou/cancelou
        if (req.method === 'GET' && req.url.indexOf('/bg/edit-status') === 0) {
            const u = new URL(req.url, 'http://localhost');
            const token = u.searchParams.get('token');
            const sess = token ? maskEditorSessions.get(token) : null;
            res.writeHead(200, { 'Content-Type': 'application/json' });
            if (!sess) {
                res.end(JSON.stringify({ status: 'expired' }));
            } else {
                res.end(JSON.stringify({ status: sess.status, processedPath: sess.processedPath }));
            }
            return;
        }

        if (req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk);
            req.on('end', async () => {
                let data = {};
                try { data = JSON.parse(body || '{}'); } catch (e) {}
                const command = req.url.replace(/^\//, '');

                // ═══════ LION SEARCH — plugin altera hotkey config ═══════
                if (command === 'lion-search/set-hotkey') {
                    const newHotkey = String(data.hotkey || '').trim();
                    const newEnabled = data.enabled !== false;
                    const newSfxFolder = (data.sfxFolder !== undefined) ? String(data.sfxFolder || '') : (_lionSettings.sfxFolder || '');
                    if (!newHotkey) {
                        res.writeHead(400, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'hotkey vazio' }));
                        return;
                    }
                    const oldHotkey2 = _lionSettings.hotkey;
                    _lionSettings = { hotkey: newHotkey, enabled: newEnabled, sfxFolder: newSfxFolder };
                    saveLionSearchSettings(_lionSettings);
                    let regResult = { ok: true };
                    if (oldHotkey2 !== newHotkey || !newEnabled) {
                        _doUnregisterHotkey();
                    }
                    if (newEnabled) {
                        startHotkeyFgWatcher();
                        try {
                            if (isWin) _checkAndRegisterHotkey(fs.readFileSync(fgFile, 'utf8').trim());
                            else if (isMac) getFgProc((fg) => _checkAndRegisterHotkey(fg));
                        } catch(e) {}
                    } else {
                        stopHotkeyFgWatcher();
                    }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({
                        ok: regResult.ok,
                        error: regResult.error,
                        hotkey: _lionSettings.hotkey,
                        enabled: _lionSettings.enabled,
                        registered: _registeredLionHotkey,
                    }));
                    return;
                }

                // ═══════ LION SEARCH — plugin submits results / catalog ═══════
                // Plugin envia o catálogo. Effects/presets/transitions REPLACE cache;
                // audio-sources ACUMULAM (porque Premiere 25.x só expõe o projeto ativo
                // — quando user troca de projeto, novo scan adiciona ao cache).
                if (command === 'lion-search/effects-update') {
                    if (Array.isArray(data.items) && data.items.length > 0) {
                        const newItems = data.items;
                        const newAudio = newItems.filter(x => x.kind === 'audio-source');
                        const newRest  = newItems.filter(x => x.kind !== 'audio-source');
                        // Library-only enforcement: drop qualquer audio-source que não seja LIB:
                        // (entries antigos de scan de projetos persistiam no cache).
                        // Mantém só os LIB: do cache antigo (eles vêm do app, não do scan de projeto).
                        const oldAudioLib = lionSearchEffectsCache.filter(x =>
                            x.kind === 'audio-source' && String(x.matchName || '').indexOf('LIB:') === 0
                        );
                        const audioMap = {};
                        for (const a of oldAudioLib) audioMap[a.matchName || a.name] = a;
                        for (const a of newAudio) {
                            // Só aceita áudio novo se for LIB: (paranoia — JSX já não envia outros)
                            if (String(a.matchName || '').indexOf('LIB:') === 0) {
                                audioMap[a.matchName || a.name] = a;
                            }
                        }
                        const mergedAudio = Object.values(audioMap);
                        lionSearchEffectsCache = [...newRest, ...mergedAudio];
                        lionSearchEffectsCachedAt = Date.now();
                        lionSearchCatalogDebug = String(data.debug || '');
                        lionSearchCatalogError = String(data.error || '');
                        console.log('[lion-search] catálogo: total=' + lionSearchEffectsCache.length + ' (effects=' + newRest.length + ', audio LIB-only=' + mergedAudio.length + ')');
                        // Persiste no disco — disponível na próxima boot do app
                        saveLionCatalogToDisk();
                    }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }
                // Limpa cache de audio (botão UI ou troca de profile)
                if (command === 'lion-search/clear-audio-cache') {
                    lionSearchEffectsCache = lionSearchEffectsCache.filter(x => x.kind !== 'audio-source');
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }
                // Plugin reporta resultado de uma execução pendente
                if (command === 'lion-search/command-result' && data.commandId) {
                    const pending = lionSearchPendingResults.get(data.commandId);
                    if (pending) {
                        clearTimeout(pending.timeout);
                        pending.resolve(data);
                        lionSearchPendingResults.delete(data.commandId);
                    }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }

                // YouTube downloader HTTP API
                if (command === 'yt/info' && data.url) {
                    try {
                        const cleanUrl = cleanYtUrl(data.url);
                        await ensureYtDlp();
                        const args = ['--no-download', '--print-json', '--no-warnings', '--no-playlist'];
                        if (ffmpegReady()) args.push('--ffmpeg-location', path.dirname(ffmpegBin));
                        args.push(cleanUrl);
                        const r = await runYtDlpRobust(args, { timeout: 45000 });
                        let out = { error: r.stderr || 'Erro' };
                        if (r.code === 0) {
                            try {
                                const info = JSON.parse(r.stdout);
                                out = { title: info.title || '', thumbnail: info.thumbnail || '', duration: info.duration || 0, uploader: info.uploader || '' };
                            } catch { out = { error: 'Parse error' }; }
                        } else if (ytIsBotError(r.stderr)) {
                            out = { error: 'YouTube bloqueou — faça login no Chrome/Edge no seu navegador e tente de novo' };
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(out));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message }));
                    }
                    return;
                }
                if (command === 'yt/download' && data.url) {
                    // Start download directly from main process
                    const cleanDlUrl = cleanYtUrl(data.url);
                    if (ytDownloadProc) {
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'Download já em andamento' }));
                        return;
                    }
                    // Reuse the IPC handler logic
                    try {
                        if (!ytDlpReady()) await ensureYtDlp();
                        if (!ffmpegReady()) await ensureFfmpeg();
                        const outputPath = (data.outputDir && fs.existsSync(data.outputDir)) ? data.outputDir : (isWin ? path.join(os.homedir(), 'Downloads') : path.join(os.homedir(), 'Downloads'));
                        const dlArgs = ['-o', path.join(outputPath, '%(title)s.%(ext)s'), '--newline', '--no-warnings', '--no-mtime', '--no-playlist', '--windows-filenames'];
                        if (ffmpegReady()) dlArgs.push('--ffmpeg-location', path.dirname(ffmpegBin));
                        // Trim info (will be applied AFTER download with ffmpeg -c copy)
                        const hasStart = data.startTime && data.startTime !== '0' && data.startTime !== '00:00' && data.startTime !== '0:00';
                        const hasEnd = data.endTime && data.endTime !== 'inf';
                        const wantTrim = ffmpegReady() && (hasStart || hasEnd);
                        const trimStart = data.startTime || '0';
                        const trimEnd = data.endTime || '';
                        const fmt = data.format || 'best';
                        const hasFf = ffmpegReady();
                        if (fmt === 'mp3') { dlArgs.push('-x', '--audio-format', 'mp3', '--audio-quality', '0'); }
                        else if (hasFf) {
                            // "best" e capped: NÃO filtra codec. H.264 cap em 720/1080p é cap real do YT.
                            // Sort (-S) ainda prefere H.264 quando disponível na mesma res.
                            if (fmt === '4k') dlArgs.push('-f', 'bestvideo[height<=2160]+bestaudio/best[height<=2160]/bestvideo+bestaudio/best');
                            else if (fmt === '1080') dlArgs.push('-f', 'bestvideo[height<=1080]+bestaudio/best[height<=1080]/bestvideo+bestaudio/best');
                            else if (fmt === '720') dlArgs.push('-f', 'bestvideo[height<=720]+bestaudio/best[height<=720]/bestvideo+bestaudio/best');
                            else dlArgs.push('-f', 'bestvideo+bestaudio/best');
                            dlArgs.push('-S', 'res,codec:h264:m4a,br', '--merge-output-format', 'mp4');
                        } else {
                            if (fmt === '4k') dlArgs.push('-f', 'best[height<=2160][ext=mp4]/best[height<=2160]/best[ext=mp4]/best');
                            else if (fmt === '1080') dlArgs.push('-f', 'best[height<=1080][ext=mp4]/best[height<=1080]/best');
                            else if (fmt === '720') dlArgs.push('-f', 'best[height<=720][ext=mp4]/best[height<=720]/best');
                            else dlArgs.push('-f', 'best[ext=mp4]/best');
                        }
                        dlArgs.push(cleanDlUrl);

                        ytProgress = { active: true, percent: 0, speed: '', eta: '', title: data.metaTitle||'', status: 'downloading', error: '', outputDir: outputPath, metaTitle: data.metaTitle||'', metaThumb: data.metaThumb||'', metaUploader: data.metaUploader||'', metaDuration: data.metaDuration||0, metaUrl: data.url||'' };
                        let lastFile = '';
                        const onStdoutLine2 = (line) => {
                            const m = line.match(/\[download\]\s+([\d.]+)%\s+of\s+~?([\d.]+\S+)\s+at\s+([\d.]+\S+)\s+ETA\s+(\S+)/);
                            if (m) { ytProgress.percent = parseFloat(m[1]); ytProgress.speed = m[3]; ytProgress.eta = m[4]; }
                            const dm = line.match(/\[download\] Destination: (.+)/);
                            if (dm) lastFile = dm[1].trim();
                            const mm = line.match(/\[Merger\] Merging formats into "(.+)"/);
                            if (mm) lastFile = mm[1].trim();
                            const am = line.match(/\[ExtractAudio\] Destination: (.+)/);
                            if (am) lastFile = am[1].trim();
                            if (line.includes('[Merger]') || line.includes('[ExtractAudio]')) { ytProgress.status = 'merging'; ytProgress.percent = 99; }
                            if (line.includes('has already been downloaded')) { ytProgress.percent = 100; ytProgress.status = 'done'; }
                            if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                        };

                        runYtDlpRobust(dlArgs, {
                            onProc: (proc) => { ytDownloadProc = proc; },
                            onStdoutLine: onStdoutLine2,
                            onStderrLine: (line) => { if (line.trim()) ytProgress.error = line.trim(); },
                        }).then(rdl => {
                            ytDownloadProc = null;
                            const code = rdl.code;
                            if (code === 0) {
                                if (!lastFile || !fs.existsSync(lastFile)) {
                                    try {
                                        const files = fs.readdirSync(outputPath)
                                            .map(f => ({ name: f, full: path.join(outputPath, f), stat: fs.statSync(path.join(outputPath, f)) }))
                                            .filter(f => f.stat.isFile() && /\.(mp4|mkv|webm|mp3|m4a)$/i.test(f.name))
                                            .sort((a, b) => b.stat.mtimeMs - a.stat.mtimeMs);
                                        if (files.length > 0) lastFile = files[0].full;
                                    } catch (e) {}
                                }

                                // Trim with ffmpeg -c copy (instant, no re-encode)
                                if (wantTrim && lastFile && fs.existsSync(lastFile)) {
                                    ytProgress.status = 'merging';
                                    ytProgress.percent = 99;
                                    if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);

                                    const ext = path.extname(lastFile);
                                    const trimmed = lastFile.replace(ext, '_cut' + ext);
                                    const ffArgs = ['-y', '-i', lastFile];
                                    if (trimStart && trimStart !== '0') ffArgs.push('-ss', trimStart);
                                    if (trimEnd) ffArgs.push('-to', trimEnd);
                                    ffArgs.push('-c', 'copy', '-avoid_negative_ts', 'make_zero', trimmed);

                                    const ffProc = spawn(ffmpegBin, ffArgs, { windowsHide: true });
                                    ffProc.on('close', fc => {
                                        if (fc === 0 && fs.existsSync(trimmed)) {
                                            try { fs.unlinkSync(lastFile); } catch (e) {}
                                            try { fs.renameSync(trimmed, lastFile); } catch (e) { lastFile = trimmed; }
                                        }
                                        ytProgress.percent = 100;
                                        ytProgress.status = 'done';
                                        ytProgress.active = false;
                                        ytProgress.file = lastFile;
                                        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                                    });
                                    ffProc.on('error', () => {
                                        ytProgress.percent = 100;
                                        ytProgress.status = 'done';
                                        ytProgress.active = false;
                                        ytProgress.file = lastFile;
                                        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                                    });
                                    return;
                                }

                                ytProgress.percent = 100; ytProgress.status = 'done'; ytProgress.active = false; ytProgress.file = lastFile;
                            } else {
                                ytProgress.status = 'error'; ytProgress.active = false;
                                if (!ytProgress.error || ytIsBotError(rdl.stderr)) {
                                    ytProgress.error = ytIsBotError(rdl.stderr)
                                        ? 'YouTube bloqueou — faça login no Chrome/Edge/Firefox e tente de novo'
                                        : (rdl.stderr || 'Falhou (código ' + code + ')');
                                }
                            }
                            if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                        }).catch(e => {
                            ytDownloadProc = null;
                            ytProgress.status = 'error'; ytProgress.active = false; ytProgress.error = e.message || String(e);
                            if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                        });
                    } catch (e) { ytProgress.status = 'error'; ytProgress.error = e.message; }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }
                // ═══════ Motion Tracker — Python + cv2 nativo ═══════
                if (command === 'motion-tracker/probe' && data.filePath) {
                    try {
                        // Probe via ffprobe — substitui o antigo motion-tracker.js
                        const args = ['-v', 'error', '-select_streams', 'v:0',
                            '-show_entries', 'stream=width,height,r_frame_rate,nb_frames,duration',
                            '-of', 'json', data.filePath];
                        const out = await new Promise((resolve, reject) => {
                            const proc = spawn(ffprobeBin, args, { windowsHide: true });
                            let stdout = '', stderr = '';
                            proc.stdout.on('data', d => { stdout += d.toString(); });
                            proc.stderr.on('data', d => { stderr += d.toString(); });
                            proc.on('error', reject);
                            proc.on('close', code => code === 0 ? resolve(stdout) : reject(new Error(`ffprobe exit ${code}: ${stderr}`)));
                        });
                        const j = JSON.parse(out);
                        const s = (j.streams && j.streams[0]) || {};
                        const fr = (s.r_frame_rate || '30/1').split('/');
                        const fps = (parseFloat(fr[0]) || 30) / (parseFloat(fr[1]) || 1);
                        const meta = {
                            width: parseInt(s.width) || 0,
                            height: parseInt(s.height) || 0,
                            fps: fps,
                            frameCount: parseInt(s.nb_frames) || Math.round((parseFloat(s.duration) || 0) * fps),
                            duration: parseFloat(s.duration) || 0,
                        };
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: true, ...meta }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                if (command === 'motion-tracker/extract-frame' && data.filePath && data.frame != null) {
                    try {
                        const fIdx = parseInt(data.frame, 10) || 0;
                        const args = ['-i', data.filePath, '-vf', `select=eq(n\\,${fIdx})`, '-vframes', '1', '-f', 'image2pipe', '-vcodec', 'png', '-'];
                        const chunks = [];
                        await new Promise((resolve, reject) => {
                            const ff = spawn(ffmpegBin, args, { stdio: ['ignore', 'pipe', 'pipe'] });
                            ff.stdout.on('data', d => chunks.push(d));
                            ff.on('close', code => code === 0 ? resolve() : reject(new Error('ffmpeg exit ' + code)));
                            ff.on('error', reject);
                        });
                        const pngBuf = Buffer.concat(chunks);
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: true, dataUrl: 'data:image/png;base64,' + pngBuf.toString('base64') }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                // ═══════ YouTube Preview Window (floating popup com embed do YT) ═══════
                // Plugin pede pra abrir/atualizar uma janela pequena always-on-top
                // que carrega o /yt-preview da própria HTTP server (origin http://).
                if (command === 'yt-preview/open' && data.id) {
                    try {
                        const ytId = String(data.id).replace(/[^a-zA-Z0-9_-]/g, '');
                        const startS = Math.max(0, parseInt(data.start || 0, 10));
                        const endS = Math.max(0, parseInt(data.end || 0, 10));
                        const url = 'http://127.0.0.1:9847/yt-preview?id=' + ytId + '&start=' + startS + (endS > startS ? '&end=' + endS : '');
                        if (_ytPreviewWin && !_ytPreviewWin.isDestroyed()) {
                            _ytPreviewWin.loadURL(url);
                            _ytPreviewWin.show();
                            _ytPreviewWin.focus();
                        } else {
                            const { BrowserWindow } = require('electron');
                            _ytPreviewWin = new BrowserWindow({
                                width: 480, height: 310,
                                title: 'YouTube Preview · Lion Workspace',
                                backgroundColor: '#000',
                                alwaysOnTop: true,
                                resizable: true,
                                minWidth: 320, minHeight: 220,
                                webPreferences: { contextIsolation: true, sandbox: true }
                            });
                            _ytPreviewWin.setMenuBarVisibility(false);
                            _ytPreviewWin.loadURL(url);
                            _ytPreviewWin.on('closed', () => { _ytPreviewWin = null; });
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: true }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }

                if (command === 'motion-tracker/open-editor' && data.mediaPath) {
                    try {
                        const { BrowserWindow } = require('electron');
                        // Singleton: se ja tem editor aberto, foca nele em vez de criar novo
                        if (_activeMtEditorWin && !_activeMtEditorWin.isDestroyed()) {
                            try {
                                if (_activeMtEditorWin.isMinimized()) _activeMtEditorWin.restore();
                                _activeMtEditorWin.show();
                                _activeMtEditorWin.focus();
                                _activeMtEditorWin.moveTop();
                            } catch(eF) {}
                            res.writeHead(200, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ ok: true, focused: true }));
                            return;
                        }
                        // Extrai token do header Authorization pra passar pro editor
                        // (editor precisa autenticar /motion-tracker/track e /queue-apply)
                        let token = '';
                        try {
                            const authH = req.headers['authorization'] || '';
                            if (authH.startsWith('Bearer ')) token = authH.slice(7);
                        } catch (e) {}
                        const session = String(data.session || ('mt-' + Date.now()));
                        const inSec = parseFloat(data.inPointSec) || 0;
                        const outSec = parseFloat(data.outPointSec) || 0;
                        const dur = Math.max(0.5, outSec - inSec);
                        // ─── Proxy MP4 H.264 — HTML5 video não decodifica ProRes/HEVC/DNxHR ───
                        // Tracking continua usando o ORIGINAL (full quality).
                        const proxyDir = path.join(currentUD, 'mt-proxies');
                        try { fs.mkdirSync(proxyDir, { recursive: true }); } catch (e) {}
                        const proxyPath = path.join(proxyDir, 'proxy-' + session + '.mp4');
                        console.log('[mt-editor] gerando proxy:', proxyPath);
                        await new Promise((resolve) => {
                            // Downscale pra max 720p — tracking não precisa de 4K.
                            // Acelera 5-10x: 4K→720p reduz frames de 8MB pra 1MB cada.
                            const ff = spawn(ffmpegBin, [
                                '-y',
                                '-ss', String(Math.max(0, inSec - 0.05)),
                                '-i', data.mediaPath,
                                '-t', String(dur + 0.1),
                                '-vf', 'scale=-2:min(720\\,ih)', // mantém aspect, limita altura a 720
                                '-c:v', 'libx264',
                                '-preset', 'ultrafast',
                                '-crf', '23',
                                '-pix_fmt', 'yuv420p',
                                '-movflags', '+faststart',
                                '-an',
                                proxyPath
                            ], { stdio: ['ignore', 'ignore', 'pipe'] });
                            let errBuf = '';
                            ff.stderr.on('data', d => { errBuf += d.toString(); });
                            ff.on('close', (code) => {
                                if (code === 0 && fs.existsSync(proxyPath)) {
                                    console.log('[mt-editor] proxy OK:', fs.statSync(proxyPath).size, 'bytes');
                                } else {
                                    console.error('[mt-editor] proxy FAIL code=' + code, errBuf.slice(-500));
                                }
                                resolve();
                            });
                            ff.on('error', e => { console.error('[mt-editor] ffmpeg spawn err:', e); resolve(); });
                        });
                        const usingProxy = fs.existsSync(proxyPath) && fs.statSync(proxyPath).size > 1024;
                        const previewPath = usingProxy ? proxyPath : data.mediaPath;
                        // No proxy o vídeo começa em 0s (trecho cortado começa do 0)
                        const editorInSec = usingProxy ? 0 : inSec;
                        const editorOutSec = usingProxy ? dur : outSec;
                        const q = new URLSearchParams({
                            session,
                            token,
                            media: encodeURIComponent(previewPath),
                            origMedia: encodeURIComponent(data.mediaPath),
                            usingProxy: usingProxy ? '1' : '0',
                            fps: String(data.fps || 30),
                            inSec: String(editorInSec),
                            outSec: String(editorOutSec),
                            origInSec: String(inSec),
                            origOutSec: String(outSec),
                            // Plugin JSX envia `inPointTicks` (não clipInPointTicks).
                            // Sem isso, keyframes vão pra source-time 0 (fora do trecho visível).
                            clipInTicks: String(data.inPointTicks || data.clipInPointTicks || '0'),
                            clipStartTicks: String(data.clipStartTicks || '0'),
                            trackIdx: String(data.trackIdx != null ? data.trackIdx : -1),
                        }).toString();
                        const win = new BrowserWindow({
                            width: 1280, height: 800,
                            title: 'Motion Tracker',
                            backgroundColor: '#0a0a0a',
                            show: true,
                            center: true,
                            alwaysOnTop: false,
                            skipTaskbar: false,
                            webPreferences: { nodeIntegration: true, contextIsolation: false }
                        });
                        try { require('@electron/remote/main').enable(win.webContents); } catch (e) { console.warn('[mt-editor] remote enable err:', e.message); }
                        _activeMtEditorWin = win;
                        win.on('closed', () => { if (_activeMtEditorWin === win) _activeMtEditorWin = null; });
                        // Debug: loga erros de load + crash
                        win.webContents.on('did-fail-load', (_e, code, desc) => console.error('[mt-editor] load fail:', code, desc));
                        win.webContents.on('render-process-gone', (_e, det) => console.error('[mt-editor] render gone:', det));
                        win.webContents.on('console-message', (_e, lvl, msg) => console.log('[mt-editor:console]', msg));
                        win.webContents.on('did-finish-load', () => {
                            console.log('[mt-editor] loaded OK');
                        });
                        const htmlPath = path.join(__dirname, 'motion-tracker-editor.html');
                        console.log('[mt-editor] loading:', htmlPath, 'q:', q);
                        win.loadFile(htmlPath, { search: q });
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: true, session }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                if (command === 'motion-tracker/queue-apply' && data.history) {
                    try {
                        // Enfileira como comando pro plugin pegar via long-poll
                        const commandId = 'mt-apply-' + Date.now();
                        pendingPluginCommands.push({
                            id: commandId,
                            type: 'motion-tracker/apply',
                            payload: data
                        });
                        // Acorda waiters
                        if (typeof _notifyLongPollWaiters === 'function') {
                            _notifyLongPollWaiters();
                        } else {
                            while (pendingPluginCommands.length > 0 && lionLongPollWaiters.length > 0) {
                                const w = lionLongPollWaiters.shift();
                                clearTimeout(w.timeoutHandle);
                                const cmds = pendingPluginCommands.splice(0, pendingPluginCommands.length);
                                try {
                                    w.res.writeHead(200, { 'Content-Type': 'application/json' });
                                    w.res.end(JSON.stringify({ commands: cmds }));
                                } catch (e) {}
                            }
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ ok: true, commandId }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                if (command === 'motion-tracker/track' && data.filePath && Array.isArray(data.points)) {
                    try {
                        // v1.4: roda via Python+cv2 nativo (10-100x mais rapido que WASM).
                        // Spawn isolado do main process — nao bloqueia HTTP server.
                        console.log('[motion-tracker] python track, points=' + data.points.length, 'algo=' + (data.options?.algorithm || 'CSRT'));
                        const t0 = Date.now();
                        const result = await _runMotionTracker(
                            data.filePath, ffmpegBin, ffprobeBin,
                            data.points, data.options || {}
                        );
                        console.log('[motion-tracker] track done in ' + (Date.now()-t0) + 'ms');
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(result));
                    } catch (e) {
                        console.error('[motion-tracker] track err:', e);
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                if (command === 'motion-tracker/check-python' && req.method === 'POST') {
                    // Verifica se Python+OpenCV estao disponiveis. Resposta: {ok, python, error?}
                    try {
                        _cachedPython = null; // forca redetect
                        const py = _findPython();
                        if (py) {
                            res.writeHead(200, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ ok: true, python: py }));
                        } else {
                            // Detecta se ao menos Python existe (mesmo sem opencv)
                            const { execSync } = require('child_process');
                            let pyOnly = null;
                            const candidates = isWin ? ['py -3', 'python', 'python3'] : ['python3', 'python'];
                            for (const c of candidates) {
                                try { execSync(`${c} --version`, { timeout: 3000, stdio: ['ignore', 'pipe', 'ignore'] }); pyOnly = c; break; } catch {}
                            }
                            res.writeHead(200, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ ok: false, pythonFound: !!pyOnly, python: pyOnly, error: pyOnly ? 'Python OK mas falta opencv-python/numpy' : 'Python nao instalado' }));
                        }
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }
                if (command === 'motion-tracker/install-deps' && req.method === 'POST') {
                    // Roda pip install opencv-python numpy scipy via Python detectado.
                    try {
                        const { execSync } = require('child_process');
                        let py = null;
                        const candidates = isWin ? ['py -3', 'python', 'python3'] : ['python3', 'python'];
                        for (const c of candidates) {
                            try { execSync(`${c} --version`, { timeout: 3000, stdio: ['ignore', 'pipe', 'ignore'] }); py = c; break; } catch {}
                        }
                        if (!py) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ ok: false, error: 'Python nao encontrado. Instale Python 3.10+ primeiro: https://python.org/downloads/' }));
                            return;
                        }
                        const { spawn } = require('child_process');
                        const useShell = isWin && py.startsWith('py ');
                        // opencv-contrib-python inclui CSRT/KCF/MOSSE/MIL via cv2.legacy.* (opencv-python nao tem)
                        const args = ['-m', 'pip', 'install', '--upgrade', 'opencv-contrib-python', 'numpy', 'scipy'];
                        const proc = useShell
                            ? spawn(py + ' ' + args.join(' '), { shell: true, windowsHide: true })
                            : spawn(py, args, { windowsHide: true });
                        let out = '', err = '';
                        proc.stdout.on('data', d => { out += d.toString(); });
                        proc.stderr.on('data', d => { err += d.toString(); });
                        proc.on('close', code => {
                            _cachedPython = null; // forca redetect
                            const verified = _findPython();
                            if (code === 0 && verified) {
                                res.writeHead(200, { 'Content-Type': 'application/json' });
                                res.end(JSON.stringify({ ok: true, python: verified, log: out.slice(-1000) }));
                            } else {
                                res.writeHead(500, { 'Content-Type': 'application/json' });
                                res.end(JSON.stringify({ ok: false, error: `pip exit ${code}`, log: (err + out).slice(-2000) }));
                            }
                        });
                        proc.on('error', e => {
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ ok: false, error: e.message || String(e) }));
                        });
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }

                // ═══════ Background Remover — IA local via BiRefNet ═══════
                if (command === 'bg/remove' && data.filePath) {
                    try {
                        const origPath = data.filePath;
                        if (!fs.existsSync(origPath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Arquivo não encontrado: ' + origPath }));
                            return;
                        }

                        // Formatos não-padrão: converte pra PNG via Electron nativeImage
                        // (sharp já suporta JPG/PNG/WEBP/TIFF/HEIC/etc — converte só pra raros)
                        const ext = path.extname(origPath).toLowerCase().substring(1);
                        const SUPPORTED = new Set(['png','jpg','jpeg','webp','tif','tiff','bmp','gif']);
                        let workingPath = origPath;
                        let tmpConverted = null;
                        if (!SUPPORTED.has(ext)) {
                            try {
                                const { nativeImage } = require('electron');
                                const img = nativeImage.createFromPath(origPath);
                                if (img.isEmpty()) throw new Error('formato não decodificável: ' + ext);
                                const pngBuf = img.toPNG();
                                tmpConverted = path.join(currentUD, 'bg-input-' + Date.now() + '.png');
                                fs.writeFileSync(tmpConverted, pngBuf);
                                workingPath = tmpConverted;
                            } catch (eConv) {
                                res.writeHead(400, { 'Content-Type': 'application/json' });
                                res.end(JSON.stringify({ error: 'Formato não suportado (' + ext + '). Use PNG, JPG, WEBP, etc.' }));
                                return;
                            }
                        }

                        // Processa via BiRefNet (helper definido em escopo de módulo)
                        let resultBuf;
                        try {
                            resultBuf = await _removeBgBiRefNet(workingPath);
                        } catch (eBR) {
                            console.error('[bg/remove] BiRefNet err:', eBR);
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'IA falhou: ' + (eBR.message || String(eBR)) }));
                            if (tmpConverted) { try { fs.unlinkSync(tmpConverted); } catch(e) {} }
                            return;
                        }

                        // Salva resultado ao lado do original
                        const dir = path.dirname(origPath);
                        const base = path.basename(origPath, path.extname(origPath));
                        let outPath = path.join(dir, base + '_nobg.png');
                        try {
                            fs.writeFileSync(outPath, resultBuf);
                        } catch (eW) {
                            outPath = path.join(currentUD, 'bgremoved-' + Date.now() + '.png');
                            fs.writeFileSync(outPath, resultBuf);
                        }

                        // Cleanup do arquivo convertido temporário
                        if (tmpConverted) { try { fs.unlinkSync(tmpConverted); } catch(e) {} }

                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        // origPath retornado pra cliente poder chamar /bg/edit-mask depois
                        res.end(JSON.stringify({ ok: true, outPath: outPath, origPath: origPath }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }

                // ═══════ BG Remove Vídeo — inicia processamento ═══════
                if (command === 'bg/remove-video' && data.filePath) {
                    try {
                        if (!fs.existsSync(data.filePath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Arquivo não encontrado: ' + data.filePath }));
                            return;
                        }
                        // Valida que é arquivo de vídeo (não imagem)
                        const ext = path.extname(data.filePath).toLowerCase();
                        const imageExts = ['.png','.jpg','.jpeg','.gif','.bmp','.webp','.tiff','.tif','.heic','.heif','.svg','.psd'];
                        const videoExts = ['.mp4','.mov','.mkv','.avi','.webm','.mxf','.m4v','.mpg','.mpeg','.wmv','.flv','.ts','.m2ts','.mts','.3gp','.dv','.prores','.dnxhd'];
                        if (imageExts.includes(ext)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Esse é um arquivo de imagem (' + ext + '). Use "Remover fundo" (acima), não a versão de vídeo.' }));
                            return;
                        }
                        if (!videoExts.includes(ext) && !data.skipExtCheck) {
                            // Extensão desconhecida — checa via ffprobe se tem stream de vídeo
                            // Avisa o user mas tenta processar mesmo assim
                            console.warn('[bg-video] extensão desconhecida:', ext, '— tentando processar mesmo assim');
                        }
                        const fmt = (data.format === 'mov') ? 'mov' : 'webm';
                        const qual = ['fast','medium','high'].includes(data.quality) ? data.quality : 'medium';
                        const dele = (String(data.delegate || 'GPU').toUpperCase() === 'CPU') ? 'CPU' : 'GPU';
                        const inSec = Math.max(0, Number(data.inSec) || 0);
                        const durSec = Math.max(0, Number(data.durationSec) || 0);

                        // Garante ffmpeg/ffprobe (auto-download se faltar)
                        if (!ffmpegReady() || !ffprobeReady()) {
                            console.log('[bg-video] baixando ffmpeg/ffprobe...');
                            await ensureFfmpeg();
                        }
                        if (!ffmpegReady() || !ffprobeReady()) {
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'FFmpeg/FFprobe não disponíveis — verifique sua conexão e tente novamente' }));
                            return;
                        }

                        // Pré-transcode: FFmpeg trim do source pra .mp4 H.264 temp.
                        // Resolve 2 problemas:
                        //   1) HTML5 <video> só decodifica H.264/VP9/AV1. ProRes/HEVC/etc
                        //      davam "Falha ao decodificar". Agora sempre vira H.264.
                        //   2) Respeita o trecho IN→OUT do clip na timeline (não processa
                        //      vídeo inteiro se user só selecionou um pedaço).
                        const transcodedPath = path.join(currentUD, 'bg-video-trim-' + Date.now() + '.mp4');
                        try {
                            await new Promise((resolve, reject) => {
                                const ffArgs = ['-y'];
                                if (durSec > 0) {
                                    // -ss antes de -i = seek rápido (fast seek). Reencode garante key frame inicial.
                                    ffArgs.push('-ss', String(inSec));
                                    ffArgs.push('-t', String(durSec));
                                }
                                ffArgs.push('-i', data.filePath);
                                if (durSec > 0) {
                                    ffArgs.push('-ss', '0', '-t', String(durSec)); // re-aplica pra garantir trim exato
                                }
                                ffArgs.push(
                                    '-c:v', 'libx264',
                                    '-preset', 'ultrafast',
                                    '-crf', '20',
                                    '-pix_fmt', 'yuv420p',
                                    // -g 1 = todo frame é keyframe (intra-only).
                                    // Arquivo fica grande mas SEEK é instantâneo.
                                    // Worker faz seek frame-a-frame, então isso é
                                    // crítico pra performance + garantir TODOS os
                                    // frames processados (rVFC pulava frames).
                                    '-g', '1',
                                    '-keyint_min', '1',
                                    '-sc_threshold', '0',
                                    '-c:a', 'aac', '-b:a', '128k',
                                    '-movflags', '+faststart',
                                    transcodedPath
                                );
                                console.log('[bg-video] pre-transcode:', ffArgs.join(' '));
                                const proc = bgvSpawn(ffmpegBin, ffArgs, { windowsHide: true });
                                let stderr = '';
                                proc.stderr.on('data', d => { stderr += d.toString(); if (stderr.length > 4096) stderr = stderr.slice(-4096); });
                                proc.on('close', (code) => {
                                    if (code === 0 && fs.existsSync(transcodedPath)) resolve();
                                    else reject(new Error('FFmpeg trim/transcode falhou (code ' + code + '): ' + stderr.slice(-300)));
                                });
                                proc.on('error', e => reject(e));
                            });
                        } catch (eTrans) {
                            // Cleanup tentativa
                            try { if (fs.existsSync(transcodedPath)) fs.unlinkSync(transcodedPath); } catch(e) {}
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Pré-processamento falhou: ' + eTrans.message }));
                            return;
                        }

                        // Worker recebe o arquivo transcodificado (não o original)
                        const token = startBgVideoWorker(transcodedPath, fmt, qual, dele);
                        // Anota original + temp na sessão pra cleanup futuro e import correto
                        const sess = bgVideoSessions.get(token);
                        if (sess) {
                            sess._origSrcPath = data.filePath;
                            sess._tempTranscoded = transcodedPath;
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ token: token, status: 'starting' }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }

                // ═══════ BG Remove Vídeo — cancel ═══════
                if (command === 'bg/remove-video/cancel' && data.token) {
                    const sess = bgVideoSessions.get(data.token);
                    if (sess) {
                        sess.status = 'cancelled';
                        try { if (sess.encoderProc) sess.encoderProc.kill('SIGKILL'); } catch(e) {}
                        try { if (sess.outPath && fs.existsSync(sess.outPath)) fs.unlinkSync(sess.outPath); } catch(e) {}
                        try { if (sess.win && !sess.win.isDestroyed()) sess.win.close(); } catch(e) {}
                    }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }

                // ═══════ Rotoscope SAM 2 — abre editor pra clip de vídeo ═══════
                if (command === 'rotoscope/start' && data.filePath) {
                    try {
                        if (!fs.existsSync(data.filePath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Arquivo não encontrado' }));
                            return;
                        }
                        // Validar que é vídeo (não imagem)
                        const ext = path.extname(data.filePath).toLowerCase();
                        const imgExts = ['.png','.jpg','.jpeg','.gif','.bmp','.webp','.tiff','.tif'];
                        if (imgExts.includes(ext)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Selecione um clip de vídeo (não imagem)' }));
                            return;
                        }
                        // Output padrão ao lado do source
                        const dir = path.dirname(data.filePath);
                        const base = path.basename(data.filePath, path.extname(data.filePath));
                        const outPath = path.join(dir, base + '_rotoscope.mov');

                        // Responde IMEDIATAMENTE com 200 — pré-transcode pode demorar
                        // alguns segundos pra vídeos grandes; cliente já pode mostrar
                        // mensagem "abrindo".
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        const tokenPromise = openRotoscopeEditor(data.filePath, outPath);
                        tokenPromise.then(token => {
                            res.end(JSON.stringify({ token, status: 'editing' }));
                        }).catch(err => {
                            res.end(JSON.stringify({ error: err.message || String(err) }));
                        });
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message }));
                    }
                    return;
                }

                // ═══════ Mask Editor — Pincel mágico pra ajustar bg-remove ═══════
                if (command === 'bg/edit-mask' && data.origPath && data.processedPath) {
                    try {
                        if (!fs.existsSync(data.origPath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Imagem original não encontrada: ' + data.origPath }));
                            return;
                        }
                        if (!fs.existsSync(data.processedPath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'PNG processado não encontrado: ' + data.processedPath }));
                            return;
                        }
                        // Abre janela do editor (não bloqueia HTTP — retorna token)
                        const token = openMaskEditor(data.origPath, data.processedPath);
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ token: token, status: 'editing' }));
                    } catch (e) {
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message || String(e) }));
                    }
                    return;
                }


                if (command === 'yt/cancel') {
                    if (ytDownloadProc) { ytDownloadProc.kill(); ytDownloadProc = null; }
                    ytProgress = { active: false, percent: 0, speed: '', eta: '', title: '', status: 'cancelled', error: '' };
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }

                // Auto-manage timer foreground check from HTTP commands
                if (command === 'timer/start') {
                    timerFgPaused = false;
                    if (!timerFgInterval) {
                        timerFgInterval = setInterval(checkTimerForeground, 1200);
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

    // Leak prevention: close idle sockets só the plugin reconnecting doesn't accumulate them
    server.keepAliveTimeout = 30000;   // 30s
    server.headersTimeout = 35000;     // must be > keepAliveTimeout
    server.requestTimeout = 60000;     // reject requests that hang >60s

    server.listen(SYNC_PORT, '127.0.0.1', () => {
        console.log('Lion Workspace sync: http://127.0.0.1:' + SYNC_PORT);
    });
    server.on('error', (err) => {
        console.error('Sync server error:', err.message);
    });
}

// Helper: collect request body with hard size limit (prevents unbounded string growth)
const MAX_BODY_BYTES = 5 * 1024 * 1024; // 5MB
function readBodyLimited(req, cb) {
    let body = '';
    let size = 0;
    let aborted = false;
    req.on('data', chunk => {
        if (aborted) return;
        size += chunk.length;
        if (size > MAX_BODY_BYTES) {
            aborted = true;
            try { req.destroy(); } catch(e) {}
            cb(new Error('body_too_large'));
            return;
        }
        body += chunk;
    });
    req.on('end', () => { if (!aborted) cb(null, body); });
    req.on('error', e => { if (!aborted) { aborted = true; cb(e); } });
}

/* ═══════ Auto-Updater ═══════ */
autoUpdater.autoInstallOnAppQuit = true;
autoUpdater.logger = console;
// macOS: Squirrel.Mac/ShipIt does native code signature validation that cannot be
// bypassed without an Apple Developer certificate ($99/yr). So on Mac we only CHECK
// for updates, then open the GitHub releases page for manual download.
if (isMac) {
    autoUpdater.autoDownload = false;
    autoUpdater.verifyUpdateCodeSignature = false;
} else {
    autoUpdater.autoDownload = true;
}

const RELEASES_URL = 'https://github.com/IlonPeruzzo/lion-workspace/releases/latest';

function setupAutoUpdater() {
    autoUpdater.on('checking-for-update', () => {
        sendUpdateStatus('checking');
    });

    autoUpdater.on('update-available', (info) => {
        if (isMac) {
            // Mac: tell frontend to show download link (no auto-install)
            sendUpdateStatus('available-mac', { version: info.version });
        } else {
            sendUpdateStatus('available', { version: info.version });
        }
    });

    autoUpdater.on('update-not-available', () => {
        sendUpdateStatus('not-available');
    });

    autoUpdater.on('download-progress', (progress) => {
        sendUpdateStatus('downloading', {
            percent: Math.round(progress.percent),
            transferred: progress.transferred,
            total: progress.total
        });
    });

    autoUpdater.on('update-downloaded', (info) => {
        sendUpdateStatus('downloaded', { version: info.version });
    });

    autoUpdater.on('error', (err) => {
        console.error('Auto-updater error:', err);
        sendUpdateStatus('error', { message: err.message });
    });

    // Check for updates after a short delay (let the app finish loading)
    setTimeout(() => {
        autoUpdater.checkForUpdates().catch(err => {
            console.log('Update check skipped:', err.message);
        });
    }, 5000);

    // Re-check every 30 minutes (frequent during early releases)
    setInterval(() => {
        autoUpdater.checkForUpdates().catch(() => {});
    }, 30 * 60 * 1000);
}

// IPC: open releases page (used for Mac manual update)
ipcMain.handle('open-releases-page', () => {
    require('electron').shell.openExternal(RELEASES_URL);
});

function sendUpdateStatus(status, data = {}) {
    if (mainWin && !mainWin.isDestroyed()) {
        mainWin.webContents.send('update-status', { status, ...data });
    }
}

// IPC: get app version
ipcMain.handle('get-app-version', () => app.getVersion());

// IPC: check for updates manually
ipcMain.handle('check-for-updates', async () => {
    try {
        const result = await autoUpdater.checkForUpdates();
        return { success: true, version: result?.updateInfo?.version };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IPC: install downloaded update (restart)
ipcMain.handle('install-update', () => {
    autoUpdater.quitAndInstall(false, true);
});

/* ═══════ Auth & License ═══════ */
const SUPABASE_URL = 'https://rxwprlqskwylhvpjiung.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InJ4d3BybHFza3d5bGh2cGppdW5nIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzU3NjYyNTcsImV4cCI6MjA5MTM0MjI1N30.0ej5GjmfzCMu9odw6R6mjU6GTAvdjUu2vQ5t6dsJYkM';

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: {
        autoRefreshToken: true,
        persistSession: true,
        storage: {
            getItem: (key) => {
                try {
                    const filePath = path.join(app.getPath('userData'), 'auth-storage.json');
                    if (!fs.existsSync(filePath)) return null;
                    const data = JSON.parse(fs.readFileSync(filePath, 'utf8'));
                    return data[key] || null;
                } catch { return null; }
            },
            setItem: (key, value) => {
                try {
                    const filePath = path.join(app.getPath('userData'), 'auth-storage.json');
                    let data = {};
                    try { data = JSON.parse(fs.readFileSync(filePath, 'utf8')); } catch {}
                    data[key] = value;
                    fs.writeFileSync(filePath, JSON.stringify(data));
                } catch {}
            },
            removeItem: (key) => {
                try {
                    const filePath = path.join(app.getPath('userData'), 'auth-storage.json');
                    let data = {};
                    try { data = JSON.parse(fs.readFileSync(filePath, 'utf8')); } catch {}
                    delete data[key];
                    fs.writeFileSync(filePath, JSON.stringify(data));
                } catch {}
            }
        }
    }
});

// Device fingerprint (stable per machine)
function getDeviceFingerprint() {
    const cpus = os.cpus();
    const cpuModel = cpus.length > 0 ? cpus[0].model : 'unknown';
    const nets = os.networkInterfaces();
    let mac = '';
    for (const ifaces of Object.values(nets)) {
        for (const iface of ifaces) {
            if (!iface.internal && iface.mac && iface.mac !== '00:00:00:00:00:00') {
                mac = iface.mac;
                break;
            }
        }
        if (mac) break;
    }
    const raw = `${os.hostname()}-${cpuModel}-${mac}-${os.platform()}-${os.arch()}`;
    return crypto.createHash('sha256').update(raw).digest('hex');
}

// License cache for offline grace period
const LICENSE_CACHE_FILE = 'license-cache.json';

function getLicenseCache() {
    try {
        const filePath = path.join(app.getPath('userData'), LICENSE_CACHE_FILE);
        return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch { return null; }
}

function setLicenseCache(data) {
    try {
        const filePath = path.join(app.getPath('userData'), LICENSE_CACHE_FILE);
        fs.writeFileSync(filePath, JSON.stringify({ ...data, cached_at: new Date().toISOString() }));
    } catch {}
}

function clearLicenseCache() {
    try {
        const filePath = path.join(app.getPath('userData'), LICENSE_CACHE_FILE);
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    } catch {}
}

// Timeout wrapper for async operations
function withTimeout(promise, ms = 15000) {
    return Promise.race([
        promise,
        new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout — verifique sua conexão')), ms))
    ]);
}

// IPC: Login
// Snapshot do usuário atualmente logado no app (alimenta auto-login do plugin)
let currentAppUser = null; // { id, email, name }
ipcMain.handle('auth-login', async (_, email, password) => {
    try {
        const { data, error } = await withTimeout(supabase.auth.signInWithPassword({ email, password }));
        if (error) return { success: false, error: error.message };

        // Save snapshot for plugin auto-login
        currentAppUser = {
            id: data.user.id,
            email: data.user.email,
            name: data.user.user_metadata?.full_name || ''
        };

        // Register/update device (don't block login for this)
        const fingerprint = getDeviceFingerprint();
        supabase.from('devices').upsert({
            user_id: data.user.id,
            fingerprint,
            name: os.hostname(),
            os: os.platform(),
            platform: os.platform(),
            os_version: os.release(),
            hostname: os.hostname(),
            cpu_model: os.cpus()[0]?.model || 'unknown',
            app_version: app.getVersion(),
            last_seen_at: new Date().toISOString(),
            last_heartbeat_at: new Date().toISOString()
        }, { onConflict: 'user_id,fingerprint' }).then(
            (res) => { if (res.error) console.error('[device-register] upsert failed:', res.error.message); else console.log('[device-register] OK'); },
            (err) => { console.error('[device-register] exception:', err.message); }
        );

        return {
            success: true,
            user: { id: data.user.id, email: data.user.email }
        };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IPC: Logout
ipcMain.handle('auth-logout', async () => {
    try {
        currentAppUser = null;
        await supabase.auth.signOut();
        clearLicenseCache();
        return { success: true };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// On startup, try to restore user from persisted Supabase session
(async () => {
    try {
        const { data: { session } } = await supabase.auth.getSession();
        if (session && session.user) {
            currentAppUser = {
                id: session.user.id,
                email: session.user.email,
                name: session.user.user_metadata?.full_name || ''
            };
            console.log('[auth] Restored user session from disk: ' + currentAppUser.email);
        }
    } catch(e) { console.warn('[auth] Could not restore session:', e.message); }
})();

// IPC: Get current session/user
ipcMain.handle('auth-get-session', async () => {
    try {
        const { data: { session } } = await supabase.auth.getSession();
        if (!session) return { authenticated: false };

        const { data: { user } } = await supabase.auth.getUser();
        if (!user) return { authenticated: false };

        // Get profile
        const { data: profile } = await supabase
            .from('profiles')
            .select('full_name, avatar_url, email')
            .eq('id', user.id)
            .single();

        return {
            authenticated: true,
            user: {
                id: user.id,
                email: user.email,
                full_name: profile?.full_name || '',
                avatar_url: profile?.avatar_url || ''
            }
        };
    } catch {
        return { authenticated: false };
    }
});

// IPC: Check license/subscription status
ipcMain.handle('auth-check-license', async () => {
    try {
        const { data: { user } } = await supabase.auth.getUser();
        if (!user) {
            // Check offline cache
            const cache = getLicenseCache();
            if (cache) {
                const cachedDate = new Date(cache.cached_at);
                const daysSince = (Date.now() - cachedDate.getTime()) / (1000 * 60 * 60 * 24);
                if (daysSince <= 7 && cache.valid) {
                    return { ...cache, offline: true };
                }
            }
            return { valid: false, reason: 'not_authenticated' };
        }

        // Check subscription
        const { data: sub } = await supabase
            .from('subscriptions')
            .select('*')
            .eq('user_id', user.id)
            .in('status', ['active', 'trialing'])
            .order('created_at', { ascending: false })
            .limit(1)
            .single();

        if (!sub) {
            return { valid: false, reason: 'no_subscription' };
        }

        // Check device
        const fingerprint = getDeviceFingerprint();
        const { data: device } = await supabase
            .from('devices')
            .select('blocked')
            .eq('user_id', user.id)
            .eq('fingerprint', fingerprint)
            .single();

        if (device?.blocked) {
            return { valid: false, reason: 'device_blocked' };
        }

        // Update heartbeat
        await supabase
            .from('devices')
            .update({
                last_heartbeat_at: new Date().toISOString(),
                last_seen_at: new Date().toISOString(),
                app_version: app.getVersion()
            })
            .eq('user_id', user.id)
            .eq('fingerprint', fingerprint);

        const result = {
            valid: true,
            plan: sub.plan,
            status: sub.status,
            current_period_end: sub.current_period_end
        };
        setLicenseCache(result);
        return result;
    } catch (err) {
        // Offline — check cache
        const cache = getLicenseCache();
        if (cache) {
            const cachedDate = new Date(cache.cached_at);
            const daysSince = (Date.now() - cachedDate.getTime()) / (1000 * 60 * 60 * 24);
            if (daysSince <= 7 && cache.valid) {
                return { ...cache, offline: true };
            }
        }
        return { valid: false, reason: 'offline_expired', error: err.message };
    }
});

// IPC: Signup
ipcMain.handle('auth-signup', async (_, email, password, fullName) => {
    try {
        const { data, error } = await withTimeout(supabase.auth.signUp({
            email,
            password,
            options: { data: { full_name: fullName } }
        }));
        if (error) return { success: false, error: error.message };

        // Ensure profile exists (don't block signup for this)
        if (data.user) {
            supabase.from('profiles').upsert({
                id: data.user.id,
                email: email,
                full_name: fullName
            }, { onConflict: 'id' }).then(() => {}, () => {});
        }

        return {
            success: true,
            needsConfirmation: !data.session,
            user: data.user ? { id: data.user.id, email: data.user.email } : null
        };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IPC: Forgot password
ipcMain.handle('auth-forgot-password', async (_, email) => {
    try {
        const { error } = await supabase.auth.resetPasswordForEmail(email, {
            redirectTo: 'https://app.lionwork.com.br/reset-password'
        });
        if (error) return { success: false, error: error.message };
        return { success: true };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IPC: Detect if Adobe Premiere Pro is installed
ipcMain.handle('detect-premiere', async () => {
    try {
        let cepDir;
        if (isMac) {
            cepDir = path.join(os.homedir(), 'Library', 'Application Support', 'Adobe', 'CEP', 'extensions');
        } else {
            cepDir = path.join(process.env.APPDATA || '', 'Adobe', 'CEP', 'extensions');
        }
        // Check if Adobe folder exists (indicates Adobe products installed)
        const adobeDir = isMac
            ? path.join(os.homedir(), 'Library', 'Application Support', 'Adobe')
            : path.join(process.env.APPDATA || '', 'Adobe');
        const adobeInstalled = fs.existsSync(adobeDir);

        // Check if our plugin is already installed
        const pluginInstalled = fs.existsSync(path.join(cepDir, 'com.lionworkspace.premiere', 'host', 'index.jsx'));
        const aePluginInstalled = fs.existsSync(path.join(cepDir, 'com.lionworkspace.after', 'host', 'index.jsx'));

        return { adobeInstalled, pluginInstalled, aePluginInstalled };
    } catch {
        return { adobeInstalled: false, pluginInstalled: false, aePluginInstalled: false };
    }
});

// IPC: Install plugins selectively. Receives { premiere: bool, ae: bool }
ipcMain.handle('install-premiere-plugin', async (event, opts) => {
    try {
        // Server-side license check before installing plugin
        const { data: { user } } = await supabase.auth.getUser();
        if (!user) return { success: false, error: 'not_authenticated' };

        const { data: sub } = await supabase
            .from('subscriptions')
            .select('plan, status')
            .eq('user_id', user.id)
            .in('status', ['active', 'trialing'])
            .limit(1)
            .single();

        if (!sub) return { success: false, error: 'no_subscription' };

        // Compat: se opts não vier, instala ambos (comportamento antigo)
        const installPr = !opts || opts.premiere !== false;
        const installAe = !opts || opts.ae !== false;
        let installed = { premiere: false, ae: false };
        if (installPr) installed.premiere = autoInstallCepPlugin('premiere-plugin', 'com.lionworkspace.premiere');
        if (installAe) installed.ae = autoInstallCepPlugin('after-plugin', 'com.lionworkspace.after');
        // Enable CEP debug mode (covers both plugins)
        if (installPr || installAe) {
            if (isMac) {
                for (let v = 9; v <= 14; v++) exec(`defaults write com.adobe.CSXS.${v} PlayerDebugMode 1`, () => {});
            } else {
                for (let v = 9; v <= 14; v++) exec(`reg add "HKCU\\SOFTWARE\\Adobe\\CSXS.${v}" /v PlayerDebugMode /t REG_SZ /d 1 /f`, () => {});
            }
        }
        return { success: true, installed };
    } catch (err) {
        return { success: false, error: err.message };
    }
});

// IPC: Skip plugin installation (write flag)
ipcMain.handle('skip-premiere-plugin', () => {
    try {
        const flagPath = path.join(path.dirname(process.execPath), 'skip-plugin.flag');
        fs.writeFileSync(flagPath, 'true');
        return { success: true };
    } catch {
        return { success: true }; // non-critical
    }
});

// Heartbeat every 5 minutes (keep device alive)
let heartbeatInterval = null;
function startHeartbeat() {
    if (heartbeatInterval) clearInterval(heartbeatInterval);
    heartbeatInterval = setInterval(async () => {
        try {
            const { data: { user } } = await supabase.auth.getUser();
            if (!user) return;
            const fingerprint = getDeviceFingerprint();
            await supabase
                .from('devices')
                .update({
                    last_heartbeat_at: new Date().toISOString(),
                    last_seen_at: new Date().toISOString(),
                    app_version: app.getVersion()
                })
                .eq('user_id', user.id)
                .eq('fingerprint', fingerprint);
        } catch {}
    }, 5 * 60 * 1000);
}

function stopHeartbeat() {
    if (heartbeatInterval) { clearInterval(heartbeatInterval); heartbeatInterval = null; }
}

/* ═══════ App lifecycle ═══════ */
app.on('before-quit', () => {
    app.isQuitting = true;
    stopFgDaemon();
    stopHeartbeat();
    // ── Clean up focus mode só it doesn't run forever ──
    if (focusActive) {
        focusActive = false;
        focusExternal = false;
        if (focusInterval) { clearInterval(focusInterval); focusInterval = null; }
        destroyOverlays();
    }
});

app.whenReady().then(() => {
    // Initialize @electron/remote pro motion-tracker-editor.html
    try { require('@electron/remote/main').initialize(); } catch (e) {}
    // Set up Edit menu só Ctrl+V / Cmd+V works in all input fields
    const menuTemplate = [
        ...(isMac ? [{ role: 'appMenu' }] : []),
        {
            label: 'Edit',
            submenu: [
                { role: 'undo' },
                { role: 'redo' },
                { type: 'separator' },
                { role: 'cut' },
                { role: 'copy' },
                { role: 'paste' },
                { role: 'selectAll' }
            ]
        }
    ];
    Menu.setApplicationMenu(Menu.buildFromTemplate(menuTemplate));

    createWindow();
    startSyncServer();
    autoInstallPremierePlugin();
    setupAutoUpdater();
    startHeartbeat();
});

// Auto-install Premiere plugin on startup (works on both Mac and Windows)
// Helper genérico: instala um plugin CEP (Premiere ou After) na pasta padrão do Adobe.
// srcFolderName: 'premiere-plugin' ou 'after-plugin' (relativo à pasta do app)
// extensionId:   'com.lionworkspace.premiere' ou 'com.lionworkspace.after'
function autoInstallCepPlugin(srcFolderName, extensionId) {
    try {
        // Find the plugin source (bundled with app in extraResources)
        let pluginSrc = path.join(process.resourcesPath, srcFolderName);
        if (!fs.existsSync(pluginSrc)) {
            // Dev mode: plugin is next to the app
            pluginSrc = path.join(__dirname, srcFolderName);
        }
        if (!fs.existsSync(pluginSrc)) return false;

        // Determine CEP extensions folder
        let cepDir;
        if (isMac) {
            cepDir = path.join(os.homedir(), 'Library', 'Application Support', 'Adobe', 'CEP', 'extensions');
        } else {
            cepDir = path.join(process.env.APPDATA || '', 'Adobe', 'CEP', 'extensions');
        }
        const pluginDest = path.join(cepDir, extensionId);

        // Check if plugin needs update (compare manifest or just overwrite)
        if (!fs.existsSync(path.join(pluginDest, 'host', 'index.jsx')) ||
            !fs.existsSync(path.join(pluginDest, 'client', 'index.html'))) {
            // Plugin missing or incomplete — install it
        } else {
            // Check if source is newer
            try {
                const srcStat = fs.statSync(path.join(pluginSrc, 'client', 'index.html'));
                const dstStat = fs.statSync(path.join(pluginDest, 'client', 'index.html'));
                if (srcStat.mtimeMs <= dstStat.mtimeMs) return false; // Already up to date
            } catch (e) { /* install anyway */ }
        }

        // Create directories
        fs.mkdirSync(path.join(pluginDest, 'CSXS'), { recursive: true });
        fs.mkdirSync(path.join(pluginDest, 'client'), { recursive: true });
        fs.mkdirSync(path.join(pluginDest, 'host'), { recursive: true });

        // Copy files recursively
        const copyDir = (src, dst) => {
            fs.mkdirSync(dst, { recursive: true });
            for (const item of fs.readdirSync(src)) {
                const s = path.join(src, item);
                const d = path.join(dst, item);
                if (fs.statSync(s).isDirectory()) {
                    copyDir(s, d);
                } else {
                    fs.copyFileSync(s, d);
                }
            }
        };
        copyDir(path.join(pluginSrc, 'CSXS'), path.join(pluginDest, 'CSXS'));
        copyDir(path.join(pluginSrc, 'client'), path.join(pluginDest, 'client'));
        copyDir(path.join(pluginSrc, 'host'), path.join(pluginDest, 'host'));

        // Copy .debug if exists
        const debugFile = path.join(pluginSrc, '.debug');
        if (fs.existsSync(debugFile)) fs.copyFileSync(debugFile, path.join(pluginDest, '.debug'));

        console.log(extensionId + ' installed to: ' + pluginDest);
        return true;
    } catch (e) {
        console.log('Plugin auto-install skipped (' + extensionId + '):', e.message);
        return false;
    }
}

function autoInstallPremierePlugin() {
    // Check if user opted out
    const skipFlag = path.join(path.dirname(process.execPath), 'skip-plugin.flag');
    if (fs.existsSync(skipFlag)) {
        console.log('Plugin installation skipped by user choice');
        return;
    }
    // Install Premiere + After plugins
    autoInstallCepPlugin('premiere-plugin', 'com.lionworkspace.premiere');
    autoInstallCepPlugin('after-plugin',    'com.lionworkspace.after');

    // Enable CEP debug mode once (covers both plugins)
    if (isMac) {
        for (let v = 9; v <= 14; v++) {
            exec(`defaults write com.adobe.CSXS.${v} PlayerDebugMode 1`, () => {});
        }
    } else {
        for (let v = 9; v <= 14; v++) {
            exec(`reg add "HKCU\\SOFTWARE\\Adobe\\CSXS.${v}" /v PlayerDebugMode /t REG_SZ /d 1 /f`, () => {});
        }
    }
}

app.on('window-all-closed', () => {
    if (!isMac) app.quit();
});

// ═══════════════════════════════════════════════════════════════════
// LION SEARCH — globalShortcut + janela do palette
// ═══════════════════════════════════════════════════════════════════

const LION_SEARCH_SETTINGS_FILE = path.join(currentUD, 'lion-search-settings.json');
// Audio Library — user-curated audio list (replaces PC scan when enabled)
const AUDIO_LIBRARY_FILE = path.join(currentUD, 'audio-library.json');
function loadAudioLibrary() {
    try {
        const raw = fs.readFileSync(AUDIO_LIBRARY_FILE, 'utf8');
        const j = JSON.parse(raw);
        return Array.isArray(j) ? j : [];
    } catch(e) { return []; }
}
function saveAudioLibrary(lib) {
    try { fs.writeFileSync(AUDIO_LIBRARY_FILE, JSON.stringify(lib, null, 2), 'utf8'); return true; }
    catch(e) { return false; }
}
const DEFAULT_LION_HOTKEY = 'Control+Shift+L';

function loadLionSearchSettings() {
    try {
        const raw = fs.readFileSync(LION_SEARCH_SETTINGS_FILE, 'utf8');
        const j = JSON.parse(raw);
        let hotkey = j.hotkey || DEFAULT_LION_HOTKEY;
        if (hotkey === 'Alt+Space') {
            console.log('[lion-search] migrando hotkey Alt+Space → ' + DEFAULT_LION_HOTKEY);
            hotkey = DEFAULT_LION_HOTKEY;
        }
        return {
            hotkey,
            enabled: j.enabled !== false,
            sfxFolder: j.sfxFolder || '',
            useLibraryOnly: j.useLibraryOnly !== false, // default TRUE
        };
    } catch (e) {
        return { hotkey: DEFAULT_LION_HOTKEY, enabled: true, sfxFolder: '', useLibraryOnly: true };
    }
}
function saveLionSearchSettings(s) {
    try { fs.writeFileSync(LION_SEARCH_SETTINGS_FILE, JSON.stringify(s), 'utf8'); } catch(e) {}
}

// Cache do foreground process no Mac (osascript é caro — ~100ms cada call)
let _macFgCache = '';
let _macFgCacheAt = 0;
let _macFgPolling = false;

function _macFgPoll() {
    if (_macFgPolling) return;
    _macFgPolling = true;
    getFgProc((fg) => {
        _macFgCache = String(fg || '').toLowerCase();
        _macFgCacheAt = Date.now();
        _macFgPolling = false;
    });
}

// Lê foreground process name (sync no Win via fgFile, cached no Mac)
function getFgProcSync() {
    if (isWin) {
        try { return fs.readFileSync(fgFile, 'utf8').trim().toLowerCase(); }
        catch { return ''; }
    } else if (isMac) {
        // Se cache é antigo (>800ms), dispara nova call async (não bloqueia)
        if (Date.now() - _macFgCacheAt > 800) _macFgPoll();
        return _macFgCache;
    }
    return '';
}

// Verifica se o Premiere Pro está em foco
function isPremiereForeground() {
    const fg = getFgProcSync();
    if (!fg) return false;
    return /\bpremiere\b|\badobe premiere\b|\badobepremiere\b/i.test(fg);
}

// Pre-cria janela escondida no boot — abertura via hotkey é INSTANTÂNEA (não cria mais BrowserWindow)
let _lionSearchShownAt = 0; // timestamp do ultimo show() — usado pra ignorar blur que dispara durante focus race no Mac
let _mainWinWasVisible = false; // memo: estava visivel quando abrimos o search? (Mac)
function preloadLionSearchWindow() {
    if (lionSearchWin && !lionSearchWin.isDestroyed()) return;
    const width = 440, height = 88;
    lionSearchWin = new BrowserWindow({
        width, height,
        frame: false,
        transparent: true,
        backgroundColor: '#00000000',
        alwaysOnTop: true,
        resizable: false,
        skipTaskbar: true,
        show: false,
        movable: true,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
        },
    });
    lionSearchWin.setMenuBarVisibility(false);
    lionSearchWin.loadFile(path.join(__dirname, 'lion-search.html'));
    // 'closed' só dispara em close() real — não em hide()
    lionSearchWin.on('closed', () => { lionSearchWin = null; });
    // Fecha em blur — mas no Mac focus() nao consegue steal do Premiere imediato, entao a primeira blur
    // dispara logo apos show() e fecha a janela. Grace period de 600ms ignora blurs durante essa janela.
    const _hideSearchAndRestore = () => {
        try { lionSearchWin.hide(); lionSearchWin.webContents.send('lion-search:reset'); } catch(e) {}
        // Mac: nao restaura mainWin (pra nao trazer o Lion Workspace pra frente).
        // mainWin permanece escondido ate o usuario clicar no dock. _mainWinWasVisible
        // serve so de memo logico.
        _mainWinWasVisible = false;
    };
    lionSearchWin.on('blur', () => {
        const sinceShow = Date.now() - _lionSearchShownAt;
        const graceMs = isMac ? 600 : 150;
        if (sinceShow < graceMs) {
            // Reagenda check apos o grace period — se ainda nao tiver foco, fecha.
            setTimeout(() => {
                if (lionSearchWin && !lionSearchWin.isDestroyed() && !lionSearchWin.isFocused() && lionSearchWin.isVisible()) {
                    _hideSearchAndRestore();
                }
            }, graceMs - sinceShow + 50);
            return;
        }
        setTimeout(() => {
            if (lionSearchWin && !lionSearchWin.isDestroyed() && !lionSearchWin.isFocused() && lionSearchWin.isVisible()) {
                _hideSearchAndRestore();
            }
        }, 150);
    });
}

function openLionSearch(forceOpen) {
    // Garante janela existe (pré-criada no boot, mas se foi destruída por algum motivo, recria)
    if (!lionSearchWin || lionSearchWin.isDestroyed()) {
        preloadLionSearchWindow();
    }
    if (!forceOpen) {
        if (!isPremiereForeground()) {
            console.log('[lion-search] hotkey ignorado — foreground não é Premiere:', getFgProcSync());
            return;
        }
    }
    // Flag pra activate handler do Mac NÃO criar mainWin enquanto abrimos LION SEARCH
    _isOpeningLionSearch = true;
    setTimeout(() => { _isOpeningLionSearch = false; }, 1500);
    // Recalcula posição com cursor atual (em caso de monitor mudou)
    let display;
    try { display = screen.getDisplayNearestPoint(screen.getCursorScreenPoint()); }
    catch(e) { display = screen.getPrimaryDisplay(); }
    const wbounds = display.workArea;
    const width = 440, height = 88;
    const x = Math.round(wbounds.x + (wbounds.width - width) / 2);
    const y = Math.round(wbounds.y + wbounds.height / 4); // ~25% do topo
    try {
        lionSearchWin.setBounds({ x, y, width, height });
        // Reset state da página (limpa input + resultados) antes de mostrar
        try { lionSearchWin.webContents.send('lion-search:reset'); } catch(e) {}
        _lionSearchShownAt = Date.now();
        // No Mac, app.focus({steal:true}) traz TODAS as janelas do Electron — incluindo
        // a main window do Lion Workspace. Pra so trazer o search, escondemos a main
        // antes do focus steal. mainWin volta quando search fechar (no _hideSearchAndRestore).
        // Nivel 'screen-saver' coloca o search acima de tudo no Mac sem precisar de focus.
        if (isMac) {
            try {
                _mainWinWasVisible = !!(mainWin && !mainWin.isDestroyed() && mainWin.isVisible());
                if (_mainWinWasVisible) mainWin.hide();
            } catch(e) {}
            try { lionSearchWin.setAlwaysOnTop(true, 'screen-saver'); } catch(e) {}
            try { app.focus({ steal: true }); } catch(e) {}
        }
        lionSearchWin.show();
        lionSearchWin.focus();
        // Reforça foco no proximo tick — show() pode ter sido pre-empted pelo OS no Mac
        if (isMac) {
            setTimeout(() => {
                if (lionSearchWin && !lionSearchWin.isDestroyed() && lionSearchWin.isVisible() && !lionSearchWin.isFocused()) {
                    try { lionSearchWin.moveTop(); lionSearchWin.focus(); } catch(e) {}
                }
            }, 50);
        }
    } catch(e) {
        // Fallback: janela pode estar em estado ruim — recria do zero
        try { lionSearchWin.destroy(); } catch(eD) {}
        lionSearchWin = null;
        preloadLionSearchWindow();
        setTimeout(() => openLionSearch(forceOpen), 50);
    }
}

let _lionSettings = loadLionSearchSettings();
let _registeredLionHotkey = null;
let _hotkeyFgWatcherTimer = null;

function _doRegisterHotkey(hotkey) {
    if (!hotkey) return false;
    try {
        const ok = globalShortcut.register(hotkey, openLionSearch);
        if (!ok) return false;
        _registeredLionHotkey = hotkey;
        return true;
    } catch (e) { return false; }
}

function _doUnregisterHotkey() {
    if (_registeredLionHotkey) {
        try { globalShortcut.unregister(_registeredLionHotkey); } catch(e) {}
        _registeredLionHotkey = null;
    }
}

// Função reusável que checa e registra/desregistra com base no foreground atual
function _checkAndRegisterHotkey(fgRaw) {
    const fg = String(fgRaw || '').toLowerCase();
    const isPremiere = /\bpremiere\b|\badobe premiere\b|\badobepremiere\b/i.test(fg);
    const lionOpen = lionSearchWin && !lionSearchWin.isDestroyed();
    let lionFocused = false;
    try { lionFocused = lionOpen && lionSearchWin.isFocused(); } catch(e) {}

    // GUARD: fecha LION SEARCH se Premiere saiu de foco E LION SEARCH também perdeu foco
    if (lionOpen && !isPremiere && !lionFocused) {
        try { lionSearchWin.close(); } catch(e) {}
    }

    if (!_lionSettings.enabled || !_lionSettings.hotkey) {
        _doUnregisterHotkey();
        return;
    }
    const wantHotkey = _lionSettings.hotkey;
    const shouldRegister = isPremiere || lionFocused;

    if (shouldRegister) {
        // Se hotkey atual é diferente do que queremos (mudou setting), re-registra
        if (_registeredLionHotkey !== wantHotkey) {
            _doUnregisterHotkey();
            const ok = _doRegisterHotkey(wantHotkey);
            if (ok) console.log('[lion-search] hotkey registrado:', wantHotkey);
        }
    } else if (_registeredLionHotkey) {
        console.log('[lion-search] hotkey desregistrado (Premiere fora de foco) — tecla volta ao normal');
        _doUnregisterHotkey();
    }
}

// Watcher: a cada 200ms, checa se Premiere tá em foco e ajusta registro do hotkey
function startHotkeyFgWatcher() {
    if (_hotkeyFgWatcherTimer) return;
    _hotkeyFgWatcherTimer = setInterval(() => {
        if (isWin) {
            try { _checkAndRegisterHotkey(fs.readFileSync(fgFile, 'utf8').trim()); }
            catch { _checkAndRegisterHotkey(''); }
        } else if (isMac) {
            // Usa cache (atualizado em background pelo _macFgPoll)
            _checkAndRegisterHotkey(_macFgCache);
            // Mantém cache fresco
            if (Date.now() - _macFgCacheAt > 800) _macFgPoll();
        }
    }, 200);
}

function stopHotkeyFgWatcher() {
    if (_hotkeyFgWatcherTimer) { clearInterval(_hotkeyFgWatcherTimer); _hotkeyFgWatcherTimer = null; }
    _doUnregisterHotkey();
}

// Wrapper público — substitui o registerLionHotkey antigo, mas agora só CONFIGURA
// o hotkey nas settings. O watcher faz register/unregister dinamicamente.
function registerLionHotkey(hotkey) {
    if (!hotkey) return { ok: false, error: 'hotkey vazio' };
    // Se já estava registrado, força re-register pra captar mudanças
    _doUnregisterHotkey();
    if (_lionSettings.enabled) startHotkeyFgWatcher();
    return { ok: true };
}

// Registra watcher na boot
app.whenReady().then(() => {
    if (_lionSettings.enabled) startHotkeyFgWatcher();
    // Mac: warm up fg cache imediato
    if (isMac) _macFgPoll();
    // Pré-cria janela LION SEARCH (escondida) — hotkey abre INSTANTÂNEO sem criar BrowserWindow
    setTimeout(() => {
        try { preloadLionSearchWindow(); } catch(e) { console.warn('[lion-search] preload err:', e); }
    }, 800); // delay pequeno pra não competir com main window load
});
app.on('will-quit', () => {
    stopHotkeyFgWatcher();
    try { globalShortcut.unregisterAll(); } catch(e) {}
});

// IPC pra UI principal configurar hotkey + SFX folder
ipcMain.handle('lion-search:get-settings', () => _lionSettings);
ipcMain.handle('lion-search:set-settings', (event, settings) => {
    const oldHotkey = _lionSettings.hotkey;
    _lionSettings = {
        hotkey: settings?.hotkey || DEFAULT_LION_HOTKEY,
        enabled: settings?.enabled !== false,
        sfxFolder: settings?.sfxFolder !== undefined ? String(settings.sfxFolder || '') : (_lionSettings.sfxFolder || ''),
        useLibraryOnly: settings?.useLibraryOnly !== undefined ? !!settings.useLibraryOnly : (_lionSettings.useLibraryOnly !== false),
    };
    saveLionSearchSettings(_lionSettings);
    // SEMPRE força unregister do antigo quando muda settings — evita hotkey órfão
    if (oldHotkey !== _lionSettings.hotkey || !_lionSettings.enabled) {
        _doUnregisterHotkey();
    }
    if (_lionSettings.enabled) {
        startHotkeyFgWatcher();
        // Force check imediato (não espera 200ms) pra registrar logo se Premiere em foco
        try {
            if (isWin) _checkAndRegisterHotkey(fs.readFileSync(fgFile, 'utf8').trim());
            else if (isMac) getFgProc((fg) => _checkAndRegisterHotkey(fg));
        } catch(e) {}
    } else {
        stopHotkeyFgWatcher();
    }
    return { ok: true, settings: _lionSettings };
});

// Invalida cache LION SEARCH automaticamente quando lib muda
function _invalidateLionCatalogAfterChange() {
    try {
        // Reset cache em memória → próxima list-effects vai disparar refresh
        lionSearchEffectsCache = [];
        lionSearchEffectsCachedAt = 0;
        try { fs.unlinkSync(LION_CATALOG_FILE); } catch(e) {}
    } catch(e) {}
}

// ═══════ Audio Library — IPC pra UI ═══════
ipcMain.handle('audio-library:get', () => loadAudioLibrary());
ipcMain.handle('audio-library:add', (event, item) => {
    if (!item || !item.path) return { ok: false, error: 'missing_path' };
    const lib = loadAudioLibrary();
    const id = (item.id) || ('aud_' + Date.now().toString(36) + Math.random().toString(36).slice(2,6));
    const entry = {
        id: id,
        path: String(item.path),
        name: String(item.name || item.path.split(/[\\/]/).pop() || 'audio'),
        inSec: (item.inSec != null && !isNaN(item.inSec)) ? parseFloat(item.inSec) : null,
        outSec: (item.outSec != null && !isNaN(item.outSec)) ? parseFloat(item.outSec) : null,
        durationSec: item.durationSec ? parseFloat(item.durationSec) : null,
        volumeDb: (item.volumeDb != null && !isNaN(parseFloat(item.volumeDb))) ? parseFloat(item.volumeDb) : null,
        addedAt: Date.now(),
    };
    lib.push(entry);
    saveAudioLibrary(lib);
    _invalidateLionCatalogAfterChange();
    return { ok: true, item: entry };
});
ipcMain.handle('audio-library:update', (event, item) => {
    if (!item || !item.id) return { ok: false, error: 'missing_id' };
    const lib = loadAudioLibrary();
    const idx = lib.findIndex(x => x.id === item.id);
    if (idx < 0) return { ok: false, error: 'not_found' };
    // Trata null explicitamente pra trims (não pega undefined)
    const upd = Object.assign({}, lib[idx]);
    if (item.name !== undefined) upd.name = String(item.name);
    if (item.inSec !== undefined) upd.inSec = item.inSec === null ? null : parseFloat(item.inSec);
    if (item.outSec !== undefined) upd.outSec = item.outSec === null ? null : parseFloat(item.outSec);
    if (item.path !== undefined) upd.path = String(item.path);
    if (item.volumeDb !== undefined) upd.volumeDb = item.volumeDb === null ? null : parseFloat(item.volumeDb);
    if (item.durationSec !== undefined) upd.durationSec = item.durationSec === null ? null : parseFloat(item.durationSec);
    lib[idx] = upd;
    saveAudioLibrary(lib);
    _invalidateLionCatalogAfterChange();
    return { ok: true, item: lib[idx] };
});
ipcMain.handle('audio-library:remove', (event, id) => {
    if (!id) return { ok: false, error: 'missing_id' };
    const lib = loadAudioLibrary();
    const next = lib.filter(x => x.id !== id);
    saveAudioLibrary(next);
    _invalidateLionCatalogAfterChange();
    return { ok: true, count: next.length };
});
// Lê arquivo como ArrayBuffer (pra renderer fazer waveform/decode no AudioContext)
ipcMain.handle('audio-library:read-buffer', async (event, filePath) => {
    try {
        if (!filePath) return null;
        const buf = fs.readFileSync(filePath);
        // Retorna Uint8Array — Electron clona via structured clone
        return new Uint8Array(buf).buffer;
    } catch(e) { return null; }
});

// File picker pra adicionar áudios
ipcMain.handle('audio-library:pick-files', async () => {
    const { dialog } = require('electron');
    const result = await dialog.showOpenDialog({
        title: 'Selecionar áudios',
        properties: ['openFile', 'multiSelections'],
        filters: [{ name: 'Áudios', extensions: ['mp3','wav','aac','m4a','flac','ogg','aif','aiff','wma','opus'] }],
    });
    if (result.canceled) return { ok: false, canceled: true };
    return { ok: true, paths: result.filePaths || [] };
});

// IPC: dialog pra selecionar SFX folder
ipcMain.handle('lion-search:pick-sfx-folder', async () => {
    try {
        const { dialog } = require('electron');
        const result = await dialog.showOpenDialog({
            properties: ['openDirectory'],
            title: 'Selecionar pasta de Sound Effects',
        });
        if (result.canceled || !result.filePaths || !result.filePaths.length) {
            return { ok: false, canceled: true };
        }
        return { ok: true, folder: result.filePaths[0] };
    } catch (e) { return { ok: false, error: e.message }; }
});
ipcMain.handle('lion-search:open', () => { openLionSearch(true); return { ok: true }; }); // testar manual = bypass FG check

// IPC: resize da janela (chamado pela lion-search.html quando troca empty/typing)
ipcMain.on('lion-search:resize', (event, payload) => {
    try {
        if (!lionSearchWin || lionSearchWin.isDestroyed()) return;
        const w = Math.max(280, Math.min(800, payload?.w || 440));
        const h = Math.max(60, Math.min(800, payload?.h || 88));
        const [curX, curY] = lionSearchWin.getPosition();
        // Temporariamente permite resize (alguns Electron travam setBounds com resizable:false)
        const wasResizable = lionSearchWin.isResizable();
        try { lionSearchWin.setResizable(true); } catch(e) {}
        // setBounds funciona melhor que setSize em janelas frame:false transparent:true
        lionSearchWin.setBounds({ x: curX, y: curY, width: w, height: h }, false);
        try { if (!wasResizable) lionSearchWin.setResizable(false); } catch(e) {}
    } catch(e) { console.warn('[lion-search] resize fail:', e); }
});

// macOS: re-create window when clicking dock icon
// CRÍTICO: ignora activate quando LION SEARCH está abrindo (senão abre janela principal
// quando user só queria o palette via hotkey)
let _isOpeningLionSearch = false;
if (isMac) {
    app.on('activate', () => {
        // Skip se LION SEARCH foi acabou de ser disparada via hotkey
        if (_isOpeningLionSearch) return;
        // Skip se está em meio a um apply de LION SEARCH (audio/preset)
        if (_isApplyInProgress) return;
        // Skip se está mostrando notificação (notificação ativa o app automaticamente)
        if (_isNotifying) return;
        // Skip se LION SEARCH VISÍVEL (não só pre-criada)
        if (lionSearchWin && !lionSearchWin.isDestroyed() && lionSearchWin.isVisible()) return;
        // Só recria mainWin se foi explicitamente fechada E user clicou no dock
        if (mainWin && !mainWin.isDestroyed() && !mainWin.isVisible()) {
            mainWin.show();
        } else if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
}
