const { app, BrowserWindow, ipcMain, screen, powerMonitor, Menu } = require('electron');
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

    // When main window closes, kill focus mode so overlays don't keep the app alive
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

        if (!isWorkApp(procName)) {
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

ipcMain.handle('check-foreground', () => {
    return new Promise(resolve => {
        getFgProc(proc => resolve(proc ? isWorkApp(proc) : false));
    });
});

ipcMain.handle('timer-fg-start', () => {
    timerFgPaused = false;
    fgMissCount = 0;
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
            'vegas': 'echo "Vegas Pro nao disponivel no macOS"',
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

// AutoCut progress tracking (polled by plugin)
let autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
let autoCutProc = null; // current FFmpeg process (so we can cancel)

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

async function ensureFfmpeg() {
    if (ffmpegReady()) return true;

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
        if (fs.existsSync(ffmpegBin)) return true;
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
                if (fs.existsSync(ffmpegBin)) return true;
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
            return fs.existsSync(ffmpegBin);
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
        return fs.existsSync(ffmpegBin);
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

ipcMain.handle('yt-ensure-deps', async () => {
    const yt = await ensureYtDlp();
    const ff = ffmpegReady() || await ensureFfmpeg();
    return { ytdlp: yt, ffmpeg: ff };
});

ipcMain.handle('yt-get-info', async (event, url) => {
    url = cleanYtUrl(url);
    if (!ytDlpReady()) { const ok = await ensureYtDlp(); if (!ok) return { error: 'yt-dlp não disponível' }; }
    return new Promise(resolve => {
        const args = ['--no-download', '--print-json', '--no-warnings', '--no-playlist'];
        if (ffmpegReady()) args.push('--ffmpeg-location', path.dirname(ffmpegBin));
        args.push(url);
        const proc = spawn(ytDlpBin, args, { windowsHide: true });
        let out = '', errOut = '';
        proc.stdout.on('data', d => out += d);
        proc.stderr.on('data', d => errOut += d);
        proc.on('close', code => {
            if (code !== 0) { resolve({ error: errOut || 'Erro ao obter info' }); return; }
            try {
                const info = JSON.parse(out);
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
                resolve({
                    title: info.title || '',
                    thumbnail: info.thumbnail || '',
                    duration: info.duration || 0,
                    uploader: info.uploader || '',
                    formats: formats
                });
            } catch (e) { resolve({ error: 'Erro ao parsear info' }); }
        });
        proc.on('error', e => resolve({ error: e.message }));
        setTimeout(() => { try { proc.kill(); } catch {} }, 30000);
    });
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
        // 4K em geral só vem em VP9/AV1 no YouTube; H.264 raramente existe acima de 1080p.
        // Por isso pra 4K pulamos o filtro avc1 e deixamos pegar o melhor disponível.
        if (format === '4k') {
            args.push('-f', 'bestvideo[height<=2160]+bestaudio/best[height<=2160]/bestvideo+bestaudio/best');
        } else if (format === '1080') {
            args.push('-f', 'bestvideo[height<=1080][vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo[height<=1080]+bestaudio/best[height<=1080]/best');
        } else if (format === '720') {
            args.push('-f', 'bestvideo[height<=720][vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo[height<=720]+bestaudio/best[height<=720]/best');
        } else {
            args.push('-f', 'bestvideo[vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo+bestaudio/best/best');
        }
        // Sort prefers H.264 quando disponível, mas não exige
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

    ytDownloadProc = spawn(ytDlpBin, args, { windowsHide: true });
    let lastFile = '';

    ytDownloadProc.stdout.on('data', d => {
        const lines = d.toString().split('\n');
        for (const line of lines) {
            // [download]  45.2% of  120.50MiB at  5.32MiB/s ETA 00:12
            const m = line.match(/\[download\]\s+([\d.]+)%\s+of\s+~?([\d.]+\S+)\s+at\s+([\d.]+\S+)\s+ETA\s+(\S+)/);
            if (m) {
                ytProgress.percent = parseFloat(m[1]);
                ytProgress.speed = m[3];
                ytProgress.eta = m[4];
            }
            // [download] Destination: /path/to/file.mp4
            const dm = line.match(/\[download\] Destination: (.+)/);
            if (dm) lastFile = dm[1].trim();
            // [Merger] Merging formats into "path/to/final.mp4" — this is the REAL final file
            const mm = line.match(/\[Merger\] Merging formats into "(.+)"/);
            if (mm) lastFile = mm[1].trim();
            // [ExtractAudio] Destination: path/to/file.mp3
            const am = line.match(/\[ExtractAudio\] Destination: (.+)/);
            if (am) lastFile = am[1].trim();
            // [Merger] or [ExtractAudio] status
            if (line.includes('[Merger]') || line.includes('[ExtractAudio]')) {
                ytProgress.status = 'merging';
                ytProgress.percent = 99;
            }
            // Already downloaded
            if (line.includes('has already been downloaded')) {
                ytProgress.percent = 100;
                ytProgress.status = 'done';
            }
        }
        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
    });

    ytDownloadProc.stderr.on('data', d => {
        const err = d.toString().trim();
        if (err) ytProgress.error = err;
    });

    ytDownloadProc.on('close', code => {
        ytDownloadProc = null;
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
            if (!ytProgress.error) ytProgress.error = 'Download falhou (código ' + code + ')';
        }
        if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
    });

    ytDownloadProc.on('error', e => {
        ytDownloadProc = null;
        ytProgress.status = 'error';
        ytProgress.active = false;
        ytProgress.error = e.message;
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

// Plugin auth tokens (persisted to disk so they survive app restarts)
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

        // YouTube downloader HTTP endpoints (for plugin)
        if (req.method === 'GET' && req.url === '/yt/progress') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(ytProgress));
            return;
        }

        // AutoCut progress polling (used by plugin for progress bar)
        if (req.method === 'GET' && req.url === '/autocut/progress') {
            res.writeHead(200, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify(autoCutProgress));
            return;
        }

        if (req.method === 'POST') {
            let body = '';
            req.on('data', chunk => body += chunk);
            req.on('end', async () => {
                let data = {};
                try { data = JSON.parse(body || '{}'); } catch (e) {}
                const command = req.url.replace(/^\//, '');

                // YouTube downloader HTTP API
                if (command === 'yt/info' && data.url) {
                    try {
                        const cleanUrl = cleanYtUrl(data.url);
                        await ensureYtDlp();
                        const result = await new Promise(resolve => {
                            const args = ['--no-download', '--print-json', '--no-warnings', '--no-playlist'];
                            if (ffmpegReady()) args.push('--ffmpeg-location', path.dirname(ffmpegBin));
                            args.push(cleanUrl);
                            const proc = spawn(ytDlpBin, args, { windowsHide: true });
                            let out = '', errOut = '';
                            proc.stdout.on('data', d => out += d);
                            proc.stderr.on('data', d => errOut += d);
                            proc.on('close', code => {
                                if (code !== 0) { resolve({ error: errOut || 'Erro' }); return; }
                                try {
                                    const info = JSON.parse(out);
                                    resolve({ title: info.title || '', thumbnail: info.thumbnail || '', duration: info.duration || 0, uploader: info.uploader || '' });
                                } catch { resolve({ error: 'Parse error' }); }
                            });
                            proc.on('error', e => resolve({ error: e.message }));
                            setTimeout(() => { try { proc.kill(); } catch {} }, 30000);
                        });
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(result));
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
                            // 4K raramente vem em H.264 — não exigir avc1 ou falha em maioria das plataformas
                            if (fmt === '4k') dlArgs.push('-f', 'bestvideo[height<=2160]+bestaudio/best[height<=2160]/bestvideo+bestaudio/best');
                            else if (fmt === '1080') dlArgs.push('-f', 'bestvideo[height<=1080][vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo[height<=1080]+bestaudio/best[height<=1080]/best');
                            else if (fmt === '720') dlArgs.push('-f', 'bestvideo[height<=720][vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo[height<=720]+bestaudio/best[height<=720]/best');
                            else dlArgs.push('-f', 'bestvideo[vcodec^=avc1]+bestaudio[acodec^=mp4a]/bestvideo+bestaudio/best/best');
                            dlArgs.push('-S', 'res,codec:h264:m4a,br', '--merge-output-format', 'mp4');
                        } else {
                            if (fmt === '4k') dlArgs.push('-f', 'best[height<=2160][ext=mp4]/best[height<=2160]/best[ext=mp4]/best');
                            else if (fmt === '1080') dlArgs.push('-f', 'best[height<=1080][ext=mp4]/best[height<=1080]/best');
                            else if (fmt === '720') dlArgs.push('-f', 'best[height<=720][ext=mp4]/best[height<=720]/best');
                            else dlArgs.push('-f', 'best[ext=mp4]/best');
                        }
                        dlArgs.push(cleanDlUrl);

                        ytProgress = { active: true, percent: 0, speed: '', eta: '', title: data.metaTitle||'', status: 'downloading', error: '', outputDir: outputPath, metaTitle: data.metaTitle||'', metaThumb: data.metaThumb||'', metaUploader: data.metaUploader||'', metaDuration: data.metaDuration||0, metaUrl: data.url||'' };
                        ytDownloadProc = spawn(ytDlpBin, dlArgs, { windowsHide: true });
                        let lastFile = '';
                        ytDownloadProc.stdout.on('data', d => {
                            const lines = d.toString().split('\n');
                            for (const line of lines) {
                                const m = line.match(/\[download\]\s+([\d.]+)%\s+of\s+~?([\d.]+\S+)\s+at\s+([\d.]+\S+)\s+ETA\s+(\S+)/);
                                if (m) { ytProgress.percent = parseFloat(m[1]); ytProgress.speed = m[3]; ytProgress.eta = m[4]; }
                                const dm = line.match(/\[download\] Destination: (.+)/);
                                if (dm) lastFile = dm[1].trim();
                                // [Merger] Merging formats into "path/to/final.mp4" — REAL final file
                                const mm = line.match(/\[Merger\] Merging formats into "(.+)"/);
                                if (mm) lastFile = mm[1].trim();
                                // [ExtractAudio] Destination: path/to/file.mp3
                                const am = line.match(/\[ExtractAudio\] Destination: (.+)/);
                                if (am) lastFile = am[1].trim();
                                if (line.includes('[Merger]') || line.includes('[ExtractAudio]')) { ytProgress.status = 'merging'; ytProgress.percent = 99; }
                                if (line.includes('has already been downloaded')) { ytProgress.percent = 100; ytProgress.status = 'done'; }
                            }
                            if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                        });
                        ytDownloadProc.stderr.on('data', d => { const err = d.toString().trim(); if (err) ytProgress.error = err; });
                        ytDownloadProc.on('close', code => {
                            ytDownloadProc = null;
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
                                    return; // Don't emit done yet, wait for ffmpeg
                                }

                                ytProgress.percent = 100; ytProgress.status = 'done'; ytProgress.active = false; ytProgress.file = lastFile;
                            } else {
                                ytProgress.status = 'error'; ytProgress.active = false;
                                if (!ytProgress.error) ytProgress.error = 'Falhou (código ' + code + ')';
                            }
                            if (mainWin && !mainWin.isDestroyed()) mainWin.webContents.send('yt-progress', ytProgress);
                        });
                        ytDownloadProc.on('error', e => { ytDownloadProc = null; ytProgress.status = 'error'; ytProgress.active = false; ytProgress.error = e.message; });
                    } catch (e) { ytProgress.status = 'error'; ytProgress.error = e.message; }
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }
                // ═══════ AutoCut — Silence Detection via FFmpeg (with progress) ═══════
                if (command === 'autocut/analyze' && data.filePath) {
                    try {
                        if (!ffmpegReady()) await ensureFfmpeg();
                        if (!ffmpegReady()) {
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'FFmpeg não disponível' }));
                            return;
                        }
                        const threshold = data.threshold || -30;
                        const minDuration = data.minDuration || 0.5;
                        const filePath = data.filePath;

                        if (!fs.existsSync(filePath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Arquivo não encontrado: ' + filePath }));
                            return;
                        }

                        // Reset progress — plugin polls /autocut/progress
                        autoCutProgress = { active: true, percent: 0, phase: 'silence', eta: '' };

                        const ffArgs = [
                            '-nostdin',
                            '-i', filePath,
                            '-vn',  // skip video decoding (HUGE speedup)
                            '-af', `silencedetect=noise=${threshold}dB:d=${minDuration}`,
                            '-f', 'null', '-'
                        ];

                        const result = await new Promise((resolve) => {
                            const proc = spawn(ffmpegBin, ffArgs, { windowsHide: true });
                            autoCutProc = proc;
                            let stderrFull = '';
                            let totalDuration = 0;
                            let durationParsed = false;

                            proc.stderr.on('data', d => {
                                const chunk = d.toString();
                                stderrFull += chunk;

                                // Parse Duration on first occurrence
                                if (!durationParsed) {
                                    const durMatch = stderrFull.match(/Duration:\s*(\d+):(\d+):([\d.]+)/);
                                    if (durMatch) {
                                        totalDuration = parseInt(durMatch[1]) * 3600 + parseInt(durMatch[2]) * 60 + parseFloat(durMatch[3]);
                                        durationParsed = true;
                                    }
                                }

                                // Parse time= progress from FFmpeg stderr
                                const timeMatch = chunk.match(/time=\s*(\d+):(\d+):([\d.]+)/);
                                if (timeMatch && totalDuration > 0) {
                                    const currentTime = parseInt(timeMatch[1]) * 3600 + parseInt(timeMatch[2]) * 60 + parseFloat(timeMatch[3]);
                                    autoCutProgress.percent = Math.min(99, Math.round((currentTime / totalDuration) * 100));
                                }
                            });

                            proc.on('close', () => {
                                autoCutProc = null;
                                autoCutProgress = { active: false, percent: 100, phase: 'done', eta: '' };

                                // Parse silences from full stderr
                                const silences = [];
                                const lines = stderrFull.split('\n');
                                let currentStart = null;
                                for (const line of lines) {
                                    const startMatch = line.match(/silence_start:\s*([\d.]+)/);
                                    if (startMatch) currentStart = parseFloat(startMatch[1]);
                                    const endMatch = line.match(/silence_end:\s*([\d.]+)\s*\|\s*silence_duration:\s*([\d.]+)/);
                                    if (endMatch) {
                                        silences.push({
                                            start: currentStart !== null ? currentStart : 0,
                                            end: parseFloat(endMatch[1]),
                                            duration: parseFloat(endMatch[2])
                                        });
                                        currentStart = null;
                                    }
                                }

                                const durMatch = stderrFull.match(/Duration:\s*(\d+):(\d+):([\d.]+)/);
                                let dur = 0;
                                if (durMatch) dur = parseInt(durMatch[1]) * 3600 + parseInt(durMatch[2]) * 60 + parseFloat(durMatch[3]);
                                resolve({ silences, totalDuration: dur || totalDuration });
                            });

                            proc.on('error', e => {
                                autoCutProc = null;
                                autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                                resolve({ silences: [], totalDuration: 0, error: e.message });
                            });

                            setTimeout(() => {
                                try { proc.kill(); } catch {}
                                autoCutProc = null;
                                autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                                resolve({ silences: [], totalDuration: 0, error: 'Timeout (10min)' });
                            }, 600000); // 10 minutes for large files
                        });

                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify(result));
                    } catch (e) {
                        autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message }));
                    }
                    return;
                }

                // ═══════ AutoCut — AI Cancel (abort running process) ═══════
                if (command === 'autocut/cancel') {
                    if (autoCutProc) { try { autoCutProc.kill(); } catch {} autoCutProc = null; }
                    autoCutProgress = { active: false, percent: 0, phase: 'cancelled', eta: '' };
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ ok: true }));
                    return;
                }

                // ═══════ AutoCut — AI Key management ═══════
                if (command === 'autocut/ai-key') {
                    if (data.key !== undefined) {
                        const cfg = getAiConfig();
                        cfg.openaiKey = data.key || '';
                        saveAiConfig(cfg);
                    }
                    const cfg = getAiConfig();
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ hasKey: !!(cfg.openaiKey), masked: cfg.openaiKey ? ('sk-...' + cfg.openaiKey.slice(-4)) : '' }));
                    return;
                }

                // ═══════ AutoCut — AI Transcription + Word Analysis ═══════
                if (command === 'autocut/transcribe' && data.filePath) {
                    try {
                        // Get API key from request or stored config
                        const apiKey = data.apiKey || (getAiConfig().openaiKey || '');
                        if (!apiKey) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Chave da API OpenAI não configurada. Clique em "API Key" para configurar.' }));
                            return;
                        }

                        if (!ffmpegReady()) await ensureFfmpeg();
                        if (!ffmpegReady()) {
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'FFmpeg não disponível' }));
                            return;
                        }

                        const filePath = data.filePath;
                        if (!fs.existsSync(filePath)) {
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Arquivo não encontrado' }));
                            return;
                        }

                        // ── Phase 1: Extract audio as MP3 (64kbps mono 16kHz = ~0.5MB/min) ──
                        autoCutProgress = { active: true, percent: 5, phase: 'transcribe-extract', eta: '' };
                        const tmpAudio = path.join(os.tmpdir(), 'lw-ac-audio-' + Date.now() + '.mp3');

                        await new Promise((resolve, reject) => {
                            const args = ['-i', filePath, '-vn', '-acodec', 'libmp3lame', '-b:a', '64k', '-ac', '1', '-ar', '16000', '-y', tmpAudio];
                            const proc = spawn(ffmpegBin, args, { windowsHide: true });
                            autoCutProc = proc;
                            let extractDur = 0;
                            let extractDurParsed = false;
                            proc.stderr.on('data', d => {
                                const chunk = d.toString();
                                if (!extractDurParsed) {
                                    const dm = chunk.match(/Duration:\s*(\d+):(\d+):([\d.]+)/);
                                    if (dm) { extractDur = parseInt(dm[1]) * 3600 + parseInt(dm[2]) * 60 + parseFloat(dm[3]); extractDurParsed = true; }
                                }
                                const tm = chunk.match(/time=\s*(\d+):(\d+):([\d.]+)/);
                                if (tm && extractDur > 0) {
                                    const ct = parseInt(tm[1]) * 3600 + parseInt(tm[2]) * 60 + parseFloat(tm[3]);
                                    autoCutProgress.percent = 5 + Math.round((ct / extractDur) * 20); // 5-25%
                                }
                            });
                            proc.on('close', code => { autoCutProc = null; code === 0 ? resolve() : reject(new Error('FFmpeg audio extract falhou (code ' + code + ')')); });
                            proc.on('error', e => { autoCutProc = null; reject(e); });
                            setTimeout(() => { try { proc.kill(); } catch {} autoCutProc = null; reject(new Error('Timeout extraindo áudio')); }, 120000);
                        });

                        // Check file size (Whisper API limit: 25MB)
                        const audioStats = fs.statSync(tmpAudio);
                        const audioSizeMB = Math.round(audioStats.size / (1024 * 1024));
                        if (audioStats.size > 25 * 1024 * 1024) {
                            safeRm(tmpAudio);
                            autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Áudio muito grande (' + audioSizeMB + 'MB). Máximo: 25MB. Use um clip menor.' }));
                            return;
                        }

                        // ── Phase 2: Send to OpenAI Whisper API ──
                        autoCutProgress = { active: true, percent: 28, phase: 'transcribe-api', eta: '' };

                        const audioData = fs.readFileSync(tmpAudio);
                        safeRm(tmpAudio);

                        const boundary = '----LWWhisper' + Date.now() + crypto.randomBytes(8).toString('hex');
                        const lang = data.language || '';

                        // Build multipart form data
                        const parts = [];
                        parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="audio.mp3"\r\nContent-Type: audio/mpeg\r\n\r\n`));
                        parts.push(audioData);
                        parts.push(Buffer.from(`\r\n--${boundary}\r\nContent-Disposition: form-data; name="model"\r\n\r\nwhisper-1\r\n`));
                        parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="response_format"\r\n\r\nverbose_json\r\n`));
                        parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="timestamp_granularities[]"\r\n\r\nword\r\n`));
                        if (lang) parts.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="language"\r\n\r\n${lang}\r\n`));
                        parts.push(Buffer.from(`--${boundary}--\r\n`));
                        const fullBody = Buffer.concat(parts);

                        // Simulate progress during API call (we can't know real %)
                        let apiPct = 28;
                        const apiProgressTimer = setInterval(() => {
                            if (apiPct < 78) { apiPct += 2; autoCutProgress.percent = apiPct; }
                        }, 1500);

                        let whisperResult;
                        try {
                            whisperResult = await new Promise((resolve, reject) => {
                                const options = {
                                    hostname: 'api.openai.com',
                                    path: '/v1/audio/transcriptions',
                                    method: 'POST',
                                    headers: {
                                        'Authorization': 'Bearer ' + apiKey,
                                        'Content-Type': 'multipart/form-data; boundary=' + boundary,
                                        'Content-Length': fullBody.length
                                    }
                                };
                                let settled = false;
                                let timeoutId = null;
                                const done = (fn, val) => { if (settled) return; settled = true; if (timeoutId) clearTimeout(timeoutId); fn(val); };
                                const httpsReq = require('https').request(options, (httpsRes) => {
                                    let resData = '';
                                    httpsRes.on('data', chunk => resData += chunk);
                                    httpsRes.on('end', () => {
                                        try { done(resolve, JSON.parse(resData)); }
                                        catch (e) { done(reject, new Error('Whisper API parse error: ' + resData.substring(0, 300))); }
                                    });
                                });
                                httpsReq.on('error', e => done(reject, e));
                                httpsReq.write(fullBody);
                                httpsReq.end();
                                timeoutId = setTimeout(() => {
                                    try { httpsReq.destroy(); } catch(e) {}
                                    done(reject, new Error('Whisper API timeout (5min)'));
                                }, 300000);
                            });
                        } finally {
                            // ALWAYS clear the progress ticker, even on error/timeout
                            clearInterval(apiProgressTimer);
                        }

                        // Check API errors
                        if (whisperResult.error) {
                            autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                            res.writeHead(400, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Whisper API: ' + (whisperResult.error.message || JSON.stringify(whisperResult.error)) }));
                            return;
                        }

                        // ── Phase 3: Analyze transcript for repeated words / fillers / stuttering ──
                        autoCutProgress = { active: true, percent: 82, phase: 'transcribe-analyze', eta: '' };

                        const words = whisperResult.words || [];
                        const segments = whisperResult.segments || [];
                        const fullText = whisperResult.text || '';

                        // Filler words (PT + EN)
                        const FILLERS_PT = ['é', 'ah', 'ãh', 'eh', 'éh', 'uh', 'uhm', 'hum', 'hmm', 'hã',
                            'tipo', 'né', 'então', 'assim', 'basicamente', 'literalmente', 'enfim',
                            'bom', 'olha', 'veja', 'cara', 'mano', 'tá', 'ok', 'certo', 'pronto',
                            'aí', 'daí', 'pois', 'sabe', 'entendeu', 'tá ligado', 'meu'];
                        const FILLERS_EN = ['uh', 'um', 'uhm', 'hmm', 'like', 'basically', 'literally',
                            'actually', 'so', 'well', 'right', 'okay', 'oh'];
                        const FILLERS = new Set([...FILLERS_PT, ...FILLERS_EN]);

                        const issues = [];

                        for (let i = 0; i < words.length; i++) {
                            const w = words[i];
                            const wClean = (w.word || '').trim().toLowerCase().replace(/[.,!?;:'"()[\]{}]/g, '');
                            if (!wClean || wClean.length < 1) continue;

                            // ─ Check filler words ─
                            if (FILLERS.has(wClean)) {
                                issues.push({
                                    type: 'filler', word: w.word.trim(),
                                    start: w.start, end: w.end,
                                    label: 'Filler: "' + w.word.trim() + '"'
                                });
                            }

                            // ─ Check consecutive repeated words ─
                            if (i > 0) {
                                const prevClean = (words[i - 1].word || '').trim().toLowerCase().replace(/[.,!?;:'"()[\]{}]/g, '');
                                if (wClean === prevClean && wClean.length > 1) {
                                    // Find extent of the repetition group
                                    let groupStart = i - 1;
                                    while (groupStart > 0) {
                                        const earlier = (words[groupStart - 1].word || '').trim().toLowerCase().replace(/[.,!?;:'"()[\]{}]/g, '');
                                        if (earlier === wClean) groupStart--;
                                        else break;
                                    }
                                    // Only add at the last word of the group (avoid duplicates)
                                    const nextClean = (i + 1 < words.length) ? (words[i + 1].word || '').trim().toLowerCase().replace(/[.,!?;:'"()[\]{}]/g, '') : '';
                                    if (nextClean !== wClean) {
                                        const count = i - groupStart + 1;
                                        issues.push({
                                            type: 'repeat', word: w.word.trim(),
                                            start: words[groupStart].start, end: w.end,
                                            count: count,
                                            label: '"' + w.word.trim() + '" repetido ' + count + 'x'
                                        });
                                    }
                                }
                            }

                            // ─ Check stuttering (partial word repetition) ─
                            if (i > 0 && wClean.length >= 3) {
                                const prevClean = (words[i - 1].word || '').trim().toLowerCase().replace(/[.,!?;:'"()[\]{}]/g, '');
                                if (prevClean.length >= 1 && prevClean.length <= 3 && prevClean !== wClean) {
                                    // Previous is a short fragment that starts the same as current word
                                    if (wClean.startsWith(prevClean)) {
                                        issues.push({
                                            type: 'stutter',
                                            word: words[i - 1].word.trim() + ' ' + w.word.trim(),
                                            start: words[i - 1].start, end: w.end,
                                            label: 'Gaguejo: "' + words[i - 1].word.trim() + ' ' + w.word.trim() + '"'
                                        });
                                    }
                                }
                            }
                        }

                        autoCutProgress = { active: false, percent: 100, phase: 'done', eta: '' };

                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({
                            text: fullText,
                            words: words,
                            segments: segments,
                            issues: issues,
                            totalDuration: whisperResult.duration || 0,
                            language: whisperResult.language || lang || 'auto'
                        }));

                    } catch (e) {
                        autoCutProgress = { active: false, percent: 0, phase: '', eta: '' };
                        res.writeHead(500, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: e.message }));
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

    // Leak prevention: close idle sockets so the plugin reconnecting doesn't accumulate them
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

        return { adobeInstalled, pluginInstalled };
    } catch {
        return { adobeInstalled: false, pluginInstalled: false };
    }
});

// IPC: Install the Premiere plugin now
ipcMain.handle('install-premiere-plugin', async () => {
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

        autoInstallPremierePlugin();
        return { success: true };
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
    // ── Clean up focus mode so it doesn't run forever ──
    if (focusActive) {
        focusActive = false;
        focusExternal = false;
        if (focusInterval) { clearInterval(focusInterval); focusInterval = null; }
        destroyOverlays();
    }
});

app.whenReady().then(() => {
    // Set up Edit menu so Ctrl+V / Cmd+V works in all input fields
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
function autoInstallPremierePlugin() {
    try {
        // Check if user opted out of plugin during installation
        const skipFlag = path.join(path.dirname(process.execPath), 'skip-plugin.flag');
        if (fs.existsSync(skipFlag)) {
            console.log('Plugin installation skipped by user choice');
            return;
        }

        // Find the plugin source (bundled with app in extraResources)
        let pluginSrc = path.join(process.resourcesPath, 'premiere-plugin');
        if (!fs.existsSync(pluginSrc)) {
            // Dev mode: plugin is next to the app
            pluginSrc = path.join(__dirname, 'premiere-plugin');
        }
        if (!fs.existsSync(pluginSrc)) return;

        // Determine CEP extensions folder
        let cepDir;
        if (isMac) {
            cepDir = path.join(os.homedir(), 'Library', 'Application Support', 'Adobe', 'CEP', 'extensions');
        } else {
            cepDir = path.join(process.env.APPDATA || '', 'Adobe', 'CEP', 'extensions');
        }
        const pluginDest = path.join(cepDir, 'com.lionworkspace.premiere');

        // Check if plugin needs update (compare manifest or just overwrite)
        if (!fs.existsSync(path.join(pluginDest, 'host', 'index.jsx')) ||
            !fs.existsSync(path.join(pluginDest, 'client', 'index.html'))) {
            // Plugin missing or incomplete — install it
        } else {
            // Check if source is newer
            try {
                const srcStat = fs.statSync(path.join(pluginSrc, 'client', 'index.html'));
                const dstStat = fs.statSync(path.join(pluginDest, 'client', 'index.html'));
                if (srcStat.mtimeMs <= dstStat.mtimeMs) return; // Already up to date
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

        // Enable CEP debug mode (unsigned extensions)
        if (isMac) {
            for (let v = 9; v <= 14; v++) {
                exec(`defaults write com.adobe.CSXS.${v} PlayerDebugMode 1`, () => {});
            }
        } else {
            for (let v = 9; v <= 14; v++) {
                exec(`reg add "HKCU\\SOFTWARE\\Adobe\\CSXS.${v}" /v PlayerDebugMode /t REG_SZ /d 1 /f`, () => {});
            }
        }

        console.log('Premiere plugin installed to:', pluginDest);
    } catch (e) {
        console.log('Plugin auto-install skipped:', e.message);
    }
}

app.on('window-all-closed', () => {
    if (!isMac) app.quit();
});

// macOS: re-create window when clicking dock icon
if (isMac) {
    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
}
