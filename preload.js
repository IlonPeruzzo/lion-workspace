const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('focusMode', {
    start: () => ipcRenderer.invoke('focus-start'),
    startExternal: () => ipcRenderer.invoke('focus-start-external'),
    stop: () => ipcRenderer.invoke('focus-stop'),
    isActive: () => ipcRenderer.invoke('focus-is-active'),
    launchApp: (name) => ipcRenderer.invoke('launch-app', name),
    openFolder: () => ipcRenderer.invoke('open-folder'),
    setWhitelist: (list) => ipcRenderer.invoke('focus-set-whitelist', list),
    listProcesses: () => ipcRenderer.invoke('list-running-processes'),
    launchByPath: (p) => ipcRenderer.invoke('launch-by-path', p),
    findAppPath: (name) => ipcRenderer.invoke('find-app-path', name)
});

contextBridge.exposeInMainWorld('electronAPI', {
    openExternal: (url) => ipcRenderer.invoke('open-external', url)
});

contextBridge.exposeInMainWorld('ytDownloader', {
    ensureDeps: () => ipcRenderer.invoke('yt-ensure-deps'),
    getInfo: (url) => ipcRenderer.invoke('yt-get-info', url),
    download: (opts) => ipcRenderer.invoke('yt-download', opts),
    cancel: () => ipcRenderer.invoke('yt-cancel'),
    getProgress: () => ipcRenderer.invoke('yt-progress'),
    openFolder: (dir) => ipcRenderer.invoke('yt-open-folder', dir),
    onProgress: (cb) => ipcRenderer.on('yt-progress', (event, data) => cb(data))
});

contextBridge.exposeInMainWorld('syncBridge', {
    pushState: (state) => ipcRenderer.invoke('sync-push-state', state),
    onCommand: (callback) => ipcRenderer.on('sync-command', (event, data) => callback(data)),
    startTimerFg: () => ipcRenderer.invoke('timer-fg-start'),
    stopTimerFg: () => ipcRenderer.invoke('timer-fg-stop'),
    onTimerFgPause: (cb) => ipcRenderer.on('timer-fg-pause', () => cb()),
    onTimerFgResume: (cb) => ipcRenderer.on('timer-fg-resume', () => cb()),
    isInWorkApp: () => ipcRenderer.invoke('check-foreground')
});
