const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('focusMode', {
    start: () => ipcRenderer.invoke('focus-start'),
    startExternal: () => ipcRenderer.invoke('focus-start-external'),
    stop: () => ipcRenderer.invoke('focus-stop'),
    isActive: () => ipcRenderer.invoke('focus-is-active'),
    launchApp: (name) => ipcRenderer.invoke('launch-app', name),
    openFolder: () => ipcRenderer.invoke('open-folder')
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
