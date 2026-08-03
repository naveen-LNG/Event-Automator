const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getConfig: () => ipcRenderer.invoke('config:get'),
  saveConfig: (partial) => ipcRenderer.invoke('config:save', partial),
  browseFolder: (startPath) => ipcRenderer.invoke('dialog:browseFolder', startPath),
  browseFile: (opts) => ipcRenderer.invoke('dialog:browseFile', opts),
  validatePath: (p) => ipcRenderer.invoke('path:validate', p),
  resolveDesignPath: (branchRoot, workspace) => ipcRenderer.invoke('path:resolveDesign', { branchRoot, workspace }),

  runStep: (params) => ipcRenderer.invoke('pipeline:runStep', params),
  onLogLine: (cb) => ipcRenderer.on('pipeline:log-line', (evt, line) => cb(line)),

  minigamesGet: (params) => ipcRenderer.invoke('minigames:get', params),
  minigamesCommit: (params) => ipcRenderer.invoke('minigames:commit', params),
  bpGet: (params) => ipcRenderer.invoke('bp:get', params),
  bpCommit: (params) => ipcRenderer.invoke('bp:commit', params),
  hbpGet: (params) => ipcRenderer.invoke('hbp:get', params),
  hbpCommit: (params) => ipcRenderer.invoke('hbp:commit', params),
  decoShopCommit: (params) => ipcRenderer.invoke('decoshop:commit', params),
  residentTaskCommit: (params) => ipcRenderer.invoke('residenttask:commit', params),
  weeklyGiftCommit: (params) => ipcRenderer.invoke('weeklygift:commit', params),
  fullScreenGiftGet: (params) => ipcRenderer.invoke('fullscreengift:get', params),
  fullScreenGiftCommit: (params) => ipcRenderer.invoke('fullscreengift:commit', params),

  eventTypesGet: (params) => ipcRenderer.invoke('eventtypes:get', params),

  refreshConvertLayoutCache: (params) => ipcRenderer.invoke('convertLayoutCache:refresh', params),
  refreshAllCaches: (params) => ipcRenderer.invoke('caches:refreshAll', params),

  triviaGet: (params) => ipcRenderer.invoke('trivia:get', params),
  triviaRun: (params) => ipcRenderer.invoke('trivia:run', params),

  endlessTreasureGet: (params) => ipcRenderer.invoke('endlesstreasure:get', params),
  endlessTreasureCommit: (params) => ipcRenderer.invoke('endlesstreasure:commit', params)
});
