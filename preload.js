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
  robotPvpCommit: (params) => ipcRenderer.invoke('robotpvp:commit', params),
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
  endlessTreasureCommit: (params) => ipcRenderer.invoke('endlesstreasure:commit', params),

  chapterGet: (params) => ipcRenderer.invoke('chapter:get', params),
  chapterRun: (params) => ipcRenderer.invoke('chapter:run', params),

  ideationGenerate: (params) => ipcRenderer.invoke('ideation:generate', params),
  ideationRegenerate: (params) => ipcRenderer.invoke('ideation:regenerate', params),
  ideationSaveFinal: (params) => ipcRenderer.invoke('ideation:saveFinal', params),
  ideationListOnePagers: (params) => ipcRenderer.invoke('ideation:listOnePagers', params),
  ideationReadOnePager: (params) => ipcRenderer.invoke('ideation:readOnePager', params),
  ideationUploadOnePager: (params) => ipcRenderer.invoke('ideation:uploadOnePager', params),

  presenceAnnounce: (params) => ipcRenderer.invoke('presence:announce', params),
  presenceList: (params) => ipcRenderer.invoke('presence:list', params),
  presenceClear: (params) => ipcRenderer.invoke('presence:clear', params),
  ideationTestOllama: (params) => ipcRenderer.invoke('ideation:testOllama', params),

  artReqGenerate: (params) => ipcRenderer.invoke('artreq:generate', params),
  artReqRegenerate: (params) => ipcRenderer.invoke('artreq:regenerate', params),
  artReqSaveFinal: (params) => ipcRenderer.invoke('artreq:saveFinal', params),
  artReqPreview: (params) => ipcRenderer.invoke('artreq:preview', params),
  artReqListDocs: (params) => ipcRenderer.invoke('artreq:listDocs', params),
  artReqReadDoc: (params) => ipcRenderer.invoke('artreq:readDoc', params),
  artReqExportDocx: (params) => ipcRenderer.invoke('artreq:exportDocx', params),
  artReqRevealDoc: (params) => ipcRenderer.invoke('artreq:revealDoc', params),

  ideationEntries: (params) => ipcRenderer.invoke('ideation:entries', params),
  ideationRetract: (params) => ipcRenderer.invoke('ideation:retract', params),
  ideationBrowseDataFolder: () => ipcRenderer.invoke('ideation:browseDataFolder'),
  ideationDataFolderInfo: () => ipcRenderer.invoke('ideation:dataFolderInfo'),

  comfyCheck: () => ipcRenderer.invoke('comfy:check'),
  artReqImages: (params) => ipcRenderer.invoke('artreq:images', params),
  artReqPlan: (params) => ipcRenderer.invoke('artreq:plan', params),
  artReqRegenImage: (params) => ipcRenderer.invoke('artreq:regenImage', params),

  sprintBoard: (params) => ipcRenderer.invoke('sprint:board', params),

  profileSave: (params) => ipcRenderer.invoke('profile:save', params),
  profileList: () => ipcRenderer.invoke('profile:list')
});
