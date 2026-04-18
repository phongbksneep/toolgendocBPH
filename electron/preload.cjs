'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // catalog & templates
  getCatalog:      ()        => ipcRenderer.invoke('get-catalog'),
  getTemplates:    ()        => ipcRenderer.invoke('get-templates'),
  // file operations
  openJson:        ()        => ipcRenderer.invoke('open-json'),
  saveJson:        (a)       => ipcRenderer.invoke('save-json', a),
  exportSampleJson:()        => ipcRenderer.invoke('export-sample-json'),
  exportSampleExcel:()       => ipcRenderer.invoke('export-sample-excel'),
  exportCurrentExcel:(a)     => ipcRenderer.invoke('export-current-excel', a),
  importExcel:     ()        => ipcRenderer.invoke('import-excel'),
  loadDemoProject: ()        => ipcRenderer.invoke('load-demo-project'),
  editLabelLocal:  (a)       => ipcRenderer.invoke('edit-label-local', a),
  exportLabelPack: ()        => ipcRenderer.invoke('export-label-pack'),
  importLabelPack: ()        => ipcRenderer.invoke('import-label-pack'),
  generateDocs:    (a)       => ipcRenderer.invoke('generate-docs', a),
  openPath:        (target)  => ipcRenderer.invoke('open-path', target),
  // app info
  getVersion:      ()        => ipcRenderer.invoke('get-version'),
  getPlatform:     ()        => ipcRenderer.invoke('get-platform'),
  refreshRuntimeAssets: ()    => ipcRenderer.invoke('refresh-runtime-assets'),
  // update
  checkUpdate:     ()        => ipcRenderer.invoke('check-update'),
  installUpdate:   (a)       => ipcRenderer.invoke('install-update', a),
  downloadUpdate:  (a)       => ipcRenderer.invoke('download-update', a),
  // autosave
  loadAutosave:    ()        => ipcRenderer.invoke('autosave-load'),
  saveAutosave:    (payload) => ipcRenderer.invoke('autosave-save', payload),
  clearAutosave:   ()        => ipcRenderer.invoke('autosave-clear'),
  // template management
  readTemplate:    (a)       => ipcRenderer.invoke('read-template', a),
  saveTemplate:    (a)       => ipcRenderer.invoke('save-template', a),
  deleteTemplate:  (a)       => ipcRenderer.invoke('delete-template', a),
  importTemplate:  ()        => ipcRenderer.invoke('import-template'),
});

