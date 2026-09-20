import { contextBridge, ipcRenderer } from 'electron';
contextBridge.exposeInMainWorld('api', {
  loadRules: () => ipcRenderer.invoke('rules:load'),
  saveRules: (content: string) => ipcRenderer.invoke('rules:save', content),
  saveDialog: (name: string) => ipcRenderer.invoke('dialog:save', name),
  writeFile: (path: string, bytes: Uint8Array) => ipcRenderer.invoke('file:write', path, bytes),
  version: () => ipcRenderer.invoke('app:version')
});
