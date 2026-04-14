const { contextBridge, ipcRenderer } = require('electron')

try {
  contextBridge.exposeInMainWorld('electronAPI', {
    minimize: () => ipcRenderer.send('minimize'),
    maximize: () => ipcRenderer.send('maximize'),
    close:    () => ipcRenderer.send('close')
  })
  console.log('✅ electronAPI expuesto correctamente')
} catch(e) {
  console.error('❌ Error en preload:', e)
}