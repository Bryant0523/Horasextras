const { app, BrowserWindow, Menu, ipcMain } = require('electron')
const path = require('path')

Menu.setApplicationMenu(null)

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    show: false,
    frame: false,
    titleBarStyle: 'hidden',
    backgroundColor: '#0e0f11',
    icon: path.join(__dirname, 'image/logo32x32.ico'),
    title: 'Rook',
    autoHideMenuBar: true,
    webPreferences: {
        nodeIntegration: false,
        contextIsolation: true,
        sandbox: false,          // ← agregar esto
        preload: path.join(__dirname, 'preload.js')
}
  })

  win.loadFile('renderer/index.html')
  console.log('preload path:', path.join(__dirname, 'preload.js'))

win.once('ready-to-show', () => {
  win.show()
  win.webContents.openDevTools() // ← abre consola automático
})

  ipcMain.on('minimize', () => win.minimize())
  ipcMain.on('maximize', () => win.isMaximized() ? win.unmaximize() : win.maximize())
  ipcMain.on('close',    () => win.close())
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  app.quit()
})