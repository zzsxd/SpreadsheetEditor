const { app, BrowserWindow } = require('electron')
const path = require('path')

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      webSecurity: false // Добавьте это для локальной загрузки файлов
    }
  })

  // Для разработки
  if (process.env.NODE_ENV === 'development') {
    win.loadURL('http://localhost:8080')
    win.webContents.openDevTools()
  } else {
    // Для production
    win.loadFile(path.join(__dirname, '../dist/index.html'))
  }

  // Обработка белого экрана
  win.webContents.on('did-fail-load', () => {
    win.loadFile(path.join(__dirname, '../dist/index.html'))
  })
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow()
})