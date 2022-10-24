// Modules to control application life and create native browser window
const {app, BrowserWindow,screen,Menu,Notification } = require('electron')
const path = require('path')
const { ipcMain } = require('electron')
const { execFile } = require('child_process');
const fs = require("fs");

let SavedUSNRange={}


let mainWindow

function createWindow () {
  // Create the browser window.
    mainWindow= new BrowserWindow({
        backgroundColor: '#000000',
        show: false,


    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      worldSafeExecuteJavaScript: true,
	  nodeIntegration: true,
	  enableRemoteModule: true,
        webSecurity: false
    }


  })


  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, './views/index.html'))

  mainWindow.maximize()
    mainWindow.once('ready-to-show', () => {

        mainWindow.show()
    })
    mainWindow.webContents.on('new-window', (event, url) => {
        event.preventDefault()
        mainWindow.loadURL(url)
    })

}

let menuTemplate = [

    {
        label: 'File',
        submenu: [
            {
                label : "Back",
                click: async () => {
                    mainWindow.webContents.goBack();

                }
            },
            {
                label: 'Reload Page',
                role: 'reload'
            },
            {
                label: 'Hard Reload',
                role: 'forceReload'
            },
            {
                label: 'Open Dev Tools',
                role: 'toggledevtools'
            },
            {
                type: 'separator'
            },
            {
                label: 'Home',
                click: async () => {
                    mainWindow.loadFile(path.join(__dirname, './views/index.html'))


                }
            },
            {
                label: 'Exit',
                role: 'close'
            },
        ]

    },

    {
        label : "Get Token/Captcha",
        click: async () => {
            mainWindow.webContents.send('getValue',{});
            new Notification({ title: "NOTE:", body: "Captcha and Token Saved" }).show()
        }
    },



    {
        label : "Next",

        click: async () => {
            mainWindow.webContents.send('next', {});

        }
    },

    {
        label : "Extract All In USN Range",

        click: async () => {
            mainWindow.webContents.send('massExtract', {});

        }
    },

    {
        label : "Send Data",
        click: async () => {
            mainWindow.webContents.send('sendData', {});

        }
    },


];

function fetchResult(dataObject,pythonExecutableFileName) {
    return new Promise((resolve , reject) => {
        let childPython
        if(pythonExecutableFileName==="fetch.exe"){
            childPython = execFile(pythonExecutableFileName, [dataObject.collegeCode, dataObject.year, dataObject.branch, dataObject.startusn, dataObject.lastusn,dataObject.path,String(dataObject.sem)]);
        }
        else if(pythonExecutableFileName==="newFetch.exe"){
            childPython = execFile(pythonExecutableFileName, [dataObject.path]);

        }
        let result='';
        childPython.stdout.on(`data` , (data) => {
            result = data.toString();
            console.log(result)
        });

        childPython.on('close' , function(code) {
            resolve("200")
        });

        childPython.on('error' , function(err){
            reject(err)
        });

    })
  };

ipcMain.on("callPython",async (event,data)=>{
    await fetchResult(data,'fetch.exe')
    mainWindow.loadFile(path.join(__dirname, './views/success.html'))
  })




  ipcMain.on("saveUSNRange",async (event,data)=>{
    SavedUSNRange=data
    console.log(SavedUSNRange,"Data on Server")
    mainWindow.loadURL("https://results.vtu.ac.in/");
  })



//new executable
ipcMain.on("callNewPython",async (event,data)=>{

    await fetchResult(data,'newFetch.exe')
    mainWindow.loadFile(path.join(__dirname, './views/success.html'))
})


ipcMain.on("openVTUResultsPage",(event,data)=>{
      mainWindow.loadURL("https://results.vtu.ac.in/");

  });


ipcMain.on("sentData",(event,data)=>{

     let writeData = JSON.stringify(data.data);

     fs.writeFileSync('data.json', writeData);

     mainWindow.loadFile(path.join(__dirname, './views/generateExcel.html'))

 })


app.whenReady().then(() => {
    let menu = Menu.buildFromTemplate(menuTemplate);



    Menu.setApplicationMenu(menu);
  createWindow()

    mainWindow.webContents.on('did-fail-load', (_event ) => {
        mainWindow.loadFile(path.join(__dirname, './views/errorPage.html')).then(()=>{})

    });


    screenElectron= screen;

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})


app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})


app.on('open-file', function () { console.log(arguments); });













