// Modules to control application life and create native browser window
const {app, BrowserWindow,screen,Menu,Notification,globalShortcut,shell } = require('electron')
const path = require('path')
const Alert = require("electron-alert");
const { ipcMain } = require('electron')
const { execFile,spawn } = require('child_process');
const fs = require("fs");
const { dialog } = require('electron')
const prompt = require('electron-prompt');



//Variables Declaration
let mainWindow

let autoNextBot;

let isBotPaused=false;

let userInputFileName="generated-result.xlsx"

const USN_RANGE_LIMIT=150;

//count for auto next
let count =0

let userSelectedFilePath = ""

//Instantiate Window
function createWindow () {
  // Create the browser window.
    mainWindow= new BrowserWindow({
        backgroundColor: '#141e26',
        show: false,
        resizable: true,



    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      worldSafeExecuteJavaScript: true,
	  nodeIntegration: true,
	  enableRemoteModule: true,
        webSecurity: false
    },
    devTools: true


  })



    mainWindow.maximize()


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

//Menu Template & Handlers
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
                    mainWindow.loadFile(path.join(__dirname, '/views/index.html'))
                }
            },
            {
                label: 'Exit',
                role: 'close'
            },
        ]

    },

    {
        label: 'Options',
        submenu: [
            {
                label : "Get Token&Captcha",
                click: async () => {
                    mainWindow.webContents.send('getValue',{});
                    showTokenCaptchaAlert()

                }
            },

            {
                label : "Manual Next",
                click: async () => {
                    mainWindow.webContents.send('next', {});
                }
            },

            {
                label : "Auto Next",

                click: async () => {


                    prompt({
                        title: 'Enter Last USN ',
                        label: 'Last USN 3 digits :',
                        menuBarVisible: false,
                        alwaysOnTop: true,
                        value: '060',
                        inputAttrs: {
                            type: 'number',
                            required: true,
                            min: 3,
                        },
                        type: 'input'
                    })
                        .then((r) => {
                            if(r === null) {
                                console.log('user cancelled');
                            } else {
                                autoNextHandler(r)
                            }
                        })
                        .catch(console.error);




                }
            },
            {
                label : "Generate Excel",
                click: async () => {
                    mainWindow.webContents.send('sendData', {});

                }
            },
            {
                label : "Refresh Screen Freeze(After Wrong Captcha)",
                click: async () => {
                    mainWindow.openDevTools()
                    setTimeout(()=>{
                        mainWindow.closeDevTools()
                    },1000)
                }
            }

        ]
    }
   ,

];

//Context Menu Template & Handlers
let contextMenuTemplate = [
     {
        label : "Manual Next",
        click: async () => {
            mainWindow.webContents.send('next', {});
        }
    },
            {
                label : "Auto Next",

                click: async () => {
                    prompt({
                        title: 'Enter Last USN ',
                        label: 'Last USN 3 digits :',
                        menuBarVisible: false,
                        alwaysOnTop: true,
                        value: '060',
                        inputAttrs: {
                            type: 'number',
                            required: true,
                            min: 3,
                        },
                        type: 'input'
                    })
                        .then((r) => {
                            if(r === null) {
                                console.log('user cancelled');
                            } else {
                                autoNextHandler(r)
                            }
                        })
                        .catch(console.error);

                }
            },
            {
                label : "Generate Excel",
                click: async () => {
                    mainWindow.webContents.send('sendData', {});

                }
            },


];

//Show Token & Captcha Alert
function showTokenCaptchaAlert(){
    let alert = new Alert();
    let token=null
    let captcha=null

    mainWindow.webContents
        .executeJavaScript('window.sessionStorage.getItem("token");', true)
        .then(result => {
            console.log(result)
            token= result

            mainWindow.webContents
                .executeJavaScript('window.sessionStorage.getItem("captcha");', true)
                .then(result => {
                    console.log(result)
                    captcha= result


                    let swalOptions = {
                        title: "Token & Captcha",
                        text: `Session Variables Recorded :\n Token: ${token} \n Captcha: ${captcha}`,
                        icon: "success",
                        showCancelButton: false,

                    };

                    let promise = alert.fireWithFrame(swalOptions,"Token & Captcha" ,null, true);
                    promise.then((result) => {
                        if (result.value) {
                            // confirmed

                        }
                    })

                });

        });
}


//Auto Next Handler
const autoNextHandler =(r)=>{
    const userInputLastUSN = Number(r);

    if(userInputLastUSN<=USN_RANGE_LIMIT){
        autoNextBot= setInterval(()=>{
            if(count<=userInputLastUSN && !isBotPaused){
                mainWindow.webContents.send('next', {});
                count++
                if(count===userInputLastUSN){
                    clearInterval(autoNextBot)
                    mainWindow.webContents.send('sendData', {});
                }
            }

        },2000)
    }
    else{
        let alert = new Alert();

        let swalOptions = {
            title: "USN Range Error",
            text: `Invalid-USN not in range`,
            icon: "Error",
            showCancelButton: false,

        };

        let promise = alert.fireWithFrame(swalOptions,"Range Error" ,null, true);
        promise.then((result) => {
            if (result.value) {
            }
        })

    }
}

//Execute Python Script ( Child Process )
function fetchResult(dataObject) {
    return new Promise((resolve , reject) => {
        userSelectedFilePath = dataObject.path
        userInputFileName = dataObject.filename
        let childPython = execFile("canaraFetch.exe", [dataObject.path,dataObject.dept,dataObject.sem,dataObject.examType,dataObject.creditPoints,dataObject.acdYear,dataObject.totalCredits,dataObject.scheme,dataObject.filename],{
            cwd: path.join(__dirname, 'python')
        });

        let result='';
        childPython.stdout.on(`data` , (data) => {
            result = data.toString();
            if(result.includes("Error")){
                userSelectedFilePath = ""
                userInputFileName = ""
                reject(result)
            }
            console.log(result)
        });

        childPython.on('close' , function(code) {
            resolve("200")
        });

        childPython.on('error' , function(err){
            userSelectedFilePath = ""
            userInputFileName = ""
            reject(err)
        });

    })
  };


ipcMain.on("serverRefreshWait",async (event,data)=>{
    isBotPaused=true
    //decrement count on server refresh
    count--;
    mainWindow.webContents.send('showWaitTimer', {});
})

ipcMain.on("serverRefreshOverResumeBot",async (event,data)=>{
    isBotPaused=false
  })

ipcMain.on("openDialog",async (event,data)=>{
    dialog.showOpenDialog(mainWindow, {
        properties: [ 'openDirectory']
      }).then(result => {

        let path = result.filePaths[0].replace(/\\/g, "/");
        console.log(path)

        mainWindow.webContents


    .executeJavaScript(`document.getElementById("filePathInput").value="${String(path)}/"`, true)
    .then(result => {
        captcha= result

   });
        console.log(result.filePaths)
      }).catch(err => {
        console.log(err)
      })
  })

ipcMain.on("callPythonExe",async (event,data)=>{
    fetchResult(data).then(()=>{
        mainWindow.loadFile(path.join(__dirname, '/views/success.html'))
        new Notification({ title: "NOTE:", body: "Excel File Generated , Check file in specified path" }).show()

    })
        .catch((err)=>{
            console.log(err)
            mainWindow.loadFile(path.join(__dirname, '/views/generationFail.html'))
        })
})

ipcMain.on("openVTUResultsPage",(event,data)=>{
    mainWindow.loadURL("https://results.vtu.ac.in/");
});

ipcMain.on("sentData",(event,data)=>{

     let writeData = JSON.stringify(data.data);

     fs.writeFileSync(path.join(__dirname, '/python/data.json'), writeData);

    mainWindow.loadFile(path.join(__dirname, '/views/generateExcel.html'))
        .then(()=>{
            mainWindow.webContents.executeJavaScript(`
            window.location.reload();
                        window.sessionStorage.setItem("extractedData", ${JSON.stringify(writeData)});
                        `)
        })




})

ipcMain.on('show-context-menu', (event) => {
    const menu = Menu.buildFromTemplate(contextMenuTemplate)
    menu.popup(BrowserWindow.fromWebContents(event.sender))
  })

ipcMain.on("showElectronAlert",(event,data)=>{
    dialog.showErrorBox("Error", data)
})

ipcMain.on("showFile",(event,data)=>{
    if(userSelectedFilePath===""){
        dialog.showErrorBox("Error", "No File Found or Generated")
    }
    else {
        shell.openPath(userSelectedFilePath+userInputFileName)
    }

})

app.whenReady().then(() => {
    let menu = Menu.buildFromTemplate(menuTemplate);

    Menu.setApplicationMenu(menu);

    createWindow()


    //add global shortcut
    globalShortcut.register('CommandOrControl+R', () => {
        mainWindow.openDevTools()
        setTimeout(()=>{
            mainWindow.closeDevTools()
        },1000)
    })

    mainWindow.webContents.on('did-fail-load', (_event ) => {
        console.log(_event)

        mainWindow.loadFile(path.join(__dirname, '/views/errorPage.html')).then(()=>{})

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













