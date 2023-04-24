//Import Modules
const { ipcRenderer, contextBridge } = require("electron")
let XLSX = require('xlsx');
const Alert = require("electron-alert");


//Window Event Listeners

//Auto Save Token & Captcha on Form Submit
window.addEventListener("submit",()=>{



    //if Token Exists!
    if(document.querySelector('[name="Token"]').value){
        let tokenValue = document.querySelector('[name="Token"]').value
        let lnsValue = document.querySelector('[name="lns"]').value
        let captchaValue = document.querySelector('[name="captchacode"]').value


        window.sessionStorage.setItem("token",tokenValue)
        window.sessionStorage.setItem("captcha",captchaValue)
        window.sessionStorage.setItem("usnFixedPart",lnsValue.substr(0,7))
        window.sessionStorage.setItem("usnNumPart",lnsValue.slice(7))
    }


})





window.alert = function(message) {
    ipcRenderer.send("showElectronAlert",message)
}


//Custom Context Menu
window.addEventListener('contextmenu', (e) => {
    e.preventDefault()
    ipcRenderer.send('show-context-menu')
})



//Variables Declaration
let allStudentData=[]

let positiveResponse=true

let next_USN;

let USN_FIX_PART,USN_NUM_PART;





//Utility functions
//Sort function
function sortByKey(array, key) {
    return array.sort(function(a, b) {
        var x = a[key]; var y = b[key];
        return ((x < y) ? -1 : ((x > y) ? 1 : 0));
    });
}

//USN update function
const updateUSN=()=> {
    USN_FIX_PART = window.sessionStorage.getItem("usnFixedPart")
    USN_NUM_PART = Number(window.sessionStorage.getItem("usnNumPart"))

    if (USN_NUM_PART <= 8) {
        next_USN = USN_FIX_PART + "00" + (USN_NUM_PART + 1)
    }
    else if (USN_NUM_PART < 100) {
        next_USN = USN_FIX_PART + "0" + +(USN_NUM_PART + 1)
    }
    else {
        next_USN = USN_FIX_PART + (USN_NUM_PART + 1)
    }

    window.sessionStorage.setItem("usnNumPart",USN_NUM_PART+1)
}


//Event Emitters Handlers
let sendGenerateNewExcel=(dataObj)=>{
    ipcRenderer.send("callPythonExe",dataObj)
}



let sendReadExcel=(data)=>{
    const HTMLOUT = document.getElementById('htmlout');
    HTMLOUT.innerHTML = "";
    let wb=XLSX.read(data, {type: 'array'})
    wb.SheetNames.forEach(function(sheetName) {
        const detailstag = document.createElement("details")
        const summaryTag = document.createElement("summary")
        detailstag.innerHTML = XLSX.utils.sheet_to_html(wb.Sheets[sheetName], {editable: false});
        detailstag.append(summaryTag)
        HTMLOUT.append(detailstag)
        summaryTag.innerText=sheetName
    });

}

let openResult=()=>{
    ipcRenderer.send("openVTUResultsPage",{})
}

let sendServerRefresh=(dataObj)=>{
    ipcRenderer.send("serverRefreshWait",dataObj)
}

let sendBrowsePath=()=>{
    ipcRenderer.send("openDialog",{})
}


let sendAlert=(dataObj)=>{
    ipcRenderer.send("showElectronAlert",dataObj)
}

let sendShowFile=(dataObj)=>{
    ipcRenderer.send("showFile",dataObj)
}


//Event Handlers

ipcRenderer.on('sendData', function (evt, data) {

    ipcRenderer.send("sentData",{
        data:allStudentData
    })
})

ipcRenderer.on('errorMessage', function (evt, data) {

    let elem= document.getElementById("errorMessage")
    elem.innerText = data.message

})

ipcRenderer.on('getValue', function (evt, data) {
    //Manual Store
    if(document.querySelector('[name="Token"]')?.value)
    {
    let tokenValue = document.querySelector('[name="Token"]').value
    let lnsValue = document.querySelector('[name="lns"]').value
    let captchaValue = document.querySelector('[name="captchacode"]').value


    window.sessionStorage.setItem("token",tokenValue)
    window.sessionStorage.setItem("captcha",captchaValue)
    window.sessionStorage.setItem("usnFixedPart",lnsValue.substr(0,7))
    window.sessionStorage.setItem("usnNumPart",lnsValue.slice(7))
    }

});

ipcRenderer.on('showWaitTimer',  function (evt, data) {

        let overlayContainer=document.createElement("div")
        overlayContainer.setAttribute("id","overlay")
        overlayContainer.setAttribute("style","position: fixed; display:flex;top:0;left:0;height:100%;width:100%;background-color:rgba(0,0,0,0.7);z-index:1000;color:white;font-size:2rem;align-items:center;justify-content:center;textAllign:center")
        overlayContainer.innerHTML=
        `
        <h1>Server Refreshing...</h1>
        <h1>Please Wait for 30 secs...</h1>
        `
        document.body.append(overlayContainer)

        let time = 30
        let timer = setInterval(() => {
            overlayContainer.innerHTML =
            `
            <h1>Server Refreshing...</h1>
            <h1>Please Wait for ${time} secs...</h1>
            <br>
            <p>It will auto resume after the stipulated time, dont exit or refresh</p>
            `
            time -= 1
            if(time===0){
                overlayContainer.remove()
                clearInterval(timer)
                ipcRenderer.send("serverRefreshOverResumeBot",{})

            }
        }, 1000);

});

ipcRenderer.on('next',  function (evt, data) {

    let studentData=[]

    let marksData=[]

    let currentStudentData={
        studentName :null,
        studentUSN  :null,
        subjects:null
    }

    let sub={
        subjectCode: null,
        subjectName: null,
        iaMarks:null,
        eaMarks:null,
        totalMarks:null,
        result:null
    }


    const readMarksData=()=>{
        let tempArr=[]
        for (let k=1;k<marksData.length-1;k++){
            sub.subjectCode = marksData[k].children[0].innerText;
            sub.subjectName = marksData[k].children[1].innerText;
            sub.iaMarks = marksData[k].children[2].innerText;
            sub.eaMarks = marksData[k].children[3].innerText;
            sub.totalMarks = marksData[k].children[4].innerText;
            sub.result = marksData[k].children[5].innerText;
            tempArr.push(sub)
            //reset sub
            sub={}
        }
        return tempArr
    }


    if(positiveResponse){
        updateUSN()



        studentData = document.getElementsByTagName('tbody')[0].children
        marksData = document.getElementsByClassName('divTableBody')[0].children

        //extract student data
        for (let i=0;i<studentData.length;i++){
            if(studentData[i].children[0].innerText ==="University Seat Number"){
                currentStudentData.studentUSN  = studentData[i].children[1].innerText.split(":")[1]
            }
            else {
                currentStudentData.studentName = studentData[i].children[1].innerText.split(":")[1]
            }

        }

        //extract marks data
        let extractedMarksData = readMarksData()

        //Sort Subject order
        sortByKey(extractedMarksData,'subjectName')

        currentStudentData.subjects= extractedMarksData

        //avoid duplicate data
        if(currentStudentData.studentUSN !== allStudentData[allStudentData.length-1]?.studentUSN){
            allStudentData.push(currentStudentData)
        }




    }

    let formData = new FormData();
    formData.append("Token",window.sessionStorage.getItem("token"))
    formData.append("lns", next_USN)
    formData.append("captchacode", window.sessionStorage.getItem("captcha"))

    fetch("resultpage.php", {
        method: "POST",
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: new URLSearchParams(formData)
    })
        .then(function (response) {
            return response.text();
        })

        .then((html)=>{
            //when session end -  'Please check website after 2 hour !!!'
            //when invalid USN -   "University Seat Number is not available or Invalid..!"





            if(html.split(">")[0].trim()!=="<!DOCTYPE html"){

                let resMessage = html.split("(")[1].split(")")[0].trim()

                let unqoutedResponseString = resMessage.replace(/["']/g,"")

                if(unqoutedResponseString==="University Seat Number is not available or Invalid..!"){
                    // alert("Invalid/Non Existent USN, Continue Next Or Exit ")
                    //Set True to skip and continue to the next USN
                    positiveResponse=true
                }
                else {
                    //alert(unqoutedResponseString)
                    sendServerRefresh()
                    positiveResponse=false
                }
            }
            else {
                positiveResponse=true

                //Update the DOM
                document.body.innerHTML = html
            }

        })

        .catch((err)=>{
            console.log(err)
        })

})



//Bridge Object
let indexBridge={
    openResult : openResult,
    sendReadExcel:sendReadExcel,
    sendGenerateNewExcel:sendGenerateNewExcel,
    sendServerRefresh : sendServerRefresh,
    sendBrowsePath:sendBrowsePath,
    sendAlert:sendAlert,
    sendShowFile:sendShowFile,
}


//Context bridge
contextBridge.exposeInMainWorld("Bridge",indexBridge)


