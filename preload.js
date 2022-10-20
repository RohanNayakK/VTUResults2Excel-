const { ipcRenderer, contextBridge } = require("electron")
var XLSX = require('xlsx');


window.setInterval(()=>{
    driver()
 },2000)




//Function Prototype def
function sortByKey(array, key) {
    return array.sort(function(a, b) {
        var x = a[key]; var y = b[key];
        return ((x < y) ? -1 : ((x > y) ? 1 : 0));
    });
}

let allStudentData=[]

let USNRangeFromServer={}

let USNList=[]


//mass result 
let allResults=[]


let sendSubmit=(dataObj)=>{
  ipcRenderer.send("callPython",dataObj)
}

let sendUSNRange=(dataObj)=>{
    ipcRenderer.send("saveUSNRange",dataObj)
  }

let sendGenerateNewExcel=(dataObj)=>{
    ipcRenderer.send("callNewPython",dataObj)
}

ipcRenderer.on('sendData', function (evt, data) {
    ipcRenderer.send("sentData",{
        data:allStudentData
    })
})

ipcRenderer.on('errorMessage', function (evt, data) {

    let elem= document.getElementById("errorMessage")
    elem.innerText = data.message

})

function sleep(ms) {
    console.log("I was called")
    return new Promise(resolve => setTimeout(resolve, ms));
}

//MassExtract
ipcRenderer.on('massExtract',async function (evt, data) {

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

    const readMarksDataMass=(marksData)=>{
        let tempArr=[]
        for (let k=1;k<marksData.length-1;k++){
            sub.subjectCode = marksData[k].children[0].innerText;
            sub.subjectName = marksData[k].children[1].innerText;
            sub.iaMarks = marksData[k].children[2].innerText.trim();
            sub.eaMarks = marksData[k].children[3].innerText.trim();
            sub.totalMarks = marksData[k].children[4].innerText.trim();
            sub.result = marksData[k].children[5].innerText;
            tempArr.push(sub)
            //reset sub
            sub={}
        }
        return tempArr
    }




    function massDataExtract(htmlPage){

        let currentStudentData={
            studentName :null,
            studentUSN  :null,
            subjects:null
        }

        let htmlDoc = document.createElement( 'html' );
        htmlDoc.innerHTML = htmlPage


        let studentData = htmlDoc.getElementsByTagName('tbody')[0].children
        let marksData = htmlDoc.getElementsByClassName('divTableBody')[0].children



        //extract student data
        for (let i=0;i<studentData.length;i++){
            
            if(studentData[i].children[0].innerText.trim() ==="University Seat Number"){
                
                currentStudentData.studentUSN  = studentData[i].children[1].innerText.split(":")[1]
            }
            else {
                currentStudentData.studentName = studentData[i].children[1].innerText.split(":")[1]
            }

        }

        //extract marks data
        let extractedMarksData = readMarksDataMass(marksData)
        sortByKey(extractedMarksData,'subjectName')


        currentStudentData.subjects= extractedMarksData
    
        return(currentStudentData)
 
       
    }

    console.log("start Extraction !",USNList)

    let fetchRequestList=[]

    let testUSNList1=['4CB19IS001', '4CB19IS002', '4CB19IS003', '4CB19IS004', '4CB19IS005', '4CB19IS006', '4CB19IS007', '4CB19IS008', '4CB19IS009', '4CB19IS010', '4CB19IS011', '4CB19IS012', '4CB19IS013', '4CB19IS014', '4CB19IS015', '4CB19IS016', '4CB19IS017', '4CB19IS018', '4CB19IS019', '4CB19IS020']
    let testUSNList2=['4CB19IS021', '4CB19IS022', '4CB19IS023', '4CB19IS024', '4CB19IS025', '4CB19IS026', '4CB19IS027', '4CB19IS028', '4CB19IS029', '4CB19IS030', '4CB19IS031', '4CB19IS032', '4CB19IS033', '4CB19IS034', '4CB19IS035', '4CB19IS036', '4CB19IS037', '4CB19IS038', '4CB19IS039', '4CB19IS040']


    let allHtml=[]

    //test set 1
    for(let curusn in testUSNList1){
    
        console.log(curusn)

        let formData = new FormData();
        formData.append("Token",window.sessionStorage.getItem("token"))
        formData.append("lns", USNList[curusn])
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


            // fetchRequestList.push(fetchReq)
        

        // console.log("fetchlist",fetchRequestList)

        // const allResponseData=Promise.all(fetchRequestList)

        // console.log(allResponseData,"allresponse")

        .then(async (html)=>{
               //when session end -  'Please check website after 2 hour !!!'
               //when invalid USN -   "University Seat Number is not available or Invalid..!"
               if(html.split(">")[0].trim()!=="<!DOCTYPE html"){
                console.log("I was called")
                   let resMessage=html.split("(")[1].split(")")[0].trim()
                   console.log(resMessage)
                   let unqoutedResponseString =resMessage.replace(/["']/g, "")
                   if(unqoutedResponseString==="Please check website after 2 hour !!!")
                   {

                    console.log("Time Out! Check after 2 hours")
                    // await sleep(40000)
            
                    // console.log("Resume")
                   }
                //    if(resMessageMessage)
               }
               else {
                     allHtml.push(html)
                     let retVal = massDataExtract(html)
                   
                     allResults.push(retVal)
               }
    
            })
          
   }




   


   










   alert("Completed Succesfully!")
   console.log(allResults,"allResults")
})







ipcRenderer.on('getValue', function (evt, data) {
    USNRangeFromServer=data
    console.log(data)
    
    for(let k=parseInt(data.startUSNValue);k<=parseInt(data.endUSNValue);k++){
        if(k<10){
            USNList.push(data.USN_FIXED_PART+"00"+String(parseInt(k)))
        }
        else{
            USNList.push(data.USN_FIXED_PART+"0"+String(k))
        }
        
    }

    console.log(USNList,"USNList")


    let tokenValue = document.querySelector('[name="Token"]').value
    let lnsValue = document.querySelector('[name="lns"]').value
    let captchaValue = document.querySelector('[name="captchacode"]').value


    window.sessionStorage.setItem("token",tokenValue)
    window.sessionStorage.setItem("captcha",captchaValue)
    window.sessionStorage.setItem("usnFixedPart",lnsValue.substr(0,7))
    window.sessionStorage.setItem("usnNumPart",lnsValue.slice(7))

});

let positiveResponse=true

let USN_FIX_PART,USN_NUM_PART,next_USN;


let count=0

function driver(){
    if(count<2){
        count++
        window.dispatchEvent(new KeyboardEvent('keydown', {'key': '39'}));
    }
}







window.addEventListener("keydown",(e)=>{
    if(e.keyCode=='39'){
        alert("ji")
    }


    // if(e.keyCode=='39'){
    //     let studentData=[]
    //     let marksData=[]
    //     let currentStudentData={
    //         studentName :null,
    //         studentUSN  :null,
    //         subjects:null
    //     }
    
    //     let sub={
    //         subjectCode: null,
    //         subjectName: null,
    //         iaMarks:null,
    //         eaMarks:null,
    //         totalMarks:null,
    //         result:null
    //     }
    

       
    
    
    //     const readMarksData=()=>{
    //         let tempArr=[]
    //         for (let k=1;k<marksData.length-1;k++){
    //             sub.subjectCode = marksData[k].children[0].innerText;
    //             sub.subjectName = marksData[k].children[1].innerText;
    //             sub.iaMarks = marksData[k].children[2].innerText;
    //             sub.eaMarks = marksData[k].children[3].innerText;
    //             sub.totalMarks = marksData[k].children[4].innerText;
    //             sub.result = marksData[k].children[5].innerText;
    //             tempArr.push(sub)
    //             //reset sub
    //             sub={}
    //         }
    //         return tempArr
    //     }
    
       
    
        
    
    
    //     if(positiveResponse){
    //         updateUSN()
    
          
    
    
    //         studentData = document.getElementsByTagName('tbody')[0].children
    //         marksData = document.getElementsByClassName('divTableBody')[0].children
    
    
    
    //         //extract student data
    //         for (let i=0;i<studentData.length;i++){
    //             if(studentData[i].children[0].innerText ==="University Seat Number"){
    //                 currentStudentData.studentUSN  = studentData[i].children[1].innerText.split(":")[1]
    //             }
    //             else {
    //                 currentStudentData.studentName = studentData[i].children[1].innerText.split(":")[1]
    //             }
    
    //         }
    
    //         //extract marks data
    //         let extractedMarksData = readMarksData()
    //         sortByKey(extractedMarksData,'subjectName')
    
    
    //         currentStudentData.subjects= extractedMarksData
    //         console.log(currentStudentData)
    //         allStudentData.push(currentStudentData)
    
    //     }
        
    
        
    
    
    
    
    
    
    
    
    
    
    
        
    
    
    //     let formData = new FormData();
    //     formData.append("Token",window.sessionStorage.getItem("token"))
    //     formData.append("lns", next_USN)
    //     formData.append("captchacode", window.sessionStorage.getItem("captcha"))
    
    //     fetch("resultpage.php", {
    //         method: "POST",
    //         headers: {
    //             'Content-Type': 'application/x-www-form-urlencoded',
    //         },
    //         body: new URLSearchParams(formData)
    //     })
    //         .then(function (response) {
    //                 return response.text();
    
    
    //         })
    //         .then((html)=>{
    //            //when session end -  'Please check website after 2 hour !!!'
    //            //when invalid USN -   "University Seat Number is not available or Invalid..!"
            
    
    //            if(html.split(">")[0].trim()!=="<!DOCTYPE html"){
    //                positiveResponse=false
    
    //                let resMessage=html.split("(")[1].split(")")[0].trim()
    //                console.log(resMessage)
    //                console.log(String(resMessage)===`"University Seat Number is not available or Invalid..!"`)
    //                if(String(resMessage)===`"University Seat Number is not available or Invalid..!"`){
    //                    alert("Invalid/Non Existent USN, Continue Next Or Exit ")
    //                    positiveResponse=true
    //                }
    //                else {
    //                    alert(resMessage)
    //                }
    //            }
    //            else {
    //                positiveResponse=true
    //                document.body.innerHTML = html
    //            }
    
    //         })



    // }
})



ipcRenderer.on('next', async function (evt, data) {



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
        sortByKey(extractedMarksData,'subjectName')


        currentStudentData.subjects= extractedMarksData
        console.log(currentStudentData)
        allStudentData.push(currentStudentData)

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
               positiveResponse=false

               let resMessage=html.split("(")[1].split(")")[0].trim()
               console.log(resMessage)
               console.log(String(resMessage)===`"University Seat Number is not available or Invalid..!"`)
               if(String(resMessage)===`"University Seat Number is not available or Invalid..!"`){
                   alert("Invalid/Non Existent USN, Continue Next Or Exit ")
                   positiveResponse=true
               }
               else {
                   alert(resMessage)
               }
           }
           else {
               positiveResponse=true
               document.body.innerHTML = html
           }

        })

})


let sendReadExcel=(data)=>{
    const HTMLOUT = document.getElementById('htmlout');
    HTMLOUT.innerHTML = "";
    let wb=XLSX.read(data, {type: 'array'})
    wb.SheetNames.forEach(function(sheetName) {
    const detailstag = document.createElement("details")
    const summaryTag = document.createElement("summary")
		const htmlstr = XLSX.utils.sheet_to_html(wb.Sheets[sheetName],{editable:false});
		detailstag.innerHTML = htmlstr;
    detailstag.append(summaryTag)
    HTMLOUT.append(detailstag)
    
    summaryTag.innerText=sheetName

    
	});

}

let openResult=()=>{
  ipcRenderer.send("openVTUResultsPage",{})
}

let indexBridge={
  sendSubmit :sendSubmit,
  sendReadExcel:sendReadExcel,
  openResult : openResult,
  sendGenerateNewExcel:sendGenerateNewExcel,
  sendUSNRange:sendUSNRange
}


contextBridge.exposeInMainWorld("Bridge",indexBridge)


