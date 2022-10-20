const puppeteer = require('puppeteer')
const fs = require('fs');
const FormData = require('form-data');
const https = require('https');

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = '0';




let baseURL="https://results.vtu.ac.in/JJEcbcs22/index.php"


function base64_encode(file) {
    // read binary data
    var bitmap = fs.readFileSync(file);
    // convert binary data to base64 encoded string
    return new Buffer(bitmap).toString('base64');
}

//TrueCaptcha API
function get_captcha(image_data){

    return new Promise((resolve, reject) => {

        const params = {			
            userid:'rnk2214@gmail.com',
            apikey:'xhoeS4kpYY0dXdBCyzkg',
            data:   image_data,
            case:   "mixed"
        }
        const url = "https://api.apitruecaptcha.org/one/gettext"
    
        fetch(url, {
            method: 'post',
            body: JSON.stringify(params)
        })
        .then((response) => response.json())
        .then((data) => {
            resolve(data.result)
        })
        .catch((error) => {
            reject(null)
        })
    });
	
}



async function scrape() {
    let allHtml=[]

    let failedUSN=[]
    
    console.log("Starting Scrape...")
    //browser instance
    const browser = await puppeteer.launch({headless:false})
    
    // Allows you to intercept a request; 
    const page = await browser.newPage()

    

    let testUSNList1=['4CB19IS001', '4CB19IS002','4CB19IS004']
    //request Page
    for (let i = 0; i < 3; i++) {

    await page.goto("https://results.vtu.ac.in/JJEcbcs22/index.php",{
        waitUntil: 'networkidle0',
      })

    let savedToken;

    //get token
    let token= await page.evaluate(() => {
        savedToken=document.querySelector('input[name="Token"]').value
        return document.querySelector('input[name="Token"]').value
    })
    console.log(token)


    //get captcha image
    let element = await page.waitForSelector(".col-md-4 > img")


    //save captcha image
    await element.screenshot({path: 'captcha.png'});

    //get base64 encoded image
    const imgData = base64_encode('./captcha.png');

    //decode captcha
    let captchaCode = await get_captcha(imgData)


    await page.type('input[name="lns"]', testUSNList1[i]);
    await page.type('input[name="captchacode"]', captchaCode)

    page.evaluate(() => {
        document.querySelector('input[id=submit]').click()
    })

    
    await page.waitForNavigation({
        waitUntil: 'networkidle0',
      });
    

    const inner_html = await page.$eval('html', element => element.innerHTML);
      

    if(inner_html.split(">")[0].trim()!=="<!DOCTYPE html"){

        failedUSN.push(testUSNList1[i])

           let resMessage=inner_html.split("(")[1].split(")")[0].trim()
           console.log(resMessage)

           let unqoutedResponseString =resMessage.replace(/["']/g, "")
           
           if(unqoutedResponseString==="Please check website after 2 hour !!!")
           {
            console.log("Time Out! Check after 2 hours")
           }

       }
       else {
             allHtml.push(inner_html)

            //  let retVal = massDataExtract(html)
           
            //  allResults.push(retVal)
       }

    console.log(allHtml);
    console.log(failedUSN);
}



    // page.evaluate(({inner_html}) => {


    //         let currentStudentData={
    //             studentName :null,
    //             studentUSN  :null,
    //             subjects:null
    //         }
    
    //         let htmlDoc = document.createElement( 'html' );
    //         htmlDoc.innerHTML = inner_html
    
    
    //         let studentData = htmlDoc.getElementsByTagName('tbody')[0].children
    //         let marksData = htmlDoc.getElementsByClassName('divTableBody')[0].children
    
    
    
    //         //extract student data
    //         for (let i=0;i<studentData.length;i++){
                
    //             if(studentData[i].children[0].innerText.trim() ==="University Seat Number"){
                    
    //                 currentStudentData.studentUSN  = studentData[i].children[1].innerText.split(":")[1]
    //             }
    //             else {
    //                 currentStudentData.studentName = studentData[i].children[1].innerText.split(":")[1]
    //             }
    
    //         }
    
    //         //extract marks data
    //         let extractedMarksData = readMarksDataMass(marksData)

    //         for (let k=1;k<marksData.length-1;k++){
    //             sub.subjectCode = marksData[k].children[0].innerText;
    //             sub.subjectName = marksData[k].children[1].innerText;
    //             sub.iaMarks = marksData[k].children[2].innerText.trim();
    //             sub.eaMarks = marksData[k].children[3].innerText.trim();
    //             sub.totalMarks = marksData[k].children[4].innerText.trim();
    //             sub.result = marksData[k].children[5].innerText;
    //             tempArr.push(sub)
    //             //reset sub
    //             sub={}
    //         }


    //         sortByKey(extractedMarksData,'subjectName')
    
    
    //         currentStudentData.subjects= extractedMarksData
        
    //         return(currentStudentData)
     
           
        

        

    // })


   

    // page.evaluate(({token,captchaCode}) => {

    //     let formData = new FormData();
    //     formData.append("Token",token)    
    //     formData.append("lns", "4CB19IS002")
    //     formData.append("captchacode", captchaCode)

        
    //     fetch("https://results.vtu.ac.in/resultpage.php", {
    //         method: "POST",
    //         headers: {
    //             'Content-Type': 'application/x-www-form-urlencoded',
    //         },
    //         body: new URLSearchParams(formData)
    //     })
    //         .then(function (response) {
    //                 return response.text();
    
    
    //         })
    //     .then(async (html)=>{
    //            //when session end -  'Please check website after 2 hour !!!'
    //            //when invalid USN -   "University Seat Number is not available or Invalid..!"
    //            if(html.split(">")[0].trim()!=="<!DOCTYPE html"){
    //             console.log("I was called")
    //                let resMessage=html.split("(")[1].split(")")[0].trim()
    //                console.log(resMessage)
    //                let unqoutedResponseString =resMessage.replace(/["']/g, "")
    //                if(unqoutedResponseString==="Please check website after 2 hour !!!")
    //                {

    //                 console.log("Time Out! Check after 2 hours")
                
    //                }
              
    //            }
    //            else {
    //                 document.querySelector('html').innerHTML=html
    //                 console.log(html)
    //                 //  allHtml.push(html)
    //                 //  let retVal = massDataExtract(html)
                   
    //                 //  allResults.push(retVal)
    //            }
    
    //         })
    // },{token,captchaCode})


    

      
    //close browser
    // browser.close()
}

//let testUSNList1=['4CB19IS001', '4CB19IS002', '4CB19IS003', '4CB19IS004', '4CB19IS005', '4CB19IS006', '4CB19IS007', '4CB19IS008', '4CB19IS009', '4CB19IS010', '4CB19IS011', '4CB19IS012', '4CB19IS013', '4CB19IS014', '4CB19IS015', '4CB19IS016', '4CB19IS017', '4CB19IS018', '4CB19IS019', '4CB19IS020']


//name=lns
//name=captchacode



const main = async () => {
    await scrape()

    
    
}


main()












