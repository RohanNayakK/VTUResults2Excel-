let detailsForm= document.getElementById('detailsForm')
let detailsFormFieldSet= document.getElementById('detailsFormFieldSet')
let detailsForm2= document.getElementById('detailsForm2')
let backBtn = document.getElementById("backBtn")
let loader = document.getElementById("loadingContainer")

const collegeCode=document.getElementById("collegeCodeInput")
const year=document.getElementById("syear")
const branch = document.getElementById("sbranch")
const startusn= document.getElementById("sfusn")
const lastusn = document.getElementById("slusn")
const path =document.getElementById("folderPath")
let latestResult;





backBtn.addEventListener("click",()=>{
  detailsFormFieldSet.style.display = "block"
  detailsForm2.style.display = "none"
})


//start screen form
detailsForm.addEventListener("submit",(e)=>{
  e.preventDefault()
  let baseURL="https://api.vtuconnect.in/result/"
  let usn;
  if(startusn.value < 10){
     usn = collegeCode.value+year.value+String(branch.value)+"00"+String(startusn.value)

  }
  else if(startusn.value < 100){
    usn = collegeCode.value+year.value+String(branch.value)+"0"+String(startusn.value)

  }
  else{
    usn = collegeCode.value+year.value+String(branch.value)+String(startusn.value)

  }
  let url= baseURL+usn
  let semArray=[]

  fetch(url,{
    headers:{
    "Accept": "application/json",
    "Authorization": "Bearer eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyIjp7InNjb3BlIjpbInVzZXIiXSwiZW1haWwiOiJuYW5kYW5AZ21haWwuY29tIn0sImlhdCI6MTU0NzMxODI5N30.ZPO8tf03azhTJ1qmgSVyGV80k9EfomXgGazdLyUC6fw",
    "Host": "api.vtuconnect.in",
    }
  
  })
  .then(response => {
    if(response.status===200){
      return response.json()
    }
    else{
      return alert("Something went wrong")
    }
      
  })
  .then((data)=>{
    console.log(data)
    for(let i=0;i<data.length;i++)
    {
      semArray.push(Number(data[i].semester))
    }
    
   



    if(data.length==0){
      let errorMessageForm1=document.getElementById("errorMessageForm1")
      console.log(errorMessageForm1)
      errorMessageForm1.classList.add("alertMessageError")
      errorMessageForm1.innerHTML=`<i class="fa-solid fa-circle-exclamation"></i> Invalid Input`
      setTimeout(()=>{
        errorMessageForm1.innerText=""
        errorMessageForm1.classList.remove("alertMessageError")
      },1000)
      return
      
    }
    else{
      latestResult=Math.max(...semArray)
      let availableResultElement=document.getElementById("availableResult")
      availableResultElement.innerText=latestResult+" Semester Result"
      detailsFormFieldSet.style.display = "none"
      detailsForm2.style.display = "block" 
    }
    
  })
  .catch(error => {
      alert("Something went wrong! Try Again")
  });



 
});


detailsForm2.addEventListener("submit",(e)=>{
    e.preventDefault()
    loader.style.display ="flex"
    let formDataObj={
      collegeCode : collegeCode.value,
      year:String(year.value),
      branch:branch.value,
      startusn:String(startusn.value),
      lastusn:String(lastusn.value),
      path:path.value,
      sem:latestResult
    }

    window.Bridge.sendSubmit(formDataObj)
  
});

collegeCode.addEventListener("input",(event)=>{
  event.target.value = event.target.value.toUpperCase()
})

branch.addEventListener("input",(event)=>{
  event.target.value = event.target.value.toUpperCase()
})




