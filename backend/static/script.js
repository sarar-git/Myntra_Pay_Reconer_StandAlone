const button=document.getElementById("uploadBtn");

const fileInput=document.getElementById("excelFile");

const status=document.getElementById("status");

button.addEventListener("click",async()=>{

    if(fileInput.files.length===0){

        alert("Select Excel File");

        return;

    }

    status.innerHTML="Uploading...";

    const formData=new FormData();

    formData.append(
        "file",
        fileInput.files[0]
    );

    const response=await fetch(

        API_URL+"/upload",

        {

            method:"POST",

            body:formData

        }

    );

    if(!response.ok){

        status.innerHTML="Upload Failed";

        return;

    }

    const blob=await response.blob();

    const url=window.URL.createObjectURL(blob);

    const a=document.createElement("a");

    a.href=url;

    a.download="Payment_Register.xlsx";

    a.click();

    status.innerHTML="Completed Successfully";

});