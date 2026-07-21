const button = document.getElementById("uploadBtn");
const fileInput = document.getElementById("excelFile");
const status = document.getElementById("status");

button.addEventListener("click", async () => {

    if (fileInput.files.length === 0) {
        alert("Please select an Excel (.xlsx) file.");
        return;
    }

    button.disabled = true;
    status.innerHTML = "Uploading and processing...";

    const formData = new FormData();
    formData.append("file", fileInput.files[0]);

    try {

        const response = await fetch("/upload", {
            method: "POST",
            body: formData
        });

        if (!response.ok) {

            const error = await response.text();

            console.log(error);
            
            alert(error.detail);;

            status.innerHTML = error.detail;

            return;
        }

        const blob = await response.blob();

        const downloadUrl = window.URL.createObjectURL(blob);

        const a = document.createElement("a");

        a.href = downloadUrl;
        a.download = "Payment_Register.xlsx";

        document.body.appendChild(a);

        a.click();

        a.remove();

        window.URL.revokeObjectURL(downloadUrl);

        status.innerHTML = "✅ Payment Register generated successfully.";

    }
    catch (err) {

        console.error(err);

        status.innerHTML = "❌ Unable to connect to the server.";

    }
    finally {

        button.disabled = false;

    }

});