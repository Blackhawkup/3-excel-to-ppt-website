// Handle Monthly PPT form
const monthlyForm = document.getElementById("monthly-form");
if (monthlyForm) {
    monthlyForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const resultDiv = document.getElementById("monthly-result");
        resultDiv.textContent = 'Processing...';

        const formData = new FormData();
        formData.append("file", document.getElementById("monthly-excel").files[0]);

        try {
            const response = await fetch("/monthly/", {
                method: "POST",
                body: formData,
            });

            if (response.ok && response.headers.get("content-type").includes("application/vnd.openxmlformats-officedocument.presentationml.presentation")) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = "generated.pptx";
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                resultDiv.textContent = "PPT generated and downloading...";
            } else {
                const err = await response.json();
                resultDiv.textContent = err.error || "Error occurred!";
            }
        } catch (err) {
            resultDiv.textContent = "Upload failed. Try again.";
        }
    });
}

// Handle RIA form
const riaForm = document.getElementById("ria-form");
if (riaForm) {
    riaForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const resultDiv = document.getElementById("ria-result");
        resultDiv.textContent = 'Processing...';

        const formData = new FormData();
        formData.append("ppt_file", document.getElementById("ria-ppt").files[0]);
        formData.append("excel_file1", document.getElementById("ria-excel1").files[0]);
        formData.append("excel_file2", document.getElementById("ria-excel2").files[0]);
        formData.append("excel_file3", document.getElementById("ria-excel3").files[0]);
        formData.append("owner_no", document.getElementById("owner-no").value);

        try {
            const response = await fetch("/process-ppt/", {
                method: "POST",
                body: formData,
            });

            if (response.ok && response.headers.get("content-type").includes("application/vnd.openxmlformats-officedocument.presentationml.presentation")) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = "modified.pptx";
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                resultDiv.textContent = "PPT processed and downloading...";
            } else {
                const err = await response.json();
                resultDiv.textContent = err.error || "Error occurred!";
            }
        } catch (err) {
            resultDiv.textContent = "Upload failed. Try again.";
        }
    });
}

// Handle Ready Reckoner form
const rrForm = document.getElementById("readyreckoner-form");
if (rrForm) {
    rrForm.addEventListener("submit", async (e) => {
        e.preventDefault();
        const resultDiv = document.getElementById("readyreckoner-result");
        resultDiv.textContent = 'Processing...';

        const formData = new FormData();
        formData.append("process_type", document.getElementById("rr-process-type").value);
        formData.append("n_pms", document.getElementById("rr-n-pms").value || 1);
        formData.append("n_hybrid", document.getElementById("rr-n-hybrid").value || 1);
        formData.append("excel_file", document.getElementById("rr-excel").files[0]);
        formData.append("pms_template", document.getElementById("rr-pms-template").files[0]);
        formData.append("hybrid_template", document.getElementById("rr-hybrid-template").files[0]);

        try {
            const response = await fetch("/generate_pptx/", {
                method: "POST",
                body: formData,
            });

            if (response.ok && response.headers.get("content-type").includes("application/vnd.openxmlformats-officedocument.presentationml.presentation")) {
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                if (formData.get("process_type") === "PMS") {
                    a.download = "Client_Associates_All_PMS_Funds.pptx";
                } else {
                    a.download = "Client_Associates_All_Hybrid_Funds.pptx";
                }
                document.body.appendChild(a);
                a.click();
                a.remove();
                window.URL.revokeObjectURL(url);
                resultDiv.textContent = "PPT generated and downloading...";
            } else {
                const err = await response.json();
                resultDiv.textContent = err.error || "Error occurred!";
            }
        } catch (err) {
            resultDiv.textContent = "Upload failed. Try again.";
        }
    });
}

// Highlight navigation for three tabs
const loc = window.location.pathname.split('/').pop();
const navHome = document.getElementById('nav-home');
const navRia = document.getElementById('nav-ria');
const navRr  = document.getElementById('nav-rr');
if (navHome && navRia && navRr) {
    if (loc === "ria.html") {
        navRia.classList.add("active");
        navHome.classList.remove("active");
        navRr.classList.remove("active");
    } else if (loc === "readyreckoner.html") {
        navRr.classList.add("active");
        navHome.classList.remove("active");
        navRia.classList.remove("active");
    } else {
        navHome.classList.add("active");
        navRia.classList.remove("active");
        navRr.classList.remove("active");
    }
}
