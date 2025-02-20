async function searchFile() {
    const styleNumber = document.getElementById("styleNumber").value.trim();
    const shrinkage = document.getElementById("shrinkage").value.trim();

    if (!styleNumber || !shrinkage) {
        alert("Please enter both Style Number and Shrinkage.");
        return;
    }

    const response = await fetch(`/open-excel?file=${styleNumber}&sheet=${shrinkage}`);
    const data = await response.json();

    if (data.error) {
        alert(data.error);
        return;
    }

    displayTable(data);
}

function displayTable(data) {
    const headerRow = document.getElementById("table-header");
    const body = document.getElementById("table-body");
    headerRow.innerHTML = "";
    body.innerHTML = "";

    if (data.length === 0) {
        alert("No data found in this sheet.");
        return;
    }

    document.getElementById("table-container").style.display = "block";

    // Add table headers
    Object.keys(data[0]).forEach(key => {
        const th = document.createElement("th");
        th.textContent = key;
        headerRow.appendChild(th);
    });

    // Add table rows
    data.forEach(row => {
        const tr = document.createElement("tr");
        Object.values(row).forEach(value => {
            const td = document.createElement("td");
            td.contentEditable = true;
            td.textContent = value;
            tr.appendChild(td);
        });
        body.appendChild(tr);
    });
}

function confirmSave() {
    const confirmBox = confirm("Are you done?");
    if (confirmBox) {
        saveToPDF();
    }
}

async function saveToPDF() {
    const styleNumber = document.getElementById("styleNumber").value.trim();
    const shrinkage = document.getElementById("shrinkage").value.trim();

    const response = await fetch(`/save-pdf?file=${styleNumber}&sheet=${shrinkage}`);
    const result = await response.blob();

    // Create a link to download the PDF
    const url = window.URL.createObjectURL(result);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${styleNumber}_${shrinkage}.pdf`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}