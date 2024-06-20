document.addEventListener("DOMContentLoaded", function() {
    const form = document.getElementById("riskForm");
    const addRowButton = document.getElementById("addRow");
    const deleteRowButton = document.getElementById("deleteRow");
    const riskInputs = document.getElementById("riskInputs");
    const riskMatrix = document.getElementById("riskMatrix");

    let rowCount = 1;

    // Function to add a new row of inputs
    const addRow = () => {
        rowCount++;
        const newRow = document.createElement("div");
        newRow.classList.add("riskRow");
        newRow.innerHTML = `
            <label for="description${rowCount}">Описание:</label>
            <input type="text" id="description${rowCount}" name="description" required>
            <label for="probability${rowCount}">Вероятность:</label>
            <input type="number" id="probability${rowCount}" name="probability" min="1" max="5" required>
            <label for="impact${rowCount}">Влияние:</label>
            <input type="number" id="impact${rowCount}" name="impact" min="1" max="5" required>
        `;
        riskInputs.appendChild(newRow);
        updateMatrix(); // Update matrix after adding row
    };

    // Function to delete the last added row
    const deleteRow = () => {
        if (rowCount > 1) {
            riskInputs.removeChild(riskInputs.lastElementChild);
            rowCount--;
            updateMatrix(); // Update matrix after deleting row
        }
    };

    const initializeMatrix = () => {
        for (let impact = 1; impact <= 5; impact++) {
            for (let probability = 1; probability <= 5; probability++) {
                const cell = document.getElementById(`${impact}-${probability}`);
                const riskValue = probability * impact;

                if (riskValue <= 2) {
                    cell.setAttribute("data-risk-level", "low");
                } else if (riskValue >= 3 && riskValue <= 6 && riskValue !== 5) {
                    cell.setAttribute("data-risk-level", "medium");
                } else if (riskValue >= 5 && riskValue <= 12) {
                    cell.setAttribute("data-risk-level", "high");
                } else {
                    cell.setAttribute("data-risk-level", "very-high");
                }
            }
        }
    };
    initializeMatrix();

    // Function to initialize or update matrix with current form data
    const updateMatrix = () => {
        // Initialize all cells with 0
        for (let impact = 1; impact <= 5; impact++) {
            for (let probability = 1; probability <= 5; probability++) {
                const cell = document.getElementById(`${impact}-${probability}`);
                cell.textContent = '0';
            }
        }

        // Update matrix with current form data
        const probabilities = form.querySelectorAll("input[name='probability']");
        const impacts = form.querySelectorAll("input[name='impact']");

        probabilities.forEach((probInput, index) => {
            const probability = parseInt(probInput.value);
            const impact = parseInt(impacts[index].value);

            if (!isNaN(impact) && !isNaN(probability) && impact >= 1 && impact <= 5 && probability >= 1 && probability <= 5) {
                const cell = document.getElementById(`${impact}-${probability}`);
                let currentCount = parseInt(cell.textContent) || 0;
                cell.textContent = currentCount + 1;
            }
        });

        riskMatrix.style.display = "table";
    };

    // Event listeners
    addRowButton.addEventListener("click", addRow);
    deleteRowButton.addEventListener("click", deleteRow);

    form.addEventListener("submit", function(event) {
        event.preventDefault();
        updateMatrix();
    });

    // Function to handle Excel file upload
    const handleExcelUpload = () => {
        const fileInput = document.getElementById('excelFileInput');
        const file = fileInput.files[0];

        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });

                // Assuming the first sheet is where data is located
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];

                // Initialize all cells with 0
                for (let impact = 1; impact <= 5; impact++) {
                    for (let probability = 1; probability <= 5; probability++) {
                        const cell = document.getElementById(`${impact}-${probability}`);
                        cell.textContent = '0';
                    }
                }

                // Process each row from Excel and populate the matrix
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                jsonData.forEach((row, index) => {
                    const description = row['Описание'];
                    const impact = parseInt(row['Влияние']);
                    const probability = parseInt(row['Вероятность']);

                    // Validate impact and probability values
                    if (!isNaN(impact) && !isNaN(probability) && impact >= 1 && impact <= 5 && probability >= 1 && probability <= 5) {
                        const cell = document.getElementById(`${impact}-${probability}`);
                        let currentCount = parseInt(cell.textContent) || 0;
                        cell.textContent = currentCount + 1;
                    }
                });

                riskMatrix.style.display = "table";
            };
            reader.readAsArrayBuffer(file);
        }
    };

    // Event listener for Excel upload button
    const uploadExcelButton = document.getElementById('uploadExcel');
    uploadExcelButton.addEventListener('click', handleExcelUpload);
});
