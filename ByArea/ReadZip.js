
var FileNames =[], Area= {}
document.getElementById('fileInput').addEventListener('change', handleFileInputChange);

async function readZipFile(file) {
    try {
        const zip = new JSZip();
        // Read the zip file
        const zipData = await zip.loadAsync(file);
        const workbooks = [];

        // Process each file in the zip
        for (const [relativePath, zipEntry] of Object.entries(zipData.files)) {
            if (zipEntry.name.endsWith('.xlsx')) {
                FileNames.push(zipEntry.name.replace('.xlsx', ''))
                console.log()
                // Read and process the Excel file
                const arrayBuffer = await zipEntry.async('arraybuffer');
                const data = new Uint8Array(arrayBuffer);
                const workbook = XLSX.read(data, { type: 'array' });
                workbooks.push(workbook);
            }
        }
        return workbooks;
    } catch (error) {
        console.error('Error reading zip file:', error);
        throw error; // Re-throw the error to propagate it to the caller
    }
}
async function handleFileInputChange() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (file) {
        try {
            if (file.name.endsWith('.zip')) {
                workbooks = await readZipFile(file);
            }
            else {
                console.log('Unsupported file type.');
            }

            // Call populateDateDropdown for each workbook
            if (workbooks && workbooks.length > 0) {
                await populateAreaDropdown(workbooks);
            } else {
                console.log('No workbooks found.');
            }
        } catch (error) {
            console.error('Error handling file input:', error);
        }
    } else {
        console.log('No file selected.');
    }
}
async function populateAreaDropdown(workbooks){
    function readSheet(worksheet) {
        return new Promise(resolve => {
            setTimeout(() => {
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                resolve(jsonData);
            }, 0);
        });
    }
    function separateColumnAndRow(cellRef) {
        const matches = cellRef.match(/([A-Z]+)(\d+)/);
        if (matches && matches.length === 3) {
            const column = matches[1];
            const row = parseInt(matches[2]);
            return { column, row };
        }
        return null;
    }
    console.log('workbooks: ',workbooks)
    // Process each workbook in the array
    let count = 0
    for (const workbook of workbooks){
        Area[FileNames[count]] = {}
        let columnField = ['sg','overall','completed', 'no splitter', 'remaining', 'remark']
         // Process each sheet in the workbook
         for (const sheetName of workbook.SheetNames){
            const worksheet = workbook.Sheets[sheetName];
            if(sheetName.toLowerCase() == 'overview'){
                let colRowArr = []
                for (columnRow in worksheet){
                    //console.log('columnRow: ',columnRow)
                    if(!columnRow.includes('!')){
                        value = worksheet[columnRow].v
                        if (typeof value == 'string'){
                            if(columnField.includes(value.toLowerCase())){
                                columnField = columnField.filter(item => item !== `${value.toLowerCase()}`);
                                const { column, row } = separateColumnAndRow(columnRow);
                                colRowArr.push([column, row, value.toLowerCase()])
                            }
                        }
                    }
                }
                let sgValue = null
                for (columnRow in worksheet){
                    if(!columnRow.includes('!')){
                        const { column, row } = separateColumnAndRow(columnRow);
                        for(let i=0; i < colRowArr.length; i++){
                            let startCol = colRowArr[i][0]
                            let startRow = colRowArr[i][1]
                            let value = colRowArr[i][2]
                            if(i == 0){
                                if(startCol == column && row != startRow){
                                    sgValue = worksheet[columnRow].v
                                    if(sgValue.toLowerCase().includes('sg')){
                                        Area[FileNames[count]][sgValue] = {}
                                    }
                                }
                            }
                            else{
                                if(sgValue != null){
                                    if(sgValue.toLowerCase().includes('sg') && startCol == column && row != startRow){
                                        cellValue = worksheet[columnRow].v
                                        Area[FileNames[count]][sgValue][value] = cellValue
                                    }
                                }
                            }
                        }
                    }
                }
                for(let SG in Area[FileNames[count]]){
                    for(let i = 1; i < colRowArr.length; i++){
                        let value = colRowArr[i][2]
                        if(Area[FileNames[count]][SG][value] == undefined){
                            if(value == 'remark'){
                                Area[FileNames[count]][SG][value] = ''
                            }
                            else{
                                Area[FileNames[count]][SG][value] = 0
                            }
                        }
                        else if (value == 'remaining'){
                            if(typeof Area[FileNames[count]][SG][value] == 'string'){
                                if(Area[FileNames[count]][SG][value].toLowerCase() == 'completed'){
                                    Area[FileNames[count]][SG][value] = 0
                                }
                            }
                        }
                    }
                }
                count += 1
                break;
            }
         }
    }
    populateDropdown() // function ni dalam dropdown.js
    displayArea(selectedArea)
}

//display table
function displayArea(selectedArea) {
    console.log('selectedArea: ',selectedArea)
    if(Object.keys(selectedArea).length === 0){
        selectedArea = Area
    }
    let grandComplete = 0, grandNotdone = 0, grandHold = 0;
    const resultBody = document.getElementById('resultBody');
    resultBody.innerHTML = '';
    
    for (let sgArea in selectedArea) {
        const keysRow = document.createElement("tr");
        keysRow.innerHTML = `<th style="padding: 10px 0; background-color: #e0e8f0; cursor: pointer;">${sgArea}</th>
        <th style="text-align: center;">Completed</th>
        <th style="text-align: center;">Not Done Yet</th>
        <th style="text-align: center;">Hold/RFI</th>
        <th style="text-align: center;">Remark</th>`;
        resultBody.appendChild(keysRow);

        // Insert SG table
        let total_completed = 0, total_notdone = 0, total_hold = 0;
        const SG = selectedArea[sgArea];
        const rows = [];
        for (let sgValue in SG) {
            let notdone = 0, hold = 0, remark = ''
            let completed = SG[sgValue].completed + SG[sgValue]["no splitter"];
            if (SG[sgValue].remaining > 0) {
                if (SG[sgValue].remark) {
                    hold = SG[sgValue].remaining;
                    remark = SG[sgValue].remark;
                    if(remark.toLowerCase().includes('done passthrough') || remark.toLowerCase().includes('passthrough done')){
                        hold = 0
                    }
                } else {
                    notdone = SG[sgValue].remaining;
                }
            }
            if(remark.toLowerCase().includes('done passthrough')){
                completed = 0
            }
            const row = document.createElement("tr");
            row.innerHTML = `
                <td style="text-align: center;">${sgValue}</td>
                <td style="text-align: center;">${completed}</td>
                <td style="text-align: center;">${notdone}</td>
                <td style="text-align: center;">${hold}</td>
                <td style="text-align: center;">${remark}</td>`;
            resultBody.appendChild(row);
            rows.push(row);
            row.style.display = 'none'
            total_completed += completed;
            total_notdone += notdone
            total_hold += hold;
            
        }

        // Add total row
        const rowTotal = document.createElement("tr");
        rowTotal.innerHTML = `
            <td style="text-align: center;">Total</td>
            <td style="text-align: center; font-weight: bold;">${total_completed}</td>
            <td style="text-align: center; font-weight: bold;">${total_notdone}</td>
            <td style="text-align: center; font-weight: bold; background-color: ${total_hold > 0 ? '#f7c6c6' : 'transparent'};">${total_hold}</td>
            <td style="text-align: center;"></td>`;
        resultBody.appendChild(rowTotal);

        // Update grand totals
        grandComplete += total_completed;
        grandNotdone += total_notdone;
        grandHold += total_hold;

        // Toggle display of rows when keysRow is clicked
        keysRow.addEventListener("click", function() {
        rows.forEach(row => {
                if (row.style.display === "none") {
                    row.style.display = "table-row";
                } else {
                    row.style.display = "none";
                }
            });
        });
    }

    // Add grand total row
    const rowGrandTotal = document.createElement("tr");
    rowGrandTotal.innerHTML = `
            <td style="text-align: center;">Grand Total</td>
            <td style="text-align: center; font-weight: bold;">${grandComplete}</td>
            <td style="text-align: center; font-weight: bold;">${grandNotdone}</td>
            <td style="text-align: center; font-weight: bold;">${grandHold}</td>
            <td style="text-align: center;"></td>`;
    resultBody.appendChild(rowGrandTotal);
}
