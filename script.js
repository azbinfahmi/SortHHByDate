var workbooks,uniqueDates =[],TotalNumColumn =[],New_TotalNumColumn=[];
var TotalY = 0, SumValue = 0, matchingRows_copy=[]
let FileNames =[], DateByArea=[], arr_totalHH =[]

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
                await populateDateDropdown(workbooks);
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

async function populateDateDropdown(workbooks) {
    const dateSelect = document.getElementById('dateSelect');
    dateSelect.style.display = "block"
    const TodateSelect = document.getElementById('dateSelect_2');
    uniqueDates = [];

    // Process each workbook in the array
    let count = 0
    for (const workbook of workbooks) {
        let DateByArea_scope = [], total_y =0, total_n =0,total_o = 0, calcTotalHH =0
        // Process each sheet in the workbook
        for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            const columnIndex = getColumnIndexByName(worksheet, 'Complete Date');

            if (columnIndex !== -1) {
                const jsonData = await readSheet(worksheet);
                jsonData.forEach(row => {
                    if(row['Passthrough']){
                        if (row['Passthrough'] === 'Y' || row['Passthrough'] === 'y') {
                            TotalY = TotalY + 1;
                            const dateValue = row['Complete Date']
                            const producerValue = row['Producer']
                            if (dateValue && !uniqueDates.includes(dateValue)) {
                                uniqueDates.push(dateValue);
                            }
                            if (dateValue && !DateByArea_scope.includes(dateValue)) {
                                DateByArea_scope.push(dateValue)
                            }
    
                            if(producerValue != "ECHO" && producerValue != "OTHER PARTIES" && producerValue != "OTHER PARTY"){
                                total_y = total_y + 1
                            }
    
                            else{
                                total_o = total_o + 1
                            }
                            calcTotalHH += 1
                        }

                        else if(row['Passthrough'] === 'n' || row['Passthrough'] === 'N'){
                            const producerValue = row['Producer']
    
                            if(producerValue != "ECHO" || producerValue != "OTHER PARTIES" || producerValue != "OTHER PARTY"){
                                total_n = total_n + 1
                            }
    
                            else{
                                total_o = total_o + 1
                            }
                            calcTotalHH += 1
                        }
                        
                    }

                    else if(row['Passthrough'] != 'n'){
                        if(row['vetro_id']){
                            calcTotalHH += 1
                        }
                        
                    }
                });
            }
        }

        
        for(var i in DateByArea_scope.sort()){
            DateByArea_scope[i] = excelDateToFormattedString(DateByArea_scope[i])
        }
        DateByArea.push([FileNames[count],[DateByArea_scope[0],DateByArea_scope[DateByArea_scope.length - 1]], total_y, total_n,total_o])
        arr_totalHH.push(calcTotalHH)
        count+= 1
    }

    console.log('arr_totalHH',arr_totalHH)
    // Sort dates in ascending order
    uniqueDates.sort();
    for(var i in uniqueDates){
        uniqueDates[i] = excelDateToFormattedString(uniqueDates[i])
    }
    // Populate the dropdown with dates
    uniqueDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        dateSelect.appendChild(option);
    });

    // Populate the dropdown with date for ToDateSelect
    uniqueDates.forEach(date => {
        const option = document.createElement('option');
        option.value = date;
        option.textContent = date;
        TodateSelect.appendChild(option);
    });

    //insert area table
    console.log('DateByArea',DateByArea)
    const AreaTableBody = document.getElementById('areaBody');
    AreaTableBody.innerHTML = ''

    if (DateByArea) {
        const totalHHColumn=[]
        for(let index = 0; index<DateByArea.length; index++){
            const newRow = AreaTableBody.insertRow();
            const cell1 = newRow.insertCell(0);
            cell1.textContent = DateByArea[index][0];
            cell1.style.textAlign = 'center';
            cell1.style.padding = '5px';
            cell1.classList.add('no-wrap');
            
            const cell2 = newRow.insertCell(1);
            if(DateByArea[index][1][0] == undefined || DateByArea[index][1][1] == undefined){
                cell2.textContent = ''
            }
            else{
                cell2.textContent = DateByArea[index][1].map(date => date.replace(/,\s\d{4}\b/, '')).join(' - ');
            }
            
            cell2.style.textAlign = 'center';
            cell2.style.padding = '5px';
            cell2.classList.add('no-wrap');

            const cell3 = newRow.insertCell(2);
            cell3.textContent = DateByArea[index][2];
            cell3.style.textAlign = 'center';
            cell3.style.padding = '5px';
            cell3.classList.add('no-wrap');

            const cell4 = newRow.insertCell(3);
            cell4.textContent = DateByArea[index][3];
            cell4.style.textAlign = 'center';
            cell4.style.padding = '5px';
            cell4.classList.add('no-wrap');

            const cell5 = newRow.insertCell(4);
            cell5.textContent = DateByArea[index][4];
            cell5.style.textAlign = 'center';
            cell5.style.padding = '5px';
            cell5.classList.add('no-wrap');

            const cell6 = newRow.insertCell(5);
            let total = DateByArea[index][2]+DateByArea[index][3] + DateByArea[index][4]
            cell6.textContent = total +' / '+ arr_totalHH[index];
            if( total != arr_totalHH[index]){
                cell6.style.backgroundColor = '#b7dec3'
            }
            cell6.style.textAlign = 'center';
            cell6.style.padding = '5px';
            cell6.classList.add('no-wrap');

            if(index == 0){
                totalHHColumn.push(DateByArea[index][2],DateByArea[index][3],DateByArea[index][4],total)
            }
            else{
                totalHHColumn[0] = totalHHColumn[0] + DateByArea[index][2]
                totalHHColumn[1] = totalHHColumn[1] + DateByArea[index][3]
                totalHHColumn[2] = totalHHColumn[2] + DateByArea[index][4]
                totalHHColumn[3] = totalHHColumn[3] + total
            }
        }

        const newRow = AreaTableBody.insertRow();
        for(let i =0; i< totalHHColumn.length +2; i++){
            const cell1 = newRow.insertCell(i);
            if(i<2){
                cell1.textContent = "";
            }
            else{
                cell1.textContent = totalHHColumn[i-2];
                cell1.style.textAlign = 'center';
                cell1.style.padding = '5px';
                cell1.style.fontWeight = 'bold';
            }
            
        } 
    }   
}

function getColumnIndexByName(sheet, columnName) {
    const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0];
    if(headerRow!= undefined){
        return headerRow.findIndex(header => header === columnName);
    }
    return -1
    
}

function readSheet(worksheet) {
    return new Promise(resolve => {
        setTimeout(() => {
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            resolve(jsonData);
        }, 0);
    });
}

function excelDateToFormattedString(serial) {
    const utcDays = Math.floor(serial - 25569);
    const utcMilliseconds = utcDays * 86400 * 1000;
    const date = new Date(utcMilliseconds);

    // Format the date as "dd MMM yyyy"
    const options = { day: 'numeric', month: 'short', year: 'numeric' };
    const formattedDate = date.toLocaleDateString('en-US', options);

    return formattedDate;
}

function formattedStringToExcelDate(formattedDate) {
    const dateParts = formattedDate.split(' ');

    // Mapping of month abbreviations to month numbers
    const monthAbbreviations = {
        'Jan': 0, 'Feb': 1, 'Mar': 2, 'Apr': 3, 'May': 4, 'Jun': 5,
        'Jul': 6, 'Aug': 7, 'Sep': 8, 'Oct': 9, 'Nov': 10, 'Dec': 11
    };

    const day = parseInt(dateParts[1], 10);
    const monthIndex = monthAbbreviations[dateParts[0]];
    const year = parseInt(dateParts[2], 10);
    const date = new Date(Date.UTC(year, monthIndex, day));

    // Convert to Excel serial number
    const utcDays = (date.getTime() / 86400000) + 25569;
    return utcDays;
}

function showRowsAndSheets() {
    TotalNumColumn =[]
    const selectedDateString = document.getElementById('dateSelect').value;
    const resultTableBody = document.getElementById('resultBody');
    resultTableBody.innerHTML = '';
    if (RangeDate.length == 0){
        RangeDate.push(formattedStringToExcelDate(selectedDateString))
    }
    // Iterate through each workbook in the array
    matchingRows_copy =[],currentRow = 0,StartingRow = 0,ColSpanValue = 0
    for(var i in RangeDate){
        totalSumValue = 0
        const selectedDateNumeric = RangeDate[i];
        for (const workbook of workbooks) {
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const columnIndex = getColumnIndexByName(worksheet, 'Complete Date');
                if (columnIndex !== -1) {
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    const matchingRows = jsonData.filter(row => {
                        const excelDateNumeric = row['Complete Date'];
                        const passthroughValue = row['Passthrough'];
                        var ProducerValue = row['Producer'];
                        if(ProducerValue == undefined || ProducerValue == ''){
                            ProducerValue = 'echo'
                        }

                        return excelDateNumeric === selectedDateNumeric && passthroughValue.toLowerCase() === 'y' && ProducerValue.toLowerCase() != 'other parties' 
                        && ProducerValue.toLowerCase() != 'other party' && ProducerValue.toLowerCase() != 'echo'
                    });
    
                    if (matchingRows.length > 0) {
                        totalSumValue += matchingRows.length;
    
                        const newRow = resultTableBody.insertRow();
                        cell1 = newRow.insertCell(0);
                        cell2 = newRow.insertCell(1);
                        cell3 = newRow.insertCell(2);
    
                        cell1.textContent = sheetName;
                        cell2.textContent = matchingRows.length;
                        cell3.textContent = matchingRows[0]['Producer']
    
                        cell1.style.textAlign = 'center';
                        cell2.style.textAlign = 'center';
                        cell3.style.textAlign = 'center';

                        cell1.classList.add('dark-left-border');
                        matchingRows_copy.push(matchingRows)
                    }
                }
            });
        }
        
        const newRow = resultTableBody.insertRow();
        cell1 = newRow.insertCell(0);
        cell2 = newRow.insertCell(1);
        cell3 = newRow.insertCell(2)

        cell1.textContent = "Total";
        cell1.style.textAlign = 'center';
        cell1.style.fontWeight = 'bold';
    
        cell2.textContent = totalSumValue ;
        cell2.style.textAlign = 'center';
        cell2.style.fontWeight = 'bold';
        
        cell3.textContent = "" ;
        //edit first row of the table
        ColSpanValue = resultTableBody.rows.length - StartingRow
        const firstRow = resultTableBody.rows[StartingRow];
        cell4 = firstRow.insertCell(3);
        cell4.textContent = excelDateToFormattedString(selectedDateNumeric);
        cell4.style.textAlign = 'center';
        cell4.rowSpan = ColSpanValue
        StartingRow = resultTableBody.rows.length

        // Add a custom class to the specific cells
        cell1.classList.add('dark-bottom-border');
        cell1.classList.add('dark-left-border');
        cell2.classList.add('dark-bottom-border');
        cell3.classList.add('dark-bottom-border');
        cell4.classList.add('dark-bottom-border');
        cell4.classList.add('dark-right-border');
        
    }
    //display total producer by creating new table
    const existingProducerTable = document.getElementById('producerTable');
    if (existingProducerTable) {
        existingProducerTable.remove();
    }

    const producerTotalsByDate  = {};
    matchingRows_copy.forEach(innerArray => {
        // Iterate through the inner array
        innerArray.forEach(item => {
          const ExcelDate = item["Complete Date"];
          const producer = item["Producer"].toUpperCase();

          // Check if both "Complete Date" and "Producer" fields are present and valid
          if (ExcelDate !== undefined && producer !== undefined) {
            Date_Complete = excelDateToFormattedString(ExcelDate)
            // Check if the producer is already in the totals object, if not, initialize it
            if (!producerTotalsByDate[producer]) {
                producerTotalsByDate[producer] = [];
            }

            const producerArray = producerTotalsByDate[producer];
            const existingEntry = producerArray.find(entry => entry.Date_Complete === Date_Complete)

            if (existingEntry) {
                // Update existing entry
                existingEntry.count += 1;
            } else {
                
                // Add a new entry
                producerArray.push({ Date_Complete, count: 1 });
            }
          }
        });
      });
      console.log('producerTotalsByDate ',producerTotalsByDate )

    const producerTable = document.createElement('table');
    producerTable.id = 'producerTable';
    producerTable.style.borderCollapse = 'collapse';
    producerTable.style.border = '2px solid black'; 
    const tableBody = document.createElement('tbody');

    // Add header row
    const headerRow = producerTable.createTHead().insertRow(0);
    const headerCell1 = headerRow.insertCell(0);
    headerCell1.textContent = 'Producer';
    headerCell1.style.textAlign = 'center';
    headerCell1.style.border = '1px solid black'; 
    headerCell1.style.fontWeight = 'bold';
    headerCell1.style.padding = '5px';

    //to display Header in table
    for(var i in RangeDate){
        const index = Number(i) + 1
        date = excelDateToFormattedString(RangeDate[i])
        const headerCell2 = headerRow.insertCell(index)
        headerCell2.textContent = date;
        headerCell2.style.textAlign = 'center';
        headerCell2.style.border = '1px solid black';
        headerCell2.style.fontWeight = 'bold';
        headerCell2.style.padding = '5px';

        if(index == RangeDate.length){
            const headerCell3 = headerRow.insertCell(index+1)
            headerCell3.textContent = 'Total HandHole Completed in ' + RangeDate.length + ' days';
            headerCell3.style.textAlign = 'center';
            headerCell3.style.border = '1px solid black';
            headerCell3.style.fontWeight = 'bold';
            headerCell3.style.padding = '5px';
        }
    }
    
    //display Producer and completed HH in next row
    let grandTotal = 0
    for (const producer in producerTotalsByDate) {
        const producerRow = tableBody.insertRow();
        const producerCell = producerRow.insertCell(0);
        producerCell.textContent = producer;
        producerCell.style.textAlign = 'center';
        producerCell.style.border = '1px solid black';
        producerCell.style.fontWeight = 'bold';
        producerCell.style.padding = '5px';
      
        let rowTotal = 0;
        const countsByDate = {};
        producerTotalsByDate[producer].forEach(entry => {
          countsByDate[entry.Date_Complete] = entry.count;
          rowTotal += entry.count;
        });

        // Populate cells with counts for each date
        CountRowArr =[]
        for (let i = 0; i < RangeDate.length; i++) {
            const index = i + 1;
            const date = excelDateToFormattedString(RangeDate[i]);
            const count = countsByDate[date] || 0; // Use 0 if the date is not present
            const countCell = producerRow.insertCell(index);
            countCell.textContent = count;
            countCell.style.textAlign = 'center';
            countCell.style.border = '1px solid black';
            countCell.style.padding = '5px';
            CountRowArr.push(count)
        }
        TotalNumColumn.push(CountRowArr)

        const totalCell = producerRow.insertCell(producerRow.cells.length);
        totalCell.textContent = rowTotal;
        totalCell.style.textAlign = 'center';
        totalCell.style.border = '1px solid black';
        totalCell.style.fontWeight = 'bold';
        totalCell.style.padding = '5px';

        // Update Grand Total
        grandTotal += rowTotal;
    }
    console.log('TotalNumColumn',TotalNumColumn)
    //Sum total value by column
    New_TotalNumColumn=[]
    for (var i=0; i<TotalNumColumn.length; i++)
    {
        len = TotalNumColumn[i].length
        for(var j=0; j<TotalNumColumn[i].length; j++){
            if (i ==0){
                New_TotalNumColumn.push(TotalNumColumn[i][j])
            }
            else{
                New_TotalNumColumn[j] = New_TotalNumColumn[j] + TotalNumColumn[i][j]
            }
        }
    }
    // Add empty cells for dates
    const grandTotalRow = tableBody.insertRow();
    for (let i = 0; i < RangeDate.length + 1; i++) {
        const index = i;
        if(index == 0){
            const grandTotalDateCell = grandTotalRow.insertCell(index);
            grandTotalDateCell.textContent = 'Total';
            grandTotalDateCell.style.textAlign = 'center';
            grandTotalDateCell.style.border = '1px solid black';
            grandTotalDateCell.style.fontWeight = 'bold';
        }
        else{
            const grandTotalDateCell = grandTotalRow.insertCell(index);
            grandTotalDateCell.textContent = New_TotalNumColumn[i-1];
            grandTotalDateCell.style.textAlign = 'center';
            grandTotalDateCell.style.border = '1px solid black';
            grandTotalDateCell.style.fontWeight = 'bold';

        }
    }

    // Add Grand Total cell
    const grandTotalCellTotal = grandTotalRow.insertCell(grandTotalRow.cells.length);
    grandTotalCellTotal.textContent = grandTotal;
    grandTotalCellTotal.style.textAlign = 'center';
    grandTotalCellTotal.style.border = '1px solid black';
    grandTotalCellTotal.style.fontWeight = 'bold';
    grandTotalCellTotal.style.padding = '5px';


    producerTable.appendChild(tableBody);
    document.body.appendChild(producerTable);
    RangeDate =[]
}
