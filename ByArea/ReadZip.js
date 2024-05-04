
let FileNames =[]
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
    console.log(workbooks)
}