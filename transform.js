// Existing Comps Transformer - JavaScript Version (with ExcelJS for full formatting)
// Converts raw MLS data to formatted, quartile-analyzed comps

let processedWorkbook = null;
let statsData = null;

// Columns to keep from raw data
const COLUMNS_TO_KEEP = [
    'Acres',
    'City',
    'DOM',
    'Garage Capacity',
    'List Price',
    'Original List Price',
    'Price Per Square Foot',
    'Sold Concessions',
    'Sold Date',
    'Sold Price',
    'Total Bedrooms',
    'Total Bathrooms',
    'Total Square Feet',
    'Year Built',
    'Property Type'
];

// Quartile colors (light to dark green)
const QUARTILE_COLORS = {
    1: 'E2EFD9',
    2: 'C5E0B3',
    3: 'A8D08D',
    4: '548135'
};

// Initialize event listeners
document.addEventListener('DOMContentLoaded', function() {
    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const downloadBtn = document.getElementById('downloadBtn');
    const processAnotherBtn = document.getElementById('processAnotherBtn');
    const tryAgainBtn = document.getElementById('tryAgainBtn');

    // File input change
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
    dropZone.addEventListener('dragover', handleDragOver);
    dropZone.addEventListener('dragleave', handleDragLeave);
    dropZone.addEventListener('drop', handleDrop);

    // Buttons
    downloadBtn.addEventListener('click', downloadFile);
    processAnotherBtn.addEventListener('click', resetApp);
    tryAgainBtn.addEventListener('click', resetApp);
});

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.add('drag-over');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.remove('drag-over');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    document.getElementById('dropZone').classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

function processFile(file) {
    // Validate file type
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
        showError('Please upload a valid Excel file (.xlsx or .xls)');
        return;
    }

    // Show file name
    document.getElementById('fileName').textContent = `Selected: ${file.name}`;

    // Show progress
    showProgress('Reading file...');
    updateProgress(10);

    // Read the file
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            updateProgress(30);
            showProgress('Processing data...');
            
            const data = e.target.result;
            transformData(data);
            
        } catch (error) {
            showError(`Error reading file: ${error.message}`);
        }
    };

    reader.onerror = function() {
        showError('Error reading file. Please try again.');
    };

    reader.readAsArrayBuffer(file);
}

async function transformData(data) {
    try {
        // Step 1: Read the workbook with ExcelJS
        showProgress('Loading raw data...');
        updateProgress(40);
        
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(data);
        
        updateProgress(50);
        
        // Find the data sheet
        let worksheet;
        let sheetName = 'Existing Comps Data';
        
        worksheet = workbook.getWorksheet(sheetName);
        
        // If "Existing Comps Data" doesn't exist, try to use the only sheet if there's just one
        if (!worksheet) {
            if (workbook.worksheets.length === 1) {
                worksheet = workbook.worksheets[0];
                sheetName = worksheet.name;
                console.log(`Using single sheet: "${sheetName}"`);
            } else {
                const sheetNames = workbook.worksheets.map(ws => ws.name).join(', ');
                throw new Error(`Sheet "Existing Comps Data" not found. Your file has ${workbook.worksheets.length} sheets: ${sheetNames}. Please make sure your raw data sheet is named "Existing Comps Data".`);
            }
        }
        
        // Step 2: Extract data
        showProgress('Extracting data...');
        updateProgress(60);
        
        const rawData = [];
        const headers = [];
        
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) {
                // Header row
                row.eachCell((cell) => {
                    headers.push(cell.value);
                });
            } else {
                // Data rows
                const rowData = {};
                row.eachCell((cell, colNumber) => {
                    const header = headers[colNumber - 1];
                    if (header) {
                        rowData[header] = cell.value;
                    }
                });
                if (Object.keys(rowData).length > 0) {
                    rawData.push(rowData);
                }
            }
        });
        
        if (rawData.length === 0) {
            throw new Error('No data found in the sheet.');
        }
        
        // Step 3: Filter columns
        showProgress('Filtering columns...');
        updateProgress(65);
        
        const filteredData = rawData.map(row => {
            const newRow = {};
            COLUMNS_TO_KEEP.forEach(col => {
                newRow[col] = row[col] !== undefined ? row[col] : null;
            });
            return newRow;
        });
        
        // Step 4: Sort by Sold Price (descending)
        showProgress('Sorting by price...');
        updateProgress(70);
        
        filteredData.sort((a, b) => {
            const priceA = parseFloat(a['Sold Price']) || 0;
            const priceB = parseFloat(b['Sold Price']) || 0;
            return priceB - priceA;
        });
        
        // Step 5: Calculate quartiles
        showProgress('Calculating quartiles...');
        updateProgress(75);
        
        const nRows = filteredData.length;
        const qSize = Math.ceil(nRows / 4);
        const q1End = qSize;
        const q2End = 2 * qSize;
        const q3End = 3 * qSize;

        // Calculate stats
        const prices = filteredData.map(row => parseFloat(row['Sold Price']) || 0);
        const minPrice = Math.min(...prices);
        const maxPrice = Math.max(...prices);

        // Calculate average prices for each quartile
        const q1Prices = prices.slice(0, q1End);
        const q2Prices = prices.slice(q1End, q2End);
        const q3Prices = prices.slice(q2End, q3End);
        const q4Prices = prices.slice(q3End);

        const avgQ1 = Math.round(q1Prices.reduce((a, b) => a + b, 0) / q1Prices.length);
        const avgQ2 = Math.round(q2Prices.reduce((a, b) => a + b, 0) / q2Prices.length);
        const avgQ3 = Math.round(q3Prices.reduce((a, b) => a + b, 0) / q3Prices.length);
        const avgQ4 = Math.round(q4Prices.reduce((a, b) => a + b, 0) / q4Prices.length);

        statsData = {
            recordCount: nRows,
            priceRange: `$${minPrice.toLocaleString()} - $${maxPrice.toLocaleString()}`,
            quartileAvgPrices: `Q1: $${avgQ1.toLocaleString()}\nQ2: $${avgQ2.toLocaleString()}\nQ3: $${avgQ3.toLocaleString()}\nQ4: $${avgQ4.toLocaleString()}`
        };

        // Step 6: Create formatted workbook
        showProgress('Creating formatted workbook...');
        updateProgress(80);
        
        processedWorkbook = await createFormattedWorkbook(filteredData, q1End, q2End, q3End);

        // Step 7: Done!
        updateProgress(100);
        showResults();

    } catch (error) {
        console.error(error);
        showError(`Error transforming data: ${error.message}`);
    }
}

async function createFormattedWorkbook(data, q1End, q2End, q3End) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Existing Comps');
    
    // Set column widths
    const columnWidths = [5, 5, 12, 15, 10, 15, 12, 18, 18, 15, 12, 12, 15, 15, 18, 12, 15];
    worksheet.columns = columnWidths.map(width => ({ width }));
    
    // Row 2: Title
    const titleCell = worksheet.getCell('C2');
    titleCell.value = 'Existing Sold Comps';
    titleCell.font = { bold: true, size: 14 };
    
    // Row 3: Subdivision and quartile headers
    const subdivCell = worksheet.getCell('C3');
    subdivCell.value = 'Pagoda Grove Circle';
    subdivCell.font = { bold: true };
    
    const quartileHeaders = [
        { cell: 'I3', value: '1st Quartile' },
        { cell: 'J3', value: '2nd Quartile' },
        { cell: 'K3', value: '3rd Quartile' },
        { cell: 'L3', value: '4th Quartile' }
    ];
    
    quartileHeaders.forEach(({ cell, value }) => {
        const c = worksheet.getCell(cell);
        c.value = value;
        c.font = { bold: true };
        c.alignment = { horizontal: 'center' };
    });
    
    // Row 5: Count
    worksheet.getCell('C5').value = 'Count';
    worksheet.getCell('D5').value = data.length;
    
    // Row 6: Quartile Size
    worksheet.getCell('C6').value = 'Quartile Size';
    worksheet.getCell('D6').value = { formula: 'D5/4' };
    
    // Row 7-11: Criteria
    worksheet.getCell('C7').value = 'Criteria';
    worksheet.getCell('C8').value = 'Sold last year';
    worksheet.getCell('C9').value = 'South of 7800, West of 2200, N';
    worksheet.getCell('C10').value = 'SFH, not manufactured';
    const sortedCell = worksheet.getCell('C11');
    sortedCell.value = 'Sorted by Sold Price';
    sortedCell.font = { bold: true };
    
    // Statistics with formulas
    const stats = [
        { label: 'Avg Sold Price', col: 'L', row: 4 },
        { label: 'Avg SF', col: 'O', row: 5 },
        { label: 'Avg Bed', col: 'M', row: 6 },
        { label: 'Avg Year Built', col: 'P', row: 7 },
        { label: 'Avg Acres', col: 'C', row: 8 },
        { label: 'Avg DOM', col: 'E', row: 9 },
        { label: 'Avg Price/SF', col: 'I', row: 10 }
    ];
    
    stats.forEach(stat => {
        const labelCell = worksheet.getCell(`H${stat.row}`);
        labelCell.value = stat.label;
        labelCell.font = { bold: true };
        
        worksheet.getCell(`I${stat.row}`).value = { formula: `AVERAGE(${stat.col}$15:${stat.col}$${14 + q1End})` };
        worksheet.getCell(`J${stat.row}`).value = { formula: `AVERAGE(${stat.col}$${15 + q1End}:${stat.col}$${14 + q2End})` };
        worksheet.getCell(`K${stat.row}`).value = { formula: `AVERAGE(${stat.col}$${15 + q2End}:${stat.col}$${14 + q3End})` };
        worksheet.getCell(`L${stat.row}`).value = { formula: `AVERAGE(${stat.col}$${15 + q3End}:${stat.col}$${14 + data.length})` };
    });
    
    // Row 14: Column headers
    worksheet.getRow(14).values = ['', '', ...COLUMNS_TO_KEEP];
    worksheet.getRow(14).font = { bold: true };
    
    // Data rows starting at row 15
    data.forEach((row, rowIdx) => {
        const excelRow = rowIdx + 15;
        
        // Determine quartile for coloring
        let quartile;
        if (rowIdx < q1End) quartile = 1;
        else if (rowIdx < q2End) quartile = 2;
        else if (rowIdx < q3End) quartile = 3;
        else quartile = 4;
        
        // Set row values
        const rowValues = ['', '', ...COLUMNS_TO_KEEP.map(col => row[col])];
        worksheet.getRow(excelRow).values = rowValues;
        
        // Apply quartile coloring (columns C through Q)
        for (let colIdx = 3; colIdx <= 3 + COLUMNS_TO_KEEP.length - 1; colIdx++) {
            const cell = worksheet.getRow(excelRow).getCell(colIdx);
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + QUARTILE_COLORS[quartile] }
            };
        }
    });
    
    return workbook;
}

async function downloadFile() {
    if (!processedWorkbook) {
        showError('No processed file available. Please process a file first.');
        return;
    }

    try {
        // Generate output filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 10);
        const filename = `Existing_Comps_Transformed_${timestamp}.xlsx`;

        // Write workbook to buffer
        const buffer = await processedWorkbook.xlsx.writeBuffer();
        
        // Create blob and download
        const blob = new Blob([buffer], {
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });
        
        saveAs(blob, filename);

    } catch (error) {
        showError(`Error downloading file: ${error.message}`);
    }
}

function showProgress(text) {
    document.querySelector('.upload-section').style.display = 'none';
    document.getElementById('progressSection').style.display = 'block';
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';
    document.getElementById('progressText').textContent = text;
}

function updateProgress(percent) {
    document.getElementById('progressFill').style.width = `${percent}%`;
}

function showResults() {
    document.querySelector('.upload-section').style.display = 'none';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'block';
    document.getElementById('errorSection').style.display = 'none';

    // Update stats
    document.getElementById('recordCount').textContent = statsData.recordCount;
    document.getElementById('priceRange').textContent = statsData.priceRange;
    document.getElementById('quartileAvgPrices').textContent = statsData.quartileAvgPrices;
}

function showError(message) {
    document.querySelector('.upload-section').style.display = 'none';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'block';
    document.getElementById('errorText').textContent = message;
}

function resetApp() {
    document.querySelector('.upload-section').style.display = 'block';
    document.getElementById('progressSection').style.display = 'none';
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('errorSection').style.display = 'none';
    document.getElementById('fileName').textContent = '';
    document.getElementById('fileInput').value = '';
    processedWorkbook = null;
    statsData = null;
    updateProgress(0);
}
