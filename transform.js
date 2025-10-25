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
    reader.onload = async function(e) {
        try {
            updateProgress(30);
            showProgress('Processing data...');
            
            const data = e.target.result;
            await transformData(data);
            
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
        // Step 1: Read the workbook with SheetJS (supports both .xls and .xlsx)
        showProgress('Loading raw data...');
        updateProgress(40);
        
        const workbook = XLSX.read(data, { type: 'array' });
        
        updateProgress(50);
        
        // Find the data sheet
        let sheetName = 'Existing Comps Data';
        
        // If "Existing Comps Data" doesn't exist, try to use the only sheet if there's just one
        if (!workbook.SheetNames.includes(sheetName)) {
            if (workbook.SheetNames.length === 1) {
                sheetName = workbook.SheetNames[0];
                console.log(`Using single sheet: "${sheetName}"`);
            } else {
                throw new Error(`Sheet "Existing Comps Data" not found. Your file has ${workbook.SheetNames.length} sheets: ${workbook.SheetNames.join(', ')}. Please make sure your raw data sheet is named "Existing Comps Data".`);
            }
        }
        
        const worksheet = workbook.Sheets[sheetName];
        
        // Step 2: Extract data
        showProgress('Extracting data...');
        updateProgress(60);
        
        const rawData = XLSX.utils.sheet_to_json(worksheet);
        
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
    
    // Row 3: Date and quartile headers
    const dateCell = worksheet.getCell('C3');
    const today = new Date();
    const formattedDate = today.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric' 
    });
    dateCell.value = formattedDate;
    dateCell.font = { bold: true };
    
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
    
    // Statistics with formulas
    const stats = [
        { label: 'Avg Sold Price', col: 'L', row: 4, format: '$#,##0' },
        { label: 'Avg SF', col: 'O', row: 5, format: '0' },
        { label: 'Avg Bed', col: 'M', row: 6, format: '0.0' },
        { label: 'Avg Year Built', col: 'P', row: 7, format: '0' },
        { label: 'Avg Acres', col: 'C', row: 8, format: '0.00' },
        { label: 'Avg DOM', col: 'E', row: 9, format: '0.00' },
        { label: 'Avg Price/SF', col: 'I', row: 10, format: '$#,##0.00' }
    ];
    
    stats.forEach(stat => {
        const labelCell = worksheet.getCell(`H${stat.row}`);
        labelCell.value = stat.label;
        labelCell.font = { bold: true };
        
        // Set formulas and apply number formatting
        const cells = [
            worksheet.getCell(`I${stat.row}`),
            worksheet.getCell(`J${stat.row}`),
            worksheet.getCell(`K${stat.row}`),
            worksheet.getCell(`L${stat.row}`)
        ];
        
        cells[0].value = { formula: `AVERAGE(${stat.col}$15:${stat.col}$${14 + q1End})` };
        cells[1].value = { formula: `AVERAGE(${stat.col}$${15 + q1End}:${stat.col}$${14 + q2End})` };
        cells[2].value = { formula: `AVERAGE(${stat.col}$${15 + q2End}:${stat.col}$${14 + q3End})` };
        cells[3].value = { formula: `AVERAGE(${stat.col}$${15 + q3End}:${stat.col}$${14 + data.length})` };
        
        // Apply number format to all quartile cells
        cells.forEach(cell => {
            cell.numFmt = stat.format;
        });
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
