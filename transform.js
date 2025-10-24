// Existing Comps Transformer - JavaScript Version
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
            
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            updateProgress(50);
            transformData(workbook);
            
        } catch (error) {
            showError(`Error reading file: ${error.message}`);
        }
    };

    reader.onerror = function() {
        showError('Error reading file. Please try again.');
    };

    reader.readAsArrayBuffer(file);
}

function transformData(workbook) {
    try {
        // Step 1: Find and read the raw data sheet
        showProgress('Loading raw data...');
        updateProgress(60);
        
        const sheetName = 'Existing Comps Data';
        if (!workbook.SheetNames.includes(sheetName)) {
            throw new Error(`Sheet "${sheetName}" not found. Please make sure your file has the correct sheet name.`);
        }

        const rawSheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(rawSheet);
        
        if (rawData.length === 0) {
            throw new Error('No data found in the sheet.');
        }

        // Step 2: Filter to relevant columns
        showProgress('Filtering columns...');
        updateProgress(70);
        
        const filteredData = rawData.map(row => {
            const newRow = {};
            COLUMNS_TO_KEEP.forEach(col => {
                newRow[col] = row[col] !== undefined ? row[col] : null;
            });
            return newRow;
        });

        // Step 3: Sort by Sold Price (descending)
        showProgress('Sorting by price...');
        updateProgress(75);
        
        filteredData.sort((a, b) => {
            const priceA = parseFloat(a['Sold Price']) || 0;
            const priceB = parseFloat(b['Sold Price']) || 0;
            return priceB - priceA;
        });

        // Step 4: Calculate quartiles
        showProgress('Calculating quartiles...');
        updateProgress(80);
        
        const nRows = filteredData.length;
        const qSize = Math.ceil(nRows / 4);
        const q1End = qSize;
        const q2End = 2 * qSize;
        const q3End = 3 * qSize;

        // Calculate stats
        const prices = filteredData.map(row => parseFloat(row['Sold Price']) || 0);
        const minPrice = Math.min(...prices);
        const maxPrice = Math.max(...prices);

        statsData = {
            recordCount: nRows,
            priceRange: `$${minPrice.toLocaleString()} - $${maxPrice.toLocaleString()}`,
            quartileSizes: `Q1:${q1End}, Q2:${q2End - q1End}, Q3:${q3End - q2End}, Q4:${nRows - q3End}`
        };

        // Step 5: Create new workbook
        showProgress('Creating formatted workbook...');
        updateProgress(85);
        
        processedWorkbook = createFormattedWorkbook(filteredData, q1End, q2End, q3End);

        // Step 6: Done!
        updateProgress(100);
        showResults();

    } catch (error) {
        showError(`Error transforming data: ${error.message}`);
    }
}

function createFormattedWorkbook(data, q1End, q2End, q3End) {
    const wb = XLSX.utils.book_new();
    const ws = {};
    
    // Set column widths
    const colWidths = [
        { wch: 5 },   // A
        { wch: 5 },   // B
        { wch: 12 },  // C - Acres
        { wch: 15 },  // D - City
        { wch: 10 },  // E - DOM
        { wch: 15 },  // F - Garage Capacity
        { wch: 12 },  // G - List Price
        { wch: 18 },  // H - Original List Price
        { wch: 18 },  // I - Price Per Square Foot
        { wch: 15 },  // J - Sold Concessions
        { wch: 12 },  // K - Sold Date
        { wch: 12 },  // L - Sold Price
        { wch: 15 },  // M - Total Bedrooms
        { wch: 15 },  // N - Total Bathrooms
        { wch: 18 },  // O - Total Square Feet
        { wch: 12 }   // P - Year Built
    ];
    ws['!cols'] = colWidths;

    // Header section (rows 1-13)
    // Row 2: Title
    setCellValue(ws, 'C2', 'Existing Sold Comps', { bold: true, fontSize: 14 });
    
    // Row 3: Subdivision and quartile headers
    setCellValue(ws, 'C3', 'Pagoda Grove Circle', { bold: true });
    setCellValue(ws, 'I3', '1st Quartile', { bold: true, align: 'center' });
    setCellValue(ws, 'J3', '2nd Quartile', { bold: true, align: 'center' });
    setCellValue(ws, 'K3', '3rd Quartile', { bold: true, align: 'center' });
    setCellValue(ws, 'L3', '4th Quartile', { bold: true, align: 'center' });

    // Row 5: Count
    setCellValue(ws, 'C5', 'Count');
    setCellValue(ws, 'D5', data.length);

    // Row 6: Quartile Size
    setCellValue(ws, 'C6', 'Quartile Size');
    setCellValue(ws, 'D6', { f: 'D5/4' });

    // Row 7-11: Criteria
    setCellValue(ws, 'C7', 'Criteria');
    setCellValue(ws, 'C8', 'Sold last year');
    setCellValue(ws, 'C9', 'South of 7800, West of 2200, N');
    setCellValue(ws, 'C10', 'SFH, not manufactured');
    setCellValue(ws, 'C11', 'Sorted by Sold Price', { bold: true });

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
        setCellValue(ws, `H${stat.row}`, stat.label, { bold: true });
        setCellValue(ws, `I${stat.row}`, { f: `AVERAGE(${stat.col}$15:${stat.col}$${14 + q1End})` });
        setCellValue(ws, `J${stat.row}`, { f: `AVERAGE(${stat.col}$${15 + q1End}:${stat.col}$${14 + q2End})` });
        setCellValue(ws, `K${stat.row}`, { f: `AVERAGE(${stat.col}$${15 + q2End}:${stat.col}$${14 + q3End})` });
        setCellValue(ws, `L${stat.row}`, { f: `AVERAGE(${stat.col}$${15 + q3End}:${stat.col}$${14 + data.length})` });
    });

    // Row 14: Column headers
    COLUMNS_TO_KEEP.forEach((colName, idx) => {
        const cellRef = XLSX.utils.encode_cell({ r: 13, c: idx + 2 }); // Start at column C (index 2)
        setCellValue(ws, cellRef, colName, { bold: true });
    });

    // Data rows starting at row 15
    data.forEach((row, rowIdx) => {
        const excelRow = rowIdx + 15; // Start at row 15 (0-indexed: row 14)
        
        // Determine quartile for coloring
        let quartile;
        if (rowIdx < q1End) quartile = 1;
        else if (rowIdx < q2End) quartile = 2;
        else if (rowIdx < q3End) quartile = 3;
        else quartile = 4;

        COLUMNS_TO_KEEP.forEach((colName, colIdx) => {
            const cellRef = XLSX.utils.encode_cell({ r: excelRow - 1, c: colIdx + 2 });
            const value = row[colName];
            setCellValue(ws, cellRef, value, { bgColor: QUARTILE_COLORS[quartile] });
        });
    });

    // Set row range
    ws['!ref'] = `A1:P${14 + data.length}`;

    XLSX.utils.book_append_sheet(wb, ws, 'Existing Comps');
    return wb;
}

function setCellValue(ws, cellRef, value, style = {}) {
    const cell = { v: value };

    // Handle formulas
    if (value && typeof value === 'object' && value.f) {
        cell.f = value.f;
        delete cell.v;
    } else {
        // Determine cell type
        if (typeof value === 'number') {
            cell.t = 'n';
        } else if (typeof value === 'boolean') {
            cell.t = 'b';
        } else if (value instanceof Date) {
            cell.t = 'd';
        } else {
            cell.t = 's';
        }
    }

    // Apply styles
    if (Object.keys(style).length > 0) {
        cell.s = {};
        
        if (style.bold) {
            cell.s.font = { bold: true };
            if (style.fontSize) {
                cell.s.font.sz = style.fontSize;
            }
        }
        
        if (style.align) {
            cell.s.alignment = { horizontal: style.align };
        }

        if (style.bgColor) {
            cell.s.fill = {
                fgColor: { rgb: style.bgColor }
            };
        }
    }

    ws[cellRef] = cell;
}

function downloadFile() {
    if (!processedWorkbook) {
        showError('No processed file available. Please process a file first.');
        return;
    }

    try {
        // Generate output filename with timestamp
        const timestamp = new Date().toISOString().slice(0, 10);
        const filename = `Existing_Comps_Transformed_${timestamp}.xlsx`;

        // Write workbook
        const wbout = XLSX.write(processedWorkbook, {
            bookType: 'xlsx',
            type: 'array',
            cellStyles: true
        });

        // Create blob and download
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = filename;
        link.click();
        URL.revokeObjectURL(url);

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
    document.getElementById('quartileSizes').textContent = statsData.quartileSizes;
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

