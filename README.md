# Existing Comps Transformer

Automates the transformation of raw MLS data (from the "Existing Comps Data" tab) into a formatted, quartile-analyzed spreadsheet.

## 🚀 Two Versions Available

### 🌐 Web Version (Recommended for Non-Technical Users)
- **Browser-based tool** - works on any device
- **No installation required** - just upload and download
- **Fully private** - processing happens in your browser
- **Deploy to GitHub Pages** - accessible from anywhere
- 👉 **See [DEPLOYMENT.md](DEPLOYMENT.md) for setup instructions**

### 🐍 Python Script (For Developers)
- Command-line tool for batch processing
- Requires Python and dependencies
- Great for automation and integration

## Features

The script performs the following transformations:

1. **Column Selection**: Extracts only the relevant columns from the raw 176-column dataset:
   - Acres, City, DOM, Garage Capacity
   - List Price, Original List Price, Price Per Square Foot
   - Sold Concessions, Sold Date, Sold Price
   - Total Bedrooms, Total Bathrooms, Total Square Feet, Year Built, Property Type

2. **Data Sorting**: Sorts all properties by Sold Price in descending order (highest to lowest)

3. **Quartile Division**: Divides the sorted data into 4 quartiles based on sold price

4. **Color Coding**: Applies distinctive green color gradients to each quartile:
   - **1st Quartile** (highest prices): Light green (#E2EFD9)
   - **2nd Quartile**: Medium-light green (#C5E0B3)
   - **3rd Quartile**: Medium green (#A8D08D)
   - **4th Quartile** (lowest prices): Dark green (#548135)

5. **Statistical Analysis**: Adds a header section with:
   - Property count and quartile size calculations
   - Average statistics for each quartile:
     - Avg Sold Price
     - Avg Square Footage
     - Avg Bedrooms
     - Avg Year Built
     - Avg Acres
     - Avg Days On Market (DOM)
     - Avg Price Per Square Foot

6. **Formatting**: Applies professional formatting including:
   - Bold headers
   - Appropriate column widths
   - Excel formulas for dynamic calculations

## Setup

### 1. Create Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

The required packages are:
- `openpyxl==3.1.2` - For Excel file manipulation with formatting
- `pandas>=2.2.0` - For data processing

## Usage

### Basic Usage

Run the script with default settings (reads from `Pagoda Grove_West Jordan_Model.xlsx` and outputs to `Existing Comps_Transformed.xlsx`):

```bash
python transform_comps.py
```

### Custom Input/Output Files

You can specify custom input and output files:

```bash
python transform_comps.py input_file.xlsx output_file.xlsx
```

### Example Output

```
================================================================================
EXISTING COMPS TRANSFORMATION SCRIPT
================================================================================

1. Reading data from 'Existing Comps Data' tab...
   Loaded 159 rows, 176 columns

2. Selecting relevant columns...
   Kept 15 columns

3. Sorting by Sold Price (descending)...
   Price range: $375,000 - $1,550,000

4. Creating output workbook...

5. Adding header section...

6. Adding column headers...

7. Writing data...

8. Applying quartile colors...
   Quartile 1: Rows 15-54 (40 rows)
   Quartile 2: Rows 55-94 (40 rows)
   Quartile 3: Rows 95-134 (40 rows)
   Quartile 4: Rows 135-173 (39 rows)

9. Adjusting column widths...

10. Saving to 'Existing Comps_Transformed.xlsx'...
   ✓ Successfully saved!

================================================================================
TRANSFORMATION COMPLETE
================================================================================
```

## 🌐 Using the Web Version

### Quick Start (Local Testing)
1. Open `index.html` in any modern browser
2. Drag and drop your Excel file or click "Browse Files"
3. Wait for processing (happens instantly in browser)
4. Click "Download Transformed File"

### Deploy to GitHub Pages
See [DEPLOYMENT.md](DEPLOYMENT.md) for complete instructions on hosting this for free on GitHub Pages.

### Web Version Features
- ✅ Beautiful drag-and-drop interface
- ✅ Real-time progress indicators
- ✅ Instant processing (no server required)
- ✅ Works on desktop, tablet, and mobile
- ✅ All processing happens locally - your files never leave your device

## File Structure

```
.
├── index.html                      # Web app UI
├── styles.css                      # Web app styling
├── transform.js                    # JavaScript transformation logic
├── DEPLOYMENT.md                   # GitHub Pages deployment guide
├── .venv/                          # Virtual environment (not in git)
├── requirements.txt                # Python dependencies
├── transform_comps.py              # Python transformation script
├── run_transform.sh                # Python convenience runner
├── README.md                       # This file
├── Pagoda Grove_West Jordan_Model.xlsx  # Sample input file
└── Existing Comps_Transformed.xlsx      # Sample output file
```

## How It Works

### Input
The script reads from the **"Existing Comps Data"** tab, which contains:
- 159 property records
- 176 columns of MLS data
- Raw, unsorted data

### Processing
1. Filters to 15 relevant columns
2. Sorts by Sold Price (descending)
3. Divides into quartiles (40, 40, 40, 39 rows for 159 total)
4. Generates Excel formulas for quartile averages
5. Uses "Existing Comps Data" sheet, or the only sheet if just one exists

### Output
Creates a new Excel file with the **"Existing Comps"** tab containing:
- Header section with title and statistics (rows 1-14)
- Formatted data table starting at row 15
- Color-coded quartiles
- Dynamic formulas that update with the data

## Customization

To customize the script for different datasets:

1. **Columns**: Edit the `columns_to_keep` list in the script (line ~48)
2. **Colors**: Modify the `colors` dictionary (line ~158)
3. **Header Info**: Update rows 2-11 for different project names or criteria
4. **Statistics**: Add or remove items from the `stats` list (line ~119)

## Notes

- The script uses ceiling division for quartiles, so if the total count isn't evenly divisible by 4, the first three quartiles get the extra row
- All Excel formulas use absolute row references (e.g., `$15`) to ensure they work correctly when copied
- The script preserves the exact formatting and color scheme from the original manually-created sheet

## Troubleshooting

**Issue**: `ModuleNotFoundError`
- **Solution**: Make sure you've activated the virtual environment and installed requirements

**Issue**: File not found
- **Solution**: Check that the input Excel file is in the same directory as the script

**Issue**: Wrong sheet name
- **Solution**: Verify the raw data is in a sheet named "Existing Comps Data"

## License

This script is provided as-is for internal use.

