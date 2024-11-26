# NKK Honda Programs
Scripts and programs to hlep NKK Honda business operations

## Generate Template Excel File from Price List

This script processes a source Excel file to generate a template price list in a specific format. It reads an input file, filters invalid rows, and writes the cleaned data to a new Excel file.

---

## Features

- Reads a source Excel file (`price_list_source.xlsx`) containing price list information.
- Filters out rows with missing `cost` values.
- Maps data into specific columns (`model`, `make`, `specialParts`, `cost`).
- Generates a new Excel file with the processed data.
- Ensures the output file name is unique and timestamped.

---

## Prerequisites

- **Node.js** (version 14 or higher)
- **Dependencies**:
  - `xlsx` for Excel file processing.
  - `path` (built-in Node.js module) for file path management.

Install the required dependency with:
```bash
npm install
```

---

## Directory Structure

```
project/
│
├── resources/       # Contains input files
│   └── price_list_source.xlsx
│
├── output/          # Stores generated output files
│
├── index.js         # Script file (place the code here)
│
└── README.md        # Documentation
```

---

## Usage Instructions

### 1. Prepare Input File
Place the source Excel file in the `resources/` directory with the name `price_list_source.xlsx`.

### 2. Execute the Script
Run the script using Node.js:
```bash
node createPriceListExcel.js
```

### 3. View the Output
The processed file will be saved in the `output/` directory with a name like:
```
2024-11-26T14_35_45.678Z_price_list.xlsx
```

---

## Code Overview

### Key Constants
- **Directories**:
  - `RESOURCE_DIR`: Input file directory (`./resources`).
  - `OUTPUT_DIR`: Output file directory (`./output`).
- **Files**:
  - `PRICE_LIST_FILE_NAME`: Name of the source file.
  - `OUTPUT_FILE_NAME`: Generates a timestamped name for the output file.

### Required Columns
The script processes the following columns from the source file:
- `รุ่น` (model)
- `ชื่อทางการตลาด` (make)
- `รุ่นพิเศษมีอะไหล่แต่ง` (specialParts)
- `ทุน` (cost)

### Main Function
- **`generateTemplateFileFromPriceList`**:
  - Reads data from `price_list_source.xlsx`.
  - Filters rows missing `cost` values.
  - Maps data to a specified format.
  - Writes processed data to a new Excel file.

---

## Example Output
The resulting Excel file will contain data structured with the following columns:

| รุ่น (model) | ชื่อทางการตลาด (make) | รุ่นพิเศษมีอะไหล่แต่ง (specialParts) | ทุน (cost) |
|--------------|-----------------------|---------------------------------------|------------|
| Model A      | Marketing Name A      | Special Part A                        | 1000       |
| Model B      | Marketing Name B      | Special Part B                        | 2000       |

---

## References

- [Writing JSON to Excel](https://stackoverflow.com/a/66403522)

---

## License
This script is provided "as is" without warranty. You may modify and use it as needed.