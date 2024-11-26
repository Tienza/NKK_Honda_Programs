'use strict'

import reader from 'xlsx';
import * as path from 'path';

// Directory constants
const RESOURCE_DIR = './resources';
const OUPUT_DIR = './output';
// File name constants
const PRICE_LIST_FILE_NAME = 'price_list_source.xlsx';
const OUTPUT_FILE_NAME = () => path.join(OUPUT_DIR, `${(new Date()).toISOString()
    .replaceAll(/[\*"/\<>:|?-]/g, '_')}_price_list.xlsx`);
// File constants
const PRICE_LIST_FILE = reader.readFile(path.join(RESOURCE_DIR, PRICE_LIST_FILE_NAME));
// Required column names
const ColumnName = {
    model: 'รุ่น',
    make: 'ชื่อทางการตลาด',
    specialParts: 'รุ่นพิเศษมีอะไหล่แต่ง',
    cost: 'ทุน',
}
/**
 * Generates a template Excel file from a price list by processing a source Excel file.
 *
 * This function performs the following steps:
 * 1. Reads the source price list Excel file and extracts its rows.
 * 2. Filters out invalid rows that do not have a defined `cost` attribute.
 * 3. Maps the filtered rows to a new format containing specific columns: 
 *    `model`, `make`, `specialParts`, and `cost`.
 * 4. Creates a new Excel workbook, appends the processed data, and writes it to an output file.
 *
 * For more information on writing JSON to Excel, refer to Stack Overflow:
 * https://stackoverflow.com/a/66403522
 */
const generateTemplateFileFromPriceList = () => {
    // Read the price list source Excel file and extract the information
    let rows = reader.utils.sheet_to_json(
        PRICE_LIST_FILE.Sheets[PRICE_LIST_FILE.SheetNames[0]]
    );

    // Filter out rows that don't have the `cost` attribute as these are invalid rows
    rows = rows.filter((row) => row[ColumnName.cost]);

    // Create an array for Excel rows containing the desired four columns
    const modelData = rows.map((row) => {
        return {
            [ColumnName.model]: row[ColumnName.model],
            [ColumnName.make]: row[ColumnName.make] ?? '',
            [ColumnName.specialParts]: row[ColumnName.specialParts] ?? '',
            [ColumnName.cost]: row[ColumnName.cost]
        };
    });

    // Write the array to an Excel sheet
    const priceListExcelWorkBook = reader.utils.book_new();
    const priceListExcelWorksheet = reader.utils.json_to_sheet(modelData);
    reader.utils.book_append_sheet(priceListExcelWorkBook, priceListExcelWorksheet, 'price_list');
    reader.writeFile(priceListExcelWorkBook, OUTPUT_FILE_NAME());
};


generateTemplateFileFromPriceList();
