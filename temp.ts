function main(workbook: ExcelScript.Workbook) {
  // Get the output worksheet
  const outputSheetName = "Pavilion Rental Report";
  let outputSheet = workbook.getWorksheet(outputSheetName);

  if (!outputSheet) {
    outputSheet = workbook.addWorksheet(outputSheetName);
  }

  // Get or create the table
  let table: ExcelScript.Table;
  const tables = outputSheet.getTables();
  if (tables.length > 0) {
    table = tables[0]; // Use existing table
  } else {
    table = outputSheet.addTable(outputSheet.getRange("A1:C1"), true);
    table.getHeaderRowRange().setValues([["Sheet Name", "Taxable Rentals", "Non-Taxable Rentals"]]);
  }

  // Clear existing rows (excluding headers)
  const dataRange = table.getRangeBetweenHeaderAndTotal();
  if (dataRange.getRowCount() > 0) {
    dataRange.clear(ExcelScript.ClearApplyTo.all);
  }

  // Get all worksheets in the workbook
  const sheets = workbook.getWorksheets();
  const tableData: (string | number)[][] = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();

    // Check if the sheet follows the naming pattern "Daily balancing report (X)"
    if (sheetName.startsWith("Daily balancing report") && sheetName.includes("(") && sheetName.includes(")")) {
      const taxableRentals = sheet.getRange("C11").getValue(); // Taxable Rentals
      const nonTaxableRentals = sheet.getRange("C33").getValue(); // Non-Taxable Rentals

      // Ensure values are numbers
      const taxableValue = typeof taxableRentals === "number" ? taxableRentals : 0;
      const nonTaxableValue = typeof nonTaxableRentals === "number" ? nonTaxableRentals : 0;

      // Add data to the table array
      tableData.push([sheetName, taxableValue, nonTaxableValue]);
    }
  });

  // Insert data into the table
  table.addRows(null, tableData);

  // Add AutoSum at the bottom
  const totalRowIndex = tableData.length + 1; // One row after last entry
  table.addRows(null, [["Total", `=SUM(B2:B${totalRowIndex})`, `=SUM(C2:C${totalRowIndex})`]]);

  // Format columns B and C as numbers with two decimal places
  table.getRange().getColumn(1).setNumberFormat("0.00"); // Taxable Rentals
  table.getRange().getColumn(2).setNumberFormat("0.00"); // Non-Taxable Rentals

  // Auto-fit columns for readability
  outputSheet.getRange("A:C").getFormat().autofitColumns();
}

