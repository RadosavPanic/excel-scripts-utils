// @ts-nocheck

function writeDataToNewSheet(
  workbook: ExcelScript.Workbook,
  filteredData: Object[],
  headersList: string[],
) {
  const today = new Date();
  const day = today.getDate().toString().padStart(2, "0");
  const month = today.toLocaleString("default", { month: "short" });
  const year = today.getFullYear();

  const sheetName = `HW Inventory ${day}-${month}-${year}`;
  let newSheet = workbook.getWorksheet(sheetName);
  let allWorksheets = workbook.getWorksheets();

  if (!newSheet) newSheet = workbook.addWorksheet(sheetName);
  else {
    newSheet.delete();
    newSheet = workbook.addWorksheet(sheetName);
  }

  if (filteredData.length === 0) return;

  allWorksheets.forEach((sheet) => {
    sheet.delete();
  });

  newSheet
    .getRangeByIndexes(0, 0, 1, headersList.length)
    .setValues([headersList]);

  const dataRows = filteredData.map((row) =>
    headersList.map((header) => row[header]),
  );

  if (dataRows.length > 0) {
    newSheet
      .getRangeByIndexes(1, 0, dataRows.length, headersList.length)
      .setValues(dataRows);
  }

  const dataRange = newSheet.getUsedRange();
  const table = workbook.addTable(dataRange, true);

  newSheet.getUsedRange()?.getFormat().autofitColumns();
}

function getFilteredMetricsGCPData(
  selectedSheet: ExcelScript.Worksheet,
  headersList: string[],
) {
  const usedRange = selectedSheet.getUsedRange();
  const headers = usedRange.getColumn(0).getValues();

  const findIndexes = headersList
    .map((field) => {
      const index = headers.findIndex((h) => h.toString() === field);

      return index;
    })
    .filter((index) => index !== -1);

  const dataColumns = usedRange.getOffsetRange(0, 1).getUsedRange().getValues();

  const transposedData: Object[] = [];
  for (let i = 0; i < dataColumns.length; i++) {
    transposedData.push(dataColumns.map((col) => col[i]));
  }

  const filteredData = transposedData
    .map((storeData) => {
      const result = {};
      findIndexes.forEach((headerIndex, i) => {
        const fieldName = headersList[i];
        let value: string = storeData[headerIndex];

        if (fieldName === "bios_date" && typeof value === "number") {
          const excelEpoch = new Date(Date.UTC(1900, 0, 1));
          const jsDate = new Date(
            excelEpoch.getTime() + (value - 2) * 24 * 60 * 60 * 1000,
          );
          value = jsDate.toLocaleDateString();
        }

        result[fieldName] = value;
      });
      return result;
    })
    .filter((row) =>
      Object.values(row).every(
        (value) => value !== "" && value !== undefined && value !== null,
      ),
    );

  return filteredData;
}

function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();

  const headersList = [
    "board_name",
    "store_name",
    "node",
    "bios_release",
    "bios_date",
    "product_version",
    "bios_version",
    "product_name",
    "product_serial",
  ];

  const filteredData = getFilteredMetricsGCPData(selectedSheet, headersList);
  writeDataToNewSheet(workbook, filteredData, headersList);
}
