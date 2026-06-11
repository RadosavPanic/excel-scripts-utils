// @ts-nocheck

function getMergedCellRowCount(
  selectedSheet: ExcelScript.Worksheet,
  targetCell: string,
) {
  let masterCell: string;
  let tempColumnLetter = targetCell.slice(0, 1);

  const allMergedAreas = selectedSheet.getUsedRange().getMergedAreas();

  const entireColumns = allMergedAreas.getEntireColumn();
  const entireColumnAddress = entireColumns
    .getAddress()
    .split(",")
    .map((s) => s.trim().split("!")[1]);

  const entireRows = allMergedAreas.getEntireRow();
  const entireRowAddress = entireRows
    .getAddress()
    .split(",")
    .map((s) => s.trim().split("!")[1]);

  const columnAIndices = entireColumnAddress
    .map((col, index) =>
      col === `${tempColumnLetter}:${tempColumnLetter}` ? index : -1,
    )
    .filter((index) => index !== -1);

  const filteredRowAddresses = columnAIndices.map(
    (index) => entireRowAddress[index],
  );

  filteredRowAddresses.forEach((rowRange) => {
    const [startRow, endRow] = rowRange.split(":").map((el) => Number(el));

    const rangeAddress = `${tempColumnLetter}${startRow}:${tempColumnLetter}${endRow}`;

    const colANumbers: number[] = [];
    for (let i = startRow; i <= endRow; i++) colANumbers.push(i);
    const addressesColA = colANumbers.map((el) => `${tempColumnLetter}${el}`);

    if (addressesColA.includes(targetCell)) {
      masterCell = `${tempColumnLetter}${startRow}`;
    }
  });

  const cell = selectedSheet.getRange(masterCell);

  const siteName: string = cell.getValue() as string;

  const mergedArea = cell.getMergedAreas();

  if (!mergedArea) {
    console.log(`The cell ${targetCell} is not merged.`);
    return;
  }

  const rowCount = mergedArea.getCellCount();

  const rangeAddress = mergedArea.getEntireRow().getAddress();
  const rowPart = rangeAddress.split("!")[1];
  const [startRow, endRow] = rowPart.split(":").map((el) => Number(el));

  const rowNumbers: number[] = [];
  for (let i = startRow; i <= endRow; i++) rowNumbers.push(i);

  let tempCellLetter = targetCell.slice(0, 1);

  const rowAddresses = rowNumbers.map((el) => `${tempCellLetter}${el}`);

  return { startRow, endRow, rowAddresses, rowNumbers, siteName };
}

function getHealthyWorkloadsCells(
  selectedSheet: ExcelScript.Worksheet,
  workloadName: string,
  targetCell: string,
  columnChar: string = "U",
) {
  const searchRange = `${columnChar}2:${columnChar}100`;

  let workloadsRange: ExcelScript.Range = selectedSheet.getRange(searchRange);
  const values = workloadsRange.getValues() as string[][];

  const cellAddresses: string[] = [];

  values.forEach((val: string[], index) => {
    if (val !== undefined) {
      const cellValue = val.toString().toLowerCase();
      if (
        cellValue.includes(workloadName.toLowerCase()) &&
        cellValue.includes("is installed and healthy")
      ) {
        cellAddresses.push(`${columnChar}${index + 2}`);
      }
    }
  });

  const cellAddressesNumbers = cellAddresses.map((el) => el.slice(1));

  return { cellAddresses, cellAddressesNumbers };
}

function getStoreNodesCount(
  selectedSheet: ExcelScript.Worksheet,
  targetCell: string,
  siteName: string,
  startRow: number,
  endRow: number,
  columnChar = "K",
): number {
  let nodesRange: ExcelScript.Range = selectedSheet.getRange(targetCell);
  let numNodes: number;

  nodesRange.setFormula(
    `=COUNTIF(${columnChar}${startRow}:${columnChar}${endRow}, "*-*")`,
  );

  numNodes = nodesRange.getValue() as number;

  return numNodes;
}

function main(workbook: ExcelScript.Workbook) {
  let selectedSheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
  let targetCell: string = "V36";
  let workloadSearchName: string = "cve-2026-31431";

  let totalNodes: number = 0;
  const totalStores: string[] = [];

  const { cellAddresses, cellAddressesNumbers } = getHealthyWorkloadsCells(
    selectedSheet,
    workloadSearchName,
    targetCell,
  );

  if (!cellAddressesNumbers) return;

  const targetStoreCells = cellAddressesNumbers.map((num) => `A${num}`);

  targetStoreCells.forEach((cell) => {
    const { startRow, endRow, siteName } = getMergedCellRowCount(
      selectedSheet,
      cell,
    );

    const numNodes = getStoreNodesCount(
      selectedSheet,
      "U37",
      siteName,
      startRow,
      endRow,
    );

    totalNodes += numNodes;

    if (!totalStores.includes(siteName)) totalStores.push(siteName);

    console.log(
      `Site name: ${siteName} (span: A${startRow}:A${endRow})\nTotal nodes: ${numNodes}\nTargeted workload: ${workloadSearchName}`,
    );
  });

  console.log(
    `Site group: ${workbook.getName().split("_")[0]}\n${totalStores.length} stores targeted\n${totalNodes} nodes targeted`,
  );
}
