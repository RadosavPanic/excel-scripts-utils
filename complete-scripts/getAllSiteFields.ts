// @ts-nocheck

function getAllSiteFields(
  selectedSheet: ExcelScript.Worksheet,
  columnLetter: string,
) {
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
      col === `${columnLetter}:${columnLetter}` ? index : -1,
    )
    .filter((index) => index !== -1);

  const filteredRowAddresses = columnAIndices.map(
    (index) => entireRowAddress[index],
  );

  const siteFullInfos = filteredRowAddresses.map((range) => {
    const [startRowNum, endRowNum] = range.split(":");

    let masterCell = `${columnLetter}${startRowNum}`;
    let cell = selectedSheet.getRange(masterCell);
    const siteName: string = cell.getValue() as string;

    cell = selectedSheet.getRange(`B${startRowNum}`);
    const clusterLabels: string[] = cell.getValue().toString().split(", ");

    cell = selectedSheet.getRange(`C${startRowNum}`);
    const siteDesc: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`D${startRowNum}`);
    const clusterId: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`E${startRowNum}`);
    const infraVersion: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`F${startRowNum}`);
    const osVersion: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`G${startRowNum}`);
    const infraStatus: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`H${startRowNum}`);
    const infraStatusMsg: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`I${startRowNum}`);
    const clusterStatus: string = cell.getValue().toString();

    cell = selectedSheet.getRange(`J${startRowNum}`);
    const clusterStatusMsg: string = cell.getValue().toString();

    const nodes = getStoreNodes(selectedSheet, "V36", +startRowNum, +endRowNum);

    return {
      siteName,
      siteDesc,
      clusterId,
      clusterLabels,
      clusterStatus,
      clusterStatusMsg,
      infraStatus,
      infraStatusMsg,
      infraVersion,
      osVersion,
      nodes,
    };
  });

  return siteFullInfos;
}

function getStoreNodes(
  selectedSheet: ExcelScript.Worksheet,
  targetCell: string,
  startRow: number,
  endRow: number,
  columnChar = "K",
) {
  let range: ExcelScript.Range;
  const nodeNames: string[] = [];

  for (let i = startRow; i <= endRow; i++) {
    range = selectedSheet.getRange(`${columnChar}${i}`);
    nodeNames.push(range.getValue().toString());
  }

  return nodeNames;
}

function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  getAllSiteFields(selectedSheet, "A");
}
