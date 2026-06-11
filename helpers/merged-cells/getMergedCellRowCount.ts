// @ts-nocheck

function getMergedCellRowCount(selectedSheet: ExcelScript.Worksheet, targetCell: string) {
    let masterCell: string;
    let tempColumnLetter = targetCell.slice(0, 1);

    const allMergedAreas = selectedSheet.getUsedRange().getMergedAreas();

    const entireColumns = allMergedAreas.getEntireColumn();
    const entireColumnAddress = entireColumns.getAddress().split(",").map(s => s.trim().split("!")[1]);

    const entireRows = allMergedAreas.getEntireRow();
    const entireRowAddress = entireRows.getAddress().split(",").map(s => s.trim().split("!")[1]);

    const columnAIndices = entireColumnAddress
        .map((col, index) => col === `${tempColumnLetter}:${tempColumnLetter}` ? index : -1)
        .filter(index => index !== -1);

    const filteredRowAddresses = columnAIndices.map(index => entireRowAddress[index]);

    filteredRowAddresses.forEach(rowRange => {
        const [startRow, endRow] = rowRange.split(":").map(el => Number(el));
        
        const rangeAddress = `${tempColumnLetter}${startRow}:${tempColumnLetter}${endRow}`;

        const colANumbers: number[] = [];
        for (let i = startRow; i <= endRow; i++) colANumbers.push(i);
        const addressesColA = colANumbers.map(el => `${tempColumnLetter}${el}`);

        if(addressesColA.includes(targetCell)){
            masterCell = `${tempColumnLetter}${startRow}`            
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
    const [startRow, endRow] = rowPart.split(":").map(el => Number(el));

    const rowNumbers: number[] = [];
    for (let i = startRow; i <= endRow; i++) rowNumbers.push(i);

    let tempCellLetter = targetCell.slice(0, 1);

    const rowAddresses = rowNumbers.map(el => `${tempCellLetter}${el}`);

    return { startRow, endRow, rowAddresses, rowNumbers, siteName };
}