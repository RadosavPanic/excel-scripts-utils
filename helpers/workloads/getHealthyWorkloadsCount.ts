// @ts-nocheck

/**
 * @function getHealthyWorkloadsCount 
 * @description Function that obtains and returns number of healthy workloads.
 * @param {ExcelScript.Worksheet} selectedSheet Sheet object selected, default: current working sheet.
 * @param {string} workloadName Name of the workload to search for.
 * @param {string} targetCell Cell where formula is written and value is returned to.
 * @param {string} columnChar Column where the query will be executed, default: 'T'
 * @returns {number} The number of healthy workloads.
 */
function getHealthyWorkloadsCount(selectedSheet: ExcelScript.Worksheet, workloadName: string, targetCell: string, columnChar: string = "T") {
    let workloadsRange: ExcelScript.Range = selectedSheet.getRange(targetCell);

    workloadsRange.setFormula(`=COUNTIFS(${columnChar}2:${columnChar}100, "*${workloadName}*", ${columnChar}2:${columnChar}100, "*is installed and healthy*")`);

    const numHealthyWorkloads: number = workloadsRange.getValue() as number;
    
    console.log(`Healthy workloads count (${workloadName}): ${numHealthyWorkloads}`);

    return numHealthyWorkloads;
}

function main(workbook: ExcelScript.Workbook) {    
    let selectedSheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();
    
    let targetCell: string = "U36";
    let workloadSearchName: string = "cve-2026-31431"; 
    
    getHealthyWorkloadsCount(selectedSheet, workloadSearchName, targetCell);

    const healthyWorkloads = getHealthyWorkloadsCount(selectedSheet, workloadSearchName, targetCell); 
}