// @ts-nocheck

/**
 * @function getNodesCount 
 * @description Function that obtains and returns number of nodes for selected site and range.
 * @param {ExcelScript.Worksheet} selectedSheet Sheet object selected, default: current working sheet.
 * @param {string} targetCell Cell where formula is written and value is returned to.
 * @param {string} siteName Name of the site for nodes to be counted for. 
 * @param {number} startRow First index row for the search range.
 * @param {number} endRow Last index row for the search range.
 * @param {string} columnChar Column where the query will be executed, default: 'J'
 * @returns {number} The number of nodes for the selected site range.
 */
function getStoreNodesCount(selectedSheet: ExcelScript.Worksheet, targetCell: string, siteName: string, startRow: number, endRow: number, columnChar = "K"): number {
    let nodesRange: ExcelScript.Range = selectedSheet.getRange(targetCell);
    let numNodes: number;

    nodesRange.setFormula(`=COUNTIF(${columnChar}${startRow}:${columnChar}${endRow}, "*-*")`);

    numNodes = nodesRange.getValue() as number;

    return numNodes;
}