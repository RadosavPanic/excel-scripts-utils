// @ts-nocheck

/**
 * @function getNodesCount 
 * @description Function that obtains and returns number of nodes for all or individual sites.
 * @param {ExcelScript.Worksheet} selectedSheet Sheet object selected, default: current working sheet.
 * @param {string} targetCell Cell where formula is written and value is returned to.
 * @param {string} siteName Name of the site for nodes to be counted for. 
 * @param {string} columnChar Column where the query will be executed, default: 'J'
 * @returns {number} The number of nodes for all sites.
 */
function getAllNodesCount(selectedSheet: ExcelScript.Worksheet, targetCell: string, siteName: string, columnChar: string = "J"): number {
    let nodesRange: ExcelScript.Range = selectedSheet.getRange(targetCell);
    let numNodes: number;    

    nodesRange.setFormula(`=COUNTIF(${columnChar}2:${columnChar}1000, "*-*")`);

    numNodes = nodesRange.getValue() as number;            

    return numNodes;
}