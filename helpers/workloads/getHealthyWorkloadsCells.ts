// @ts-nocheck

function getHealthyWorkloadsCells(selectedSheet: ExcelScript.Worksheet, workloadName: string, targetCell: string, columnChar: string = "T") {
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
    
    const cellAddressesNumbers = cellAddresses.map(el => el.slice(1));

    console.log(`Healthy workloads cell positions (${workloadName}): ${cellAddresses.join(", ")}`);

    return {cellAddresses, cellAddressesNumbers};
}

function main(workbook: ExcelScript.Workbook) {    
    let selectedSheet: ExcelScript.Worksheet = workbook.getActiveWorksheet(); 
    let targetCell: string = "U36";    
    let workloadSearchName: string = "cve-2026-31431";
   
    getHealthyWorkloadsCells(selectedSheet, workloadSearchName, targetCell);

    const {cellAddresses, cellAddressesNumbers} = getHealthyWorkloadsCells(selectedSheet, workloadSearchName, targetCell);
}