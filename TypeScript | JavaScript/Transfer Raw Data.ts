
function main(workbook: ExcelScript.Workbook, datastartIndex: number, batchSize: number): { batch: string[][], totalRows: number} {
  const sheet = workbook.getWorksheet("Query Output");
  // const batchSize = 5000;
  const skipHeaderRow = true;
  const usedRange = sheet.getUsedRange();
  const allData = usedRange.getValues();
  const totalRows = allData.length - 1;
  const adjustedStart = datastartIndex + 1;
  const endRow = Math.min(datastartIndex + batchSize, allData.length);
    
  const processedData: string[][] = [];
  for (let i = adjustedStart; i < endRow; i++) {
    const row = allData[i];

    const newRow = row.map((cell, idx) => {
        if (idx === 12 || idx === 13) {
          return formatDate(cell);
        }
        return cell?.toString() ?? "";
      });
      processedData.push(newRow);
    }
    return {
      batch: processedData,
      totalRows: totalRows
    };  
  }
function formatDate(cell: string | Date): string {
  if (cell instanceof Date) {
    const d = cell as Date;
    return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
    }
    if (typeof cell === "string") {
      const parsed = new Date(cell);
      if (!isNaN(parsed.getTime())) {
        return `${String(parsed.getDate()).padStart(2, "0")}/${String(parsed.getMonth() + 1).padStart(2, "0")}/${parsed.getFullYear()}`;
      }
    }
  return cell?.toString() ?? "";
}
