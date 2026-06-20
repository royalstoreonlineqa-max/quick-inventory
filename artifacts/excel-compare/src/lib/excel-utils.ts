import * as XLSX from "xlsx";

export type ComparisonResult = {
  summary: {
    totalA: number;
    totalB: number;
    matchedCount: number;
    onlyACount: number;
    onlyBCount: number;
    diffCount: number;
  };
  differences: DiffRow[];
  matched: any[];
  onlyA: any[];
  onlyB: any[];
  columns: string[];
  keyCol: string;
};

export type DiffRow = {
  key: string;
  rowA: any;
  rowB: any;
  changedCols: string[];
};

export async function parseExcelFile(file: File): Promise<{ data: any[]; columns: string[] }> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: "array" });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        let columns: string[] = [];
        if (jsonData.length > 0) {
          columns = Object.keys(jsonData[0] as object);
        } else {
          // Fallback if empty but has headers
          const range = XLSX.utils.decode_range(worksheet["!ref"] || "A1:A1");
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const cell = worksheet[XLSX.utils.encode_cell({ c: C, r: range.s.r })];
            if (cell && cell.v) columns.push(String(cell.v));
          }
        }
        
        resolve({ data: jsonData, columns });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
}

export function compareSheets(
  dataA: any[],
  dataB: any[],
  keyCol: string,
  compareCols: string[]
): ComparisonResult {
  const mapA = new Map<string, any>();
  const mapB = new Map<string, any>();

  dataA.forEach(row => {
    const key = String(row[keyCol] || "").trim();
    if (key) mapA.set(key, row);
  });

  dataB.forEach(row => {
    const key = String(row[keyCol] || "").trim();
    if (key) mapB.set(key, row);
  });

  const matched: any[] = [];
  const onlyA: any[] = [];
  const onlyB: any[] = [];
  const differences: DiffRow[] = [];

  for (const [key, rowA] of mapA.entries()) {
    const rowB = mapB.get(key);
    if (!rowB) {
      onlyA.push(rowA);
    } else {
      const changedCols: string[] = [];
      for (const col of compareCols) {
        const valA = String(rowA[col] || "").trim();
        const valB = String(rowB[col] || "").trim();
        if (valA !== valB) {
          changedCols.push(col);
        }
      }
      
      if (changedCols.length > 0) {
        differences.push({ key, rowA, rowB, changedCols });
      } else {
        matched.push(rowA); // Just use A since they are identical on compared cols
      }
    }
  }

  for (const [key, rowB] of mapB.entries()) {
    if (!mapA.has(key)) {
      onlyB.push(rowB);
    }
  }

  return {
    summary: {
      totalA: dataA.length,
      totalB: dataB.length,
      matchedCount: matched.length,
      onlyACount: onlyA.length,
      onlyBCount: onlyB.length,
      diffCount: differences.length,
    },
    differences,
    matched,
    onlyA,
    onlyB,
    columns: compareCols,
    keyCol
  };
}

export function exportResultsToExcel(result: ComparisonResult) {
  const wb = XLSX.utils.book_new();

  // Export Summary
  const summaryData = [
    ["Metric", "Value"],
    ["Total Rows in Sheet A", result.summary.totalA],
    ["Total Rows in Sheet B", result.summary.totalB],
    ["Matched Rows (No changes)", result.summary.matchedCount],
    ["Rows with Differences", result.summary.diffCount],
    ["Rows Only in Sheet A", result.summary.onlyACount],
    ["Rows Only in Sheet B", result.summary.onlyBCount],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  // Export Differences
  if (result.differences.length > 0) {
    const diffExportData: any[] = [];
    result.differences.forEach(diff => {
      const row: any = { [result.keyCol]: diff.key };
      result.columns.forEach(col => {
        if (diff.changedCols.includes(col)) {
          row[`${col} (Sheet A)`] = diff.rowA[col] || "";
          row[`${col} (Sheet B)`] = diff.rowB[col] || "";
        } else {
          row[col] = diff.rowA[col] || "";
        }
      });
      diffExportData.push(row);
    });
    const diffSheet = XLSX.utils.json_to_sheet(diffExportData);
    XLSX.utils.book_append_sheet(wb, diffSheet, "Differences");
  }

  // Export Matched
  if (result.matched.length > 0) {
    const matchedSheet = XLSX.utils.json_to_sheet(result.matched);
    XLSX.utils.book_append_sheet(wb, matchedSheet, "Matched");
  }

  // Export Only in A
  if (result.onlyA.length > 0) {
    const onlyASheet = XLSX.utils.json_to_sheet(result.onlyA);
    XLSX.utils.book_append_sheet(wb, onlyASheet, "Only in A");
  }

  // Export Only in B
  if (result.onlyB.length > 0) {
    const onlyBSheet = XLSX.utils.json_to_sheet(result.onlyB);
    XLSX.utils.book_append_sheet(wb, onlyBSheet, "Only in B");
  }

  XLSX.writeFile(wb, "comparison_results.xlsx");
}
