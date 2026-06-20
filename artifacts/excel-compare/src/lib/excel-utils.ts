import * as XLSX from "xlsx";

export type RuleOperator =
  | "equals"
  | "not_equals"
  | "contains"
  | "not_contains"
  | "greater_than"
  | "less_than"
  | "is_empty"
  | "is_not_empty";

export type RuleTarget = "A" | "B" | "both" | "diff_only";

export type Rule = {
  id: string;
  column: string;
  operator: RuleOperator;
  value: string;
  target: RuleTarget;
};

export type ComparisonResult = {
  summary: {
    totalA: number;
    totalB: number;
    matchedCount: number;
    onlyACount: number;
    onlyBCount: number;
    diffCount: number;
    filteredCount: number;
  };
  differences: DiffRow[];
  matched: any[];
  onlyA: any[];
  onlyB: any[];
  columns: string[];
  keyCols: string[];
  appliedRules: Rule[];
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

function makeCompositeKey(row: any, keyCols: string[]): string {
  return keyCols.map(col => String(row[col] ?? "").trim()).join("|||");
}

function applyRule(row: any, rule: Rule): boolean {
  const rawVal = row[rule.column];
  const cellStr = String(rawVal ?? "").trim().toLowerCase();
  const ruleVal = rule.value.trim().toLowerCase();
  const numCell = parseFloat(cellStr);
  const numRule = parseFloat(ruleVal);

  switch (rule.operator) {
    case "equals":
      return cellStr === ruleVal;
    case "not_equals":
      return cellStr !== ruleVal;
    case "contains":
      return cellStr.includes(ruleVal);
    case "not_contains":
      return !cellStr.includes(ruleVal);
    case "greater_than":
      return !isNaN(numCell) && !isNaN(numRule) && numCell > numRule;
    case "less_than":
      return !isNaN(numCell) && !isNaN(numRule) && numCell < numRule;
    case "is_empty":
      return cellStr === "" || rawVal === null || rawVal === undefined;
    case "is_not_empty":
      return cellStr !== "" && rawVal !== null && rawVal !== undefined;
    default:
      return true;
  }
}

function rowPassesRules(row: any, rules: Rule[], side: "A" | "B"): boolean {
  return rules.every(rule => {
    if (rule.target === "both") return applyRule(row, rule);
    if (rule.target === side) return applyRule(row, rule);
    return true;
  });
}

export function compareSheets(
  dataA: any[],
  dataB: any[],
  keyCols: string[],
  compareCols: string[],
  rules: Rule[]
): ComparisonResult {
  const mapA = new Map<string, any>();
  const mapB = new Map<string, any>();

  dataA.forEach(row => {
    const key = makeCompositeKey(row, keyCols);
    if (key.replace(/\|{3}/g, "").trim()) mapA.set(key, row);
  });

  dataB.forEach(row => {
    const key = makeCompositeKey(row, keyCols);
    if (key.replace(/\|{3}/g, "").trim()) mapB.set(key, row);
  });

  const matched: any[] = [];
  const onlyA: any[] = [];
  const onlyB: any[] = [];
  const differences: DiffRow[] = [];
  let filteredCount = 0;

  const rowRules = rules.filter(r => r.target !== "diff_only");
  const diffRules = rules.filter(r => r.target === "diff_only");

  for (const [key, rowA] of mapA.entries()) {
    if (!rowPassesRules(rowA, rowRules, "A")) { filteredCount++; continue; }

    const rowB = mapB.get(key);
    if (!rowB) {
      onlyA.push(rowA);
    } else {
      if (!rowPassesRules(rowB, rowRules, "B")) { filteredCount++; continue; }

      const changedCols: string[] = [];
      for (const col of compareCols) {
        const valA = String(rowA[col] ?? "").trim();
        const valB = String(rowB[col] ?? "").trim();
        if (valA !== valB) changedCols.push(col);
      }

      if (changedCols.length > 0) {
        const diffRow: DiffRow = { key, rowA, rowB, changedCols };
        const passesDiffRules = diffRules.every(rule => {
          const colVal = changedCols.includes(rule.column);
          if (rule.operator === "is_not_empty") return colVal;
          if (rule.operator === "is_empty") return !colVal;
          return applyRule(rowA, rule) || applyRule(rowB, rule);
        });
        if (passesDiffRules) {
          differences.push(diffRow);
        } else {
          filteredCount++;
        }
      } else {
        matched.push(rowA);
      }
    }
  }

  for (const [key, rowB] of mapB.entries()) {
    if (!mapA.has(key)) {
      if (rowPassesRules(rowB, rowRules, "B")) {
        onlyB.push(rowB);
      } else {
        filteredCount++;
      }
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
      filteredCount,
    },
    differences,
    matched,
    onlyA,
    onlyB,
    columns: compareCols,
    keyCols,
    appliedRules: rules,
  };
}

export function exportResultsToExcel(result: ComparisonResult, nameA: string, nameB: string) {
  const wb = XLSX.utils.book_new();

  const summaryData = [
    ["Metric", "Value"],
    ["Total Rows in " + nameA, result.summary.totalA],
    ["Total Rows in " + nameB, result.summary.totalB],
    ["Matched (No Changes)", result.summary.matchedCount],
    ["Rows with Differences", result.summary.diffCount],
    ["Only in " + nameA, result.summary.onlyACount],
    ["Only in " + nameB, result.summary.onlyBCount],
    ["Filtered by Rules", result.summary.filteredCount],
    [],
    ["Key Columns", result.keyCols.join(", ")],
    ["Applied Rules", result.appliedRules.length],
    ...result.appliedRules.map(r => [
      `  Rule`, `${r.column} ${r.operator.replace("_", " ")} "${r.value}" (on: ${r.target})`
    ]),
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  if (result.differences.length > 0) {
    const diffExportData: any[] = [];
    result.differences.forEach(diff => {
      const keyParts = diff.key.split("|||");
      const row: any = {};
      result.keyCols.forEach((col, i) => { row[col] = keyParts[i] ?? ""; });
      result.columns.forEach(col => {
        if (diff.changedCols.includes(col)) {
          row[`${col} (${nameA})`] = diff.rowA[col] ?? "";
          row[`${col} (${nameB})`] = diff.rowB[col] ?? "";
        } else {
          row[col] = diff.rowA[col] ?? "";
        }
      });
      diffExportData.push(row);
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(diffExportData), "Differences");
  }

  if (result.matched.length > 0)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.matched), "Matched");

  if (result.onlyA.length > 0)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.onlyA), `Only in ${nameA.slice(0, 20)}`);

  if (result.onlyB.length > 0)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.onlyB), `Only in ${nameB.slice(0, 20)}`);

  XLSX.writeFile(wb, "inventory_comparison.xlsx");
}
