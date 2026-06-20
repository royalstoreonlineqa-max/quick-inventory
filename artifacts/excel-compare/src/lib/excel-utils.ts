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

export type RuleSheet = "A" | "B";

export type Rule = {
  id: string;
  sheet: RuleSheet;
  column: string;
  operator: RuleOperator;
  value: string;
};

export type ColMapping = {
  id: string;
  colA: string;
  colB: string;
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
  matched: MatchedRow[];
  onlyA: any[];
  onlyB: any[];
  keyMappings: ColMapping[];
  compareMappings: ColMapping[];
  appliedRules: Rule[];
};

export type DiffRow = {
  key: string;
  rowA: any;
  rowB: any;
  changedMappings: ColMapping[];
};

export type MatchedRow = {
  key: string;
  rowA: any;
  rowB: any;
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
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function makeCompositeKey(row: any, mappings: ColMapping[], side: "A" | "B"): string {
  return mappings
    .map(m => String(row[side === "A" ? m.colA : m.colB] ?? "").trim())
    .join("|||");
}

function applyRule(row: any, rule: Rule): boolean {
  const rawVal = row[rule.column];
  const cellStr = String(rawVal ?? "").trim().toLowerCase();
  const ruleVal = rule.value.trim().toLowerCase();
  const numCell = parseFloat(cellStr);
  const numRule = parseFloat(ruleVal);
  switch (rule.operator) {
    case "equals": return cellStr === ruleVal;
    case "not_equals": return cellStr !== ruleVal;
    case "contains": return cellStr.includes(ruleVal);
    case "not_contains": return !cellStr.includes(ruleVal);
    case "greater_than": return !isNaN(numCell) && !isNaN(numRule) && numCell > numRule;
    case "less_than": return !isNaN(numCell) && !isNaN(numRule) && numCell < numRule;
    case "is_empty": return cellStr === "" || rawVal === null || rawVal === undefined;
    case "is_not_empty": return cellStr !== "" && rawVal !== null && rawVal !== undefined;
    default: return true;
  }
}

export function compareSheets(
  dataA: any[],
  dataB: any[],
  keyMappings: ColMapping[],
  compareMappings: ColMapping[],
  rules: Rule[]
): ComparisonResult {
  const rulesA = rules.filter(r => r.sheet === "A");
  const rulesB = rules.filter(r => r.sheet === "B");

  const mapA = new Map<string, any>();
  const mapB = new Map<string, any>();
  let filteredCount = 0;

  dataA.forEach(row => {
    if (rulesA.every(r => applyRule(row, r))) {
      const key = makeCompositeKey(row, keyMappings, "A");
      if (key.replace(/\|{3}/g, "").trim()) mapA.set(key, row);
    } else {
      filteredCount++;
    }
  });

  dataB.forEach(row => {
    if (rulesB.every(r => applyRule(row, r))) {
      const key = makeCompositeKey(row, keyMappings, "B");
      if (key.replace(/\|{3}/g, "").trim()) mapB.set(key, row);
    } else {
      filteredCount++;
    }
  });

  const matched: MatchedRow[] = [];
  const onlyA: any[] = [];
  const onlyB: any[] = [];
  const differences: DiffRow[] = [];

  for (const [key, rowA] of mapA.entries()) {
    const rowB = mapB.get(key);
    if (!rowB) {
      onlyA.push(rowA);
    } else {
      const changedMappings: ColMapping[] = [];
      for (const m of compareMappings) {
        const valA = String(rowA[m.colA] ?? "").trim();
        const valB = String(rowB[m.colB] ?? "").trim();
        if (valA !== valB) changedMappings.push(m);
      }
      if (changedMappings.length > 0) {
        differences.push({ key, rowA, rowB, changedMappings });
      } else {
        matched.push({ key, rowA, rowB });
      }
    }
  }

  for (const [key, rowB] of mapB.entries()) {
    if (!mapA.has(key)) onlyB.push(rowB);
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
    keyMappings,
    compareMappings,
    appliedRules: rules,
  };
}

export function exportResultsToExcel(result: ComparisonResult, nameA: string, nameB: string) {
  const wb = XLSX.utils.book_new();

  const summaryData: any[][] = [
    ["Metric", "Value"],
    ["Total rows in " + nameA, result.summary.totalA],
    ["Total rows in " + nameB, result.summary.totalB],
    ["Matched (no changes)", result.summary.matchedCount],
    ["Rows with differences", result.summary.diffCount],
    ["Only in " + nameA, result.summary.onlyACount],
    ["Only in " + nameB, result.summary.onlyBCount],
    ["Filtered by rules", result.summary.filteredCount],
    [],
    ["Key Mappings"],
    ...result.keyMappings.map(m => [`  ${nameA}: ${m.colA}`, `→ ${nameB}: ${m.colB}`]),
    [],
    ["Compare Mappings"],
    ...result.compareMappings.map(m => [`  ${nameA}: ${m.colA}`, `→ ${nameB}: ${m.colB}`]),
    [],
    ["Applied Rules"],
    ...result.appliedRules.map(r => [`  Sheet ${r.sheet}: ${r.column}`, `${r.operator.replace("_", " ")} "${r.value}"`]),
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), "Summary");

  if (result.differences.length > 0) {
    const diffExportData: any[] = result.differences.map(diff => {
      const row: any = {};
      result.keyMappings.forEach(m => {
        row[`KEY: ${m.colA} (Online)`] = diff.rowA[m.colA] ?? "";
      });
      result.compareMappings.forEach(m => {
        const changed = diff.changedMappings.some(c => c.id === m.id);
        row[`${m.colA} (Online)`] = diff.rowA[m.colA] ?? "";
        row[`${m.colB} (In-Store)`] = diff.rowB[m.colB] ?? "";
        row[`Changed?`] = changed ? "YES" : "";
      });
      return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(diffExportData), "Differences");
  }

  if (result.matched.length > 0) {
    const matchedData = result.matched.map(m => {
      const row: any = {};
      result.keyMappings.forEach(km => { row[`KEY: ${km.colA}`] = m.rowA[km.colA] ?? ""; });
      result.compareMappings.forEach(cm => { row[`${cm.colA} (Online)`] = m.rowA[cm.colA] ?? ""; row[`${cm.colB} (In-Store)`] = m.rowB[cm.colB] ?? ""; });
      return row;
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(matchedData), "Matched");
  }

  if (result.onlyA.length > 0)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.onlyA), `Only in ${nameA.slice(0, 20)}`);
  if (result.onlyB.length > 0)
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(result.onlyB), `Only in ${nameB.slice(0, 20)}`);

  XLSX.writeFile(wb, "inventory_comparison.xlsx");
}
