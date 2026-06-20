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

export type CompareOptions = {
  caseInsensitive: boolean;
  ignoreWhitespace: boolean;
};

export type NumericDiff = {
  delta: number;
  pct: number | null;
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
  options: CompareOptions;
};

export type DiffRow = {
  key: string;
  rowA: any;
  rowB: any;
  changedMappings: ColMapping[];
  numericDiffs: Record<string, NumericDiff>;
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

function normalizeVal(raw: any, opts: CompareOptions): string {
  let s = String(raw ?? "");
  if (opts.ignoreWhitespace) s = s.trim();
  if (opts.caseInsensitive) s = s.toLowerCase();
  return s;
}

function makeCompositeKey(row: any, mappings: ColMapping[], side: "A" | "B", opts: CompareOptions): string {
  return mappings
    .map(m => normalizeVal(row[side === "A" ? m.colA : m.colB], opts))
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

function computeNumericDiff(valA: string, valB: string): NumericDiff | null {
  const nA = parseFloat(valA.replace(/,/g, ""));
  const nB = parseFloat(valB.replace(/,/g, ""));
  if (isNaN(nA) || isNaN(nB)) return null;
  const delta = nB - nA;
  const pct = nA !== 0 ? ((delta / Math.abs(nA)) * 100) : null;
  return { delta, pct };
}

/** Suggest column mappings between two column lists based on name similarity */
export function autoSuggestMappings(colsA: string[], colsB: string[]): ColMapping[] {
  const suggestions: ColMapping[] = [];
  const usedB = new Set<string>();

  function similarity(a: string, b: string): number {
    const na = a.toLowerCase().replace(/[^a-z0-9]/g, "");
    const nb = b.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (na === nb) return 1;
    if (na.includes(nb) || nb.includes(na)) return 0.8;
    // Count shared words
    const wordsA = new Set(a.toLowerCase().split(/[\s_\-]+/));
    const wordsB = b.toLowerCase().split(/[\s_\-]+/);
    const shared = wordsB.filter(w => wordsA.has(w)).length;
    const total = Math.max(wordsA.size, wordsB.length);
    return shared / total;
  }

  for (const ca of colsA) {
    let bestScore = 0.3; // minimum threshold
    let bestB: string | null = null;
    for (const cb of colsB) {
      if (usedB.has(cb)) continue;
      const score = similarity(ca, cb);
      if (score > bestScore) { bestScore = score; bestB = cb; }
    }
    if (bestB) {
      suggestions.push({ id: Math.random().toString(36).slice(2), colA: ca, colB: bestB });
      usedB.add(bestB);
    }
  }
  return suggestions;
}

export function compareSheets(
  dataA: any[],
  dataB: any[],
  keyMappings: ColMapping[],
  compareMappings: ColMapping[],
  rules: Rule[],
  options: CompareOptions
): ComparisonResult {
  const rulesA = rules.filter(r => r.sheet === "A");
  const rulesB = rules.filter(r => r.sheet === "B");
  let filteredCount = 0;

  const mapA = new Map<string, any>();
  const mapB = new Map<string, any>();

  dataA.forEach(row => {
    if (rulesA.every(r => applyRule(row, r))) {
      const key = makeCompositeKey(row, keyMappings, "A", options);
      if (key.replace(/\|{3}/g, "").trim()) mapA.set(key, row);
    } else filteredCount++;
  });

  dataB.forEach(row => {
    if (rulesB.every(r => applyRule(row, r))) {
      const key = makeCompositeKey(row, keyMappings, "B", options);
      if (key.replace(/\|{3}/g, "").trim()) mapB.set(key, row);
    } else filteredCount++;
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
      const numericDiffs: Record<string, NumericDiff> = {};

      for (const m of compareMappings) {
        const valA = normalizeVal(rowA[m.colA], options);
        const valB = normalizeVal(rowB[m.colB], options);
        if (valA !== valB) {
          changedMappings.push(m);
          const nd = computeNumericDiff(
            String(rowA[m.colA] ?? ""),
            String(rowB[m.colB] ?? "")
          );
          if (nd) numericDiffs[m.id] = nd;
        }
      }

      if (changedMappings.length > 0) {
        differences.push({ key, rowA, rowB, changedMappings, numericDiffs });
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
    options,
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
    ["Options"],
    ["  Case insensitive", result.options.caseInsensitive ? "Yes" : "No"],
    ["  Ignore whitespace", result.options.ignoreWhitespace ? "Yes" : "No"],
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
      result.keyMappings.forEach(m => { row[`KEY: ${m.colA}`] = diff.rowA[m.colA] ?? ""; });
      result.compareMappings.forEach(m => {
        const changed = diff.changedMappings.some(c => c.id === m.id);
        row[`${m.colA} (${nameA})`] = diff.rowA[m.colA] ?? "";
        row[`${m.colB} (${nameB})`] = diff.rowB[m.colB] ?? "";
        if (changed && diff.numericDiffs[m.id]) {
          const nd = diff.numericDiffs[m.id];
          row[`${m.colA} Delta`] = nd.delta.toFixed(2);
          if (nd.pct !== null) row[`${m.colA} Change %`] = nd.pct.toFixed(1) + "%";
        }
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
      result.compareMappings.forEach(cm => {
        row[`${cm.colA} (${nameA})`] = m.rowA[cm.colA] ?? "";
        row[`${cm.colB} (${nameB})`] = m.rowB[cm.colB] ?? "";
      });
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
