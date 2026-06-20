import * as XLSX from "xlsx";

export type RuleOperator =
  | "equals" | "not_equals" | "contains" | "not_contains"
  | "greater_than" | "less_than" | "is_empty" | "is_not_empty";

export type RuleSheet = "A" | "B";

export type Rule = {
  id: string;
  sheet: RuleSheet;
  column: string;
  operator: RuleOperator;
  value: string;
};

export type KeyMapping = {
  id: string;
  colA: string;
  colB: string;
};

export type CompareMapping = {
  id: string;
  colA: string;
  colB: string;
};

export type OutputCol = {
  id: string;
  source: "A" | "B";
  col: string;
  label: string;
};

export type Operation = "common" | "only_a" | "only_b" | "diff_values";

export type RunConfig = {
  operation: Operation;
  keyMappings: KeyMapping[];
  compareMappings: CompareMapping[];   // only used for diff_values
  outputCols: OutputCol[];
  rules: Rule[];
  caseInsensitive: boolean;
  ignoreWhitespace: boolean;
};

export type ResultRow = Record<string, string>;

export type RunResult = {
  rows: ResultRow[];
  columns: { key: string; label: string; isDiff?: boolean }[];
  summary: string;
  totalA: number;
  totalB: number;
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
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function normalize(val: any, cfg: Pick<RunConfig, "caseInsensitive" | "ignoreWhitespace">): string {
  let s = String(val ?? "");
  if (cfg.ignoreWhitespace) s = s.trim();
  if (cfg.caseInsensitive) s = s.toLowerCase();
  return s;
}

function makeKey(row: any, keys: KeyMapping[], side: "A" | "B", cfg: RunConfig): string {
  return keys.map(k => normalize(row[side === "A" ? k.colA : k.colB], cfg)).join("|||");
}

function applyRule(row: any, rule: Rule): boolean {
  const raw = row[rule.column];
  const cell = String(raw ?? "").trim().toLowerCase();
  const val = rule.value.trim().toLowerCase();
  const nc = parseFloat(cell), nv = parseFloat(val);
  switch (rule.operator) {
    case "equals": return cell === val;
    case "not_equals": return cell !== val;
    case "contains": return cell.includes(val);
    case "not_contains": return !cell.includes(val);
    case "greater_than": return !isNaN(nc) && !isNaN(nv) && nc > nv;
    case "less_than": return !isNaN(nc) && !isNaN(nv) && nc < nv;
    case "is_empty": return cell === "" || raw === null || raw === undefined;
    case "is_not_empty": return cell !== "";
    default: return true;
  }
}

function filterRows(rows: any[], rules: Rule[], sheet: RuleSheet): any[] {
  const sheetRules = rules.filter(r => r.sheet === sheet);
  if (sheetRules.length === 0) return rows;
  return rows.filter(row => sheetRules.every(r => applyRule(row, r)));
}

export function runOperation(dataA: any[], dataB: any[], cfg: RunConfig): RunResult {
  const filtA = filterRows(dataA, cfg.rules, "A");
  const filtB = filterRows(dataB, cfg.rules, "B");

  const mapA = new Map<string, any>();
  const mapB = new Map<string, any>();
  filtA.forEach(row => { const k = makeKey(row, cfg.keyMappings, "A", cfg); if (k.replace(/\|{3}/g, "").trim()) mapA.set(k, row); });
  filtB.forEach(row => { const k = makeKey(row, cfg.keyMappings, "B", cfg); if (k.replace(/\|{3}/g, "").trim()) mapB.set(k, row); });

  // Build output column schema
  const columns: RunResult["columns"] = cfg.outputCols.map(oc => ({ key: oc.id, label: oc.label || oc.col }));

  let rows: ResultRow[] = [];

  if (cfg.operation === "common") {
    for (const [key, rowA] of mapA.entries()) {
      const rowB = mapB.get(key);
      if (!rowB) continue;
      const out: ResultRow = {};
      cfg.outputCols.forEach(oc => { out[oc.id] = String((oc.source === "A" ? rowA : rowB)[oc.col] ?? ""); });
      rows.push(out);
    }
    return { rows, columns, summary: `${rows.length} products found in both sheets`, totalA: dataA.length, totalB: dataB.length };
  }

  if (cfg.operation === "only_a") {
    for (const [key, rowA] of mapA.entries()) {
      if (mapB.has(key)) continue;
      const out: ResultRow = {};
      cfg.outputCols.forEach(oc => { out[oc.id] = String(rowA[oc.col] ?? ""); });
      rows.push(out);
    }
    return { rows, columns, summary: `${rows.length} products only in Online (not in In-Store)`, totalA: dataA.length, totalB: dataB.length };
  }

  if (cfg.operation === "only_b") {
    for (const [key, rowB] of mapB.entries()) {
      if (mapA.has(key)) continue;
      const out: ResultRow = {};
      cfg.outputCols.forEach(oc => { out[oc.id] = String(rowB[oc.col] ?? ""); });
      rows.push(out);
    }
    return { rows, columns, summary: `${rows.length} products only in In-Store (not Online)`, totalA: dataA.length, totalB: dataB.length };
  }

  if (cfg.operation === "diff_values") {
    // Build extended columns: output cols + diff pairs
    const diffCols: RunResult["columns"] = [];
    cfg.compareMappings.forEach(m => {
      diffCols.push({ key: `diff_a_${m.id}`, label: `${m.colA} (Online)`, isDiff: false });
      diffCols.push({ key: `diff_b_${m.id}`, label: `${m.colB} (In-Store)`, isDiff: true });
    });
    const allCols = [...columns, ...diffCols];

    for (const [key, rowA] of mapA.entries()) {
      const rowB = mapB.get(key);
      if (!rowB) continue;
      // Check if any compare column differs
      const hasDiff = cfg.compareMappings.some(m => {
        return normalize(rowA[m.colA], cfg) !== normalize(rowB[m.colB], cfg);
      });
      if (!hasDiff) continue;
      const out: ResultRow = {};
      cfg.outputCols.forEach(oc => { out[oc.id] = String((oc.source === "A" ? rowA : rowB)[oc.col] ?? ""); });
      cfg.compareMappings.forEach(m => {
        out[`diff_a_${m.id}`] = String(rowA[m.colA] ?? "");
        out[`diff_b_${m.id}`] = String(rowB[m.colB] ?? "");
      });
      rows.push(out);
    }
    return { rows, columns: allCols, summary: `${rows.length} products with value differences`, totalA: dataA.length, totalB: dataB.length };
  }

  return { rows: [], columns: [], summary: "No operation selected", totalA: dataA.length, totalB: dataB.length };
}

/** Suggest column mappings by name similarity */
export function autoSuggestMappings(colsA: string[], colsB: string[]): { colA: string; colB: string }[] {
  const used = new Set<string>();
  return colsA.flatMap(ca => {
    const na = ca.toLowerCase().replace(/[^a-z0-9]/g, "");
    let best = 0.3, bestB: string | null = null;
    for (const cb of colsB) {
      if (used.has(cb)) continue;
      const nb = cb.toLowerCase().replace(/[^a-z0-9]/g, "");
      let score = na === nb ? 1 : na.includes(nb) || nb.includes(na) ? 0.8 : 0;
      if (score === 0) {
        const wa = new Set(ca.toLowerCase().split(/[\s_\-]+/));
        const wb = cb.toLowerCase().split(/[\s_\-]+/);
        score = wb.filter(w => wa.has(w)).length / Math.max(wa.size, wb.length);
      }
      if (score > best) { best = score; bestB = cb; }
    }
    if (bestB) { used.add(bestB); return [{ colA: ca, colB: bestB }]; }
    return [];
  });
}

export function exportToExcel(result: RunResult, operation: Operation) {
  const wb = XLSX.utils.book_new();
  const headers = result.columns.map(c => c.label);
  const dataRows = result.rows.map(row => result.columns.map(c => row[c.key] ?? ""));
  const ws = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
  const sheetName = operation === "common" ? "Common Products" : operation === "only_a" ? "Only Online" : operation === "only_b" ? "Only In-Store" : "Value Differences";
  XLSX.utils.book_append_sheet(wb, ws, sheetName);

  // Summary sheet
  const summary = XLSX.utils.aoa_to_sheet([["Result"], [result.summary]]);
  XLSX.utils.book_append_sheet(wb, summary, "Summary");

  XLSX.writeFile(wb, "inventory_results.xlsx");
}
