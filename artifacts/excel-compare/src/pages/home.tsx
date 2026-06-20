import React, { useState, useCallback, useMemo } from "react";
import {
  UploadCloud, FileSpreadsheet, ArrowRight, Play, RefreshCcw, Download,
  CheckCircle2, AlertTriangle, FileWarning, PlusCircle, MinusCircle,
  Key, SlidersHorizontal, Trash2, Plus, X, ShoppingCart, Store, ArrowLeftRight
} from "lucide-react";
import {
  parseExcelFile, compareSheets, exportResultsToExcel,
  ComparisonResult, Rule, RuleOperator, RuleSheet, ColMapping
} from "@/lib/excel-utils";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";

type Step = "upload" | "config" | "loading" | "results";

const OPERATORS: { value: RuleOperator; label: string; needsValue: boolean }[] = [
  { value: "equals", label: "equals", needsValue: true },
  { value: "not_equals", label: "does not equal", needsValue: true },
  { value: "contains", label: "contains", needsValue: true },
  { value: "not_contains", label: "does not contain", needsValue: true },
  { value: "greater_than", label: "greater than (>)", needsValue: true },
  { value: "less_than", label: "less than (<)", needsValue: true },
  { value: "is_empty", label: "is empty", needsValue: false },
  { value: "is_not_empty", label: "is not empty", needsValue: false },
];

function makeId() { return Math.random().toString(36).slice(2); }

type FileData = { name: string; data: any[]; columns: string[] };

export default function Home() {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");

  const [fileA, setFileA] = useState<FileData | null>(null);
  const [fileB, setFileB] = useState<FileData | null>(null);

  // Key mappings: colA (Online) ↔ colB (In-Store)
  const [keyMappings, setKeyMappings] = useState<ColMapping[]>([]);
  // Compare mappings: compare colA (Online) vs colB (In-Store)
  const [compareMappings, setCompareMappings] = useState<ColMapping[]>([]);
  // Rules
  const [rules, setRules] = useState<Rule[]>([]);

  const [results, setResults] = useState<ComparisonResult | null>(null);

  const handleFileUpload = async (file: File, side: "A" | "B") => {
    try {
      const parsed = await parseExcelFile(file);
      if (side === "A") setFileA({ name: file.name, data: parsed.data, columns: parsed.columns });
      else setFileB({ name: file.name, data: parsed.data, columns: parsed.columns });
    } catch {
      toast({ title: "Error parsing file", description: "Please upload a valid Excel or CSV file.", variant: "destructive" });
    }
  };

  const handleDrop = useCallback((e: React.DragEvent, side: "A" | "B") => {
    e.preventDefault();
    if (e.dataTransfer.files?.[0]) handleFileUpload(e.dataTransfer.files[0], side);
  }, []);

  const proceedToConfig = () => {
    if (!fileA || !fileB) return;
    // Start with one empty key mapping and one empty compare mapping
    setKeyMappings([{ id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
    setCompareMappings([{ id: makeId(), colA: fileA.columns[1] ?? fileA.columns[0] ?? "", colB: fileB.columns[1] ?? fileB.columns[0] ?? "" }]);
    setRules([]);
    setStep("config");
  };

  // Key mapping helpers
  const addKeyMapping = () => {
    if (!fileA || !fileB) return;
    setKeyMappings(p => [...p, { id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
  };
  const updateKeyMapping = (id: string, patch: Partial<ColMapping>) =>
    setKeyMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeKeyMapping = (id: string) => setKeyMappings(p => p.filter(m => m.id !== id));

  // Compare mapping helpers
  const addCompareMapping = () => {
    if (!fileA || !fileB) return;
    setCompareMappings(p => [...p, { id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
  };
  const updateCompareMapping = (id: string, patch: Partial<ColMapping>) =>
    setCompareMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeCompareMapping = (id: string) => setCompareMappings(p => p.filter(m => m.id !== id));

  // Rule helpers
  const addRule = () => {
    if (!fileA) return;
    setRules(p => [...p, { id: makeId(), sheet: "A", column: fileA.columns[0] ?? "", operator: "equals", value: "" }]);
  };
  const updateRule = (id: string, patch: Partial<Rule>) => setRules(p => p.map(r => r.id === id ? { ...r, ...patch } : r));
  const removeRule = (id: string) => setRules(p => p.filter(r => r.id !== id));

  const runComparison = () => {
    if (!fileA || !fileB || keyMappings.length === 0 || compareMappings.length === 0) return;
    setStep("loading");
    setTimeout(() => {
      try {
        const validRules = rules.filter(r => {
          const op = OPERATORS.find(o => o.value === r.operator);
          return op && (!op.needsValue || r.value.trim() !== "");
        });
        const res = compareSheets(fileA.data, fileB.data, keyMappings, compareMappings, validRules);
        setResults(res);
        setStep("results");
      } catch {
        toast({ title: "Comparison Failed", description: "An error occurred during comparison.", variant: "destructive" });
        setStep("config");
      }
    }, 100);
  };

  const reset = () => {
    setFileA(null); setFileB(null);
    setKeyMappings([]); setCompareMappings([]);
    setRules([]); setResults(null);
    setStep("upload");
  };

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col font-sans">
      <header className="border-b border-border bg-card px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-2">
          <FileSpreadsheet className="w-6 h-6 text-primary" />
          <h1 className="text-xl font-bold tracking-tight">ReconcileX</h1>
          <Badge variant="secondary" className="ml-2 text-xs font-normal">Inventory</Badge>
        </div>
        {step === "results" && (
          <div className="flex items-center gap-3">
            <Button variant="outline" onClick={reset} size="sm" data-testid="button-reset">
              <RefreshCcw className="w-4 h-4 mr-2" />New Comparison
            </Button>
            <Button
              onClick={() => results && exportResultsToExcel(results, fileA?.name ?? "Online", fileB?.name ?? "In-Store")}
              size="sm"
              className="bg-primary hover:bg-primary/90 text-primary-foreground"
              data-testid="button-export"
            >
              <Download className="w-4 h-4 mr-2" />Export Results
            </Button>
          </div>
        )}
      </header>

      <main className="flex-1 p-6 max-w-7xl w-full mx-auto">

        {/* ── UPLOAD ── */}
        {step === "upload" && (
          <div className="animate-in fade-in zoom-in-95 duration-300">
            <div className="mb-8">
              <h2 className="text-3xl font-semibold mb-2">Load Inventory Sheets</h2>
              <p className="text-muted-foreground text-lg">Upload your online and in-store inventory files to begin reconciliation.</p>
            </div>
            <div className="grid md:grid-cols-2 gap-6">
              <UploadZone side="A" label="Online Inventory" icon={<ShoppingCart className="w-8 h-8" />} file={fileA} onDrop={handleDrop} onUpload={handleFileUpload} onClear={() => setFileA(null)} />
              <UploadZone side="B" label="In-Store Inventory" icon={<Store className="w-8 h-8" />} file={fileB} onDrop={handleDrop} onUpload={handleFileUpload} onClear={() => setFileB(null)} />
            </div>
            <div className="mt-8 flex justify-end">
              <Button size="lg" onClick={proceedToConfig} disabled={!fileA || !fileB} className="w-full md:w-auto font-semibold px-8" data-testid="button-proceed-config">
                Configure Comparison <ArrowRight className="ml-2 w-5 h-5" />
              </Button>
            </div>
          </div>
        )}

        {/* ── CONFIG ── */}
        {step === "config" && fileA && fileB && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <div className="flex items-center gap-4 mb-8">
              <Button variant="ghost" size="icon" onClick={() => setStep("upload")} className="shrink-0 rounded-full h-8 w-8">
                <ArrowRight className="w-4 h-4 rotate-180" />
              </Button>
              <div>
                <h2 className="text-3xl font-semibold">Configure Comparison</h2>
                <p className="text-muted-foreground text-lg">Map columns between the two files and set filter rules.</p>
              </div>
            </div>

            {/* Column header labels */}
            <div className="grid grid-cols-[1fr_auto_1fr_auto] gap-3 mb-2 px-1 items-center">
              <div className="flex items-center gap-2">
                <ShoppingCart className="w-4 h-4 text-blue-400 shrink-0" />
                <span className="text-xs font-semibold text-blue-400 uppercase tracking-wide truncate">{fileA.name}</span>
              </div>
              <div className="w-8" />
              <div className="flex items-center gap-2">
                <Store className="w-4 h-4 text-violet-400 shrink-0" />
                <span className="text-xs font-semibold text-violet-400 uppercase tracking-wide truncate">{fileB.name}</span>
              </div>
              <div className="w-8" />
            </div>

            {/* Step 1: Match Keys */}
            <Card className="border-border/50 shadow-md mb-6 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Key className="w-5 h-5 text-primary" />
                    <CardTitle className="text-lg">1. Match Keys</CardTitle>
                  </div>
                  <Button variant="outline" size="sm" onClick={addKeyMapping} data-testid="button-add-key">
                    <Plus className="w-4 h-4 mr-1" />Add Key
                  </Button>
                </div>
                <CardDescription>
                  Pick the column in each sheet that identifies the same product. They can have different names — e.g. "Product Code" in Online matches "Item ID" in In-Store.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-3">
                  {keyMappings.map((m, idx) => (
                    <MappingRow
                      key={m.id}
                      mapping={m}
                      colsA={fileA.columns}
                      colsB={fileB.columns}
                      idx={idx}
                      prefix="key"
                      accent="primary"
                      onUpdateA={v => updateKeyMapping(m.id, { colA: v })}
                      onUpdateB={v => updateKeyMapping(m.id, { colB: v })}
                      onRemove={() => removeKeyMapping(m.id)}
                      canRemove={keyMappings.length > 1}
                    />
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* Step 2: Columns to Compare */}
            <Card className="border-border/50 shadow-md mb-6 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <ArrowLeftRight className="w-5 h-5 text-primary" />
                    <CardTitle className="text-lg">2. Columns to Compare</CardTitle>
                  </div>
                  <Button variant="outline" size="sm" onClick={addCompareMapping} data-testid="button-add-compare">
                    <Plus className="w-4 h-4 mr-1" />Add Column
                  </Button>
                </div>
                <CardDescription>
                  Map each value column you want to compare across the two files. e.g. "Online Price" vs "Store Price", "Stock Online" vs "Stock Offline".
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-3">
                  {compareMappings.map((m, idx) => (
                    <MappingRow
                      key={m.id}
                      mapping={m}
                      colsA={fileA.columns}
                      colsB={fileB.columns}
                      idx={idx}
                      prefix="cmp"
                      accent="cyan"
                      onUpdateA={v => updateCompareMapping(m.id, { colA: v })}
                      onUpdateB={v => updateCompareMapping(m.id, { colB: v })}
                      onRemove={() => removeCompareMapping(m.id)}
                      canRemove={compareMappings.length > 1}
                    />
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* Step 3: Rules */}
            <Card className="border-border/50 shadow-md mb-8 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <SlidersHorizontal className="w-5 h-5 text-primary" />
                    <CardTitle className="text-lg">3. Filter Rules <span className="text-sm font-normal text-muted-foreground ml-1">(optional)</span></CardTitle>
                  </div>
                  <Button variant="outline" size="sm" onClick={addRule} data-testid="button-add-rule">
                    <Plus className="w-4 h-4 mr-1" />Add Rule
                  </Button>
                </div>
                <CardDescription>
                  Narrow results to rows that meet conditions — e.g. "Stock less than 10", "Category contains Electronics".
                </CardDescription>
              </CardHeader>
              <CardContent>
                {rules.length === 0 ? (
                  <div className="border border-dashed border-border/50 rounded-md p-6 text-center">
                    <p className="text-sm text-muted-foreground">No rules set — all rows will be included.</p>
                    <Button variant="ghost" size="sm" className="mt-2 text-primary" onClick={addRule}>
                      <Plus className="w-3 h-3 mr-1" />Add your first rule
                    </Button>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {rules.map((rule, idx) => {
                      const opMeta = OPERATORS.find(o => o.value === rule.operator);
                      const cols = rule.sheet === "A" ? fileA.columns : fileB.columns;
                      return (
                        <div key={rule.id} className="flex items-center gap-2 flex-wrap p-3 bg-background rounded-md border border-border/50" data-testid={`rule-row-${idx}`}>
                          <span className="text-xs font-medium text-muted-foreground w-5 shrink-0">#{idx + 1}</span>

                          {/* Sheet picker */}
                          <Select value={rule.sheet} onValueChange={v => updateRule(rule.id, { sheet: v as RuleSheet, column: (v === "A" ? fileA : fileB).columns[0] ?? "" })}>
                            <SelectTrigger className="w-28 h-8 text-xs" data-testid={`rule-sheet-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              <SelectItem value="A" className="text-xs">
                                <span className="flex items-center gap-1"><ShoppingCart className="w-3 h-3 text-blue-400" />Online</span>
                              </SelectItem>
                              <SelectItem value="B" className="text-xs">
                                <span className="flex items-center gap-1"><Store className="w-3 h-3 text-violet-400" />In-Store</span>
                              </SelectItem>
                            </SelectContent>
                          </Select>

                          {/* Column */}
                          <Select value={rule.column} onValueChange={v => updateRule(rule.id, { column: v })}>
                            <SelectTrigger className="w-40 h-8 text-xs" data-testid={`rule-col-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              {cols.map(c => <SelectItem key={c} value={c} className="text-xs">{c}</SelectItem>)}
                            </SelectContent>
                          </Select>

                          {/* Operator */}
                          <Select value={rule.operator} onValueChange={v => updateRule(rule.id, { operator: v as RuleOperator, value: "" })}>
                            <SelectTrigger className="w-44 h-8 text-xs" data-testid={`rule-op-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              {OPERATORS.map(op => <SelectItem key={op.value} value={op.value} className="text-xs">{op.label}</SelectItem>)}
                            </SelectContent>
                          </Select>

                          {/* Value */}
                          {opMeta?.needsValue && (
                            <Input
                              className="w-32 h-8 text-xs font-mono"
                              placeholder="value..."
                              value={rule.value}
                              onChange={e => updateRule(rule.id, { value: e.target.value })}
                              data-testid={`rule-val-${idx}`}
                            />
                          )}

                          <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-destructive shrink-0 ml-auto" onClick={() => removeRule(rule.id)}>
                            <Trash2 className="w-4 h-4" />
                          </Button>
                        </div>
                      );
                    })}
                    <p className="text-xs text-muted-foreground pl-1">Rules are applied per-sheet (AND logic). Each sheet's rows must pass its rules before matching.</p>
                  </div>
                )}
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button
                size="lg"
                onClick={runComparison}
                disabled={keyMappings.length === 0 || compareMappings.length === 0}
                className="w-full md:w-auto font-bold px-10 h-14 text-lg bg-primary hover:bg-primary/90 text-primary-foreground"
                data-testid="button-run-comparison"
              >
                <Play className="mr-2 w-5 h-5 fill-current" />
                Run Analysis
              </Button>
            </div>
          </div>
        )}

        {/* ── LOADING ── */}
        {step === "loading" && (
          <div className="h-[60vh] flex flex-col items-center justify-center animate-in fade-in duration-500">
            <div className="w-16 h-16 border-4 border-primary border-t-transparent rounded-full animate-spin mb-6" />
            <h2 className="text-2xl font-semibold mb-2">Processing Inventory</h2>
            <p className="text-muted-foreground">Matching rows and applying rules...</p>
          </div>
        )}

        {/* ── RESULTS ── */}
        {step === "results" && results && (
          <div className="animate-in fade-in zoom-in-95 duration-500">

            {/* Active filters */}
            {(results.appliedRules.length > 0 || results.summary.filteredCount > 0) && (
              <div className="flex flex-wrap gap-2 mb-4 items-center">
                <span className="text-xs text-muted-foreground font-medium">Active filters:</span>
                {results.appliedRules.map(r => {
                  const op = OPERATORS.find(o => o.value === r.operator);
                  return (
                    <Badge key={r.id} variant="secondary" className="font-mono text-xs gap-1">
                      <span className="text-muted-foreground opacity-70">{r.sheet === "A" ? "Online" : "In-Store"}:</span>
                      <span className="text-primary font-semibold">{r.column}</span>
                      <span className="text-muted-foreground">{op?.label}</span>
                      {op?.needsValue && <span>"{r.value}"</span>}
                    </Badge>
                  );
                })}
                {results.summary.filteredCount > 0 && (
                  <Badge variant="outline" className="text-xs text-muted-foreground gap-1">
                    <X className="w-3 h-3" />{results.summary.filteredCount} rows filtered out
                  </Badge>
                )}
              </div>
            )}

            {/* Metrics */}
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-6">
              <MetricCard title="Online Total" value={results.summary.totalA} icon={<ShoppingCart className="w-4 h-4 text-blue-400" />} className="border-blue-400/20 bg-blue-400/5" />
              <MetricCard title="In-Store Total" value={results.summary.totalB} icon={<Store className="w-4 h-4 text-violet-400" />} className="border-violet-400/20 bg-violet-400/5" />
              <MetricCard title="Matched" value={results.summary.matchedCount} icon={<CheckCircle2 className="w-4 h-4 text-emerald-500" />} className="border-emerald-500/20 bg-emerald-500/5" />
              <MetricCard title="Differences" value={results.summary.diffCount} icon={<AlertTriangle className="w-4 h-4 text-accent" />} className="border-accent/20 bg-accent/5" />
              <MetricCard title="Only Online" value={results.summary.onlyACount} icon={<MinusCircle className="w-4 h-4 text-destructive" />} className="border-destructive/20 bg-destructive/5" />
              <MetricCard title="Only In-Store" value={results.summary.onlyBCount} icon={<PlusCircle className="w-4 h-4 text-cyan-400" />} className="border-cyan-400/20 bg-cyan-400/5" />
            </div>

            {/* Mapping summary chips */}
            <div className="flex flex-wrap gap-3 mb-5 items-start">
              <div className="flex flex-wrap gap-2 items-center">
                <span className="text-xs text-muted-foreground font-medium">Keys:</span>
                {results.keyMappings.map(m => (
                  <Badge key={m.id} variant="outline" className="font-mono text-xs border-primary/30">
                    <span className="text-blue-400">{m.colA}</span>
                    <ArrowLeftRight className="w-3 h-3 mx-1 text-muted-foreground" />
                    <span className="text-violet-400">{m.colB}</span>
                  </Badge>
                ))}
              </div>
              <div className="flex flex-wrap gap-2 items-center">
                <span className="text-xs text-muted-foreground font-medium">Compared:</span>
                {results.compareMappings.map(m => (
                  <Badge key={m.id} variant="secondary" className="font-mono text-xs">
                    <span className="text-blue-400">{m.colA}</span>
                    <span className="text-muted-foreground mx-1">vs</span>
                    <span className="text-violet-400">{m.colB}</span>
                  </Badge>
                ))}
              </div>
            </div>

            <Tabs defaultValue="diffs" className="w-full">
              <TabsList className="grid w-full grid-cols-4 mb-6 h-12 bg-card/50">
                <TabsTrigger value="diffs" className="data-[state=active]:bg-primary data-[state=active]:text-primary-foreground font-medium" data-testid="tab-diffs">
                  Differences ({results.summary.diffCount})
                </TabsTrigger>
                <TabsTrigger value="matches" className="font-medium" data-testid="tab-matches">
                  Matched ({results.summary.matchedCount})
                </TabsTrigger>
                <TabsTrigger value="onlya" className="font-medium" data-testid="tab-onlya">
                  Only Online ({results.summary.onlyACount})
                </TabsTrigger>
                <TabsTrigger value="onlyb" className="font-medium" data-testid="tab-onlyb">
                  Only In-Store ({results.summary.onlyBCount})
                </TabsTrigger>
              </TabsList>

              {/* Differences Tab */}
              <TabsContent value="diffs" className="m-0">
                <Card className="border-border overflow-hidden">
                  <ScrollArea className="h-[60vh] w-full">
                    {results.differences.length === 0 ? (
                      <EmptyState icon={<CheckCircle2 className="w-12 h-12 text-emerald-500 opacity-80" />} title="No differences found" subtitle="All matched rows are identical in the compared columns." />
                    ) : (
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            {/* Key columns */}
                            {results.keyMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px] uppercase tracking-wide">
                                <span className="flex items-center gap-1">
                                  <Key className="w-3 h-3 text-primary shrink-0" />
                                  <span className="text-blue-400">{m.colA}</span>
                                </span>
                              </th>
                            ))}
                            {/* Compare columns */}
                            {results.compareMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[240px] uppercase tracking-wide">
                                <span className="flex items-center gap-1">
                                  <span className="text-blue-400">{m.colA}</span>
                                  <span className="text-muted-foreground opacity-60 text-[10px]">vs</span>
                                  <span className="text-violet-400">{m.colB}</span>
                                </span>
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {results.differences.map((diff, idx) => {
                            const keyParts = diff.key.split("|||");
                            return (
                              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors" data-testid={`diff-row-${idx}`}>
                                {results.keyMappings.map((m, i) => (
                                  <td key={m.id} className="px-4 py-3 font-mono text-xs font-medium border-r bg-background/50">{keyParts[i] ?? ""}</td>
                                ))}
                                {results.compareMappings.map(m => {
                                  const isChanged = diff.changedMappings.some(c => c.id === m.id);
                                  const valA = diff.rowA[m.colA] !== undefined ? String(diff.rowA[m.colA]) : "(empty)";
                                  const valB = diff.rowB[m.colB] !== undefined ? String(diff.rowB[m.colB]) : "(empty)";
                                  return (
                                    <td key={m.id} className={`px-4 py-3 border-r ${isChanged ? "bg-accent/5" : ""}`}>
                                      {isChanged ? (
                                        <div className="flex flex-col gap-1.5">
                                          <div className="flex items-center gap-1.5 text-destructive font-mono bg-destructive/10 px-2 py-1 rounded text-xs">
                                            <ShoppingCart className="w-3 h-3 shrink-0 opacity-60" />
                                            <span className="truncate" title={valA}>{valA}</span>
                                          </div>
                                          <div className="flex items-center gap-1.5 text-emerald-400 font-mono bg-emerald-500/10 px-2 py-1 rounded text-xs">
                                            <Store className="w-3 h-3 shrink-0 opacity-60" />
                                            <span className="truncate" title={valB}>{valB}</span>
                                          </div>
                                        </div>
                                      ) : (
                                        <span className="text-muted-foreground font-mono text-xs truncate block" title={valA}>{valA}</span>
                                      )}
                                    </td>
                                  );
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}
                  </ScrollArea>
                </Card>
              </TabsContent>

              {/* Matched Tab */}
              <TabsContent value="matches" className="m-0">
                {results.matched.length === 0 ? (
                  <EmptyCard message="No perfectly matched rows." />
                ) : (
                  <Card className="border-border overflow-hidden">
                    <ScrollArea className="h-[60vh] w-full">
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            {results.keyMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px]">
                                <Key className="w-3 h-3 text-primary inline mr-1" />{m.colA}
                              </th>
                            ))}
                            {results.compareMappings.map(m => (
                              <th key={`a-${m.id}`} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[130px] text-blue-400">{m.colA}</th>
                            ))}
                            {results.compareMappings.map(m => (
                              <th key={`b-${m.id}`} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[130px] text-violet-400">{m.colB}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {results.matched.map((row, idx) => {
                            const keyParts = row.key.split("|||");
                            return (
                              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors">
                                {results.keyMappings.map((m, i) => (
                                  <td key={m.id} className="px-4 py-3 font-mono text-xs font-medium border-r">{keyParts[i] ?? ""}</td>
                                ))}
                                {results.compareMappings.map(m => (
                                  <td key={`a-${m.id}`} className="px-4 py-3 font-mono text-xs text-muted-foreground border-r">{String(row.rowA[m.colA] ?? "")}</td>
                                ))}
                                {results.compareMappings.map(m => (
                                  <td key={`b-${m.id}`} className="px-4 py-3 font-mono text-xs text-muted-foreground border-r">{String(row.rowB[m.colB] ?? "")}</td>
                                ))}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </ScrollArea>
                  </Card>
                )}
              </TabsContent>

              {/* Only A */}
              <TabsContent value="onlya" className="m-0">
                <SimpleDataTable rows={results.onlyA} emptyMessage="No rows exclusive to Online." />
              </TabsContent>

              {/* Only B */}
              <TabsContent value="onlyb" className="m-0">
                <SimpleDataTable rows={results.onlyB} emptyMessage="No rows exclusive to In-Store." />
              </TabsContent>
            </Tabs>
          </div>
        )}
      </main>
    </div>
  );
}

/* ── Mapping Row ── */
function MappingRow({ mapping, colsA, colsB, idx, prefix, accent, onUpdateA, onUpdateB, onRemove, canRemove }: {
  mapping: ColMapping; colsA: string[]; colsB: string[];
  idx: number; prefix: string; accent: string;
  onUpdateA: (v: string) => void; onUpdateB: (v: string) => void;
  onRemove: () => void; canRemove: boolean;
}) {
  return (
    <div className="grid grid-cols-[1fr_auto_1fr_auto] gap-3 items-center" data-testid={`${prefix}-row-${idx}`}>
      <Select value={mapping.colA} onValueChange={onUpdateA}>
        <SelectTrigger className="h-9 text-sm border-blue-400/30 bg-blue-400/5 text-blue-300" data-testid={`${prefix}-colA-${idx}`}>
          <SelectValue placeholder="Online column..." />
        </SelectTrigger>
        <SelectContent>
          {colsA.map(c => <SelectItem key={c} value={c} className="text-sm">{c}</SelectItem>)}
        </SelectContent>
      </Select>

      <div className="flex items-center justify-center">
        <ArrowLeftRight className="w-4 h-4 text-muted-foreground" />
      </div>

      <Select value={mapping.colB} onValueChange={onUpdateB}>
        <SelectTrigger className="h-9 text-sm border-violet-400/30 bg-violet-400/5 text-violet-300" data-testid={`${prefix}-colB-${idx}`}>
          <SelectValue placeholder="In-Store column..." />
        </SelectTrigger>
        <SelectContent>
          {colsB.map(c => <SelectItem key={c} value={c} className="text-sm">{c}</SelectItem>)}
        </SelectContent>
      </Select>

      <Button
        variant="ghost" size="icon"
        className="h-9 w-9 text-muted-foreground hover:text-destructive"
        onClick={onRemove}
        disabled={!canRemove}
        data-testid={`${prefix}-remove-${idx}`}
      >
        <Trash2 className="w-4 h-4" />
      </Button>
    </div>
  );
}

/* ── Upload Zone ── */
function UploadZone({ side, label, icon, file, onDrop, onUpload, onClear }: {
  side: "A" | "B"; label: string; icon: React.ReactNode;
  file: FileData | null;
  onDrop: (e: React.DragEvent, side: "A" | "B") => void;
  onUpload: (file: File, side: "A" | "B") => void;
  onClear: () => void;
}) {
  const inputId = `file-upload-${side}`;
  if (file) {
    return (
      <Card className="border-primary/40 bg-primary/5 shadow-md">
        <CardContent className="p-6">
          <div className="flex justify-between items-start mb-4">
            <div>
              <Badge variant="outline" className="mb-2 bg-background text-xs">{label}</Badge>
              <h3 className="font-semibold text-base truncate pr-6" title={file.name}>{file.name}</h3>
              <p className="text-sm text-muted-foreground font-mono mt-0.5">{file.data.length.toLocaleString()} rows · {file.columns.length} cols</p>
            </div>
            <Button variant="ghost" size="sm" onClick={onClear} className="text-muted-foreground hover:text-destructive text-xs">Clear</Button>
          </div>
          <div className="bg-background rounded border border-border overflow-hidden">
            <table className="w-full text-left text-xs whitespace-nowrap">
              <thead className="bg-muted/50 text-muted-foreground">
                <tr>
                  {file.columns.slice(0, 4).map(c => (
                    <th key={c} className="px-3 py-2 border-b font-medium max-w-[130px] truncate">{c}</th>
                  ))}
                  {file.columns.length > 4 && <th className="px-3 py-2 border-b text-muted-foreground/60">+{file.columns.length - 4} more</th>}
                </tr>
              </thead>
              <tbody className="font-mono">
                {file.data.slice(0, 3).map((row, i) => (
                  <tr key={i} className="border-b last:border-0 hover:bg-muted/20">
                    {file.columns.slice(0, 4).map(c => (
                      <td key={c} className="px-3 py-2 max-w-[130px] truncate opacity-75">{String(row[c] ?? "")}</td>
                    ))}
                    {file.columns.length > 4 && <td className="px-3 py-2 text-muted-foreground/40">...</td>}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </CardContent>
      </Card>
    );
  }
  return (
    <Card
      className="border-dashed border-2 hover:border-primary/50 hover:bg-card/80 transition-colors cursor-pointer group bg-card/30"
      onDragOver={e => e.preventDefault()}
      onDrop={e => onDrop(e, side)}
      onClick={() => document.getElementById(inputId)?.click()}
    >
      <CardContent className="flex flex-col items-center justify-center p-12 text-center h-full min-h-[280px]">
        <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mb-5 group-hover:scale-110 transition-transform duration-300 group-hover:bg-primary/10 text-muted-foreground group-hover:text-primary">
          {icon}
        </div>
        <h3 className="text-xl font-medium mb-1">{label}</h3>
        <p className="text-muted-foreground text-sm mb-5 max-w-[220px]">Drag & drop your .xlsx or .csv file here, or click to browse.</p>
        <Button variant="secondary" className="pointer-events-none group-hover:bg-primary group-hover:text-primary-foreground transition-colors">Select File</Button>
        <input type="file" id={inputId} className="hidden" accept=".xlsx,.xls,.csv"
          onChange={e => { if (e.target.files?.[0]) onUpload(e.target.files[0], side); }} />
      </CardContent>
    </Card>
  );
}

function MetricCard({ title, value, icon, className }: { title: string; value: number; icon: React.ReactNode; className?: string }) {
  return (
    <Card className={`overflow-hidden transition-all hover:shadow-md ${className}`}>
      <CardContent className="p-4 flex flex-col justify-center">
        <div className="flex items-center justify-between mb-3">
          <p className="text-xs font-medium text-muted-foreground leading-tight">{title}</p>
          <div className="bg-background/80 p-1.5 rounded-full">{icon}</div>
        </div>
        <div className="text-2xl font-bold font-mono tracking-tight">{value.toLocaleString()}</div>
      </CardContent>
    </Card>
  );
}

function EmptyState({ icon, title, subtitle }: { icon: React.ReactNode; title: string; subtitle: string }) {
  return (
    <div className="p-12 text-center flex flex-col items-center">
      <div className="mb-4">{icon}</div>
      <p className="text-lg font-medium mb-1">{title}</p>
      <p className="text-muted-foreground text-sm">{subtitle}</p>
    </div>
  );
}

function EmptyCard({ message }: { message: string }) {
  return (
    <Card className="p-12 text-center border-dashed">
      <FileWarning className="w-12 h-12 text-muted-foreground mx-auto mb-4 opacity-50" />
      <p className="text-lg text-muted-foreground font-medium">{message}</p>
    </Card>
  );
}

function SimpleDataTable({ rows, emptyMessage }: { rows: any[]; emptyMessage: string }) {
  if (rows.length === 0) return <EmptyCard message={emptyMessage} />;
  const columns = Object.keys(rows[0]);
  return (
    <Card className="border-border overflow-hidden">
      <ScrollArea className="h-[60vh] w-full">
        <table className="w-full text-sm text-left whitespace-nowrap">
          <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
            <tr>
              {columns.map(col => (
                <th key={col} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px]">{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors">
                {columns.map(col => (
                  <td key={col} className="px-4 py-3 border-r font-mono text-xs text-muted-foreground">{String(row[col] ?? "")}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </ScrollArea>
    </Card>
  );
}
