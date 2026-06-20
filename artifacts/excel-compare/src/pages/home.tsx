import React, { useState, useCallback, useMemo } from "react";
import {
  UploadCloud, FileSpreadsheet, ArrowRight, Play, RefreshCcw, Download,
  CheckCircle2, AlertTriangle, FileWarning, PlusCircle, MinusCircle,
  Key, SlidersHorizontal, Trash2, Plus, X, ShoppingCart, Store
} from "lucide-react";
import { parseExcelFile, compareSheets, exportResultsToExcel, ComparisonResult, Rule, RuleOperator, RuleTarget } from "@/lib/excel-utils";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Checkbox } from "@/components/ui/checkbox";
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

const TARGETS: { value: RuleTarget; label: string }[] = [
  { value: "both", label: "Both sheets" },
  { value: "A", label: "Online only" },
  { value: "B", label: "In-Store only" },
  { value: "diff_only", label: "Differences only" },
];

function makeId() {
  return Math.random().toString(36).slice(2);
}

export default function Home() {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");

  const [fileA, setFileA] = useState<{ name: string; data: any[]; columns: string[] } | null>(null);
  const [fileB, setFileB] = useState<{ name: string; data: any[]; columns: string[] } | null>(null);

  const [keyCols, setKeyCols] = useState<Set<string>>(new Set());
  const [selectedCols, setSelectedCols] = useState<Set<string>>(new Set());
  const [rules, setRules] = useState<Rule[]>([]);

  const [results, setResults] = useState<ComparisonResult | null>(null);

  const sharedColumns = useMemo(() => {
    if (!fileA || !fileB) return [];
    return fileA.columns.filter(c => fileB.columns.includes(c));
  }, [fileA, fileB]);

  const allColumns = useMemo(() => {
    if (!fileA && !fileB) return [];
    const setA = new Set(fileA?.columns ?? []);
    const setB = new Set(fileB?.columns ?? []);
    return Array.from(new Set([...(fileA?.columns ?? []), ...(fileB?.columns ?? [])])).map(c => ({
      col: c,
      inA: setA.has(c),
      inB: setB.has(c),
    }));
  }, [fileA, fileB]);

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
    if (sharedColumns.length > 0) {
      setKeyCols(new Set([sharedColumns[0]]));
      setSelectedCols(new Set(sharedColumns));
    }
    setRules([]);
    setStep("config");
  };

  const toggleKeyCol = (col: string) => {
    const next = new Set(keyCols);
    next.has(col) ? next.delete(col) : next.add(col);
    setKeyCols(next);
  };

  const toggleCompareCols = (col: string) => {
    const next = new Set(selectedCols);
    next.has(col) ? next.delete(col) : next.add(col);
    setSelectedCols(next);
  };

  const addRule = () => {
    if (sharedColumns.length === 0) return;
    setRules(prev => [...prev, {
      id: makeId(),
      column: sharedColumns[0],
      operator: "equals",
      value: "",
      target: "both",
    }]);
  };

  const updateRule = (id: string, patch: Partial<Rule>) => {
    setRules(prev => prev.map(r => r.id === id ? { ...r, ...patch } : r));
  };

  const removeRule = (id: string) => {
    setRules(prev => prev.filter(r => r.id !== id));
  };

  const runComparison = () => {
    if (!fileA || !fileB || keyCols.size === 0) return;
    setStep("loading");
    setTimeout(() => {
      try {
        const compareColsArr = Array.from(selectedCols).filter(c => !keyCols.has(c));
        const validRules = rules.filter(r => {
          const op = OPERATORS.find(o => o.value === r.operator);
          return op && (!op.needsValue || r.value.trim() !== "");
        });
        const res = compareSheets(fileA.data, fileB.data, Array.from(keyCols), compareColsArr, validRules);
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
    setKeyCols(new Set()); setSelectedCols(new Set());
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

        {/* ── UPLOAD STEP ── */}
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

        {/* ── CONFIG STEP ── */}
        {step === "config" && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <div className="flex items-center gap-4 mb-8">
              <Button variant="ghost" size="icon" onClick={() => setStep("upload")} className="shrink-0 rounded-full h-8 w-8">
                <ArrowRight className="w-4 h-4 rotate-180" />
              </Button>
              <div>
                <h2 className="text-3xl font-semibold">Configure Comparison</h2>
                <p className="text-muted-foreground text-lg">Set match keys, comparison columns, and filter rules.</p>
              </div>
            </div>

            {/* Step 1: Match Keys */}
            <Card className="border-border/50 shadow-md mb-6 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center gap-2">
                  <Key className="w-5 h-5 text-primary" />
                  <CardTitle className="text-lg">1. Match Keys</CardTitle>
                </div>
                <CardDescription>
                  Select one or more columns that uniquely identify a product across both sheets. Rows are joined by combining all selected keys.
                  <span className="block mt-1 text-xs text-primary/80">e.g. select "SKU" + "Warehouse" to match on both fields together.</span>
                </CardDescription>
              </CardHeader>
              <CardContent>
                {sharedColumns.length === 0 ? (
                  <p className="text-sm text-destructive">No shared columns found between the two files.</p>
                ) : (
                  <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                    {sharedColumns.map(col => (
                      <div
                        key={col}
                        onClick={() => toggleKeyCol(col)}
                        className={`flex items-center gap-2 p-3 rounded-md border cursor-pointer transition-colors select-none ${keyCols.has(col) ? "border-primary bg-primary/10 text-primary" : "border-border/50 bg-background hover:border-primary/40"}`}
                        data-testid={`key-col-${col}`}
                      >
                        <Checkbox checked={keyCols.has(col)} onCheckedChange={() => toggleKeyCol(col)} id={`key-${col}`} className="pointer-events-none" />
                        <label htmlFor={`key-${col}`} className="text-sm font-medium truncate cursor-pointer" title={col}>{col}</label>
                      </div>
                    ))}
                  </div>
                )}
                {keyCols.size > 1 && (
                  <div className="mt-3 flex flex-wrap gap-2">
                    <span className="text-xs text-muted-foreground mt-1">Composite key:</span>
                    {Array.from(keyCols).map((k, i) => (
                      <React.Fragment key={k}>
                        <Badge variant="outline" className="font-mono text-xs border-primary/40 text-primary">{k}</Badge>
                        {i < keyCols.size - 1 && <span className="text-xs text-muted-foreground mt-1">+</span>}
                      </React.Fragment>
                    ))}
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Step 2: Columns to Compare */}
            <Card className="border-border/50 shadow-md mb-6 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <CardTitle className="text-lg">2. Columns to Compare</CardTitle>
                <CardDescription>Choose which fields to check for differences. Key columns are excluded automatically.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="flex gap-3 mb-3">
                  <Button variant="ghost" size="sm" className="text-xs h-7 px-3" onClick={() => setSelectedCols(new Set(sharedColumns))}>Select All</Button>
                  <Button variant="ghost" size="sm" className="text-xs h-7 px-3" onClick={() => setSelectedCols(new Set())}>Clear All</Button>
                </div>
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-3">
                  {allColumns.map(({ col, inA, inB }) => {
                    const isKey = keyCols.has(col);
                    const shared = inA && inB;
                    return (
                      <div
                        key={col}
                        className={`flex items-center gap-2 p-3 rounded-md border transition-colors select-none ${isKey ? "opacity-40 cursor-not-allowed border-border/30 bg-muted/30" : shared ? "cursor-pointer border-border/50 bg-background hover:border-primary/40" : "opacity-50 cursor-not-allowed border-dashed border-border/30"}`}
                        onClick={() => { if (!isKey && shared) toggleCompareCols(col); }}
                        title={!shared ? (inA ? "Only in Online sheet" : "Only in In-Store sheet") : isKey ? "Used as match key" : col}
                      >
                        <Checkbox
                          checked={shared && !isKey && selectedCols.has(col)}
                          disabled={isKey || !shared}
                          onCheckedChange={() => { if (!isKey && shared) toggleCompareCols(col); }}
                          id={`cmp-${col}`}
                          className="pointer-events-none"
                        />
                        <div className="flex-1 min-w-0">
                          <label htmlFor={`cmp-${col}`} className="text-sm font-medium truncate block cursor-pointer" title={col}>{col}</label>
                          {!shared && (
                            <span className="text-xs text-muted-foreground">{inA ? "Online only" : "In-Store only"}</span>
                          )}
                          {isKey && <span className="text-xs text-primary/60">Match key</span>}
                        </div>
                      </div>
                    );
                  })}
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
                  Set conditions to focus only on rows that matter — e.g. "Stock less than 10", "Category contains Electronics", "Price greater than 500".
                </CardDescription>
              </CardHeader>
              <CardContent>
                {rules.length === 0 ? (
                  <div className="border border-dashed border-border/50 rounded-md p-6 text-center">
                    <p className="text-sm text-muted-foreground">No rules set. All rows will be included.</p>
                    <Button variant="ghost" size="sm" className="mt-2 text-primary" onClick={addRule}>
                      <Plus className="w-3 h-3 mr-1" />Add your first rule
                    </Button>
                  </div>
                ) : (
                  <div className="space-y-3">
                    {rules.map((rule, idx) => {
                      const opMeta = OPERATORS.find(o => o.value === rule.operator);
                      return (
                        <div key={rule.id} className="flex items-center gap-2 flex-wrap p-3 bg-background rounded-md border border-border/50" data-testid={`rule-row-${idx}`}>
                          <span className="text-xs font-medium text-muted-foreground w-6 shrink-0">#{idx + 1}</span>

                          {/* Column */}
                          <Select value={rule.column} onValueChange={v => updateRule(rule.id, { column: v })}>
                            <SelectTrigger className="w-40 h-8 text-xs" data-testid={`rule-col-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              {(fileA?.columns ?? []).map(c => (
                                <SelectItem key={c} value={c} className="text-xs">{c}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>

                          {/* Operator */}
                          <Select value={rule.operator} onValueChange={v => updateRule(rule.id, { operator: v as RuleOperator, value: "" })}>
                            <SelectTrigger className="w-44 h-8 text-xs" data-testid={`rule-op-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              {OPERATORS.map(op => (
                                <SelectItem key={op.value} value={op.value} className="text-xs">{op.label}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>

                          {/* Value */}
                          {opMeta?.needsValue && (
                            <Input
                              className="w-36 h-8 text-xs font-mono"
                              placeholder="value..."
                              value={rule.value}
                              onChange={e => updateRule(rule.id, { value: e.target.value })}
                              data-testid={`rule-val-${idx}`}
                            />
                          )}

                          {/* Target */}
                          <Select value={rule.target} onValueChange={v => updateRule(rule.id, { target: v as RuleTarget })}>
                            <SelectTrigger className="w-36 h-8 text-xs" data-testid={`rule-target-${idx}`}>
                              <SelectValue />
                            </SelectTrigger>
                            <SelectContent>
                              {TARGETS.map(t => (
                                <SelectItem key={t.value} value={t.value} className="text-xs">{t.label}</SelectItem>
                              ))}
                            </SelectContent>
                          </Select>

                          <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-destructive shrink-0 ml-auto" onClick={() => removeRule(rule.id)} data-testid={`rule-remove-${idx}`}>
                            <Trash2 className="w-4 h-4" />
                          </Button>
                        </div>
                      );
                    })}
                    {rules.length > 0 && (
                      <p className="text-xs text-muted-foreground pl-1">All rules are applied together (AND logic). Rows must pass every rule to appear in results.</p>
                    )}
                  </div>
                )}
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button
                size="lg"
                onClick={runComparison}
                disabled={keyCols.size === 0 || selectedCols.size === 0}
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

            {/* Active rules pill bar */}
            {results.appliedRules.length > 0 && (
              <div className="flex flex-wrap gap-2 mb-4 items-center">
                <span className="text-xs text-muted-foreground font-medium">Active filters:</span>
                {results.appliedRules.map(r => {
                  const op = OPERATORS.find(o => o.value === r.operator);
                  return (
                    <Badge key={r.id} variant="secondary" className="font-mono text-xs gap-1">
                      <span className="text-primary font-semibold">{r.column}</span>
                      <span className="text-muted-foreground">{op?.label}</span>
                      {op?.needsValue && <span>"{r.value}"</span>}
                      <span className="text-muted-foreground opacity-60">on {r.target === "both" ? "both" : r.target === "A" ? "Online" : r.target === "B" ? "In-Store" : "diffs"}</span>
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
            <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-4 mb-8">
              <MetricCard title="Online Total" value={results.summary.totalA} icon={<ShoppingCart className="w-4 h-4 text-blue-400" />} className="border-blue-400/20 bg-blue-400/5" />
              <MetricCard title="In-Store Total" value={results.summary.totalB} icon={<Store className="w-4 h-4 text-violet-400" />} className="border-violet-400/20 bg-violet-400/5" />
              <MetricCard title="Matched" value={results.summary.matchedCount} icon={<CheckCircle2 className="w-4 h-4 text-emerald-500" />} className="border-emerald-500/20 bg-emerald-500/5" />
              <MetricCard title="Differences" value={results.summary.diffCount} icon={<AlertTriangle className="w-4 h-4 text-accent" />} className="border-accent/20 bg-accent/5" />
              <MetricCard title="Only Online" value={results.summary.onlyACount} icon={<MinusCircle className="w-4 h-4 text-destructive" />} className="border-destructive/20 bg-destructive/5" />
              <MetricCard title="Only In-Store" value={results.summary.onlyBCount} icon={<PlusCircle className="w-4 h-4 text-cyan-400" />} className="border-cyan-400/20 bg-cyan-400/5" />
            </div>

            {/* Key used */}
            <div className="flex flex-wrap gap-2 items-center mb-5">
              <span className="text-xs text-muted-foreground font-medium">Matched by:</span>
              {results.keyCols.map(k => (
                <Badge key={k} variant="outline" className="font-mono text-xs border-primary/30 text-primary"><Key className="w-3 h-3 mr-1" />{k}</Badge>
              ))}
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

              <TabsContent value="diffs" className="m-0">
                <Card className="border-border overflow-hidden">
                  <ScrollArea className="h-[60vh] w-full">
                    {results.differences.length === 0 ? (
                      <EmptyState icon={<CheckCircle2 className="w-12 h-12 text-emerald-500 opacity-80" />} title="No differences found" subtitle="All matched rows are identical in the compared columns." />
                    ) : (
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            {results.keyCols.map(k => (
                              <th key={k} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px]">
                                <span className="flex items-center gap-1"><Key className="w-3 h-3 text-primary" />{k}</span>
                              </th>
                            ))}
                            {results.columns.map(col => (
                              <th key={col} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[200px]">{col}</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {results.differences.map((diff, idx) => {
                            const keyParts = diff.key.split("|||");
                            return (
                              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors">
                                {results.keyCols.map((k, i) => (
                                  <td key={k} className="px-4 py-3 font-mono font-medium border-r bg-background/50 text-xs">{keyParts[i] ?? ""}</td>
                                ))}
                                {results.columns.map(col => {
                                  const isChanged = diff.changedCols.includes(col);
                                  const valA = diff.rowA[col] !== undefined ? String(diff.rowA[col]) : "(empty)";
                                  const valB = diff.rowB[col] !== undefined ? String(diff.rowB[col]) : "(empty)";
                                  return (
                                    <td key={col} className={`px-4 py-3 border-r ${isChanged ? "bg-accent/5" : ""}`}>
                                      {isChanged ? (
                                        <div className="flex flex-col gap-1.5">
                                          <div className="flex items-center gap-1.5 text-destructive font-mono bg-destructive/10 px-2 py-1 rounded text-xs">
                                            <span className="font-bold opacity-60 shrink-0 uppercase text-[10px]">Online</span>
                                            <span className="truncate" title={valA}>{valA}</span>
                                          </div>
                                          <div className="flex items-center gap-1.5 text-emerald-400 font-mono bg-emerald-500/10 px-2 py-1 rounded text-xs">
                                            <span className="font-bold opacity-60 shrink-0 uppercase text-[10px]">Store</span>
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

              <TabsContent value="matches" className="m-0">
                <SimpleDataTable rows={results.matched} columns={[...results.keyCols, ...results.columns]} emptyMessage="No perfectly matched rows." />
              </TabsContent>
              <TabsContent value="onlya" className="m-0">
                <SimpleDataTable rows={results.onlyA} columns={[...results.keyCols, ...results.columns.filter(c => results.onlyA[0] && c in results.onlyA[0])]} emptyMessage="No rows exclusive to Online." />
              </TabsContent>
              <TabsContent value="onlyb" className="m-0">
                <SimpleDataTable rows={results.onlyB} columns={[...results.keyCols, ...results.columns.filter(c => results.onlyB[0] && c in results.onlyB[0])]} emptyMessage="No rows exclusive to In-Store." />
              </TabsContent>
            </Tabs>
          </div>
        )}
      </main>
    </div>
  );
}

function UploadZone({ side, label, icon, file, onDrop, onUpload, onClear }: {
  side: "A" | "B";
  label: string;
  icon: React.ReactNode;
  file: { name: string; data: any[]; columns: string[] } | null;
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

function SimpleDataTable({ rows, columns, emptyMessage }: { rows: any[]; columns: string[]; emptyMessage: string }) {
  if (rows.length === 0) {
    return (
      <Card className="p-12 text-center border-dashed">
        <FileWarning className="w-12 h-12 text-muted-foreground mx-auto mb-4 opacity-50" />
        <p className="text-lg text-muted-foreground font-medium">{emptyMessage}</p>
      </Card>
    );
  }
  const visibleCols = columns.filter(c => rows[0] && (c in rows[0]));
  return (
    <Card className="border-border overflow-hidden">
      <ScrollArea className="h-[60vh] w-full">
        <table className="w-full text-sm text-left whitespace-nowrap">
          <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
            <tr>
              {visibleCols.map(col => (
                <th key={col} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px]">{col}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors">
                {visibleCols.map(col => (
                  <td key={col} className="px-4 py-3 border-r font-mono text-xs text-muted-foreground">
                    {String(row[col] ?? "")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </ScrollArea>
    </Card>
  );
}
