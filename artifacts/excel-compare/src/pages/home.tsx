import React, { useState, useCallback, useMemo } from "react";
import {
  FileSpreadsheet, ArrowRight, RefreshCcw, Download, Plus, Trash2,
  ShoppingCart, Store, ArrowLeftRight, Wand2, Search, X,
  ChevronsUpDown, Check, SlidersHorizontal, Play,
  GitMerge, Minus, PlusCircle, Zap, ChevronLeft
} from "lucide-react";
import {
  parseExcelFile, runOperation, exportToExcel,
  RunResult, Rule, RuleOperator, RuleSheet,
  KeyMapping, CompareMapping, OutputCol, Operation, RunConfig,
  autoSuggestMappings
} from "@/lib/excel-utils";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Badge } from "@/components/ui/badge";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList } from "@/components/ui/command";
import { Switch } from "@/components/ui/switch";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import { useToast } from "@/hooks/use-toast";
import { cn } from "@/lib/utils";

type Step = "upload" | "operation" | "config" | "loading" | "results";
type FileData = { name: string; data: any[]; columns: string[] };

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

const OPERATIONS: { value: Operation; label: string; desc: string; icon: React.ReactNode; color: string }[] = [
  { value: "common", label: "Find Common Products", desc: "Products that exist in BOTH Online and In-Store inventory", icon: <GitMerge className="w-6 h-6" />, color: "border-emerald-500/40 hover:border-emerald-500 hover:bg-emerald-500/5 data-[selected=true]:border-emerald-500 data-[selected=true]:bg-emerald-500/10" },
  { value: "only_a", label: "Only in Online", desc: "Products in Online inventory that are NOT in In-Store", icon: <ShoppingCart className="w-6 h-6" />, color: "border-blue-400/40 hover:border-blue-400 hover:bg-blue-400/5 data-[selected=true]:border-blue-400 data-[selected=true]:bg-blue-400/10" },
  { value: "only_b", label: "Only in In-Store", desc: "Products in In-Store inventory that are NOT Online", icon: <Store className="w-6 h-6" />, color: "border-violet-400/40 hover:border-violet-400 hover:bg-violet-400/5 data-[selected=true]:border-violet-400 data-[selected=true]:bg-violet-400/10" },
  { value: "diff_values", label: "Compare Values", desc: "Products in both sheets where selected values differ", icon: <Zap className="w-6 h-6" />, color: "border-accent/40 hover:border-accent hover:bg-accent/5 data-[selected=true]:border-accent data-[selected=true]:bg-accent/10" },
];

function makeId() { return Math.random().toString(36).slice(2); }

/* ── Searchable Column Picker ── */
function ColumnPicker({ value, columns, onChange, placeholder, accent, testId }: {
  value: string; columns: string[]; onChange: (v: string) => void;
  placeholder?: string; accent?: "blue" | "violet" | "default"; testId?: string;
}) {
  const [open, setOpen] = useState(false);
  const borderCls = accent === "violet" ? "border-violet-400/40 bg-violet-400/5 text-violet-300"
    : accent === "blue" ? "border-blue-400/40 bg-blue-400/5 text-blue-300"
    : "";
  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button variant="outline" role="combobox" aria-expanded={open}
          className={cn("w-full justify-between h-9 text-sm font-normal truncate", borderCls)} data-testid={testId}>
          <span className="truncate">{value || placeholder || "Select column..."}</span>
          <ChevronsUpDown className="ml-2 h-4 w-4 shrink-0 opacity-40" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-64 p-0" align="start">
        <Command>
          <CommandInput placeholder="Search columns..." className="h-9 text-sm" />
          <CommandList>
            <CommandEmpty>No column found.</CommandEmpty>
            <CommandGroup>
              {columns.map(col => (
                <CommandItem key={col} value={col} onSelect={() => { onChange(col); setOpen(false); }} className="text-sm cursor-pointer">
                  <Check className={cn("mr-2 h-4 w-4 shrink-0", value === col ? "opacity-100" : "opacity-0")} />
                  <span className="truncate" title={col}>{col}</span>
                </CommandItem>
              ))}
            </CommandGroup>
          </CommandList>
        </Command>
      </PopoverContent>
    </Popover>
  );
}

export default function Home() {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");
  const [fileA, setFileA] = useState<FileData | null>(null);
  const [fileB, setFileB] = useState<FileData | null>(null);

  const [operation, setOperation] = useState<Operation | null>(null);
  const [keyMappings, setKeyMappings] = useState<KeyMapping[]>([]);
  const [compareMappings, setCompareMappings] = useState<CompareMapping[]>([]);
  const [outputCols, setOutputCols] = useState<OutputCol[]>([]);
  const [rules, setRules] = useState<Rule[]>([]);
  const [caseInsensitive, setCaseInsensitive] = useState(false);
  const [ignoreWS, setIgnoreWS] = useState(true);

  const [result, setResult] = useState<RunResult | null>(null);
  const [resultSearch, setResultSearch] = useState("");

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

  const proceedToOperation = () => {
    if (!fileA || !fileB) return;
    setOperation(null);
    setStep("operation");
  };

  const selectOperation = (op: Operation) => {
    setOperation(op);
    if (!fileA || !fileB) return;
    // Auto-init key mappings
    setKeyMappings([{ id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
    // Init output cols: all from A by default (or B for only_b)
    const src = op === "only_b" ? "B" : "A";
    const srcCols = op === "only_b" ? fileB.columns : fileA.columns;
    setOutputCols([]);
    // Compare mappings for diff_values
    if (op === "diff_values") {
      setCompareMappings([{ id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
    } else {
      setCompareMappings([]);
    }
    setRules([]);
    setStep("config");
  };

  // Key mapping helpers
  const addKeyMapping = () => fileA && fileB && setKeyMappings(p => [...p, { id: makeId(), colA: fileA.columns[0], colB: fileB.columns[0] }]);
  const updateKeyMapping = (id: string, patch: Partial<KeyMapping>) => setKeyMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeKeyMapping = (id: string) => setKeyMappings(p => p.filter(m => m.id !== id));

  // Compare mapping helpers
  const addCompareMapping = () => fileA && fileB && setCompareMappings(p => [...p, { id: makeId(), colA: fileA.columns[0], colB: fileB.columns[0] }]);
  const updateCompareMapping = (id: string, patch: Partial<CompareMapping>) => setCompareMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeCompareMapping = (id: string) => setCompareMappings(p => p.filter(m => m.id !== id));

  // Output col helpers
  const toggleOutputCol = (source: "A" | "B", col: string) => {
    const existing = outputCols.find(o => o.source === source && o.col === col);
    if (existing) {
      setOutputCols(p => p.filter(o => !(o.source === source && o.col === col)));
    } else {
      setOutputCols(p => [...p, { id: makeId(), source, col, label: col }]);
    }
  };
  const isOutputSelected = (source: "A" | "B", col: string) => outputCols.some(o => o.source === source && o.col === col);

  // Rule helpers
  const addRule = () => fileA && setRules(p => [...p, { id: makeId(), sheet: "A", column: fileA.columns[0] ?? "", operator: "equals", value: "" }]);
  const updateRule = (id: string, patch: Partial<Rule>) => setRules(p => p.map(r => r.id === id ? { ...r, ...patch } : r));
  const removeRule = (id: string) => setRules(p => p.filter(r => r.id !== id));

  const handleAutoMap = () => {
    if (!fileA || !fileB) return;
    const suggestions = autoSuggestMappings(fileA.columns, fileB.columns);
    if (suggestions.length === 0) { toast({ title: "No close matches found", description: "Try mapping columns manually." }); return; }
    setKeyMappings(suggestions.slice(0, 1).map(s => ({ id: makeId(), ...s })));
    toast({ title: `Auto-mapped ${suggestions.length} column pair${suggestions.length > 1 ? "s" : ""}` });
  };

  const runAnalysis = () => {
    if (!fileA || !fileB || !operation || keyMappings.length === 0) return;
    if (outputCols.length === 0) { toast({ title: "Select at least one output column", variant: "destructive" }); return; }
    setStep("loading");
    setTimeout(() => {
      try {
        const validRules = rules.filter(r => {
          const op = OPERATORS.find(o => o.value === r.operator);
          return op && (!op.needsValue || r.value.trim() !== "");
        });
        const cfg: RunConfig = { operation, keyMappings, compareMappings, outputCols, rules: validRules, caseInsensitive, ignoreWhitespace: ignoreWS };
        const res = runOperation(fileA.data, fileB.data, cfg);
        setResult(res);
        setResultSearch("");
        setStep("results");
      } catch (e) {
        toast({ title: "Analysis Failed", description: "An error occurred.", variant: "destructive" });
        setStep("config");
      }
    }, 80);
  };

  const reset = () => {
    setFileA(null); setFileB(null); setOperation(null);
    setKeyMappings([]); setCompareMappings([]); setOutputCols([]);
    setRules([]); setResult(null); setResultSearch("");
    setStep("upload");
  };

  const filteredRows = useMemo(() => {
    if (!result) return [];
    if (!resultSearch.trim()) return result.rows;
    const q = resultSearch.toLowerCase();
    return result.rows.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(q)));
  }, [result, resultSearch]);

  const opMeta = OPERATIONS.find(o => o.value === operation);

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col font-sans">
      {/* Header */}
      <header className="border-b border-border bg-card px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-2">
          <FileSpreadsheet className="w-6 h-6 text-primary" />
          <h1 className="text-xl font-bold tracking-tight">ROYAL SHEETS</h1>
          {opMeta && step !== "upload" && step !== "operation" && (
            <Badge variant="secondary" className="ml-2 text-xs font-normal gap-1">
              {opMeta.label}
            </Badge>
          )}
        </div>
        <div className="flex items-center gap-2">
          {step === "results" && (
            <>
              <Button variant="outline" size="sm" onClick={() => setStep("config")}>
                <SlidersHorizontal className="w-4 h-4 mr-2" />Edit
              </Button>
              <Button variant="outline" size="sm" onClick={() => setStep("operation")}>
                <RefreshCcw className="w-4 h-4 mr-2" />Change Mode
              </Button>
              <Button size="sm" className="bg-primary hover:bg-primary/90 text-primary-foreground"
                onClick={() => result && operation && exportToExcel(result, operation)} data-testid="button-export">
                <Download className="w-4 h-4 mr-2" />Export
              </Button>
            </>
          )}
          {(step === "config" || step === "operation") && (
            <Button variant="ghost" size="sm" onClick={reset} className="text-muted-foreground">
              <X className="w-4 h-4 mr-1" />Start Over
            </Button>
          )}
        </div>
      </header>

      <main className="flex-1 p-6 max-w-6xl w-full mx-auto">

        {/* ── STEP 1: UPLOAD ── */}
        {step === "upload" && (
          <div className="animate-in fade-in zoom-in-95 duration-300">
            <div className="mb-8">
              <h2 className="text-3xl font-semibold mb-2">Load Inventory Sheets</h2>
              <p className="text-muted-foreground text-lg">Upload your Online and In-Store inventory files to get started.</p>
            </div>
            <div className="grid md:grid-cols-2 gap-6 mb-8">
              <UploadZone side="A" label="Online Inventory" icon={<ShoppingCart className="w-8 h-8" />} file={fileA} onDrop={handleDrop} onUpload={handleFileUpload} onClear={() => setFileA(null)} />
              <UploadZone side="B" label="In-Store Inventory" icon={<Store className="w-8 h-8" />} file={fileB} onDrop={handleDrop} onUpload={handleFileUpload} onClear={() => setFileB(null)} />
            </div>
            <div className="flex justify-end">
              <Button size="lg" onClick={proceedToOperation} disabled={!fileA || !fileB} className="font-semibold px-8" data-testid="button-proceed">
                Choose What to Find <ArrowRight className="ml-2 w-5 h-5" />
              </Button>
            </div>
          </div>
        )}

        {/* ── STEP 2: CHOOSE OPERATION ── */}
        {step === "operation" && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <div className="flex items-center gap-3 mb-8">
              <Button variant="ghost" size="icon" onClick={() => setStep("upload")} className="rounded-full h-8 w-8 shrink-0">
                <ChevronLeft className="w-4 h-4" />
              </Button>
              <div>
                <h2 className="text-3xl font-semibold">What do you want to find?</h2>
                <p className="text-muted-foreground">Choose an operation — each produces a clean result table.</p>
              </div>
            </div>
            <div className="grid md:grid-cols-2 gap-4">
              {OPERATIONS.map(op => (
                <button
                  key={op.value}
                  data-selected={operation === op.value}
                  onClick={() => selectOperation(op.value)}
                  className={cn(
                    "text-left p-6 rounded-xl border-2 transition-all duration-200 group",
                    op.color
                  )}
                  data-testid={`op-${op.value}`}
                >
                  <div className="flex items-start gap-4">
                    <div className="mt-0.5 text-muted-foreground group-hover:text-foreground transition-colors">{op.icon}</div>
                    <div>
                      <h3 className="font-semibold text-lg mb-1">{op.label}</h3>
                      <p className="text-sm text-muted-foreground leading-relaxed">{op.desc}</p>
                    </div>
                  </div>
                </button>
              ))}
            </div>
          </div>
        )}

        {/* ── STEP 3: CONFIG ── */}
        {step === "config" && fileA && fileB && operation && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <div className="flex items-center gap-3 mb-6">
              <Button variant="ghost" size="icon" onClick={() => setStep("operation")} className="rounded-full h-8 w-8 shrink-0">
                <ChevronLeft className="w-4 h-4" />
              </Button>
              <div className="flex-1 min-w-0">
                <h2 className="text-2xl font-semibold">{opMeta?.label}</h2>
                <p className="text-muted-foreground text-sm">{opMeta?.desc}</p>
              </div>
              <Button variant="outline" size="sm" onClick={handleAutoMap} className="shrink-0 gap-1.5 text-primary border-primary/40 hover:bg-primary/10">
                <Wand2 className="w-4 h-4" />Auto-Match Key
              </Button>
            </div>

            {/* Column headers */}
            <div className="grid grid-cols-[1fr_auto_1fr_auto] gap-2 px-1 mb-2">
              <div className="flex items-center gap-1.5">
                <ShoppingCart className="w-3 h-3 text-blue-400" />
                <span className="text-xs font-semibold text-blue-400 uppercase tracking-wide truncate">{fileA.name}</span>
              </div>
              <div className="w-8" />
              <div className="flex items-center gap-1.5">
                <Store className="w-3 h-3 text-violet-400" />
                <span className="text-xs font-semibold text-violet-400 uppercase tracking-wide truncate">{fileB.name}</span>
              </div>
              <div className="w-9" />
            </div>

            {/* 1. Match Key */}
            <Card className="border-border/50 mb-4 bg-card/60">
              <CardHeader className="pb-2 pt-4 px-4">
                <div className="flex items-center justify-between">
                  <div>
                    <CardTitle className="text-sm font-semibold">1. Match Key</CardTitle>
                    <CardDescription className="text-xs mt-0.5">Column(s) that identify the same product in both files</CardDescription>
                  </div>
                  <Button variant="outline" size="sm" className="h-7 px-2 text-xs" onClick={addKeyMapping}>
                    <Plus className="w-3 h-3 mr-1" />Add
                  </Button>
                </div>
              </CardHeader>
              <CardContent className="px-4 pb-4 space-y-2">
                {keyMappings.map((m, idx) => (
                  <div key={m.id} className="grid grid-cols-[1fr_auto_1fr_auto] gap-2 items-center">
                    <ColumnPicker value={m.colA} columns={fileA.columns} onChange={v => updateKeyMapping(m.id, { colA: v })} accent="blue" testId={`key-colA-${idx}`} />
                    <ArrowLeftRight className="w-4 h-4 text-muted-foreground shrink-0" />
                    <ColumnPicker value={m.colB} columns={fileB.columns} onChange={v => updateKeyMapping(m.id, { colB: v })} accent="violet" testId={`key-colB-${idx}`} />
                    <Button variant="ghost" size="icon" className="h-9 w-9 text-muted-foreground hover:text-destructive" onClick={() => removeKeyMapping(m.id)} disabled={keyMappings.length === 1}>
                      <Trash2 className="w-3.5 h-3.5" />
                    </Button>
                  </div>
                ))}
              </CardContent>
            </Card>

            {/* 2. Compare Columns (diff_values only) */}
            {operation === "diff_values" && (
              <Card className="border-border/50 mb-4 bg-card/60">
                <CardHeader className="pb-2 pt-4 px-4">
                  <div className="flex items-center justify-between">
                    <div>
                      <CardTitle className="text-sm font-semibold">2. Columns to Compare</CardTitle>
                      <CardDescription className="text-xs mt-0.5">Map which values to check for differences</CardDescription>
                    </div>
                    <Button variant="outline" size="sm" className="h-7 px-2 text-xs" onClick={addCompareMapping}>
                      <Plus className="w-3 h-3 mr-1" />Add
                    </Button>
                  </div>
                </CardHeader>
                <CardContent className="px-4 pb-4 space-y-2">
                  {compareMappings.map((m, idx) => (
                    <div key={m.id} className="grid grid-cols-[1fr_auto_1fr_auto] gap-2 items-center">
                      <ColumnPicker value={m.colA} columns={fileA.columns} onChange={v => updateCompareMapping(m.id, { colA: v })} accent="blue" testId={`cmp-colA-${idx}`} />
                      <ArrowLeftRight className="w-4 h-4 text-muted-foreground shrink-0" />
                      <ColumnPicker value={m.colB} columns={fileB.columns} onChange={v => updateCompareMapping(m.id, { colB: v })} accent="violet" testId={`cmp-colB-${idx}`} />
                      <Button variant="ghost" size="icon" className="h-9 w-9 text-muted-foreground hover:text-destructive" onClick={() => removeCompareMapping(m.id)} disabled={compareMappings.length === 1}>
                        <Trash2 className="w-3.5 h-3.5" />
                      </Button>
                    </div>
                  ))}
                </CardContent>
              </Card>
            )}

            {/* 3. Output Columns */}
            <Card className="border-border/50 mb-4 bg-card/60">
              <CardHeader className="pb-2 pt-4 px-4">
                <div className="flex items-center justify-between">
                  <div>
                    <CardTitle className="text-sm font-semibold">{operation === "diff_values" ? "3" : "2"}. Output Columns</CardTitle>
                    <CardDescription className="text-xs mt-0.5">Choose which columns appear in the result table</CardDescription>
                  </div>
                  <div className="flex gap-1">
                    <Button variant="ghost" size="sm" className="h-7 px-2 text-xs text-muted-foreground"
                      onClick={() => {
                        const src = operation === "only_b" ? "B" : "A";
                        const cols = operation === "only_b" ? fileB.columns : fileA.columns;
                        setOutputCols(cols.map(c => ({ id: makeId(), source: src, col: c, label: c })));
                      }}>All A</Button>
                    <Button variant="ghost" size="sm" className="h-7 px-2 text-xs text-muted-foreground"
                      onClick={() => setOutputCols([])}>Clear</Button>
                  </div>
                </div>
              </CardHeader>
              <CardContent className="px-4 pb-4">
                <div className="grid md:grid-cols-2 gap-3">
                  {/* Sheet A columns */}
                  {(operation !== "only_b") && (
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <ShoppingCart className="w-3 h-3 text-blue-400" />
                        <span className="text-xs font-semibold text-blue-400 uppercase tracking-wide">Online</span>
                        <span className="text-xs text-muted-foreground ml-1">({fileA.columns.length} cols)</span>
                      </div>
                      <div className="space-y-1 max-h-48 overflow-y-auto pr-1">
                        {fileA.columns.map(col => (
                          <label key={col} className="flex items-center gap-2 px-2 py-1.5 rounded hover:bg-muted/40 cursor-pointer group">
                            <Checkbox
                              checked={isOutputSelected("A", col)}
                              onCheckedChange={() => toggleOutputCol("A", col)}
                              className="shrink-0"
                            />
                            <span className="text-sm truncate" title={col}>{col}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                  )}
                  {/* Sheet B columns */}
                  {(operation !== "only_a") && (
                    <div>
                      <div className="flex items-center gap-1.5 mb-2">
                        <Store className="w-3 h-3 text-violet-400" />
                        <span className="text-xs font-semibold text-violet-400 uppercase tracking-wide">In-Store</span>
                        <span className="text-xs text-muted-foreground ml-1">({fileB.columns.length} cols)</span>
                      </div>
                      <div className="space-y-1 max-h-48 overflow-y-auto pr-1">
                        {fileB.columns.map(col => (
                          <label key={col} className="flex items-center gap-2 px-2 py-1.5 rounded hover:bg-muted/40 cursor-pointer group">
                            <Checkbox
                              checked={isOutputSelected("B", col)}
                              onCheckedChange={() => toggleOutputCol("B", col)}
                              className="shrink-0"
                            />
                            <span className="text-sm truncate" title={col}>{col}</span>
                          </label>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
                {outputCols.length > 0 && (
                  <p className="text-xs text-muted-foreground mt-3">{outputCols.length} column{outputCols.length > 1 ? "s" : ""} selected</p>
                )}
              </CardContent>
            </Card>

            {/* 4. Filters */}
            <Card className="border-border/50 mb-4 bg-card/60">
              <CardHeader className="pb-2 pt-4 px-4">
                <div className="flex items-center justify-between">
                  <div>
                    <CardTitle className="text-sm font-semibold">{operation === "diff_values" ? "4" : "3"}. Filters <span className="font-normal text-muted-foreground">(optional)</span></CardTitle>
                    <CardDescription className="text-xs mt-0.5">Narrow rows before searching — e.g. Category = Electronics, Stock &lt; 5</CardDescription>
                  </div>
                  <Button variant="outline" size="sm" className="h-7 px-2 text-xs" onClick={addRule}>
                    <Plus className="w-3 h-3 mr-1" />Add Filter
                  </Button>
                </div>
              </CardHeader>
              <CardContent className="px-4 pb-4">
                {rules.length === 0 ? (
                  <p className="text-sm text-muted-foreground italic">No filters — all rows included.</p>
                ) : (
                  <div className="space-y-2">
                    {rules.map((rule, idx) => {
                      const opMeta = OPERATORS.find(o => o.value === rule.operator);
                      const cols = rule.sheet === "A" ? fileA.columns : fileB.columns;
                      return (
                        <div key={rule.id} className="flex items-center gap-2 flex-wrap p-2.5 bg-background rounded-md border border-border/40">
                          <Select value={rule.sheet} onValueChange={v => updateRule(rule.id, { sheet: v as RuleSheet, column: (v === "A" ? fileA : fileB).columns[0] ?? "" })}>
                            <SelectTrigger className="w-28 h-8 text-xs"><SelectValue /></SelectTrigger>
                            <SelectContent>
                              <SelectItem value="A" className="text-xs"><span className="flex items-center gap-1"><ShoppingCart className="w-3 h-3 text-blue-400" />Online</span></SelectItem>
                              <SelectItem value="B" className="text-xs"><span className="flex items-center gap-1"><Store className="w-3 h-3 text-violet-400" />In-Store</span></SelectItem>
                            </SelectContent>
                          </Select>
                          <div className="w-40"><ColumnPicker value={rule.column} columns={cols} onChange={v => updateRule(rule.id, { column: v })} accent={rule.sheet === "A" ? "blue" : "violet"} testId={`rule-col-${idx}`} /></div>
                          <Select value={rule.operator} onValueChange={v => updateRule(rule.id, { operator: v as RuleOperator, value: "" })}>
                            <SelectTrigger className="w-40 h-8 text-xs"><SelectValue /></SelectTrigger>
                            <SelectContent>{OPERATORS.map(op => <SelectItem key={op.value} value={op.value} className="text-xs">{op.label}</SelectItem>)}</SelectContent>
                          </Select>
                          {opMeta?.needsValue && (
                            <Input className="w-28 h-8 text-xs font-mono" placeholder="value..." value={rule.value} onChange={e => updateRule(rule.id, { value: e.target.value })} />
                          )}
                          <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-destructive ml-auto" onClick={() => removeRule(rule.id)}>
                            <Trash2 className="w-3.5 h-3.5" />
                          </Button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Options */}
            <Card className="border-border/50 mb-6 bg-card/60">
              <CardContent className="px-4 py-3 flex flex-wrap gap-6">
                <div className="flex items-center gap-2.5">
                  <Switch id="ci" checked={caseInsensitive} onCheckedChange={setCaseInsensitive} />
                  <Label htmlFor="ci" className="text-sm cursor-pointer">Case-insensitive <span className="text-xs text-muted-foreground">("Apple" = "apple")</span></Label>
                </div>
                <div className="flex items-center gap-2.5">
                  <Switch id="ws" checked={ignoreWS} onCheckedChange={setIgnoreWS} />
                  <Label htmlFor="ws" className="text-sm cursor-pointer">Ignore extra spaces</Label>
                </div>
              </CardContent>
            </Card>

            <Button size="lg" onClick={runAnalysis} disabled={keyMappings.length === 0 || outputCols.length === 0}
              className="w-full font-bold h-13 text-base bg-primary hover:bg-primary/90 text-primary-foreground" data-testid="button-run">
              <Play className="mr-2 w-5 h-5 fill-current" />Run — {opMeta?.label}
            </Button>
          </div>
        )}

        {/* ── LOADING ── */}
        {step === "loading" && (
          <div className="h-[60vh] flex flex-col items-center justify-center animate-in fade-in duration-300">
            <div className="w-14 h-14 border-4 border-primary border-t-transparent rounded-full animate-spin mb-6" />
            <h2 className="text-2xl font-semibold mb-1">Running Analysis</h2>
            <p className="text-muted-foreground">{opMeta?.desc}</p>
          </div>
        )}

        {/* ── RESULTS ── */}
        {step === "results" && result && (
          <div className="animate-in fade-in zoom-in-95 duration-300">

            {/* Result header */}
            <div className="flex items-center justify-between mb-5">
              <div>
                <div className="flex items-center gap-2 mb-1">
                  <div className="text-muted-foreground">{opMeta?.icon}</div>
                  <h2 className="text-2xl font-semibold">{opMeta?.label}</h2>
                </div>
                <p className="text-muted-foreground text-sm">{result.summary}</p>
              </div>
              <div className="flex items-center gap-3">
                <div className="text-right">
                  <p className="text-3xl font-bold font-mono text-primary">{result.rows.length.toLocaleString()}</p>
                  <p className="text-xs text-muted-foreground">rows found</p>
                </div>
              </div>
            </div>

            {/* Stats row */}
            <div className="flex flex-wrap gap-3 mb-5">
              <Badge variant="outline" className="gap-1 text-xs font-mono">
                <ShoppingCart className="w-3 h-3 text-blue-400" />{result.totalA.toLocaleString()} Online rows
              </Badge>
              <Badge variant="outline" className="gap-1 text-xs font-mono">
                <Store className="w-3 h-3 text-violet-400" />{result.totalB.toLocaleString()} In-Store rows
              </Badge>
              {result.columns.length > 0 && (
                <Badge variant="outline" className="gap-1 text-xs font-mono">
                  {result.columns.length} column{result.columns.length > 1 ? "s" : ""}
                </Badge>
              )}
            </div>

            {/* Search */}
            <div className="relative mb-4">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <Input className="pl-9 h-9 text-sm bg-card border-border/50" placeholder="Search in results..."
                value={resultSearch} onChange={e => setResultSearch(e.target.value)} data-testid="input-search" />
              {resultSearch && (
                <Button variant="ghost" size="icon" className="absolute right-1 top-1/2 -translate-y-1/2 h-7 w-7"
                  onClick={() => setResultSearch("")}><X className="w-3 h-3" /></Button>
              )}
            </div>
            {resultSearch && (
              <p className="text-xs text-muted-foreground mb-3">{filteredRows.length} of {result.rows.length} rows match "{resultSearch}"</p>
            )}

            {/* Result table */}
            {result.rows.length === 0 ? (
              <Card className="p-16 text-center border-dashed">
                <div className="text-muted-foreground mb-3">{opMeta?.icon}</div>
                <p className="text-lg font-medium mb-1">No results found</p>
                <p className="text-sm text-muted-foreground">Try adjusting your match key, filters, or operation type.</p>
                <Button variant="outline" className="mt-4" onClick={() => setStep("config")}>
                  <ChevronLeft className="w-4 h-4 mr-1" />Edit Configuration
                </Button>
              </Card>
            ) : (
              <Card className="border-border overflow-hidden">
                <ScrollArea className="h-[62vh] w-full">
                  <table className="w-full text-sm text-left whitespace-nowrap">
                    <thead className="text-xs bg-muted/60 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                      <tr>
                        <th className="px-3 py-3 font-medium border-b border-r bg-card/90 text-muted-foreground/60 w-12 text-center">#</th>
                        {result.columns.map(col => (
                          <th key={col.key} className={cn(
                            "px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px] uppercase tracking-wide",
                            col.isDiff ? "text-amber-400" : ""
                          )}>
                            {col.label}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRows.map((row, idx) => (
                        <tr key={idx} className="border-b border-border/50 hover:bg-muted/20 transition-colors" data-testid={`result-row-${idx}`}>
                          <td className="px-3 py-3 text-xs text-muted-foreground/40 text-center border-r font-mono">{idx + 1}</td>
                          {result.columns.map((col, ci) => {
                            const val = row[col.key] ?? "";
                            // For diff_values, alternate coloring for A/B pairs
                            const isDiffA = col.isDiff === false && result.columns[ci + 1]?.isDiff === true;
                            const isDiffB = col.isDiff === true;
                            return (
                              <td key={col.key} className={cn(
                                "px-4 py-3 border-r font-mono text-xs",
                                isDiffA ? "text-blue-300 bg-blue-400/5" : isDiffB ? "text-violet-300 bg-violet-400/5" : "text-foreground/80"
                              )}>
                                <span className="truncate block max-w-[200px]" title={val}>{val || <span className="text-muted-foreground/40 italic">—</span>}</span>
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </ScrollArea>
                <div className="px-4 py-2 border-t border-border/40 bg-card/50 flex items-center justify-between">
                  <span className="text-xs text-muted-foreground">
                    {filteredRows.length < result.rows.length
                      ? `Showing ${filteredRows.length} of ${result.rows.length} rows`
                      : `${result.rows.length} row${result.rows.length !== 1 ? "s" : ""}`}
                  </span>
                  <Button size="sm" variant="ghost" className="text-xs h-7 gap-1 text-muted-foreground"
                    onClick={() => result && operation && exportToExcel(result, operation)}>
                    <Download className="w-3 h-3" />Export table
                  </Button>
                </div>
              </Card>
            )}
          </div>
        )}
      </main>
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
        <CardContent className="p-5">
          <div className="flex justify-between items-start mb-3">
            <div>
              <Badge variant="outline" className="mb-1.5 bg-background text-xs">{label}</Badge>
              <h3 className="font-semibold text-base truncate pr-6" title={file.name}>{file.name}</h3>
              <p className="text-xs text-muted-foreground font-mono mt-0.5">{file.data.length.toLocaleString()} rows · {file.columns.length} cols</p>
            </div>
            <Button variant="ghost" size="sm" onClick={onClear} className="text-muted-foreground hover:text-destructive text-xs h-7">Clear</Button>
          </div>
          <div className="bg-background rounded border border-border overflow-hidden">
            <table className="w-full text-left text-xs whitespace-nowrap">
              <thead className="bg-muted/50 text-muted-foreground">
                <tr>
                  {file.columns.slice(0, 4).map(c => <th key={c} className="px-3 py-2 border-b font-medium max-w-[120px] truncate">{c}</th>)}
                  {file.columns.length > 4 && <th className="px-3 py-2 border-b text-muted-foreground/50">+{file.columns.length - 4}</th>}
                </tr>
              </thead>
              <tbody className="font-mono">
                {file.data.slice(0, 3).map((row, i) => (
                  <tr key={i} className="border-b last:border-0 hover:bg-muted/20">
                    {file.columns.slice(0, 4).map(c => <td key={c} className="px-3 py-1.5 max-w-[120px] truncate opacity-70">{String(row[c] ?? "")}</td>)}
                    {file.columns.length > 4 && <td className="px-3 py-1.5 opacity-30">...</td>}
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
    <Card className="border-dashed border-2 hover:border-primary/50 hover:bg-card/80 transition-all cursor-pointer group bg-card/30"
      onDragOver={e => e.preventDefault()} onDrop={e => onDrop(e, side)} onClick={() => document.getElementById(inputId)?.click()}>
      <CardContent className="flex flex-col items-center justify-center p-12 text-center min-h-[260px]">
        <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mb-4 group-hover:scale-110 transition-transform group-hover:bg-primary/10 text-muted-foreground group-hover:text-primary">{icon}</div>
        <h3 className="text-xl font-medium mb-1">{label}</h3>
        <p className="text-muted-foreground text-sm mb-5 max-w-[200px]">Drag & drop .xlsx or .csv, or click to browse</p>
        <Button variant="secondary" className="pointer-events-none group-hover:bg-primary group-hover:text-primary-foreground transition-colors">Select File</Button>
        <input type="file" id={inputId} className="hidden" accept=".xlsx,.xls,.csv" onChange={e => { if (e.target.files?.[0]) onUpload(e.target.files[0], side); }} />
      </CardContent>
    </Card>
  );
}
