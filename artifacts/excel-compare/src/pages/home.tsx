import React, { useState, useCallback, useMemo } from "react";
import {
  UploadCloud, FileSpreadsheet, ArrowRight, Play, RefreshCcw, Download,
  CheckCircle2, AlertTriangle, FileWarning, PlusCircle, MinusCircle,
  Key, SlidersHorizontal, Trash2, Plus, X, ShoppingCart, Store,
  ArrowLeftRight, Wand2, Search, TrendingUp, TrendingDown, Minus,
  ChevronsUpDown, Check
} from "lucide-react";
import {
  parseExcelFile, compareSheets, exportResultsToExcel,
  ComparisonResult, Rule, RuleOperator, RuleSheet, ColMapping,
  CompareOptions, autoSuggestMappings
} from "@/lib/excel-utils";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Input } from "@/components/ui/input";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Badge } from "@/components/ui/badge";
import { Popover, PopoverContent, PopoverTrigger } from "@/components/ui/popover";
import { Command, CommandEmpty, CommandGroup, CommandInput, CommandItem, CommandList } from "@/components/ui/command";
import { Switch } from "@/components/ui/switch";

import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { cn } from "@/lib/utils";

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

/* ── Searchable Column Picker ── */
function ColumnPicker({
  value, columns, onChange, placeholder, accent, testId
}: {
  value: string; columns: string[]; onChange: (v: string) => void;
  placeholder?: string; accent?: "blue" | "violet"; testId?: string;
}) {
  const [open, setOpen] = useState(false);
  const borderCls = accent === "violet" ? "border-violet-400/40 bg-violet-400/5" : "border-blue-400/40 bg-blue-400/5";
  const textCls = accent === "violet" ? "text-violet-300" : "text-blue-300";

  return (
    <Popover open={open} onOpenChange={setOpen}>
      <PopoverTrigger asChild>
        <Button
          variant="outline"
          role="combobox"
          aria-expanded={open}
          className={cn("w-full justify-between h-9 text-sm font-normal", borderCls, textCls)}
          data-testid={testId}
        >
          <span className="truncate">{value || placeholder || "Select column..."}</span>
          <ChevronsUpDown className="ml-2 h-4 w-4 shrink-0 opacity-50" />
        </Button>
      </PopoverTrigger>
      <PopoverContent className="w-64 p-0" align="start">
        <Command>
          <CommandInput placeholder="Search columns..." className="h-9 text-sm" />
          <CommandList>
            <CommandEmpty>No column found.</CommandEmpty>
            <CommandGroup>
              {columns.map(col => (
                <CommandItem
                  key={col}
                  value={col}
                  onSelect={() => { onChange(col); setOpen(false); }}
                  className="text-sm cursor-pointer"
                >
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

/* ── Mapping Row ── */
function MappingRow({ mapping, colsA, colsB, idx, prefix, onUpdateA, onUpdateB, onRemove, canRemove }: {
  mapping: ColMapping; colsA: string[]; colsB: string[];
  idx: number; prefix: string;
  onUpdateA: (v: string) => void; onUpdateB: (v: string) => void;
  onRemove: () => void; canRemove: boolean;
}) {
  return (
    <div className="grid grid-cols-[1fr_auto_1fr_auto] gap-2 items-center" data-testid={`${prefix}-row-${idx}`}>
      <ColumnPicker value={mapping.colA} columns={colsA} onChange={onUpdateA} accent="blue" placeholder="Online column..." testId={`${prefix}-colA-${idx}`} />
      <ArrowLeftRight className="w-4 h-4 text-muted-foreground shrink-0" />
      <ColumnPicker value={mapping.colB} columns={colsB} onChange={onUpdateB} accent="violet" placeholder="In-Store column..." testId={`${prefix}-colB-${idx}`} />
      <Button
        variant="ghost" size="icon" className="h-9 w-9 text-muted-foreground hover:text-destructive shrink-0"
        onClick={onRemove} disabled={!canRemove} data-testid={`${prefix}-remove-${idx}`}
      >
        <Trash2 className="w-4 h-4" />
      </Button>
    </div>
  );
}

export default function Home() {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");

  const [fileA, setFileA] = useState<FileData | null>(null);
  const [fileB, setFileB] = useState<FileData | null>(null);

  const [keyMappings, setKeyMappings] = useState<ColMapping[]>([]);
  const [compareMappings, setCompareMappings] = useState<ColMapping[]>([]);
  const [rules, setRules] = useState<Rule[]>([]);
  const [options, setOptions] = useState<CompareOptions>({ caseInsensitive: false, ignoreWhitespace: true });

  const [results, setResults] = useState<ComparisonResult | null>(null);
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

  const proceedToConfig = () => {
    if (!fileA || !fileB) return;
    setKeyMappings([{ id: makeId(), colA: fileA.columns[0] ?? "", colB: fileB.columns[0] ?? "" }]);
    setCompareMappings([{ id: makeId(), colA: fileA.columns[1] ?? fileA.columns[0] ?? "", colB: fileB.columns[1] ?? fileB.columns[0] ?? "" }]);
    setRules([]);
    setResultSearch("");
    setStep("config");
  };

  const handleAutoMap = () => {
    if (!fileA || !fileB) return;
    const suggestions = autoSuggestMappings(fileA.columns, fileB.columns);
    if (suggestions.length === 0) {
      toast({ title: "No matches found", description: "Column names are too different to auto-suggest mappings." });
      return;
    }
    // First suggestion → key, rest → compare
    setKeyMappings([suggestions[0]]);
    setCompareMappings(suggestions.slice(1).length > 0 ? suggestions.slice(1) : [{ id: makeId(), colA: fileA.columns[0], colB: fileB.columns[0] }]);
    toast({ title: `Auto-mapped ${suggestions.length} column${suggestions.length > 1 ? "s" : ""}`, description: "Review the suggestions and adjust as needed." });
  };

  const addKeyMapping = () => fileA && fileB && setKeyMappings(p => [...p, { id: makeId(), colA: fileA.columns[0], colB: fileB.columns[0] }]);
  const updateKeyMapping = (id: string, patch: Partial<ColMapping>) => setKeyMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeKeyMapping = (id: string) => setKeyMappings(p => p.filter(m => m.id !== id));

  const addCompareMapping = () => fileA && fileB && setCompareMappings(p => [...p, { id: makeId(), colA: fileA.columns[0], colB: fileB.columns[0] }]);
  const updateCompareMapping = (id: string, patch: Partial<ColMapping>) => setCompareMappings(p => p.map(m => m.id === id ? { ...m, ...patch } : m));
  const removeCompareMapping = (id: string) => setCompareMappings(p => p.filter(m => m.id !== id));

  const addRule = () => fileA && setRules(p => [...p, { id: makeId(), sheet: "A", column: fileA.columns[0] ?? "", operator: "equals", value: "" }]);
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
        const res = compareSheets(fileA.data, fileB.data, keyMappings, compareMappings, validRules, options);
        setResults(res);
        setResultSearch("");
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
    setRules([]); setResults(null); setResultSearch("");
    setStep("upload");
  };

  // Filtered results for search
  const filteredDiffs = useMemo(() => {
    if (!results) return [];
    if (!resultSearch.trim()) return results.differences;
    const q = resultSearch.toLowerCase();
    return results.differences.filter(diff => {
      const keyMatch = diff.key.toLowerCase().includes(q);
      const valMatch = diff.changedMappings.some(m => {
        const va = String(diff.rowA[m.colA] ?? "").toLowerCase();
        const vb = String(diff.rowB[m.colB] ?? "").toLowerCase();
        return va.includes(q) || vb.includes(q);
      });
      return keyMatch || valMatch;
    });
  }, [results, resultSearch]);

  const filteredOnlyA = useMemo(() => {
    if (!results || !resultSearch.trim()) return results?.onlyA ?? [];
    const q = resultSearch.toLowerCase();
    return results.onlyA.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(q)));
  }, [results, resultSearch]);

  const filteredOnlyB = useMemo(() => {
    if (!results || !resultSearch.trim()) return results?.onlyB ?? [];
    const q = resultSearch.toLowerCase();
    return results.onlyB.filter(row => Object.values(row).some(v => String(v).toLowerCase().includes(q)));
  }, [results, resultSearch]);

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
            <Button variant="outline" onClick={() => setStep("config")} size="sm">
              <SlidersHorizontal className="w-4 h-4 mr-2" />Edit Config
            </Button>
            <Button variant="outline" onClick={reset} size="sm" data-testid="button-reset">
              <RefreshCcw className="w-4 h-4 mr-2" />New
            </Button>
            <Button
              onClick={() => results && exportResultsToExcel(results, fileA?.name ?? "Online", fileB?.name ?? "In-Store")}
              size="sm" className="bg-primary hover:bg-primary/90 text-primary-foreground" data-testid="button-export"
            >
              <Download className="w-4 h-4 mr-2" />Export
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
            <div className="flex items-center gap-4 mb-6">
              <Button variant="ghost" size="icon" onClick={() => setStep("upload")} className="shrink-0 rounded-full h-8 w-8">
                <ArrowRight className="w-4 h-4 rotate-180" />
              </Button>
              <div className="flex-1">
                <h2 className="text-3xl font-semibold">Configure Comparison</h2>
                <p className="text-muted-foreground">Map columns between files, set rules, and choose options.</p>
              </div>
              <Button variant="outline" onClick={handleAutoMap} className="shrink-0 gap-2 border-primary/40 text-primary hover:bg-primary/10" data-testid="button-auto-map">
                <Wand2 className="w-4 h-4" />Auto-Map Columns
              </Button>
            </div>

            {/* Column header labels */}
            <div className="grid grid-cols-[1fr_auto_1fr_auto] gap-2 px-1 mb-2 items-center">
              <div className="flex items-center gap-1.5">
                <ShoppingCart className="w-3.5 h-3.5 text-blue-400 shrink-0" />
                <span className="text-xs font-semibold text-blue-400 uppercase tracking-wide truncate">{fileA.name}</span>
              </div>
              <div className="w-8" />
              <div className="flex items-center gap-1.5">
                <Store className="w-3.5 h-3.5 text-violet-400 shrink-0" />
                <span className="text-xs font-semibold text-violet-400 uppercase tracking-wide truncate">{fileB.name}</span>
              </div>
              <div className="w-9" />
            </div>

            {/* Step 1: Match Keys */}
            <Card className="border-border/50 shadow-md mb-5 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <Key className="w-4 h-4 text-primary" />
                    <CardTitle className="text-base">1. Match Keys</CardTitle>
                  </div>
                  <Button variant="outline" size="sm" className="h-7 px-3 text-xs" onClick={addKeyMapping} data-testid="button-add-key">
                    <Plus className="w-3 h-3 mr-1" />Add Key
                  </Button>
                </div>
                <CardDescription className="text-xs">
                  Column that identifies the same product in both files. Names can differ — e.g. "Product Code" ↔ "Item ID". Search by typing.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-2">
                  {keyMappings.map((m, idx) => (
                    <MappingRow key={m.id} mapping={m} colsA={fileA.columns} colsB={fileB.columns} idx={idx} prefix="key"
                      onUpdateA={v => updateKeyMapping(m.id, { colA: v })} onUpdateB={v => updateKeyMapping(m.id, { colB: v })}
                      onRemove={() => removeKeyMapping(m.id)} canRemove={keyMappings.length > 1} />
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* Step 2: Columns to Compare */}
            <Card className="border-border/50 shadow-md mb-5 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <ArrowLeftRight className="w-4 h-4 text-primary" />
                    <CardTitle className="text-base">2. Columns to Compare</CardTitle>
                  </div>
                  <Button variant="outline" size="sm" className="h-7 px-3 text-xs" onClick={addCompareMapping} data-testid="button-add-compare">
                    <Plus className="w-3 h-3 mr-1" />Add Column
                  </Button>
                </div>
                <CardDescription className="text-xs">
                  Map value fields to compare across both files — e.g. "Online Price" vs "Store Price", "Web Stock" vs "Shelf Count".
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-2">
                  {compareMappings.map((m, idx) => (
                    <MappingRow key={m.id} mapping={m} colsA={fileA.columns} colsB={fileB.columns} idx={idx} prefix="cmp"
                      onUpdateA={v => updateCompareMapping(m.id, { colA: v })} onUpdateB={v => updateCompareMapping(m.id, { colB: v })}
                      onRemove={() => removeCompareMapping(m.id)} canRemove={compareMappings.length > 1} />
                  ))}
                </div>
              </CardContent>
            </Card>

            {/* Step 3: Rules */}
            <Card className="border-border/50 shadow-md mb-5 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <SlidersHorizontal className="w-4 h-4 text-primary" />
                    <CardTitle className="text-base">3. Filter Rules <span className="text-sm font-normal text-muted-foreground ml-1">(optional)</span></CardTitle>
                  </div>
                  <Button variant="outline" size="sm" className="h-7 px-3 text-xs" onClick={addRule} data-testid="button-add-rule">
                    <Plus className="w-3 h-3 mr-1" />Add Rule
                  </Button>
                </div>
                <CardDescription className="text-xs">Filter rows before comparing — e.g. "Stock less than 10", "Category contains Electronics".</CardDescription>
              </CardHeader>
              <CardContent>
                {rules.length === 0 ? (
                  <div className="border border-dashed border-border/40 rounded-md p-5 text-center">
                    <p className="text-sm text-muted-foreground">No rules — all rows included.</p>
                    <Button variant="ghost" size="sm" className="mt-1 text-primary h-7 text-xs" onClick={addRule}><Plus className="w-3 h-3 mr-1" />Add first rule</Button>
                  </div>
                ) : (
                  <div className="space-y-2">
                    {rules.map((rule, idx) => {
                      const opMeta = OPERATORS.find(o => o.value === rule.operator);
                      const cols = rule.sheet === "A" ? fileA.columns : fileB.columns;
                      return (
                        <div key={rule.id} className="flex items-center gap-2 flex-wrap p-3 bg-background rounded-md border border-border/40" data-testid={`rule-row-${idx}`}>
                          <span className="text-xs text-muted-foreground w-5 shrink-0">#{idx + 1}</span>
                          {/* Sheet */}
                          <Select value={rule.sheet} onValueChange={v => updateRule(rule.id, { sheet: v as RuleSheet, column: (v === "A" ? fileA : fileB).columns[0] ?? "" })}>
                            <SelectTrigger className="w-28 h-8 text-xs"><SelectValue /></SelectTrigger>
                            <SelectContent>
                              <SelectItem value="A" className="text-xs"><span className="flex items-center gap-1"><ShoppingCart className="w-3 h-3 text-blue-400" />Online</span></SelectItem>
                              <SelectItem value="B" className="text-xs"><span className="flex items-center gap-1"><Store className="w-3 h-3 text-violet-400" />In-Store</span></SelectItem>
                            </SelectContent>
                          </Select>
                          {/* Column — searchable */}
                          <div className="w-44">
                            <ColumnPicker value={rule.column} columns={cols} onChange={v => updateRule(rule.id, { column: v })} accent={rule.sheet === "A" ? "blue" : "violet"} testId={`rule-col-${idx}`} />
                          </div>
                          {/* Operator */}
                          <Select value={rule.operator} onValueChange={v => updateRule(rule.id, { operator: v as RuleOperator, value: "" })}>
                            <SelectTrigger className="w-44 h-8 text-xs"><SelectValue /></SelectTrigger>
                            <SelectContent>{OPERATORS.map(op => <SelectItem key={op.value} value={op.value} className="text-xs">{op.label}</SelectItem>)}</SelectContent>
                          </Select>
                          {/* Value */}
                          {opMeta?.needsValue && (
                            <Input className="w-28 h-8 text-xs font-mono" placeholder="value..." value={rule.value} onChange={e => updateRule(rule.id, { value: e.target.value })} data-testid={`rule-val-${idx}`} />
                          )}
                          <Button variant="ghost" size="icon" className="h-8 w-8 text-muted-foreground hover:text-destructive shrink-0 ml-auto" onClick={() => removeRule(rule.id)}><Trash2 className="w-4 h-4" /></Button>
                        </div>
                      );
                    })}
                    <p className="text-xs text-muted-foreground pl-1">All rules applied with AND logic.</p>
                  </div>
                )}
              </CardContent>
            </Card>

            {/* Step 4: Options */}
            <Card className="border-border/50 shadow-md mb-8 bg-card/50 backdrop-blur">
              <CardHeader className="pb-3">
                <CardTitle className="text-base">4. Comparison Options</CardTitle>
              </CardHeader>
              <CardContent>
                <div className="flex flex-wrap gap-6">
                  <div className="flex items-center gap-3">
                    <Switch id="case-insensitive" checked={options.caseInsensitive} onCheckedChange={v => setOptions(o => ({ ...o, caseInsensitive: v }))} data-testid="toggle-case" />
                    <div>
                      <Label htmlFor="case-insensitive" className="text-sm cursor-pointer">Case-insensitive</Label>
                      <p className="text-xs text-muted-foreground">"Apple" = "apple"</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-3">
                    <Switch id="ignore-ws" checked={options.ignoreWhitespace} onCheckedChange={v => setOptions(o => ({ ...o, ignoreWhitespace: v }))} data-testid="toggle-ws" />
                    <div>
                      <Label htmlFor="ignore-ws" className="text-sm cursor-pointer">Ignore whitespace</Label>
                      <p className="text-xs text-muted-foreground">Trims leading/trailing spaces</p>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button size="lg" onClick={runComparison} disabled={keyMappings.length === 0 || compareMappings.length === 0}
                className="w-full md:w-auto font-bold px-10 h-14 text-lg bg-primary hover:bg-primary/90 text-primary-foreground" data-testid="button-run-comparison">
                <Play className="mr-2 w-5 h-5 fill-current" />Run Analysis
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
            {results.appliedRules.length > 0 && (
              <div className="flex flex-wrap gap-2 mb-4 items-center">
                <span className="text-xs text-muted-foreground font-medium">Filters:</span>
                {results.appliedRules.map(r => {
                  const op = OPERATORS.find(o => o.value === r.operator);
                  return (
                    <Badge key={r.id} variant="secondary" className="font-mono text-xs gap-1">
                      <span className="text-muted-foreground opacity-70">{r.sheet === "A" ? "Online" : "In-Store"}:</span>
                      <span className="text-primary font-semibold">{r.column}</span>
                      <span>{op?.label}</span>
                      {op?.needsValue && <span>"{r.value}"</span>}
                    </Badge>
                  );
                })}
                {results.summary.filteredCount > 0 && (
                  <Badge variant="outline" className="text-xs text-muted-foreground gap-1">
                    <X className="w-3 h-3" />{results.summary.filteredCount} filtered out
                  </Badge>
                )}
              </div>
            )}

            {/* Options used */}
            <div className="flex flex-wrap gap-2 mb-4 items-center">
              <span className="text-xs text-muted-foreground">Options:</span>
              {results.options.caseInsensitive && <Badge variant="outline" className="text-xs">Case-insensitive</Badge>}
              {results.options.ignoreWhitespace && <Badge variant="outline" className="text-xs">Trim whitespace</Badge>}
              {!results.options.caseInsensitive && !results.options.ignoreWhitespace && (
                <span className="text-xs text-muted-foreground">Exact match</span>
              )}
            </div>

            {/* Metrics */}
            <div className="grid grid-cols-3 md:grid-cols-6 gap-3 mb-6">
              <MetricCard title="Online" value={results.summary.totalA} icon={<ShoppingCart className="w-4 h-4 text-blue-400" />} className="border-blue-400/20 bg-blue-400/5" />
              <MetricCard title="In-Store" value={results.summary.totalB} icon={<Store className="w-4 h-4 text-violet-400" />} className="border-violet-400/20 bg-violet-400/5" />
              <MetricCard title="Matched" value={results.summary.matchedCount} icon={<CheckCircle2 className="w-4 h-4 text-emerald-500" />} className="border-emerald-500/20 bg-emerald-500/5" />
              <MetricCard title="Differences" value={results.summary.diffCount} icon={<AlertTriangle className="w-4 h-4 text-accent" />} className="border-accent/20 bg-accent/5" />
              <MetricCard title="Only Online" value={results.summary.onlyACount} icon={<MinusCircle className="w-4 h-4 text-destructive" />} className="border-destructive/20 bg-destructive/5" />
              <MetricCard title="Only In-Store" value={results.summary.onlyBCount} icon={<PlusCircle className="w-4 h-4 text-cyan-400" />} className="border-cyan-400/20 bg-cyan-400/5" />
            </div>

            {/* Mapping summary */}
            <div className="flex flex-wrap gap-3 mb-5 items-center">
              <div className="flex flex-wrap gap-1.5 items-center">
                <span className="text-xs text-muted-foreground">Keys:</span>
                {results.keyMappings.map(m => (
                  <Badge key={m.id} variant="outline" className="font-mono text-xs border-primary/30">
                    <Key className="w-3 h-3 mr-1 text-primary" />
                    <span className="text-blue-400">{m.colA}</span>
                    <ArrowLeftRight className="w-3 h-3 mx-1 text-muted-foreground" />
                    <span className="text-violet-400">{m.colB}</span>
                  </Badge>
                ))}
              </div>
              <div className="flex flex-wrap gap-1.5 items-center">
                <span className="text-xs text-muted-foreground">Comparing:</span>
                {results.compareMappings.map(m => (
                  <Badge key={m.id} variant="secondary" className="font-mono text-xs">
                    <span className="text-blue-400">{m.colA}</span>
                    <span className="text-muted-foreground mx-1">vs</span>
                    <span className="text-violet-400">{m.colB}</span>
                  </Badge>
                ))}
              </div>
            </div>

            {/* Search bar */}
            <div className="relative mb-5">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-muted-foreground" />
              <Input
                className="pl-9 h-9 text-sm bg-card border-border/50"
                placeholder="Search results by key value or cell content..."
                value={resultSearch}
                onChange={e => setResultSearch(e.target.value)}
                data-testid="input-result-search"
              />
              {resultSearch && (
                <Button variant="ghost" size="icon" className="absolute right-1 top-1/2 -translate-y-1/2 h-7 w-7 text-muted-foreground hover:text-foreground" onClick={() => setResultSearch("")}>
                  <X className="w-3 h-3" />
                </Button>
              )}
            </div>

            <Tabs defaultValue="diffs" className="w-full">
              <TabsList className="grid w-full grid-cols-4 mb-5 h-11 bg-card/50">
                <TabsTrigger value="diffs" className="data-[state=active]:bg-primary data-[state=active]:text-primary-foreground font-medium text-sm" data-testid="tab-diffs">
                  Differences ({filteredDiffs.length}{resultSearch ? `/${results.summary.diffCount}` : ""})
                </TabsTrigger>
                <TabsTrigger value="matches" className="font-medium text-sm" data-testid="tab-matches">
                  Matched ({results.summary.matchedCount})
                </TabsTrigger>
                <TabsTrigger value="onlya" className="font-medium text-sm" data-testid="tab-onlya">
                  Only Online ({filteredOnlyA.length}{resultSearch ? `/${results.summary.onlyACount}` : ""})
                </TabsTrigger>
                <TabsTrigger value="onlyb" className="font-medium text-sm" data-testid="tab-onlyb">
                  Only In-Store ({filteredOnlyB.length}{resultSearch ? `/${results.summary.onlyBCount}` : ""})
                </TabsTrigger>
              </TabsList>

              {/* Differences Tab */}
              <TabsContent value="diffs" className="m-0">
                <Card className="border-border overflow-hidden">
                  <ScrollArea className="h-[58vh] w-full">
                    {filteredDiffs.length === 0 ? (
                      <EmptyState icon={<CheckCircle2 className="w-12 h-12 text-emerald-500 opacity-80" />}
                        title={resultSearch ? "No matching differences" : "No differences found"}
                        subtitle={resultSearch ? "Try a different search term." : "All matched rows are identical."} />
                    ) : (
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            {results.keyMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[110px] uppercase tracking-wide">
                                <span className="flex items-center gap-1"><Key className="w-3 h-3 text-primary shrink-0" /><span className="text-blue-400">{m.colA}</span></span>
                              </th>
                            ))}
                            {results.compareMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[220px] uppercase tracking-wide">
                                <span className="flex items-center gap-1 flex-wrap">
                                  <span className="text-blue-400">{m.colA}</span>
                                  <span className="text-muted-foreground opacity-50 text-[10px]">vs</span>
                                  <span className="text-violet-400">{m.colB}</span>
                                </span>
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {filteredDiffs.map((diff, idx) => {
                            const keyParts = diff.key.split("|||");
                            return (
                              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors" data-testid={`diff-row-${idx}`}>
                                {results.keyMappings.map((m, i) => (
                                  <td key={m.id} className="px-4 py-3 font-mono text-xs font-semibold border-r bg-background/40">{keyParts[i] ?? ""}</td>
                                ))}
                                {results.compareMappings.map(m => {
                                  const isChanged = diff.changedMappings.some(c => c.id === m.id);
                                  const valA = diff.rowA[m.colA] !== undefined ? String(diff.rowA[m.colA]) : "(empty)";
                                  const valB = diff.rowB[m.colB] !== undefined ? String(diff.rowB[m.colB]) : "(empty)";
                                  const nd = diff.numericDiffs[m.id];
                                  return (
                                    <td key={m.id} className={`px-4 py-3 border-r ${isChanged ? "bg-accent/5" : ""}`}>
                                      {isChanged ? (
                                        <div className="flex flex-col gap-1.5">
                                          <div className="flex items-center gap-1.5 text-destructive font-mono bg-destructive/10 px-2 py-1 rounded text-xs">
                                            <ShoppingCart className="w-3 h-3 shrink-0 opacity-50" />
                                            <span className="truncate" title={valA}>{valA}</span>
                                          </div>
                                          <div className="flex items-center gap-1.5 text-emerald-400 font-mono bg-emerald-500/10 px-2 py-1 rounded text-xs">
                                            <Store className="w-3 h-3 shrink-0 opacity-50" />
                                            <span className="truncate" title={valB}>{valB}</span>
                                          </div>
                                          {nd && (
                                            <div className={`flex items-center gap-1 text-[11px] font-mono px-2 ${nd.delta > 0 ? "text-emerald-400" : nd.delta < 0 ? "text-destructive" : "text-muted-foreground"}`}>
                                              {nd.delta > 0 ? <TrendingUp className="w-3 h-3 shrink-0" /> : nd.delta < 0 ? <TrendingDown className="w-3 h-3 shrink-0" /> : <Minus className="w-3 h-3 shrink-0" />}
                                              <span>{nd.delta > 0 ? "+" : ""}{nd.delta.toLocaleString(undefined, { maximumFractionDigits: 2 })}</span>
                                              {nd.pct !== null && <span className="opacity-70">({nd.pct > 0 ? "+" : ""}{nd.pct.toFixed(1)}%)</span>}
                                            </div>
                                          )}
                                        </div>
                                      ) : (
                                        <span className="text-muted-foreground font-mono text-xs" title={valA}>{valA}</span>
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
                {results.matched.length === 0 ? <EmptyCard message="No perfectly matched rows." /> : (
                  <Card className="border-border overflow-hidden">
                    <ScrollArea className="h-[58vh] w-full">
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            {results.keyMappings.map(m => (
                              <th key={m.id} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[110px]">
                                <Key className="w-3 h-3 text-primary inline mr-1" />{m.colA}
                              </th>
                            ))}
                            {results.compareMappings.flatMap(m => [
                              <th key={`a-${m.id}`} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px] text-blue-400">{m.colA}</th>,
                              <th key={`b-${m.id}`} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px] text-violet-400">{m.colB}</th>,
                            ])}
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
                                {results.compareMappings.flatMap(m => [
                                  <td key={`a-${m.id}`} className="px-4 py-3 font-mono text-xs text-muted-foreground border-r">{String(row.rowA[m.colA] ?? "")}</td>,
                                  <td key={`b-${m.id}`} className="px-4 py-3 font-mono text-xs text-muted-foreground border-r">{String(row.rowB[m.colB] ?? "")}</td>,
                                ])}
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </ScrollArea>
                  </Card>
                )}
              </TabsContent>

              <TabsContent value="onlya" className="m-0">
                <SimpleDataTable rows={filteredOnlyA} emptyMessage={resultSearch ? "No matching rows." : "No rows exclusive to Online."} />
              </TabsContent>
              <TabsContent value="onlyb" className="m-0">
                <SimpleDataTable rows={filteredOnlyB} emptyMessage={resultSearch ? "No matching rows." : "No rows exclusive to In-Store."} />
              </TabsContent>
            </Tabs>
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
                  {file.columns.slice(0, 4).map(c => <th key={c} className="px-3 py-2 border-b font-medium max-w-[130px] truncate">{c}</th>)}
                  {file.columns.length > 4 && <th className="px-3 py-2 border-b text-muted-foreground/60">+{file.columns.length - 4} more</th>}
                </tr>
              </thead>
              <tbody className="font-mono">
                {file.data.slice(0, 3).map((row, i) => (
                  <tr key={i} className="border-b last:border-0 hover:bg-muted/20">
                    {file.columns.slice(0, 4).map(c => <td key={c} className="px-3 py-2 max-w-[130px] truncate opacity-75">{String(row[c] ?? "")}</td>)}
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
    <Card className="border-dashed border-2 hover:border-primary/50 hover:bg-card/80 transition-colors cursor-pointer group bg-card/30"
      onDragOver={e => e.preventDefault()} onDrop={e => onDrop(e, side)} onClick={() => document.getElementById(inputId)?.click()}>
      <CardContent className="flex flex-col items-center justify-center p-12 text-center h-full min-h-[280px]">
        <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mb-5 group-hover:scale-110 transition-transform duration-300 group-hover:bg-primary/10 text-muted-foreground group-hover:text-primary">{icon}</div>
        <h3 className="text-xl font-medium mb-1">{label}</h3>
        <p className="text-muted-foreground text-sm mb-5 max-w-[220px]">Drag & drop your .xlsx or .csv file here, or click to browse.</p>
        <Button variant="secondary" className="pointer-events-none group-hover:bg-primary group-hover:text-primary-foreground transition-colors">Select File</Button>
        <input type="file" id={inputId} className="hidden" accept=".xlsx,.xls,.csv" onChange={e => { if (e.target.files?.[0]) onUpload(e.target.files[0], side); }} />
      </CardContent>
    </Card>
  );
}

function MetricCard({ title, value, icon, className }: { title: string; value: number; icon: React.ReactNode; className?: string }) {
  return (
    <Card className={`overflow-hidden transition-all hover:shadow-md ${className}`}>
      <CardContent className="p-4">
        <div className="flex items-center justify-between mb-2">
          <p className="text-xs font-medium text-muted-foreground">{title}</p>
          <div className="bg-background/80 p-1 rounded-full">{icon}</div>
        </div>
        <div className="text-2xl font-bold font-mono">{value.toLocaleString()}</div>
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
      <ScrollArea className="h-[58vh] w-full">
        <table className="w-full text-sm text-left whitespace-nowrap">
          <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
            <tr>{columns.map(col => <th key={col} className="px-4 py-3 font-medium border-b border-r bg-card/90 min-w-[120px]">{col}</th>)}</tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-border hover:bg-muted/20 transition-colors">
                {columns.map(col => <td key={col} className="px-4 py-3 border-r font-mono text-xs text-muted-foreground">{String(row[col] ?? "")}</td>)}
              </tr>
            ))}
          </tbody>
        </table>
      </ScrollArea>
    </Card>
  );
}
