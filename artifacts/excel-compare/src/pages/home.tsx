import React, { useState, useCallback, useMemo } from "react";
import { UploadCloud, FileSpreadsheet, ArrowRight, Play, RefreshCcw, Download, CheckCircle2, AlertTriangle, FileWarning, PlusCircle, MinusCircle } from "lucide-react";
import { parseExcelFile, compareSheets, exportResultsToExcel, ComparisonResult } from "@/lib/excel-utils";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Checkbox } from "@/components/ui/checkbox";
import { Label } from "@/components/ui/label";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";

type Step = "upload" | "config" | "loading" | "results";

export default function Home() {
  const { toast } = useToast();
  const [step, setStep] = useState<Step>("upload");

  // Files & Data
  const [fileA, setFileA] = useState<{ name: string; data: any[]; columns: string[] } | null>(null);
  const [fileB, setFileB] = useState<{ name: string; data: any[]; columns: string[] } | null>(null);

  // Config
  const [keyCol, setKeyCol] = useState<string>("");
  const [selectedCols, setSelectedCols] = useState<Set<string>>(new Set());

  // Results
  const [results, setResults] = useState<ComparisonResult | null>(null);

  const sharedColumns = useMemo(() => {
    if (!fileA || !fileB) return [];
    return fileA.columns.filter(c => fileB.columns.includes(c));
  }, [fileA, fileB]);

  const handleFileUpload = async (file: File, side: "A" | "B") => {
    try {
      const parsed = await parseExcelFile(file);
      if (side === "A") {
        setFileA({ name: file.name, data: parsed.data, columns: parsed.columns });
      } else {
        setFileB({ name: file.name, data: parsed.data, columns: parsed.columns });
      }
    } catch (err) {
      toast({
        title: "Error parsing file",
        description: "Please ensure you uploaded a valid Excel or CSV file.",
        variant: "destructive"
      });
    }
  };

  const handleDragOver = (e: React.DragEvent) => e.preventDefault();
  const handleDrop = useCallback((e: React.DragEvent, side: "A" | "B") => {
    e.preventDefault();
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleFileUpload(e.dataTransfer.files[0], side);
    }
  }, []);

  const proceedToConfig = () => {
    if (!fileA || !fileB) return;
    
    // Auto-select first shared column as key if possible
    if (sharedColumns.length > 0) {
      setKeyCol(sharedColumns[0]);
      // Select all shared columns by default
      setSelectedCols(new Set(sharedColumns));
    }
    
    setStep("config");
  };

  const runComparison = () => {
    if (!fileA || !fileB || !keyCol) return;
    setStep("loading");
    
    // Use timeout to allow UI to render loading state
    setTimeout(() => {
      try {
        const res = compareSheets(
          fileA.data,
          fileB.data,
          keyCol,
          Array.from(selectedCols)
        );
        setResults(res);
        setStep("results");
      } catch (err) {
        toast({
          title: "Comparison Failed",
          description: "An error occurred during comparison.",
          variant: "destructive"
        });
        setStep("config");
      }
    }, 100);
  };

  const reset = () => {
    setFileA(null);
    setFileB(null);
    setKeyCol("");
    setSelectedCols(new Set());
    setResults(null);
    setStep("upload");
  };

  return (
    <div className="min-h-screen bg-background text-foreground flex flex-col font-sans">
      <header className="border-b border-border bg-card px-6 py-4 flex items-center justify-between sticky top-0 z-10 shadow-sm">
        <div className="flex items-center gap-2">
          <FileSpreadsheet className="w-6 h-6 text-primary" />
          <h1 className="text-xl font-bold tracking-tight">ReconcileX</h1>
        </div>
        {step === "results" && (
          <div className="flex items-center gap-3">
            <Button variant="outline" onClick={reset} size="sm" data-testid="button-reset">
              <RefreshCcw className="w-4 h-4 mr-2" />
              New Comparison
            </Button>
            <Button 
              onClick={() => results && exportResultsToExcel(results)} 
              size="sm"
              className="bg-primary hover:bg-primary/90 text-primary-foreground"
              data-testid="button-export"
            >
              <Download className="w-4 h-4 mr-2" />
              Export Results
            </Button>
          </div>
        )}
      </header>

      <main className="flex-1 p-6 max-w-7xl w-full mx-auto">
        {step === "upload" && (
          <div className="animate-in fade-in zoom-in-95 duration-300">
            <div className="mb-8">
              <h2 className="text-3xl font-semibold mb-2">Select files to compare</h2>
              <p className="text-muted-foreground text-lg">Upload two spreadsheets to find differences, matches, and missing rows.</p>
            </div>
            
            <div className="grid md:grid-cols-2 gap-6">
              <UploadZone 
                side="A" 
                file={fileA} 
                onDrop={handleDrop} 
                onUpload={handleFileUpload} 
                onClear={() => setFileA(null)}
              />
              <UploadZone 
                side="B" 
                file={fileB} 
                onDrop={handleDrop} 
                onUpload={handleFileUpload} 
                onClear={() => setFileB(null)}
              />
            </div>

            <div className="mt-8 flex justify-end">
              <Button 
                size="lg" 
                onClick={proceedToConfig} 
                disabled={!fileA || !fileB}
                className="w-full md:w-auto font-semibold px-8"
                data-testid="button-proceed-config"
              >
                Configure Comparison
                <ArrowRight className="ml-2 w-5 h-5" />
              </Button>
            </div>
          </div>
        )}

        {step === "config" && (
          <div className="animate-in fade-in slide-in-from-bottom-4 duration-300">
            <div className="flex items-center gap-4 mb-8">
              <Button variant="ghost" size="icon" onClick={() => setStep("upload")} className="shrink-0 rounded-full h-8 w-8">
                <ArrowRight className="w-4 h-4 rotate-180" />
              </Button>
              <div>
                <h2 className="text-3xl font-semibold">Configure Comparison</h2>
                <p className="text-muted-foreground text-lg">Define how the datasets should be joined and compared.</p>
              </div>
            </div>

            <Card className="border-border/50 shadow-lg mb-6 bg-card/50 backdrop-blur">
              <CardHeader>
                <CardTitle className="text-xl">1. Primary Key</CardTitle>
                <CardDescription className="text-base">
                  Select the unique identifier column present in both sheets to align rows.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="max-w-md">
                  <Select value={keyCol} onValueChange={setKeyCol}>
                    <SelectTrigger data-testid="select-keycol">
                      <SelectValue placeholder="Select a column..." />
                    </SelectTrigger>
                    <SelectContent>
                      {sharedColumns.map(col => (
                        <SelectItem key={col} value={col}>{col}</SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </CardContent>
            </Card>

            <Card className="border-border/50 shadow-lg mb-8 bg-card/50 backdrop-blur">
              <CardHeader>
                <CardTitle className="text-xl">2. Columns to Compare</CardTitle>
                <CardDescription className="text-base">
                  Select which columns to evaluate for differences across matched rows.
                </CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
                  {sharedColumns.map(col => (
                    <div key={col} className="flex items-center space-x-2 bg-background p-3 rounded-md border border-border/50">
                      <Checkbox 
                        id={`col-${col}`} 
                        checked={selectedCols.has(col)}
                        onCheckedChange={(checked) => {
                          const newSet = new Set(selectedCols);
                          checked ? newSet.add(col) : newSet.delete(col);
                          setSelectedCols(newSet);
                        }}
                      />
                      <label 
                        htmlFor={`col-${col}`}
                        className="text-sm font-medium leading-none peer-disabled:cursor-not-allowed peer-disabled:opacity-70 truncate cursor-pointer select-none"
                        title={col}
                      >
                        {col}
                      </label>
                    </div>
                  ))}
                </div>
              </CardContent>
            </Card>

            <div className="flex justify-end">
              <Button 
                size="lg" 
                onClick={runComparison} 
                disabled={!keyCol || selectedCols.size === 0}
                className="w-full md:w-auto font-bold px-10 h-14 text-lg bg-primary hover:bg-primary/90 text-primary-foreground"
                data-testid="button-run-comparison"
              >
                <Play className="mr-2 w-5 h-5 fill-current" />
                Run Analysis
              </Button>
            </div>
          </div>
        )}

        {step === "loading" && (
          <div className="h-[60vh] flex flex-col items-center justify-center animate-in fade-in duration-500">
            <div className="w-16 h-16 border-4 border-primary border-t-transparent rounded-full animate-spin mb-6"></div>
            <h2 className="text-2xl font-semibold mb-2">Processing Datasets</h2>
            <p className="text-muted-foreground">Comparing rows and identifying differences...</p>
          </div>
        )}

        {step === "results" && results && (
          <div className="animate-in fade-in zoom-in-95 duration-500">
            
            {/* Summary Metrics */}
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-8">
              <MetricCard 
                title="Matched Identical" 
                value={results.summary.matchedCount} 
                icon={<CheckCircle2 className="w-5 h-5 text-emerald-500" />} 
                className="border-emerald-500/20 bg-emerald-500/5"
              />
              <MetricCard 
                title="Differences Found" 
                value={results.summary.diffCount} 
                icon={<AlertTriangle className="w-5 h-5 text-accent" />} 
                className="border-accent/20 bg-accent/5"
              />
              <MetricCard 
                title={`Only in ${fileA?.name}`} 
                value={results.summary.onlyACount} 
                icon={<MinusCircle className="w-5 h-5 text-destructive" />} 
                className="border-destructive/20 bg-destructive/5"
              />
              <MetricCard 
                title={`Only in ${fileB?.name}`} 
                value={results.summary.onlyBCount} 
                icon={<PlusCircle className="w-5 h-5 text-blue-400" />} 
                className="border-blue-400/20 bg-blue-400/5"
              />
            </div>

            <Tabs defaultValue="diffs" className="w-full">
              <TabsList className="grid w-full grid-cols-4 mb-6 h-12 bg-card/50">
                <TabsTrigger value="diffs" className="data-[state=active]:bg-primary data-[state=active]:text-primary-foreground font-medium" data-testid="tab-diffs">
                  Differences ({results.summary.diffCount})
                </TabsTrigger>
                <TabsTrigger value="matches" className="data-[state=active]:bg-card data-[state=active]:text-foreground font-medium" data-testid="tab-matches">
                  Matched ({results.summary.matchedCount})
                </TabsTrigger>
                <TabsTrigger value="onlya" className="data-[state=active]:bg-card data-[state=active]:text-foreground font-medium" data-testid="tab-onlya">
                  Missing in B ({results.summary.onlyACount})
                </TabsTrigger>
                <TabsTrigger value="onlyb" className="data-[state=active]:bg-card data-[state=active]:text-foreground font-medium" data-testid="tab-onlyb">
                  New in B ({results.summary.onlyBCount})
                </TabsTrigger>
              </TabsList>

              <TabsContent value="diffs" className="m-0">
                <Card className="border-border overflow-hidden">
                  <ScrollArea className="h-[60vh] w-full relative">
                    {results.differences.length === 0 ? (
                      <div className="p-12 text-center text-muted-foreground flex flex-col items-center">
                        <CheckCircle2 className="w-12 h-12 mb-4 text-emerald-500 opacity-80" />
                        <p className="text-lg font-medium">No differences found!</p>
                        <p>All matched rows have identical values in the compared columns.</p>
                      </div>
                    ) : (
                      <table className="w-full text-sm text-left whitespace-nowrap">
                        <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
                          <tr>
                            <th className="px-6 py-4 font-medium border-b border-r bg-card/90">{results.keyCol}</th>
                            {results.columns.map(col => (
                              <th key={col} className="px-6 py-4 font-medium border-b border-r bg-card/90 min-w-[200px] max-w-[300px] truncate" title={col}>
                                {col}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {results.differences.map((diff, idx) => (
                            <tr key={diff.key} className="border-b border-border hover:bg-muted/30 transition-colors">
                              <td className="px-6 py-4 font-mono font-medium border-r bg-background/50">
                                {diff.key}
                              </td>
                              {results.columns.map(col => {
                                const isChanged = diff.changedCols.includes(col);
                                const valA = diff.rowA[col] !== undefined ? String(diff.rowA[col]) : "(empty)";
                                const valB = diff.rowB[col] !== undefined ? String(diff.rowB[col]) : "(empty)";

                                return (
                                  <td key={col} className={`px-6 py-4 border-r ${isChanged ? 'bg-accent/10 relative' : ''}`}>
                                    {isChanged ? (
                                      <div className="flex flex-col gap-2">
                                        <div className="flex items-center gap-2 text-destructive font-mono bg-destructive/10 px-2 py-1 rounded">
                                          <span className="text-xs font-bold uppercase w-4 opacity-50 shrink-0">A</span> 
                                          <span className="truncate" title={valA}>{valA}</span>
                                        </div>
                                        <div className="flex items-center gap-2 text-emerald-500 font-mono bg-emerald-500/10 px-2 py-1 rounded">
                                          <span className="text-xs font-bold uppercase w-4 opacity-50 shrink-0">B</span> 
                                          <span className="truncate" title={valB}>{valB}</span>
                                        </div>
                                      </div>
                                    ) : (
                                      <span className="text-muted-foreground font-mono truncate block" title={valA}>{valA}</span>
                                    )}
                                  </td>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </ScrollArea>
                </Card>
              </TabsContent>

              <TabsContent value="matches" className="m-0">
                <SimpleDataTable rows={results.matched} columns={[results.keyCol, ...results.columns]} emptyMessage="No perfect matches found." />
              </TabsContent>
              
              <TabsContent value="onlya" className="m-0">
                <SimpleDataTable rows={results.onlyA} columns={[results.keyCol, ...results.columns]} emptyMessage="No missing rows found." />
              </TabsContent>
              
              <TabsContent value="onlyb" className="m-0">
                <SimpleDataTable rows={results.onlyB} columns={[results.keyCol, ...results.columns]} emptyMessage="No new rows found." />
              </TabsContent>
            </Tabs>
          </div>
        )}
      </main>
    </div>
  );
}

// Subcomponents

function UploadZone({ 
  side, 
  file, 
  onDrop, 
  onUpload, 
  onClear 
}: { 
  side: "A" | "B"; 
  file: { name: string; data: any[]; columns: string[] } | null; 
  onDrop: (e: React.DragEvent, side: "A" | "B") => void; 
  onUpload: (file: File, side: "A" | "B") => void; 
  onClear: () => void;
}) {
  const inputId = `file-upload-${side}`;
  
  if (file) {
    return (
      <Card className="border-primary/50 bg-primary/5 relative overflow-hidden shadow-lg">
        <CardContent className="p-6">
          <div className="flex justify-between items-start mb-6">
            <div>
              <Badge variant="outline" className="mb-2 bg-background">Sheet {side}</Badge>
              <h3 className="font-semibold text-lg truncate pr-8" title={file.name}>{file.name}</h3>
              <p className="text-sm text-muted-foreground font-mono">{file.data.length.toLocaleString()} rows • {file.columns.length} columns</p>
            </div>
            <Button variant="ghost" size="sm" onClick={onClear} className="text-muted-foreground hover:text-destructive">Clear</Button>
          </div>
          
          <div className="text-sm">
            <p className="text-xs font-semibold text-muted-foreground uppercase tracking-wider mb-2">Preview (Top 3 rows)</p>
            <div className="bg-background rounded border border-border overflow-hidden">
              <table className="w-full text-left text-xs whitespace-nowrap">
                <thead className="bg-muted/50 text-muted-foreground">
                  <tr>
                    {file.columns.slice(0, 4).map(c => (
                      <th key={c} className="px-3 py-2 border-b font-medium max-w-[150px] truncate">{c}</th>
                    ))}
                    {file.columns.length > 4 && <th className="px-3 py-2 border-b font-medium">...</th>}
                  </tr>
                </thead>
                <tbody className="font-mono">
                  {file.data.slice(0, 3).map((row, i) => (
                    <tr key={i} className="border-b last:border-0 hover:bg-muted/20">
                      {file.columns.slice(0, 4).map(c => (
                        <td key={c} className="px-3 py-2 max-w-[150px] truncate opacity-80" title={String(row[c] || "")}>
                          {String(row[c] || "")}
                        </td>
                      ))}
                      {file.columns.length > 4 && <td className="px-3 py-2 text-muted-foreground">...</td>}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card 
      className="border-dashed border-2 hover:border-primary/50 hover:bg-card/80 transition-colors cursor-pointer group bg-card/30"
      onDragOver={(e) => e.preventDefault()}
      onDrop={(e) => onDrop(e, side)}
      onClick={() => document.getElementById(inputId)?.click()}
    >
      <CardContent className="flex flex-col items-center justify-center p-12 text-center h-full min-h-[300px]">
        <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mb-6 group-hover:scale-110 transition-transform duration-300 group-hover:bg-primary/10">
          <UploadCloud className="w-8 h-8 text-muted-foreground group-hover:text-primary transition-colors" />
        </div>
        <h3 className="text-xl font-medium mb-2">Upload Sheet {side}</h3>
        <p className="text-muted-foreground mb-6 max-w-[250px]">
          Drag and drop your Excel or CSV file here, or click to browse.
        </p>
        <Button variant="secondary" className="pointer-events-none group-hover:bg-primary group-hover:text-primary-foreground transition-colors">
          Select File
        </Button>
        <input 
          type="file" 
          id={inputId} 
          className="hidden" 
          accept=".xlsx,.xls,.csv"
          onChange={(e) => {
            if (e.target.files && e.target.files.length > 0) {
              onUpload(e.target.files[0], side);
            }
          }}
        />
      </CardContent>
    </Card>
  );
}

function MetricCard({ title, value, icon, className }: { title: string, value: number, icon: React.ReactNode, className?: string }) {
  return (
    <Card className={`overflow-hidden transition-all hover:shadow-md ${className}`}>
      <CardContent className="p-6 flex flex-col justify-center">
        <div className="flex items-center justify-between mb-4">
          <p className="text-sm font-medium text-muted-foreground">{title}</p>
          <div className="bg-background/80 p-2 rounded-full backdrop-blur-sm shadow-sm">{icon}</div>
        </div>
        <div className="text-3xl font-bold font-mono tracking-tight">{value.toLocaleString()}</div>
      </CardContent>
    </Card>
  );
}

function SimpleDataTable({ rows, columns, emptyMessage }: { rows: any[], columns: string[], emptyMessage: string }) {
  if (rows.length === 0) {
    return (
      <Card className="p-12 text-center border-dashed">
        <FileWarning className="w-12 h-12 text-muted-foreground mx-auto mb-4 opacity-50" />
        <p className="text-lg text-muted-foreground font-medium">{emptyMessage}</p>
      </Card>
    );
  }

  return (
    <Card className="border-border overflow-hidden">
      <ScrollArea className="h-[60vh] w-full relative">
        <table className="w-full text-sm text-left whitespace-nowrap">
          <thead className="text-xs uppercase bg-muted/50 text-muted-foreground sticky top-0 z-10 backdrop-blur-md">
            <tr>
              {columns.map(col => (
                <th key={col} className="px-6 py-4 font-medium border-b border-r bg-card/90 min-w-[150px] max-w-[300px] truncate" title={col}>
                  {col}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, idx) => (
              <tr key={idx} className="border-b border-border hover:bg-muted/30 transition-colors">
                {columns.map((col, cIdx) => {
                  const val = row[col] !== undefined ? String(row[col]) : "";
                  return (
                    <td key={col} className={`px-6 py-4 border-r font-mono opacity-80 ${cIdx === 0 ? 'font-medium opacity-100 bg-background/50' : ''}`} title={val}>
                      <span className="truncate block max-w-[300px]">{val}</span>
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </ScrollArea>
    </Card>
  );
}
