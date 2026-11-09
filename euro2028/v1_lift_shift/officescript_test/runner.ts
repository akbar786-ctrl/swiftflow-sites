// Runs the generated Office Script against the template model and writes the resulting cells as JSON.
declare const require: (m: string) => { readFileSync(p: string, e: string): string; writeFileSync(p: string, d: string): void };
declare const process: { argv: string[] };

(function run() {
  const fs = require("fs");
  const model: { [sheet: string]: [number, number, string | number | boolean | null, string | null, string | null][] } =
    JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
  const wb = new ExcelScript.Workbook();
  for (const name of Object.keys(model)) {
    if (name.startsWith("__meta__")) { continue; }
    const ws = wb.addWorksheet(name);
    for (const [, kind, v] of model["__meta__" + name]) {
      if (kind === 0 && v === "hidden") { ws.visibility = ExcelScript.SheetVisibility.hidden; }
      if (kind === 1 && typeof v === "string") { ws.merges.add(v); }
    }
    for (const [r, c, v, f, fill] of model[name]) {
      const cell = ws.cell(r, c);
      cell.v = v; cell.f = f; cell.fill = fill;
    }
  }
  main(wb);
  const out: { [sheet: string]: { visibility: string; merges: string[]; cells: [number, number, string | number | boolean | null, string | null, string | null][] } } = {};
  for (const ws of wb.sheets) {
    const cells: [number, number, string | number | boolean | null, string | null, string | null][] = [];
    for (const [k, c] of ws.cells) {
      const [r, col] = k.split(",").map(Number);
      if (c.v !== null || c.f !== null || c.fill !== null) { cells.push([r, col, c.v, c.f, c.fill]); }
    }
    out[ws.getName()] = { visibility: ws.visibility, merges: Array.from(ws.merges), cells };
  }
  fs.writeFileSync(process.argv[3], JSON.stringify(out));
})();
