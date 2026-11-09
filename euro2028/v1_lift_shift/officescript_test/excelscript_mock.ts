declare const console: { log(...data: (string | number)[]): void };
// Minimal stand-in for the Office Scripts ExcelScript API, used only to type-check the generated
// script under strict TypeScript and to run it against a model of the blank template. Method names and
// signatures follow the published ExcelScript API (only the members the script uses).
namespace ExcelScript {
  export enum BorderIndex { edgeTop = "EdgeTop", edgeBottom = "EdgeBottom", edgeLeft = "EdgeLeft", edgeRight = "EdgeRight", insideVertical = "InsideVertical", insideHorizontal = "InsideHorizontal" }
  export enum BorderLineStyle { continuous = "Continuous" }
  export enum HorizontalAlignment { general = "General", left = "Left", center = "Center", right = "Right", centerAcrossSelection = "CenterAcrossSelection" }
  export enum VerticalAlignment { center = "Center" }
  export enum InsertShiftDirection { down = "Down", right = "Right" }
  export enum RangeCopyType { all = "All", formulas = "Formulas", values = "Values", formats = "Formats" }
  export enum ClearApplyTo { all = "All", formats = "Formats", contents = "Contents" }
  export enum SheetVisibility { visible = "Visible", hidden = "Hidden", veryHidden = "VeryHidden" }

  export interface Cell { v: string | number | boolean | null; f: string | null; fill: string | null; }

  function colNum(s: string): number { let n = 0; for (const ch of s) { n = n * 26 + ch.charCodeAt(0) - 64; } return n; }
  function parse(addr: string): [number, number, number, number] {
    const parts = addr.replace(/\$/g, "").split(":");
    const one = (p: string, end: boolean): [number, number] => {
      const m = /^([A-Z]*)(\d*)$/.exec(p);
      if (!m) { throw new Error("bad address " + addr); }
      const c = m[1] ? colNum(m[1]) : (end ? 16384 : 1);
      const r = m[2] ? Number(m[2]) : (end ? 1048576 : 1);
      return [r, c];
    };
    const a = one(parts[0], false), b = parts.length > 1 ? one(parts[1], true) : a;
    return [a[0], a[1], b[0], b[1]];
  }

  export class RangeFill {
    constructor(private rg: Range) { }
    setColor(color: string): void { this.rg.eachCell(c => { c.fill = "FF" + color.replace("#", "").toUpperCase(); }); }
    clear(): void { this.rg.eachCell(c => { c.fill = null; }); }
  }
  export class RangeFont {
    setBold(bold: boolean): void { } setColor(color: string): void { } setItalic(italic: boolean): void { }
    setName(name: string): void { } setSize(size: number): void { }
  }
  export class RangeBorder { setStyle(style: BorderLineStyle): void { } setColor(color: string): void { } }
  export class RangeFormat {
    constructor(private rg: Range) { }
    getFill(): RangeFill { return new RangeFill(this.rg); }
    getFont(): RangeFont { return new RangeFont(); }
    getRangeBorder(index: BorderIndex): RangeBorder { return new RangeBorder(); }
    setColumnWidth(columnWidth: number): void { }
    setHorizontalAlignment(horizontalAlignment: HorizontalAlignment): void { }
    setVerticalAlignment(verticalAlignment: VerticalAlignment): void { }
    setWrapText(wrapText: boolean): void { }
  }

  export class Range {
    r1: number; c1: number; r2: number; c2: number;
    constructor(public ws: Worksheet, addr: string) { [this.r1, this.c1, this.r2, this.c2] = parse(addr); }
    eachCell(fn: (c: Cell, r: number, col: number) => void): void {
      if (this.r2 - this.r1 > 2000 || this.c2 - this.c1 > 2000) {        // whole rows/columns: touch existing cells only
        for (const [k, c] of this.ws.cells) {
          const [r, col] = k.split(",").map(Number);
          if (r >= this.r1 && r <= this.r2 && col >= this.c1 && col <= this.c2) { fn(c, r, col); }
        }
        return;
      }
      for (let r = this.r1; r <= this.r2; r++) { for (let c = this.c1; c <= this.c2; c++) { fn(this.ws.cell(r, c), r, c); } }
    }
    private put(r: number, c: number, x: string | number | boolean, asFormula: boolean): void {
      const cell = this.ws.cell(this.r1 + r, this.c1 + c);
      if (typeof x === "string" && x.startsWith("=")) { cell.f = x; cell.v = null; return; }
      cell.f = null;
      if (typeof x === "string" && x.startsWith("'")) { cell.v = x.substring(1); return; }
      if (x === "") { cell.v = null; return; }
      if (asFormula && typeof x === "string" && x !== "" && !isNaN(Number(x))) { cell.v = Number(x); return; }
      cell.v = x;
    }
    private check(rows: number, cols: number): void {
      if (rows !== this.r2 - this.r1 + 1 || cols !== this.c2 - this.c1 + 1) { throw new Error("size mismatch " + rows + "x" + cols); }
    }
    setValue(value: string | number | boolean): void { this.eachCell((c, r, col) => { this.put(r - this.r1, col - this.c1, value, false); }); }
    setValues(values: (string | number | boolean)[][]): void {
      this.check(values.length, values[0].length);
      values.forEach((row, i) => row.forEach((v, j) => this.put(i, j, v, false)));
    }
    setFormula(formula: string): void { this.eachCell((c, r, col) => { this.put(r - this.r1, col - this.c1, formula, true); }); }
    setFormulas(formulas: string[][]): void {
      this.check(formulas.length, formulas[0].length);
      formulas.forEach((row, i) => row.forEach((v, j) => this.put(i, j, v, true)));
    }
    setNumberFormat(numberFormat: string): void { }
    getFormat(): RangeFormat { return new RangeFormat(this); }
    merge(across?: boolean): void { this.ws.merges.add(this.key()); }
    unmerge(): void {
      for (const m of Array.from(this.ws.merges)) {
        const [a, b, c, d] = parse(m);
        if (a <= this.r2 && c >= this.r1 && b <= this.c2 && d >= this.c1) { this.ws.merges.delete(m); }
      }
    }
    key(): string { return colLetters(this.c1) + this.r1 + ":" + colLetters(this.c2) + this.r2; }
    clear(applyTo?: ClearApplyTo): void {
      this.eachCell(c => { if (applyTo !== ClearApplyTo.formats) { c.v = null; c.f = null; } if (applyTo !== ClearApplyTo.contents) { c.fill = null; } });
    }
    insert(shift: InsertShiftDirection): Range {
      if (shift !== InsertShiftDirection.right || this.r1 !== 1 || this.r2 !== 1048576) { throw new Error("mock supports whole-column insert only"); }
      const n = this.c2 - this.c1 + 1;
      this.ws.shiftColumns(this.c1, n);
      return this;
    }
    moveTo(destinationRange: Range | string): void {
      const d = typeof destinationRange === "string" ? new Range(this.ws, destinationRange) : destinationRange;
      const snap: [number, number, Cell][] = [];
      for (let r = this.r1; r <= this.r2; r++) {
        for (let c = this.c1; c <= this.c2; c++) {
          const k = r + "," + c;
          const cell = this.ws.cells.get(k);
          if (cell) { snap.push([r - this.r1, c - this.c1, cell]); this.ws.cells.delete(k); }
        }
      }
      for (let r = 0; r <= this.r2 - this.r1; r++) { for (let c = 0; c <= this.c2 - this.c1; c++) { this.ws.cells.delete((d.r1 + r) + "," + (d.c1 + c)); } }
      for (const [r, c, cell] of snap) { this.ws.cells.set((d.r1 + r) + "," + (d.c1 + c), cell); }
    }
    copyFrom(sourceRange: Range | string, copyType?: RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void {
      const s = typeof sourceRange === "string" ? new Range(this.ws, sourceRange) : sourceRange;
      const h = s.r2 - s.r1 + 1, w = s.c2 - s.c1 + 1;
      if (this.r2 - this.r1 + 1 !== h || this.c2 - this.c1 + 1 !== w) { throw new Error("mock copyFrom needs equal sizes: " + this.key() + " <- " + s.key()); }
      const fills: (string | null)[][] = [];
      for (let r = 0; r < h; r++) { fills.push([]); for (let c = 0; c < w; c++) { const x = this.ws.cells.get((s.r1 + r) + "," + (s.c1 + c)); fills[r].push(x ? x.fill : null); } }
      for (let r = 0; r < h; r++) { for (let c = 0; c < w; c++) { this.ws.cell(this.r1 + r, this.c1 + c).fill = fills[r][c]; } }
    }
  }

  export function colLetters(n: number): string { let s = ""; while (n > 0) { s = String.fromCharCode(65 + ((n - 1) % 26)) + s; n = Math.floor((n - 1) / 26); } return s; }

  export class WorksheetFreezePanes { freezeAt(frozenRange: Range | string): void { } unfreeze(): void { } }
  export class AutoFilter { apply(range: Range | string, columnIndex?: number): void { } remove(): void { } }

  export class Worksheet {
    cells = new Map<string, Cell>();
    merges = new Set<string>();
    visibility: SheetVisibility = SheetVisibility.visible;
    constructor(private name: string) { }
    cell(r: number, c: number): Cell {
      const k = r + "," + c;
      let x = this.cells.get(k);
      if (!x) { x = { v: null, f: null, fill: null }; this.cells.set(k, x); }
      return x;
    }
    shiftColumns(first: number, n: number): void {
      const moved = new Map<string, Cell>();
      for (const [k, c] of this.cells) {
        const [r, col] = k.split(",").map(Number);
        moved.set(col >= first ? r + "," + (col + n) : k, c);
      }
      // inserted cells take the format of the column on the left (Excel's default)
      for (const [k, c] of this.cells) {
        const [r, col] = k.split(",").map(Number);
        if (col === first - 1 && c.fill) { for (let i = 0; i < n; i++) { moved.set(r + "," + (first + i), { v: null, f: null, fill: c.fill }); } }
      }
      this.cells = moved;
      const ms = new Set<string>();
      for (const m of this.merges) {
        const [a, b, c, d] = parse(m);
        const nb = b >= first ? b + n : b, nd = d >= first ? d + n : d;
        ms.add(colLetters(nb) + a + ":" + colLetters(nd) + c);
      }
      this.merges = ms;
    }
    getRange(address?: string): Range { return new Range(this, address || "A1"); }
    getName(): string { return this.name; }
    setShowGridlines(showGridlines: boolean): void { }
    setVisibility(visibility: SheetVisibility): void { this.visibility = visibility; }
    getFreezePanes(): WorksheetFreezePanes { return new WorksheetFreezePanes(); }
    getAutoFilter(): AutoFilter { return new AutoFilter(); }
    activate(): void { }
  }

  export class Workbook {
    sheets: Worksheet[] = [];
    getWorksheet(name: string): Worksheet | undefined { return this.sheets.find(s => s.getName() === name); }
    addWorksheet(name?: string): Worksheet { const w = new Worksheet(name || "Sheet"); this.sheets.push(w); return w; }
  }
}
