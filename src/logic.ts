// Motor puro del FDC Selector (sin dependencias de Office.js: testeable en Node).
// Reglas extraidas del flujo real Maritime Octubre 2026.

export type SheetKind = "REC" | "SHORE" | "RIGREC" | "RIG";

export interface SourceRow {
  sheet: string;
  row: number; // 1-based
  rdm: string;
  desc: string;
  kind: SheetKind;
  section: string;
}

export interface FdcRow {
  row: number; // fila en la principal
  sheet: string; // hoja fuente referenciada
  src: number; // fila fuente referenciada
  week: number | null; // 41|42|43|44
}

export const WEEK_OF: Record<string, number> = {
  K: 41, L: 41, M: 42, N: 42, O: 43, P: 43, Q: 44, R: 44,
};

export const WEEK_USD_COL: Record<number, string> = { 41: "L", 42: "N", 43: "P" };

// Encabezados/subtotales: nunca son filas dato.
const HEADER_LABELS = [
  "MANTENIMIENTO DE FLOTA", "MEDICAMENTOS", "COMPRAS DE EPP",
  "COMPRAS DE MISCELANEOS", "COMPRAS DE CONSUMIBLES", "COMPRA DE CONSUMIBLES",
  "COMPRAS DE INSUMOS", "OPERACIONES RIG", "COMPRA DE MEDICAMENTOS",
  "REQUERIMIENTO HSE", "REQUERIMIENT0", "REQUERIMIENTO RRHH",
  "SERVICIOS RECURRENTE", "SEVICIOS RECURRENTE", "COMPRAS DE LUBRICANTES",
  "COMPRAS SSGG", "KIT DE MANTENIMIENTO", "ILUMINACION",
];

export function isHeaderLabel(text: string): boolean {
  const t = (text || "").trim().toUpperCase();
  if (!t) return true;
  return HEADER_LABELS.some((h) => t.includes(h));
}

export function kindOfSheet(name: string): SheetKind | null {
  const n = (name || "").toUpperCase();
  if (/FDC/.test(n)) return null;
  if (/COMPRAS/.test(n) && /RIG/.test(n)) return "RIG";
  if (/RECURENTE/.test(n) && /RIG/.test(n)) return "RIGREC";
  if (/COMPRAS/.test(n)) return "SHORE";
  if (/RECURENTE/.test(n)) return "REC";
  return null;
}

export function isFdcSheet(name: string): boolean {
  return /FDC/.test((name || "").toUpperCase());
}

// Columnas de monto/imp por tipo: [colMonto, colImp] en la fuente.
export function amountCols(kind: SheetKind): [string, string] {
  switch (kind) {
    case "REC":
    case "SHORE":
      return ["J", "M"];
    case "RIGREC":
      return ["H", "K"];
    case "RIG":
      return ["I", "L"];
  }
}

// Referencia ='Hoja'!COLnum  -> {sheet, col, row} | null
export function parseRef(formula: unknown): { sheet: string; col: string; row: number } | null {
  if (typeof formula !== "string" || formula.indexOf("!") < 0) return null;
  const m = formula.match(/'?([^'!]+)'?!\$?([A-Z]+)\$?(\d+)/);
  if (!m) return null;
  return { sheet: m[1].trim(), col: m[2], row: parseInt(m[3], 10) };
}

// Construye las formulas de una fila FDC nueva.
export function buildFdcFormulas(sheet: string, src: number, kind: SheetKind, fdcRow: number) {
  const [gc, hc] = amountCols(kind);
  return {
    A: `='${sheet}'!A${src}`,
    B: `='${sheet}'!B${src}`,
    C: `='${sheet}'!C${src}`,
    G: `='${sheet}'!${gc}${src}`,
    H: `='${sheet}'!${hc}${src}`,
    I: `=G${fdcRow}+H${fdcRow}`,
  };
}

// Orden canonico de secciones en la principal.
export function sectionRank(kind: SheetKind): number {
  return { REC: 0, SHORE: 1, RIGREC: 2, RIG: 3 }[kind];
}

// Chequea un vinculo FDC: ok | apunta a encabezado/vacio | #REF!
export function checkLink(
  getCell: (sheet: string, row: number, col: string) => unknown,
  fdcRefC: unknown
): { ok: boolean; reason: string } {
  const p = parseRef(fdcRefC);
  if (!p) return { ok: false, reason: "sin referencia" };
  const c = getCell(p.sheet, p.row, "C");
  if (c === null || c === undefined || String(c).trim() === "")
    return { ok: false, reason: `apunta a vacio ${p.sheet}!C${p.row}` };
  if (typeof c === "string" && (c.indexOf("#REF!") >= 0 || isHeaderLabel(c)))
    return { ok: false, reason: `apunta a encabezado/error ${p.sheet}!C${p.row}` };
  return { ok: true, reason: "" };
}
