// UI del panel + operaciones Excel via Office.js
import {
  SourceRow, SheetKind, WEEK_OF,
  isHeaderLabel, kindOfSheet, isFdcSheet,
  parseRef, buildFdcFormulas, sectionRank,
} from "./logic";

declare const Excel: any;
declare const Office: any;

let model: { sources: SourceRow[]; fdc: { row: number; sheet: string; src: number; week: number | null }[]; fdcSheet: string; totalRow: number } | null = null;

const $ = (id: string) => document.getElementById(id) as any;
function status(t: string) { $("status").textContent = t; }

async function loadModel() {
  status("Leyendo libro...");
  await Excel.run(async (ctx: any) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const sources: SourceRow[] = [];
    let fdcSheet = "";
    const fdcCells: { row: number; f: any[] }[] = [];
    let totalRow = 0;
    for (const sh of sheets.items) {
      const kind: SheetKind | null = kindOfSheet(sh.name);
      const rng = sh.getUsedRange();
      rng.load("values,formulas,rowCount,rowIndex");
      await ctx.sync();
      const vals: any[][] = rng.values;
      const start = rng.rowIndex + 1;
      if (kind) {
        let section = "";
        for (let i = 0; i < vals.length; i++) {
          const c = String(vals[i][2] || "");
          if (!c.trim()) continue;
          if (isHeaderLabel(c)) { section = c.trim().slice(0, 60); continue; }
          sources.push({ sheet: sh.name, row: start + i, rdm: String(vals[i][0] || ""), desc: c.trim().slice(0, 80), kind, section });
        }
      } else if (isFdcSheet(sh.name) && !fdcSheet) {
        fdcSheet = sh.name;
        const forms: any[][] = rng.formulas;
        for (let i = 0; i < forms.length; i++) {
          const g = String(forms[i][6] || "");
          if (/^=SUM\(/i.test(g)) { totalRow = start + i; continue; }
          const p = parseRef(forms[i][2]);
          if (!p) continue;
          let week: number | null = null;
          for (const col of ["K", "L", "M", "N", "O", "P", "Q", "R"]) {
            const v = String(forms[col.charCodeAt(0) - 65] || "");
            if (v.startsWith("=")) { week = WEEK_OF[col]; break; }
          }
          fdcCells.push({ row: start + i, f: forms[i] });
          void week;
        }
      }
    }
    // segunda pasada FDC ya con semanas (reusa fdcCells)
    const fdc = fdcCells.map(({ row, f }) => {
      const p = parseRef(f[2])!;
      let week: number | null = null;
      for (const col of ["K", "L", "M", "N", "O", "P", "Q", "R"]) {
        if (String(f[col.charCodeAt(0) - 65] || "").startsWith("=")) { week = WEEK_OF[col]; break; }
      }
      return { row, sheet: p.sheet, src: p.row, week };
    });
    model = { sources, fdc, fdcSheet, totalRow };
  });
  renderFilters();
  renderList();
  status(`Listo: ${model!.sources.length} filas fuente, ${model!.fdc.length} en principal.`);
}

function covered(sheet: string, src: number): boolean {
  return !!model!.fdc.find((f) => f.sheet === sheet && f.src === src);
}

function renderFilters() {
  const ss = [...new Set(model!.sources.map((s) => s.sheet))];
  const fs = $("f-sheet"); fs.innerHTML = '<option value="">Todas las hojas</option>';
  ss.forEach((s) => { const o = document.createElement("option"); o.value = s; o.textContent = s; fs.appendChild(o); });
  const secs = [...new Set(model!.sources.map((s) => s.section).filter(Boolean))];
  const fx = $("f-sec"); fx.innerHTML = '<option value="">Todas las secciones</option>';
  secs.forEach((s) => { const o = document.createElement("option"); o.value = s; o.textContent = s; fx.appendChild(o); });
}

function filtered(): SourceRow[] {
  const q = ($("q").value || "").toUpperCase();
  const sh = $("f-sheet").value, sec = $("f-sec").value;
  return model!.sources
    .filter((s) => (!sh || s.sheet === sh) && (!sec || s.section === sec))
    .filter((s) => !q || s.desc.toUpperCase().includes(q) || s.rdm.toUpperCase().includes(q))
    .sort((a, b) => sectionRank(a.kind) - sectionRank(b.kind) || a.sheet.localeCompare(b.sheet) || a.row - b.row);
}

function renderList() {
  const box = $("list"); box.innerHTML = "";
  let lastSec = "";
  for (const s of filtered()) {
    const key = s.sheet + " / " + (s.section || "general");
    if (key !== lastSec) {
      const h = document.createElement("div");
      h.className = "sec-h"; h.textContent = key; box.appendChild(h);
      lastSec = key;
    }
    const hit = covered(s.sheet, s.row);
    const lab = document.createElement("label");
    lab.className = hit ? "row-present" : "row-missing";
    const cb = document.createElement("input");
    cb.type = "checkbox"; cb.checked = !hit;
    cb.dataset.sheet = s.sheet; cb.dataset.src = String(s.row);
    lab.appendChild(cb);
    lab.appendChild(document.createTextNode(` ${hit ? "✓" : "○"} [${s.rdm || "s/n"}] ${s.desc}`));
    box.appendChild(lab); box.appendChild(document.createElement("br"));
  }
}

function selected(): { sheet: string; src: number }[] {
  return [...document.querySelectorAll("#list input:checked")].map((cb: any) => ({
    sheet: cb.dataset.sheet, src: parseInt(cb.dataset.src, 10),
  }));
}

async function readFdcMap(ctx: any, fdc: any): Promise<Map<string, number>> {
  const rng = fdc.getUsedRange();
  rng.load("formulas,rowIndex");
  await ctx.sync();
  const vals = rng.formulas as any[][];
  const start = rng.rowIndex + 1;
  const pos = new Map<string, number>();
  for (let i = 0; i < vals.length; i++) {
    const p = parseRef(vals[i][2]);
    if (p) pos.set(p.sheet + "|" + p.row, start + i);
  }
  return pos;
}

async function applyAdd(sel: { sheet: string; src: number }[]) {
  await Excel.run(async (ctx: any) => {
    const fdc = ctx.workbook.worksheets.getItem(model!.fdcSheet);
    const pos = await readFdcMap(ctx, fdc);
    const key = (sh: string, r: number) => sh + "|" + r;
    const coveredSrc = (sh: string) => [...pos.keys()].filter((k) => k.startsWith(sh + "|")).map((k) => parseInt(k.split("|")[1], 10));
    const jobs = sel
      .filter((s) => !pos.has(key(s.sheet, s.src)))
      .map((s) => ({ ...s, kind: kindOfSheet(s.sheet)! }))
      .sort((a, b) => sectionRank(b.kind) - sectionRank(a.kind) || b.src - a.src);
    for (const j of jobs) {
      const below = coveredSrc(j.sheet).filter((r) => r < j.src).sort((x, y) => y - x)[0];
      if (below === undefined) continue;
      const nrow = pos.get(key(j.sheet, below))! + 1;
      fdc.getRange(`${nrow}:${nrow}`).insert(Excel.InsertShiftDirection.down);
      const f = buildFdcFormulas(j.sheet, j.src, j.kind, nrow);
      fdc.getRange(`A${nrow}`).formulas = [[f.A]];
      fdc.getRange(`B${nrow}`).formulas = [[f.B]];
      fdc.getRange(`C${nrow}`).formulas = [[f.C]];
      fdc.getRange(`G${nrow}`).formulas = [[f.G]];
      fdc.getRange(`H${nrow}`).formulas = [[f.H]];
      fdc.getRange(`I${nrow}`).formulas = [[f.I]];
      fdc.getRange(`J${nrow}`).values = [["USD Pagaderos en VES al cambio"]];
      fdc.getRange(`S${nrow}`).values = [[1]];
      for (const cc of ["T", "U", "V", "W"]) fdc.getRange(`${cc}${nrow}`).values = [[false]];
      fdc.getRange(`X${nrow}`).values = [["En Proceso"]];
      fdc.getRange(`Y${nrow}`).values = [["No"]];
      fdc.getRange(`Z${nrow}`).values = [["Por iniciar la solicitud de cotizacion "]];
      const kr = fdc.getRange(`K${nrow}:R${nrow}`);
      kr.clear();
      kr.format.fill.color = "FFFF00";
      await ctx.sync();
      // re-mapear tras insertar
      const r2 = fdc.getUsedRange(); r2.load("formulas,rowIndex"); await ctx.sync();
      pos.clear();
      const v2 = r2.formulas as any[][];
      const st = r2.rowIndex + 1;
      for (let i = 0; i < v2.length; i++) {
        const p = parseRef(v2[i][2]);
        if (p) pos.set(p.sheet + "|" + p.row, st + i);
      }
    }
    // totales
    const rt = fdc.getUsedRange(); rt.load("formulas,rowIndex"); await ctx.sync();
    const vt = rt.formulas as any[][];
    let total = 0;
    for (let i = 0; i < vt.length; i++)
      if (/^=SUM\(/i.test(String(vt[i][6] || ""))) { total = rt.rowIndex + 1 + i; break; }
    if (total) {
      const last = total - 1;
      for (const c of ["G", "H", "I", "K", "L", "M", "N", "O", "P", "Q", "R"])
        fdc.getRange(`${c}${total}`).formulas = [[`=SUM(${c}6:${c}${last})`]];
      await ctx.sync();
    }
  });
  status("Insertadas. Pulsa Actualizar para ver el estado.");
}

async function applyDel(sel: { sheet: string; src: number }[], fromSource: boolean) {
  await Excel.run(async (ctx: any) => {
    const fdc = ctx.workbook.worksheets.getItem(model!.fdcSheet);
    const pos = await readFdcMap(ctx, fdc);
    const targets = sel
      .map((s) => ({ ...s, fdcRow: pos.get(s.sheet + "|" + s.src) }))
      .filter((t) => t.fdcRow !== undefined) as { sheet: string; src: number; fdcRow: number }[];
    targets.sort((a, b) => b.fdcRow - a.fdcRow);
    for (const t of targets) {
      fdc.getRange(`${t.fdcRow}:${t.fdcRow}`).delete(Excel.DeleteShiftDirection.up);
      if (fromSource) {
        const ws = ctx.workbook.worksheets.getItem(t.sheet);
        ws.getRange(`${t.src}:${t.src}`).delete(Excel.DeleteShiftDirection.up);
      }
    }
    await ctx.sync();
  });
  status("Eliminadas. Pulsa Actualizar.");
}

async function repair() {
  const box = $("orphans"); box.innerHTML = "";
  let found = 0;
  await Excel.run(async (ctx: any) => {
    const fdc = ctx.workbook.worksheets.Item ? null : null;
    void fdc;
    const ws = ctx.workbook.worksheets.getItem(model!.fdcSheet);
    const rng = ws.getUsedRange(); rng.load("formulas,rowIndex"); await ctx.sync();
    const vals = rng.formulas as any[][];
    const start = rng.rowIndex + 1;
    const names = new Set<string>();
    {
      const sheets = ctx.workbook.worksheets; sheets.load("items/name"); await ctx.sync();
      sheets.items.forEach((s: any) => names.add(s.name));
    }
    for (let i = 0; i < vals.length; i++) {
      const f = String(vals[i][2] || "");
      if (!f.includes("!")) continue;
      const p = parseRef(f);
      if (!p) continue;
      const broken = f.includes("#REF!") || !names.has(p.sheet);
      if (!broken) continue;
      found++;
      const d = document.createElement("div");
      d.className = "orphan";
      const r = start + i;
      d.textContent = `Fila ${r}: ref rota (${f.slice(0, 70)}). `;
      const btn = document.createElement("button");
      btn.textContent = "Borrar fila";
      btn.onclick = async () => {
        await Excel.run(async (c2: any) => {
          const w = c2.workbooks ?? null;
          void w;
          const w2 = c2.workbook.worksheets.getItem(model!.fdcSheet);
          w2.getRange(`${r}:${r}`).delete(Excel.DeleteShiftDirection.up);
          await c2.sync();
        });
        loadModel();
      };
      d.appendChild(btn);
      box.appendChild(d);
    }
    await ctx.sync();
  });
  status(found ? `${found} huérfanas listadas.` : "Sin vínculos rotos.");
}

function wire() {
  const all = (v: boolean) => document.querySelectorAll("#list input").forEach((c: any) => (c.checked = v));
  ($("b-all") as any).onclick = () => all(true);
  ($("b-none") as any).onclick = () => all(false);
  ($("b-missing") as any).onclick = () => {
    document.querySelectorAll("#list input").forEach((c: any) => {
      c.checked = (c.parentElement as HTMLElement).className.includes("missing");
    });
  };
  ($("b-refresh") as any).onclick = () => loadModel();
  ($("b-repair") as any).onclick = () => repair();
  ($("q") as any).oninput = () => renderList();
  ($("f-sheet") as any).onchange = () => renderList();
  ($("f-sec") as any).onchange = () => renderList();
  document.querySelectorAll('input[name="mode"]').forEach((r: any) => (r.onchange = () => {
    ($("del-scope") as any).style.display =
      (document.querySelector('input[name="mode"]:checked') as any).value === "del" ? "inline" : "none";
  }));
  ($("b-apply") as any).onclick = async () => {
    const mode = (document.querySelector('input[name="mode"]:checked') as any).value;
    const sel = selected();
    if (!sel.length) { status("Nada seleccionado."); return; }
    if (mode === "del" && !confirm(`Eliminar ${sel.length} fila(s)${($("del-source") as any).checked ? " (incluye hojas fuente)" : ""}?`)) return;
    status("Aplicando...");
    try {
      if (mode === "add") await applyAdd(sel);
      else await applyDel(sel, ($("del-source") as any).checked);
    } catch (e: any) {
      status("Error: " + (e.message || e));
    }
  };
}

Office.onReady(() => { wire(); loadModel(); });
