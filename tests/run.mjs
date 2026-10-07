// Tests del motor puro (node). Se empaqueta logic.ts con esbuild y se valida.
import { buildSync } from "esbuild";
import { writeFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { pathToFileURL } from "url";

const out = join(tmpdir(), "fdc-logic-test.mjs");
buildSync({
  entryPoints: ["src/logic.ts"],
  outfile: out,
  bundle: true,
  format: "esm",
  platform: "node",
  logLevel: "silent",
});
const L = await import(pathToFileURL(out).href);

let n = 0;
function ok(cond, name) {
  n++;
  if (!cond) { console.error("FALLA:", name); process.exitCode = 1; }
  else console.log("ok:", name);
}

ok(L.kindOfSheet("RECURENTE REGULAR OCTUBRE 2026") === "REC", "kind REC");
ok(L.kindOfSheet("COMPRAS SHORE BASE OCTUBRE 2026") === "SHORE", "kind SHORE");
ok(L.kindOfSheet("RECURENTE RIG 42 OCTUBRE 2026") === "RIGREC", "kind RIGREC");
ok(L.kindOfSheet("COMPRAS RIG-42 OCTUBRE 2026") === "RIG", "kind RIG");
ok(L.kindOfSheet("Compras - FDC  10 2026") === null, "FDC no es fuente");
ok(L.isFdcSheet("Compras - FDC  10 2026"), "detecta FDC");
ok(L.isHeaderLabel("MANTENIMIENTO DE FLOTA"), "header flota");
ok(L.isHeaderLabel("Sevicios Recurrente  SSGG"), "header recurrente typo");
ok(!L.isHeaderLabel("SUMINISTRO DE CAMIÓN CISTERNA"), "dato no es header");
const p = L.parseRef("='COMPRAS RIG-42 OCTUBRE 2026'!C12");
ok(p && p.sheet === "COMPRAS RIG-42 OCTUBRE 2026" && p.col === "C" && p.row === 12, "parseRef");
ok(L.parseRef("texto") === null, "parseRef null");
const f = L.buildFdcFormulas("COMPRAS RIG-42 OCTUBRE 2026", 12, "RIG", 51);
ok(f.G === "='COMPRAS RIG-42 OCTUBRE 2026'!I12" && f.H.endsWith("!L12") && f.I === "=G51+H51", "buildFdcFormulas RIG");
const f2 = L.buildFdcFormulas("RECURENTE REGULAR OCTUBRE 2026", 7, "REC", 7);
ok(f2.G.endsWith("!J7") && f2.H.endsWith("!M7"), "buildFdcFormulas REC");
ok(L.WEEK_OF["L"] === 41 && WEEK_OF_SAFE(), "weeks");
function WEEK_OF_SAFE() { return L.WEEK_OF["N"] === 42 && L.WEEK_OF["P"] === 43; }
ok(L.sectionRank("REC") < L.sectionRank("SHORE"), "orden secciones");
console.log(`\n${n} checks, exit=${process.exitCode || 0}`);
