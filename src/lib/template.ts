import * as XLSX from "xlsx-js-style";
import { safeString, uid } from "./utils";

export type TemplateField = {
  id: string;
  group: string;     // bv "Roosendaal 2"
  label: string;     // bv "6" of "7 TR" of "Roosendaal | 9 Con"
  addr: string;      // Excel celadres waar de naam moet komen
  sheetName: string;
  sortKey: string;   // voor nette volgorde
};

type FoundCell = { addr: string; r: number; c: number };

function isCellAddress(k: string) {
  return /^[A-Z]+[0-9]+$/.test(k);
}

function decodeAddr(addr: string) {
  const p = XLSX.utils.decode_cell(addr);
  return { r: p.r, c: p.c };
}

function encodeAddr(r: number, c: number) {
  return XLSX.utils.encode_cell({ r, c });
}

function findCellByExactText(ws: any, text: string): FoundCell | null {
  const target = text.trim();
  for (const k of Object.keys(ws)) {
    if (!isCellAddress(k)) continue;
    const cell = ws[k];
    const v = safeString(cell?.v).trim();
    if (v === target) {
      const { r, c } = decodeAddr(k);
      return { addr: k, r, c };
    }
  }
  return null;
}

function findMergeForCell(ws: any, r: number, c: number) {
  const merges = (ws["!merges"] ?? []) as Array<{ s: { r: number; c: number }, e: { r: number; c: number } }>;
  for (const m of merges) {
    if (r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c) return m;
  }
  return null;
}

function extractTwoColBlock(ws: any, sheetName: string, headerText: string, maxRows = 28): TemplateField[] {
  const header = findCellByExactText(ws, headerText);
  if (!header) return [];

  const fields: TemplateField[] = [];
  const timeCol = header.c;       // label staat onder header in dezelfde kolom
  const nameCol = header.c + 1;   // naam direct ernaast
  const startRow = header.r + 1;

  let emptyStreak = 0;

  for (let r = startRow; r < startRow + maxRows; r++) {
    const timeAddr = encodeAddr(r, timeCol);
    const nameAddr = encodeAddr(r, nameCol);

    const timeVal = safeString(ws[timeAddr]?.v).trim();
    const nameVal = safeString(ws[nameAddr]?.v).trim();

    if (!timeVal && !nameVal) {
      emptyStreak++;
      if (emptyStreak >= 8) break;
      continue;
    }
    emptyStreak = 0;

    if (!timeVal) continue;

    fields.push({
      id: uid("f"),
      group: headerText,
      label: timeVal,
      addr: nameAddr,
      sheetName,
      sortKey: `${headerText}|${r.toString().padStart(3, "0")}`
    });
  }

  return fields;
}

function extractKeyValue(ws: any, sheetName: string, labelText: string, group: string): TemplateField[] {
  const cell = findCellByExactText(ws, labelText);
  if (!cell) return [];
  const valueAddr = encodeAddr(cell.r, cell.c + 1);

  return [{
    id: uid("kv"),
    group,
    label: labelText,
    addr: valueAddr,
    sheetName,
    sortKey: `${group}|${labelText}`
  }];
}

function extractZwareBergingRows(ws: any, sheetName: string, group: string, rowLabels: string[]): TemplateField[] {
  const out: TemplateField[] = [];

  for (const label of rowLabels) {
    const areaCell = findCellByExactText(ws, label);
    if (!areaCell) continue;

    const merge = findMergeForCell(ws, areaCell.r, areaCell.c);
    const areaEndCol = merge ? merge.e.c : areaCell.c;

    const timeAddr = encodeAddr(areaCell.r, areaEndCol + 1);
    const nameAddr = encodeAddr(areaCell.r, areaEndCol + 2);

    const timeVal = safeString(ws[timeAddr]?.v).trim();
    const pretty = timeVal ? `${label} | ${timeVal}` : label;

    out.push({
      id: uid("zb"),
      group,
      label: pretty,
      addr: nameAddr,
      sheetName,
      sortKey: `${group}|${label}|${areaCell.r.toString().padStart(3, "0")}`
    });
  }

  return out;
}

export async function loadTemplateWorkbook(): Promise<any> {
  const res = await fetch("/template.xlsx", { cache: "no-store" });
  if (!res.ok) throw new Error("template.xlsx niet gevonden in /public.");
  const buf = await res.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  return wb;
}

export function extractFieldsFromTemplate(wb: any) {
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error("Geen sheet gevonden in template.");

  const twoColHeaders = [
    "Roosendaal 2",
    "Raamsdonksveer 2",
    "Rosmalen 3",
    "Eindhoven 1",
    "Duiven 4",
    "Breda 2",
    "Hulten 2",
    "Oss 3",
    "Boxmeer 3",
    "Ede 4",
    "Andelst 4"
  ];

  const keyValues: Array<{ label: string; group: string }> = [
    { label: "DATUM:", group: "Algemeen" },
    { label: "Vroeg Roosendaal", group: "Algemeen" },
    { label: "Laat Roosendaal", group: "Algemeen" },
    { label: "Vroeg Veer", group: "Algemeen" },
    { label: "Laat Veer", group: "Algemeen" }
  ];

  const zwareLabels = [
    "Roosendaal",
    "Raamsdonksveer",
    "Breda",
    "Hulten",
    "Eindhoven",
    "Duiven",
    "Ede",
    "Internationaal"
  ];

  let fields: TemplateField[] = [];

  for (const h of twoColHeaders) fields = fields.concat(extractTwoColBlock(ws, sheetName, h));
  for (const kv of keyValues) fields = fields.concat(extractKeyValue(ws, sheetName, kv.label, kv.group));
  fields = fields.concat(extractZwareBergingRows(ws, sheetName, "Zware Berging", zwareLabels));

  fields.sort((a, b) => a.sortKey.localeCompare(b.sortKey));

  const seen = new Set<string>();
  fields = fields.filter(f => {
    const k = `${f.sheetName}!${f.addr}`;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  return { sheetName, fields };
}

export function setCellString(ws: any, addr: string, value: string) {
  if (!ws[addr]) ws[addr] = { t: "s", v: "" };
  ws[addr].t = "s";
  ws[addr].v = value ?? "";
}
