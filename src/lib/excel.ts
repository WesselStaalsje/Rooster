import dayjs from "dayjs";
import * as XLSX from "xlsx-js-style";
import { loadTemplateWorkbook, extractFieldsFromTemplate, setCellString } from "./template";

function downloadArrayBuffer(buf: ArrayBuffer, filename: string) {
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

export async function exportFilledTemplate(params: {
  dateISO: string;
  valuesByAddr: Record<string, string>;
}) {
  const wb = await loadTemplateWorkbook();
  const { sheetName, fields } = extractFieldsFromTemplate(wb);
  const ws = wb.Sheets[sheetName];

  // Datum invullen op het datumveld (indien aanwezig)
  const dateString = dayjs(params.dateISO).format("D-M-YYYY");
  const datumField = fields.find(f => f.label === "DATUM:");
  if (datumField) setCellString(ws, datumField.addr, dateString, datumField.styleFromAddr);

  // overige velden
  for (const f of fields) {
    if (f.label === "DATUM:") continue;
    const v = params.valuesByAddr[f.addr] ?? "";
    setCellString(ws, f.addr, v, f.styleFromAddr);
  }

  // Cruciaal: cellStyles mee wegschrijven, anders verlies je opmaak.
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true }) as ArrayBuffer;
  downloadArrayBuffer(out, `Dagrooster_${params.dateISO}.xlsx`);
}
