import dayjs from "dayjs";
import "dayjs/locale/nl";
import { useEffect, useMemo, useState } from "react";
import { exportFilledTemplate } from "./lib/excel";
import { loadTemplateWorkbook, extractFieldsFromTemplate, TemplateField } from "./lib/template";
import { loadValues, saveValues } from "./lib/storage";

dayjs.locale("nl");

function todayISO() {
  return dayjs().format("YYYY-MM-DD");
}

export default function App() {
  const [dateISO, setDateISO] = useState(todayISO());
  const [fields, setFields] = useState<TemplateField[]>([]);
  const [valuesByAddr, setValuesByAddr] = useState<Record<string, string>>(() => loadValues(todayISO()));
  const [error, setError] = useState("");

  useEffect(() => {
    setValuesByAddr(loadValues(dateISO));
  }, [dateISO]);

  useEffect(() => {
    (async () => {
      try {
        setError("");
        const wb = await loadTemplateWorkbook();
        const extracted = extractFieldsFromTemplate(wb);
        setFields(extracted.fields);
      } catch (e: any) {
        setError(e?.message ?? "Template laden mislukt.");
      }
    })();
  }, []);

  const groups = useMemo(() => {
    const m = new Map<string, TemplateField[]>();
    for (const f of fields) {
      const key = f.group;
      if (!m.has(key)) m.set(key, []);
      m.get(key)!.push(f);
    }
    return Array.from(m.entries());
  }, [fields]);

  const setValue = (addr: string, value: string) => {
    const next = { ...valuesByAddr, [addr]: value };
    setValuesByAddr(next);
    saveValues(dateISO, next);
  };

  const exportXlsx = async () => {
    try {
      setError("");
      await exportFilledTemplate({ dateISO, valuesByAddr });
    } catch (e: any) {
      setError(e?.message ?? "Export mislukt.");
    }
  };

  return (
    <div className="container">
      <div className="card">
        <h1>Dagrooster → Excel (exact template)</h1>

        <div className="row" style={{ justifyContent: "space-between" }}>
          <div className="row">
            <div className="col">
              <div className="small muted">Datum</div>
              <input type="date" value={dateISO} onChange={(e) => setDateISO(e.target.value)} />
            </div>
            <button onClick={exportXlsx} disabled={!fields.length || !!error}>
              Export Excel
            </button>
            <button
              className="secondary"
              onClick={() => {
                const next: Record<string, string> = {};
                setValuesByAddr(next);
                saveValues(dateISO, next);
              }}
              disabled={!fields.length}
            >
              Leegmaken (alle velden)
            </button>
          </div>

          <div className="badge">
            {fields.length ? `Velden gevonden: ${fields.length}` : "Template nog niet geladen"}
          </div>
        </div>

        {error && <div className="warning" style={{ marginTop: 12 }}>{error}</div>}

        {!error && !fields.length && (
          <div className="muted" style={{ marginTop: 12 }}>
            Let op: je moet `public/template.xlsx` uploaden in de repo.
          </div>
        )}

        {groups.map(([groupName, groupFields]) => (
          <div key={groupName} className="group" style={{ marginTop: 14 }}>
            <div className="group-title">
              <h2>{groupName}</h2>
              <div className="small muted">{groupFields.length} velden</div>
            </div>

            <table className="table">
              <thead>
                <tr>
                  <th style={{ width: 260 }}>Label</th>
                  <th>Waarde</th>
                </tr>
              </thead>
              <tbody>
                {groupFields.map((f) => (
                  <tr key={f.id}>
                    <td className="muted">{f.label}</td>
                    <td>
                      <input
                        value={valuesByAddr[f.addr] ?? ""}
                        onChange={(e) => setValue(f.addr, e.target.value)}
                        placeholder="Naam / tekst"
                        style={{ width: "100%" }}
                      />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>

            <div className="small muted">
              Deze velden worden direct in de template-cellen geschreven (adres-based), layout blijft exact hetzelfde.
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
