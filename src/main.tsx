import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import ExcelJS from "exceljs";
import "./styles.css";
import { norm, dateVal } from "./review/reviewUtils";
import type { Row, Rule } from "./review/reviewUtils";
import { findMatchingPatient, findEvidence, evaluateEvidence } from "./review/reviewEngine";
import * as XLSX from "xlsx";


function App() {
  const [payer, setPayer] = useState<Row[]>([]),
    [payerHeaders, setPayerHeaders] = useState<string[]>([]),
    [emr, setEmr] = useState<Row[]>([]),
    [emrHeaders, setEmrHeaders] = useState<string[]>([]),
    [mapping, setMapping] = useState<any>({
      name: "",
      dob: "",
      provider: "",
      description: "",
      date: "",
      code: "",
      result: "",
    }),
    [rules, setRules] = useState<any>(null),
    [selected, setSelected] = useState<string[]>([]),
    [results, setResults] = useState<any[]>([]),
    [group, setGroup] = useState("Provider"),
    [filter, setFilter] = useState("All"),
    [status, setStatus] = useState("");
  useEffect(() => {
    window.api.loadRules().then((x) => setRules(JSON.parse(x)));
  }, []);
  const payerCols = useMemo(() => payerHeaders, [payerHeaders]);
  function parse(
    file: File,
    setter: (r: Row[]) => void,
    setHeaders: (h: string[]) => void,
  ) {
    const fr = new FileReader();
    fr.onload = () => {
      const wb = XLSX.read(fr.result, { type: "array", cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<Row>(ws, { defval: "" });
      setter(rows);
      setHeaders(rows.length ? Object.keys(rows[0]) : []);
    };
    fr.readAsArrayBuffer(file);
  }
  const autoMap = (headers: string[]) => {
    const pick = (patterns: string[]) =>
      patterns
    .map((p) => headers.find((h) => norm(h).includes(p)))
    .find((h) => h !== undefined) || "";
    return {
      name: pick(["membername", "patientname", "name"]),
      dob: pick(["dob", "dateofbirth", "birth"]),
      provider: pick(["provider", "pcp", "physician"]),
      description: pick([
        "description",
        "test",
        "procedure",
        "service",
        "clinical",
        "exam",
        "measure",
      ]),
      date: pick(["servicedate", "date", "performed", "resultdate"]),
      code: pick(["code", "cpt", "hcpcs", "icd"]),
      result: pick(["result", "value", "finding"]),
    };
  };
  function uploadPayer(f: File) {
    parse(f, setPayer, (h) => {
      setPayerHeaders(h);
    });
  }
  function uploadEmr(f: File) {
    parse(f, setEmr, (h) => {
      setEmrHeaders(h);
      setMapping(autoMap(h));
    });
  }
  function runReview() {
    if (!rules || !payer.length || !emr.length) {
      setStatus("Upload both files first.");
      return;
    }
    const complianceKey = payerHeaders.find((h) =>
      norm(h).includes("compliance"),
    );
    const nameKey = payerHeaders.find((h) =>
      ["membername", "patientname", "name"].some((p) => norm(h).includes(p)),
    );
    const dobKey = payerHeaders.find((h) =>
      ["dob", "dateofbirth", "birth"].some((p) => norm(h).includes(p)),
    );
    const measureKey = payerHeaders.find((h) => norm(h).includes("measure"));
    const active = payer
      .filter((r) =>
        complianceKey ? String(r[complianceKey]).trim() !== "1" : true,
      )
      .filter((r) => {
        if (!selected.length) return true;
        const m = String(r[measureKey || ""] || "").trim();
        return selected.includes(m);
      });
    const out = active.map((p: any, i) => {
      const pname = norm(p[nameKey || ""]);
      const measure = String(p[measureKey || ""] || "").trim();
      const rule =
        rules.measures[measure] ||
        Object.values(rules.measures).find(
          (x: any) => norm(x.label) === norm(measure),
        );
      const match = findMatchingPatient(
        p,
        emr,
        mapping,
        nameKey || "",
        dobKey || "",
      );
      if (!match)
        return {
          sourceIndex: payer.indexOf(p),
          name: p[nameKey || ""],
          dob: p[dobKey || ""],
          measure,
          status: "Not Found",
          notes: "Patient not found in EMR",
          provider:
            p[payerHeaders.find((h) => norm(h).includes("provider")) || ""],
        };
      const evid = findEvidence(emr, pname, mapping, rule);
      const evaluation = evaluateEvidence(evid, measure, rule);

      return {
        sourceIndex: payer.indexOf(p),
        name: p[nameKey || ''],
        dob: p[dobKey || ''],
        measure,
        status: evaluation.status,
        notes: evaluation.notes,
        provider: p[
          payerHeaders.find(h => norm(h).includes('provider')) || ''
        ]
      };
    });
    setResults(out);
    setStatus(`Reviewed ${out.length} care-gap rows.`);
  }
  async function exportFile() {
    if (!results.length) {
      setStatus("Run Automated Review first.");
      return;
    }
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Payer List");
    ws.addRow([...payerHeaders, "Notes"]);
    payer.forEach((r, i) => {
      const row = ws.addRow([...payerHeaders.map((h) => r[h]), ""]);
      const result = results.find((x) => x.sourceIndex === i);
      if (result) {
        const c = row.getCell(payerHeaders.length + 1);
        c.value = result.notes;
        if (result.status === "Closed")
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFC6EFCE" },
          };
        else if (result.status === "Open")
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFFFEB9C" },
          };
        else
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "FFD9D9D9" },
          };
      }
    });
    ws.getRow(1).font = { bold: true };
    ws.columns.forEach(
      (c) =>
        (c.width = Math.min(
          Math.max((c.header?.toString().length || 10) + 2, 12),
          35,
        )),
    );
    const buf = await wb.xlsx.writeBuffer();
    const d = await window.api.saveDialog("care-gap-results.xlsx");
    if (!d.canceled && d.filePath) {
      await window.api.writeFile(d.filePath, new Uint8Array(buf));
      setStatus("Export complete.");
    }
  }
  const shown = results
    .filter((r) => filter === "All" || r.status === filter)
    .sort((a, b) =>
      String(
        a[
          group === "Provider"
            ? "provider"
            : group === "Measure"
              ? "measure"
              : "status"
        ],
      ).localeCompare(
        String(
          b[
            group === "Provider"
              ? "provider"
              : group === "Measure"
                ? "measure"
                : "status"
          ],
        ),
      ),
    );
  return (
    <div className="app">
      <header>
        <div>
          <h1>Care Gap Coordinator</h1>
          <p>Local-first automated chart review</p>
        </div>
        <span className="badge">v0.1 MVP</span>
      </header>
      <nav>
        <b>1 Upload</b>
        <b>2 Map & Configure</b>
        <b>3 Automated Review</b>
        <b>4 Export</b>
      </nav>
      <main>
        <section>
          <h2>1. Upload</h2>
          <div className="cards">
            <div className="drop">
              Payer Excel / CSV
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) =>
                  e.target.files && uploadPayer(e.target.files[0])
                }
              />
              <small>
                {payer.length
                  ? `${payer.length} rows loaded`
                  : "Choose payer file"}
              </small>
            </div>
            <div className="drop">
              EMR Export
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => e.target.files && uploadEmr(e.target.files[0])}
              />
              <small>
                {emr.length ? `${emr.length} rows loaded` : "Choose EMR export"}
              </small>
            </div>
          </div>
        </section>
        <section>
          <h2>2. Map & Configure</h2>
          {emrHeaders.length > 0 && (
            <div className="mapping">
              {Object.keys(mapping).map((k) => (
                <label key={k}>
                  {{
                    name: 'Patient Name',
                    dob: 'Date of Birth',
                    provider: 'Provider',
                    description: 'Clinical Description',
                    date: 'Service Date',
                    code: 'CPT Code',
                    result: 'Result'
                  }[k] || k}
                  <select
                    value={mapping[k]}
                    onChange={(e) =>
                      setMapping({ ...mapping, [k]: e.target.value })
                    }
                  >
                    <option value="">— ignore —</option>
                    {emrHeaders.map((h) => (
                      <option key={h}>{h}</option>
                    ))}
                  </select>
                </label>
              ))}
            </div>
          )}
          <div className="rules">
            <strong>Measures</strong>
            {rules &&
              Object.entries(rules.measures).map(([k, v]: any) => (
                <label key={k}>
                  <input
                    type="checkbox"
                    checked={selected.includes(k)}
                    onChange={(e) =>
                      setSelected(
                        e.target.checked
                          ? [...selected, k]
                          : selected.filter((x) => x !== k),
                      )
                    }
                  />
                  {v.label}
                </label>
              ))}
            <small>{rules?.warning}</small>
          </div>
          <button onClick={runReview}>Run Automated Review</button>
        </section>
        <section>
          <div className="sectionTitle">
            <h2>3. Automated Review</h2>
            <div>
              <select value={group} onChange={(e) => setGroup(e.target.value)}>
                <option>Provider</option>
                <option>Measure</option>
                <option>Status</option>
                <option>None</option>
              </select>
              <select
                value={filter}
                onChange={(e) => setFilter(e.target.value)}
              >
                <option>All</option>
                <option>Closed</option>
                <option>Open</option>
                <option>Not Found</option>
              </select>
            </div>
          </div>
          <div className="tableWrap">
            <table>
              <thead>
                <tr>
                  <th>Member</th>
                  <th>DOB</th>
                  <th>Measure</th>
                  <th>Provider</th>
                  <th>Status</th>
                  <th>Notes</th>
                </tr>
              </thead>
              <tbody>
                {shown.map((r, i) => (
                  <tr key={i}>
                    <td>{r.name}</td>
                    <td>{String(r.dob)}</td>
                    <td>{r.measure}</td>
                    <td>{r.provider}</td>
                    <td>
                      <span
                        className={`status ${r.status.replace(" ", "-").toLowerCase()}`}
                      >
                        {r.status}
                      </span>
                    </td>
                    <td>{r.notes}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
        <section>
          <h2>4. Export</h2>
          <p>
            Exports the original payer rows and columns, in their original
            order, with one <b>Notes</b> column added at the end. Only care-gap
            rows receive notes.
          </p>
          <button className="primary" onClick={exportFile}>
            Export Excel
          </button>
          {status && <p className="statusMsg">{status}</p>}
        </section>
      </main>
    </div>
  );
}
createRoot(document.getElementById("root")!).render(<App />);
