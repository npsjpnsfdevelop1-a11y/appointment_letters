"use client";

import { useEffect, useMemo, useState } from "react";
import JSZip from "jszip";

const templateUrl = "/Sample.docx";

const defaultRows = [
  { label: "Letter reference", find: "NPS/APPT.LETTER/JPN/2025-274", value: "NPS/APPT.LETTER/JPN/2026-001", mode: "replace" },
  { label: "Letter date", find: "4th March 2026", value: "2nd June 2026", mode: "replace" },
  { label: "School address", find: "National Public School, JP Nagar- Survey No:22, Anjanapura Main Road, JP Nagar 9th Phase, Anjanapura Township, Bengaluru, Karnataka – 560062", value: "National Public School, JP Nagar- Survey No:22, Anjanapura Main Road, JP Nagar 9th Phase, Anjanapura Township, Bengaluru, Karnataka – 560062", mode: "replace" },
  { label: "Employee name", find: "ABC", value: "", mode: "replace" },
  { label: "Employee detail", find: "XVZ", value: "", mode: "replace" },
  { label: "Designation in agreement", find: "Senior Mistress", value: "", mode: "replace" },
  { label: "Work location", find: "Bangalore", value: "Bangalore", mode: "replace" },
  { label: "Annexure C name", find: "Name:", value: "", mode: "nextCell" },
  { label: "Annexure C designation", find: "Designation:", value: "", mode: "nextCell" },
  { label: "EmployeeID", find: "EmployeeID:", value: "", mode: "nextCell" },
  { label: "Date of joining", find: "Date of joining:", value: "", mode: "nextCell" },
  { label: "Basic", find: "Basic:", value: "", mode: "nextCell" },
  { label: "Dearness allowance", find: "Dearness Allowance:", value: "", mode: "nextCell" },
  { label: "House rent allowance", find: "House Rent Allowance:", value: "", mode: "nextCell" },
  { label: "Special allowance", find: "Special Allowance:", value: "", mode: "nextCell" },
  { label: "Total gross salary", find: "Total Gross Salary:", value: "", mode: "nextCell" },
  { label: "TDS", find: "TDS:", value: "", mode: "nextCell" },
  { label: "Provident fund employee", find: "Provident Fund (Employee):", value: "", mode: "nextCell" },
  { label: "ESI employee", find: "ESI(Employee):", value: "", mode: "nextCell" },
  { label: "Professional tax", find: "Professional Tax:", value: "", mode: "nextCell" },
  { label: "Total net salary", find: "Total Net Salary (Take-Home Salary):", value: "", mode: "nextCell" },
  { label: "Provident fund employer", find: "Provident Fund (Employer):", value: "", mode: "nextCell" },
  { label: "Total CTC", find: "Total Cost to Company:", value: "", mode: "nextCell" },
  { label: "Total annual CTC", find: "Total Annual Cost to Company:", value: "", mode: "nextCell" }
];

export default function Home() {
  const [templateBytes, setTemplateBytes] = useState(null);
  const [status, setStatus] = useState("Loading Sample.docx");
  const [rows, setRows] = useState(defaultRows);
  const [employeeName, setEmployeeName] = useState("");
  const [designation, setDesignation] = useState("");
  const [fileName, setFileName] = useState("{{name}}_appointment_letter.docx");
  const [isGenerating, setIsGenerating] = useState(false);
  const [logs, setLogs] = useState([]);

  useEffect(() => {
    fetch(templateUrl, { cache: "no-store" })
      .then((response) => {
        if (!response.ok) throw new Error(`Template request failed: ${response.status}`);
        return response.arrayBuffer();
      })
      .then((bytes) => {
        setTemplateBytes(bytes);
        setStatus("Sample.docx ready");
      })
      .catch((error) => {
        console.error(error);
        setStatus("Could not load Sample.docx");
      });
  }, []);

  const generationRows = useMemo(() => getRowsForGeneration(rows, employeeName, designation), [rows, employeeName, designation]);
  const activeRows = useMemo(
    () => generationRows
      .map((row) => ({ find: row.find.trim(), value: sanitizeReplacement(row.value), mode: row.mode || "replace", label: row.label }))
      .filter((row) => row.find && row.value),
    [generationRows]
  );
  const outputName = normalizeDocxName(renderFileName(fileName, getEmployeeName(rows, employeeName)));

  async function generateDocument(event) {
    event.preventDefault();
    if (!templateBytes || isGenerating) return;

    const runLogs = [];
    const log = (message) => {
      const line = `${new Date().toLocaleTimeString()} - ${message}`;
      runLogs.push(line);
      setLogs([...runLogs]);
      console.log(`[offer-generator] ${message}`);
    };

    setIsGenerating(true);
    setStatus("Generating file");
    setLogs([]);
    log("Export started");

    try {
      log(`Template size: ${templateBytes.byteLength} bytes`);
      const zip = await JSZip.loadAsync(templateBytes);
      log(`DOCX zip loaded with ${Object.keys(zip.files).length} entries`);

      const replacements = activeRows.filter((row) => row.mode !== "nextCell" && row.find !== row.value);
      const skippedNoOps = activeRows.filter((row) => row.mode !== "nextCell" && row.find === row.value);
      const nextCellFields = activeRows.filter((row) => row.mode === "nextCell");
      skippedNoOps.forEach((row) => log(`Skipped no-op replacement: "${row.label}"`));
      log(`Text replacements: ${replacements.length}`);
      log(`Table cell fills: ${nextCellFields.length}`);

      if (!activeRows.length) {
        setStatus("Fill at least one replacement value");
        log("Stopped: no active rows");
        return;
      }

      const counts = Object.fromEntries(activeRows.map((row) => [row.find, 0]));
      const documentPath = findZipPath(zip, "word/document.xml");
      if (!documentPath) throw new Error("Template is missing word/document.xml");
      log(`Main document path: ${documentPath}`);

      const xmlFiles = Object.keys(zip.files).filter((path) =>
        /^word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$/.test(normalizeZipPath(path))
      );
      log(`Word XML files to process: ${xmlFiles.length}`);

      for (const path of xmlFiles) {
        log(`Processing ${path}`);
        const file = zip.file(path);
        if (!file) continue;

        let xml = await file.async("string");
        log(`Loaded ${path}: ${xml.length} chars`);
        const textResult = replaceXmlText(xml, replacements);
        xml = textResult.xml;
        addCounts(counts, textResult.counts);
        log(`Text replacement counts for ${path}: ${formatCounts(textResult.counts)}`);

        if (path === documentPath && nextCellFields.length) {
          const tableResult = fillNextTableCells(xml, nextCellFields);
          xml = tableResult.xml;
          addCounts(counts, tableResult.counts);
          log(`Table fill counts: ${formatCounts(tableResult.counts)}`);
          const missedTableFields = nextCellFields.filter((field) => !tableResult.counts[field.find]);
          missedTableFields.forEach((field) => log(`Table label not found: "${field.find}"`));
        }

        zip.file(path, xml);
        log(`Saved ${path}`);
      }

      log("Generating final DOCX blob");
      const blob = await zip.generateAsync({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });

      log(`Blob ready: ${blob.size} bytes`);
      downloadBlob(blob, outputName);
      setStatus(summarizeCounts(counts, activeRows));
      log(`Download triggered: ${outputName}`);
      log(`Final counts: ${formatCounts(counts)}`);
    } catch (error) {
      console.error(error);
      setStatus("Generation failed");
      log(`Error: ${error.message}`);
    } finally {
      setIsGenerating(false);
      log("Export finished");
    }
  }

  function updateRow(index, key, value) {
    setRows((current) => current.map((row, rowIndex) => (rowIndex === index ? { ...row, [key]: value } : row)));
  }

  function addRow() {
    setRows((current) => [...current, { label: "Custom field", find: "", value: "", mode: "replace" }]);
  }

  function removeRow(index) {
    setRows((current) => current.filter((_, rowIndex) => rowIndex !== index));
  }

  function resetRows() {
    setRows(defaultRows);
    setEmployeeName("");
    setDesignation("");
    setFileName("{{name}}_appointment_letter.docx");
  }

  const preview = makePreview(generationRows, rows, employeeName, designation, outputName);

  return (
    <main className="app">
      <section className="hero">
        <div>
          <p className="eyebrow">Word template generator</p>
          <h1>Offer Letter Generator</h1>
          <p className="lede">Fill the fields, preview the exact data, and export a Word file from Sample.docx.</p>
        </div>
        <div className="template-status">{status}</div>
      </section>

      <section className="panel">
        <div className="panel-header">
          <div>
            <h2>Template Values</h2>
            <p>Name and designation/details fill every matching location in the document.</p>
          </div>
          <button type="button" className="secondary" onClick={addRow}>Add field</button>
        </div>

        <form onSubmit={generateDocument}>
          <div className="quick-fields">
            <label>
              Employee name
              <input value={employeeName} onChange={(event) => setEmployeeName(event.target.value)} placeholder="Akhil" autoComplete="name" />
            </label>
            <label>
              Designation / details
              <input value={designation} onChange={(event) => setDesignation(event.target.value)} placeholder="Software Developer" autoComplete="off" />
            </label>
          </div>

          <div className="workspace-grid">
            <section className="editor-column">
              <div className="field-grid">
                {rows.map((row, index) => (
                  <div className="field-row" key={`${row.label}-${index}`}>
                    <label className="field-label">
                      Label
                      <input value={row.label} onChange={(event) => updateRow(index, "label", event.target.value)} autoComplete="off" />
                    </label>
                    <label className="field-find">
                      Find in template
                      <input value={row.find} onChange={(event) => updateRow(index, "find", event.target.value)} autoComplete="off" />
                    </label>
                    <label className="field-value">
                      Replacement
                      <input value={row.value} onChange={(event) => updateRow(index, "value", event.target.value)} autoComplete="off" />
                    </label>
                    <label className="field-mode">
                      Mode
                      <select value={row.mode || "replace"} onChange={(event) => updateRow(index, "mode", event.target.value)}>
                        <option value="replace">Replace text</option>
                        <option value="nextCell">Fill next cell</option>
                      </select>
                    </label>
                    <button type="button" className="icon-button" onClick={() => removeRow(index)} aria-label="Remove field">x</button>
                  </div>
                ))}
              </div>
            </section>

            <aside className="preview-column">
              <DataPreview preview={preview} rows={activeRows} />
              <ExportLogs logs={logs} />
            </aside>
          </div>

          <label className="file-name">
            Output file name
            <input value={fileName} onChange={(event) => setFileName(event.target.value)} autoComplete="off" />
          </label>

          <div className="actions">
            <button type="submit" disabled={!templateBytes || isGenerating}>{isGenerating ? "Generating..." : "Generate Word File"}</button>
            <button type="button" className="secondary" onClick={resetRows}>Reset</button>
          </div>
        </form>
      </section>
    </main>
  );
}

function ExportLogs({ logs }) {
  return (
    <section className="log-panel">
      <div className="panel-header compact">
        <div>
          <h2>Export Logs</h2>
          <p>These messages show exactly where export is spending time.</p>
        </div>
      </div>
      <pre>{logs.length ? logs.join("\n") : "No export run yet."}</pre>
    </section>
  );
}

function DataPreview({ preview, rows }) {
  return (
    <section className="preview-panel" aria-live="polite">
      <div className="panel-header compact">
        <div>
          <h2>Data Preview</h2>
          <p>This is the data the export will use.</p>
        </div>
      </div>
      <div className="preview-summary">
        <div><span>Output file</span><strong>{preview.outputName}</strong></div>
        <div><span>Employee name source</span><strong>{preview.employeeName || "Not set"}</strong></div>
        <div><span>Designation/details source</span><strong>{preview.designation || "Not set"}</strong></div>
        <div><span>Body line preview</span><strong>{preview.documentLine}</strong></div>
        <div><span>Annexure C Name</span><strong>{preview.annexureName || "Not set"}</strong></div>
        <div><span>Annexure C Designation</span><strong>{preview.annexureDesignation || "Not set"}</strong></div>
      </div>
      <div className="preview-table-wrap">
        <table className="preview-table">
          <thead>
            <tr>
              <th>Field</th>
              <th>Mode</th>
              <th>Find</th>
              <th>Value</th>
            </tr>
          </thead>
          <tbody>
            {rows.length ? rows.map((row, index) => (
              <tr key={`${row.label}-${index}`}>
                <td>{row.label}</td>
                <td>{row.mode === "nextCell" ? "Fill next cell" : "Replace text"}</td>
                <td>{row.find}</td>
                <td>{row.value}</td>
              </tr>
            )) : (
              <tr><td colSpan={4}>No export values set yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function getRowsForGeneration(rows, employeeName, designation) {
  const name = getEmployeeName(rows, employeeName);
  const role = getDesignation(rows, designation);
  return rows.map((row) => {
    if (normalizeText(row.label) === "employee name" && name) return { ...row, value: name };
    if (normalizeText(row.label) === "annexure c name" && name) return { ...row, value: name };
    if (normalizeText(row.label) === "employee detail" && role) return { ...row, value: role };
    if (normalizeText(row.label) === "designation in agreement" && role) return { ...row, value: role };
    if (normalizeText(row.label) === "annexure c designation" && role) return { ...row, value: role };
    return row;
  });
}

function getEmployeeName(rows, employeeName) {
  if (employeeName.trim()) return employeeName.trim();
  const nameRow = rows.find((row) => normalizeText(row.label) === "employee name" && row.value.trim())
    || rows.find((row) => normalizeText(row.label) === "annexure c name" && row.value.trim())
    || rows.find((row) => row.find.trim() === "ABC" && row.value.trim());
  return nameRow?.value?.trim() || "";
}

function getDesignation(rows, designation) {
  if (designation.trim()) return designation.trim();
  const designationRow = rows.find((row) => normalizeText(row.label) === "annexure c designation" && row.value.trim())
    || rows.find((row) => normalizeText(row.label) === "designation in agreement" && row.value.trim())
    || rows.find((row) => normalizeText(row.label) === "employee detail" && row.value.trim())
    || rows.find((row) => row.find.trim() === "XVZ" && row.value.trim())
    || rows.find((row) => row.find.trim() === "Senior Mistress" && row.value.trim());
  return designationRow?.value?.trim() || "";
}

function makePreview(generationRows, rows, employeeName, designation, outputName) {
  const name = getEmployeeName(rows, employeeName);
  const role = getDesignation(rows, designation);
  const bodyName = generationRows.find((row) => normalizeText(row.label) === "employee name")?.value || "";
  const annexureName = generationRows.find((row) => normalizeText(row.label) === "annexure c name")?.value || "";
  const employeeDetail = generationRows.find((row) => normalizeText(row.label) === "employee detail")?.value || "XVZ";
  const annexureDesignation = generationRows.find((row) => normalizeText(row.label) === "annexure c designation")?.value || "";

  return {
    outputName,
    employeeName: name,
    designation: role,
    documentLine: `${bodyName || "ABC"} - ${employeeDetail} ("Employee")`,
    annexureName,
    annexureDesignation
  };
}

function replaceXmlText(xml, replacements) {
  const directResult = replaceExactTextNodes(xml, replacements);
  xml = directResult.xml;

  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("Could not parse Word XML");

  const textNodes = Array.from(doc.getElementsByTagNameNS("*", "t"));
  const counts = {};

  for (const replacement of replacements) {
    if (replacement.find === replacement.value) {
      counts[replacement.find] = directResult.counts[replacement.find] || 0;
      continue;
    }

    counts[replacement.find] = directResult.counts[replacement.find] || 0;
    let found = true;
    while (found) {
      const map = buildTextMap(textNodes);
      const start = map.text.indexOf(replacement.find);
      if (start === -1) {
        found = false;
        continue;
      }
      replaceRange(map.parts, start, start + replacement.find.length, replacement.value);
      counts[replacement.find] += 1;
    }
  }

  return { xml: new XMLSerializer().serializeToString(doc), counts };
}

function replaceExactTextNodes(xml, replacements) {
  const counts = Object.fromEntries(replacements.map((replacement) => [replacement.find, 0]));

  for (const replacement of replacements) {
    const escapedFind = escapeXmlText(replacement.find);
    const escapedValue = escapeXmlText(replacement.value);
    const pattern = new RegExp(`(<w:t(?:\\s+[^>]*)?>)${escapeRegExp(escapedFind)}(<\\/w:t>)`, "g");
    xml = xml.replace(pattern, (...args) => {
      counts[replacement.find] += 1;
      return `${args[1]}${escapedValue}${args[2]}`;
    });
  }

  return { xml, counts };
}

function fillNextTableCells(xml, fields) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("Could not parse Word XML");

  const counts = Object.fromEntries(fields.map((field) => [field.find, 0]));
  const rows = Array.from(doc.getElementsByTagNameNS("*", "tr"));

  rows.forEach((row) => {
    const cells = Array.from(row.getElementsByTagNameNS("*", "tc"));
    cells.forEach((cell, index) => {
      const cellText = getCellText(cell);
      const field = fields.find((item) => normalizeFieldLabel(cellText) === normalizeFieldLabel(item.find));
      const nextCell = cells[index + 1];
      if (!field || !nextCell) return;
      setCellText(doc, nextCell, field.value);
      counts[field.find] += 1;
    });
  });

  return { xml: new XMLSerializer().serializeToString(doc), counts };
}

function getCellText(cell) {
  return Array.from(cell.getElementsByTagNameNS("*", "t")).map((node) => node.textContent || "").join("");
}

function setCellText(doc, cell, value) {
  const textNodes = Array.from(cell.getElementsByTagNameNS("*", "t"));
  if (textNodes.length) {
    textNodes[0].textContent = value;
    textNodes.slice(1).forEach((node) => { node.textContent = ""; });
    return;
  }

  const paragraph = cell.getElementsByTagNameNS("*", "p")[0] || createWordElement(doc, "p");
  if (!paragraph.parentNode) cell.appendChild(paragraph);
  const run = createWordElement(doc, "r");
  const text = createWordElement(doc, "t");
  text.textContent = value;
  run.appendChild(text);
  paragraph.appendChild(run);
}

function createWordElement(doc, tagName) {
  return doc.createElementNS("http://schemas.openxmlformats.org/wordprocessingml/2006/main", `w:${tagName}`);
}

function buildTextMap(nodes) {
  let cursor = 0;
  const parts = [];
  const text = nodes.map((node) => {
    const value = node.textContent || "";
    parts.push({ node, start: cursor, end: cursor + value.length });
    cursor += value.length;
    return value;
  }).join("");
  return { text, parts };
}

function replaceRange(parts, start, end, value) {
  const touched = parts.filter((part) => part.end > start && part.start < end);
  if (!touched.length) return;

  const first = touched[0];
  const last = touched[touched.length - 1];
  const firstText = first.node.textContent || "";
  const lastText = last.node.textContent || "";
  const prefix = firstText.slice(0, start - first.start);
  const suffix = lastText.slice(end - last.start);

  first.node.textContent = prefix + value + suffix;
  for (let index = 1; index < touched.length; index += 1) touched[index].node.textContent = "";
}

function findZipPath(zip, wantedPath) {
  const normalizedWanted = normalizeZipPath(wantedPath);
  return Object.keys(zip.files).find((path) => normalizeZipPath(path) === normalizedWanted);
}

function normalizeZipPath(path) {
  return path.toLowerCase().replaceAll("\\", "/");
}

function sanitizeReplacement(value) {
  return String(value)
    .trim()
    .replace(/[\u2018\u2019\u02bc\u2032]/g, "'")
    .replace(/[\u201c\u201d]/g, '"')
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
}

function normalizeText(value) {
  return String(value).replace(/\s+/g, " ").trim().toLowerCase();
}

function normalizeFieldLabel(value) {
  return normalizeText(value)
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .replace(/\s*:\s*$/, ":")
    .replace(/\s*\(\s*/g, "(")
    .replace(/\s*\)\s*/g, ")")
    .trim();
}

function escapeXmlText(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function renderFileName(pattern, employeeName) {
  return pattern.replaceAll("{{name}}", slugFilePart(employeeName || "employee"));
}

function slugFilePart(value) {
  return value.trim().replace(/[\\/:*?"<>|]+/g, "").replace(/\s+/g, "_").toLowerCase();
}

function normalizeDocxName(name) {
  const cleaned = name.trim() || "Appointment Letter.docx";
  return cleaned.toLowerCase().endsWith(".docx") ? cleaned : `${cleaned}.docx`;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function summarizeCounts(counts, rows) {
  const total = Object.values(counts).reduce((sum, count) => sum + count, 0);
  const missed = rows.filter((row) => !counts[row.find]).length;
  if (total === 0) return "No matching text found";
  if (missed) return `Generated file. ${total} replacements made, ${missed} field(s) not found.`;
  return `Generated file. ${total} replacements made.`;
}

function addCounts(target, source) {
  Object.entries(source).forEach(([key, count]) => {
    target[key] = (target[key] || 0) + count;
  });
}

function formatCounts(counts) {
  const entries = Object.entries(counts).filter(([, count]) => count);
  if (!entries.length) return "none";
  return entries.map(([key, count]) => `${key}: ${count}`).join(", ");
}
