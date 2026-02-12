import { useMemo, useState } from "react";
import type { ChangeEvent } from "react";
import * as XLSX from "xlsx";

const DEFAULT_FILTER_COLUMN = "Studium";
const FILTERABLE_COLUMNS = ["Hörerstatus", "Studium"];
const DIMMED_COLUMNS = [
  "Titel vor",
  "Titel nach",
  "Geburtsdatum",
  "Heimat_PLZ",
  "Heimat_Ort",
  "Heimat_Straße",
  "Studien_PLZ",
  "Studien_Ort",
  "Studien_Straße",
];

type Cell = string | number | boolean | null;
type Row = Cell[];

function unique<T>(values: T[]): T[] {
  return Array.from(new Set(values));
}

function generateShortId() {
  const bytes = new Uint8Array(4);
  crypto.getRandomValues(bytes);
  let value = 0;
  bytes.forEach((byte) => {
    value = (value << 8) | byte;
  });
  return value.toString(36).padStart(7, "0").slice(0, 7);
}

function toCellText(cell: Cell) {
  if (cell === null || cell === undefined) return "";
  return String(cell);
}

function buildColumnWidths(rows: Row[]) {
  if (rows.length === 0) return [];
  const columnCount = rows[0]?.length ?? 0;
  return Array.from({ length: columnCount }, (_, columnIndex) => {
    const maxLength = rows.reduce((longest, row) => {
      const length = toCellText(row[columnIndex] ?? "").length;
      return Math.max(longest, length);
    }, 0);
    return { wch: Math.max(10, Math.min(maxLength + 2, 60)) };
  });
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [rows, setRows] = useState<Row[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [filterColumn, setFilterColumn] = useState("");
  const [selectedStudium, setSelectedStudium] = useState<string[]>([]);
  const [error, setError] = useState("");
  const [exportMode, setExportMode] = useState<"student" | "statistics">(
    "student",
  );

  const columnIndex = useMemo(() => {
    if (!filterColumn) return -1;
    return headers.indexOf(filterColumn);
  }, [headers, filterColumn]);

  const studiumOptions = useMemo(() => {
    if (columnIndex < 0) return [];
    return unique(rows.map((row) => String(row[columnIndex] ?? ""))).sort();
  }, [rows, columnIndex]);

  const studiumCounts = useMemo(() => {
    if (columnIndex < 0) return new Map<string, number>();
    const counts = new Map<string, number>();
    rows.forEach((row) => {
      const value = String(row[columnIndex] ?? "");
      counts.set(value, (counts.get(value) ?? 0) + 1);
    });
    return counts;
  }, [rows, columnIndex]);

  const filteredRows = useMemo(() => {
    if (rows.length === 0) return [];
    if (columnIndex < 0 || selectedStudium.length === 0) return rows;
    const allowed = new Set(selectedStudium);
    return rows.filter((row) => allowed.has(String(row[columnIndex] ?? "")));
  }, [rows, columnIndex, selectedStudium]);

  const selectedColumnIndexes = useMemo(() => {
    const set = new Set(selectedColumns);
    return headers
      .map((header, index) => ({ header, index }))
      .filter((item) => set.has(item.header));
  }, [headers, selectedColumns]);

  const dimmedLookup = useMemo(
    () => new Set(DIMMED_COLUMNS.map((name) => name.toLowerCase())),
    [],
  );

  const trimmedRows = useMemo(() => {
    if (selectedColumnIndexes.length === 0) return [];
    return filteredRows.map((row) =>
      selectedColumnIndexes.map((item) => row[item.index] ?? ""),
    );
  }, [filteredRows, selectedColumnIndexes]);

  const uniqueTrimmedRows = useMemo(() => {
    if (trimmedRows.length === 0) return [];
    const studiumIndex = selectedColumnIndexes.findIndex(
      (item) => item.header === "Studium",
    );
    if (studiumIndex < 0) {
      const seen = new Set<string>();
      const deduped: Row[] = [];
      trimmedRows.forEach((row) => {
        const key = JSON.stringify(row.map((cell) => String(cell ?? "")));
        if (!seen.has(key)) {
          seen.add(key);
          deduped.push(row);
        }
      });
      return deduped;
    }

    const merged = new Map<string, { row: Row; values: Set<string> }>();
    const order: string[] = [];
    trimmedRows.forEach((row) => {
      const keyParts = row.map((cell, index) => {
        if (index === studiumIndex) return "";
        return String(cell ?? "");
      });
      const key = JSON.stringify(keyParts);
      if (!merged.has(key)) {
        merged.set(key, {
          row: row.map((cell) => cell ?? ""),
          values: new Set(),
        });
        order.push(key);
      }
      const entry = merged.get(key);
      if (entry) {
        entry.values.add(String(row[studiumIndex] ?? "").trim());
      }
    });

    return order.map((key) => {
      const entry = merged.get(key)!;
      const values = Array.from(entry.values).filter((value) => value.length);
      entry.row[studiumIndex] = values.join(", ");
      return entry.row;
    });
  }, [trimmedRows, selectedColumnIndexes]);

  const processedRows = useMemo(() => {
    const baseRows = exportMode === "student" ? uniqueTrimmedRows : trimmedRows;
    if (exportMode !== "statistics") return baseRows;

    const matrikelnummerIndex = selectedColumnIndexes.findIndex(
      (item) => item.header === "Matrikelnummer",
    );
    if (matrikelnummerIndex < 0) return baseRows;

    const ids = new Map<string, string>();
    const usedIds = new Set<string>();

    return baseRows.map((row) => {
      const original = String(row[matrikelnummerIndex] ?? "").trim();
      if (!original) return row;

      let id = ids.get(original);
      if (!id) {
        do {
          id = generateShortId();
        } while (usedIds.has(id));
        ids.set(original, id);
        usedIds.add(id);
      }

      const next = row.slice();
      next[matrikelnummerIndex] = id;
      return next;
    });
  }, [exportMode, uniqueTrimmedRows, trimmedRows, selectedColumnIndexes]);

  function resetState() {
    setHeaders([]);
    setRows([]);
    setSelectedColumns([]);
    setFilterColumn("");
    setSelectedStudium([]);
  }

  async function handleFileChange(event: ChangeEvent<HTMLInputElement>) {
    const file = event.currentTarget.files?.[0];
    if (!file) return;
    setError("");

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        blankrows: false,
      }) as Cell[][];

      if (!data || data.length === 0) {
        throw new Error("The sheet is empty.");
      }

      const nonEmptyCounts = data.map(
        (row) => row.filter((cell) => String(cell ?? "").trim()).length,
      );
      const maxNonEmpty = Math.max(...nonEmptyCounts);
      if (maxNonEmpty === 0) {
        throw new Error("No populated rows were found.");
      }

      const startIndex = nonEmptyCounts.findIndex(
        (count) => count === maxNonEmpty,
      );
      const trimmedData = data.slice(startIndex);
      const [headerRow, ...dataRows] = trimmedData;
      const maxColumns = headerRow.length;
      const cleanHeaders = headerRow
        .slice(0, maxColumns)
        .map((value, index) => {
          const text = String(value ?? "").trim();
          return text.length > 0 ? text : `Column ${index + 1}`;
        });

      setFileName(file.name);
      setHeaders(cleanHeaders);
      setRows(dataRows as Row[]);
      const preferredColumns = [
        // "Matrikelnummer",
        "Vorname",
        "Zuname",
        // "Nationalität",
        // "Heimat_Land",
        // "Studien_Land",
        // "Hörerstatus",
        "Studium",
        "Email",
      ];
      const preferredLookup = new Set(
        preferredColumns.map((name) => name.toLowerCase()),
      );
      const defaultSelected = cleanHeaders.filter((header) =>
        preferredLookup.has(header.toLowerCase()),
      );
      setSelectedColumns(defaultSelected.length > 0 ? defaultSelected : []);

      const defaultFilter = cleanHeaders.includes(DEFAULT_FILTER_COLUMN)
        ? DEFAULT_FILTER_COLUMN
        : FILTERABLE_COLUMNS.find((column) => cleanHeaders.includes(column)) ||
          "";
      setFilterColumn(defaultFilter);
      setSelectedStudium([]);
    } catch (err: unknown) {
      resetState();
      setFileName("");
      setError(err instanceof Error ? err.message : "Failed to read file.");
    }
  }

  function toggleColumn(header: string) {
    setSelectedColumns((current) => {
      const set = new Set(current);
      if (set.has(header)) {
        set.delete(header);
      } else {
        set.add(header);
      }
      return Array.from(set);
    });
  }

  function toggleStudium(value: string) {
    setSelectedStudium((current) => {
      const set = new Set(current);
      if (set.has(value)) {
        set.delete(value);
      } else {
        set.add(value);
      }
      return Array.from(set);
    });
  }

  function selectAllColumns(checked: boolean) {
    setSelectedColumns(checked ? headers : []);
  }

  function handleDownload() {
    if (selectedColumnIndexes.length === 0) return;
    const output: Row[] = [
      selectedColumnIndexes.map((item) => item.header),
      ...processedRows,
    ];

    const worksheet = XLSX.utils.aoa_to_sheet(output);
    worksheet["!cols"] = buildColumnWidths(output);
    worksheet["!autofilter"] = {
      ref: XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: output.length - 1, c: selectedColumnIndexes.length - 1 },
      }),
    };
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Filtered");

    const baseName = fileName ? fileName.replace(/\.xlsx$/i, "") : "output";
    XLSX.writeFile(workbook, `${baseName}-filtered.xlsx`);
  }

  const canDownload =
    headers.length > 0 &&
    selectedColumnIndexes.length > 0 &&
    (filterColumn === "" || selectedStudium.length > 0);

  return (
    <div className="mx-auto flex w-full max-w-[1100px] flex-col gap-8 px-6 pb-20 pt-12">
      <header className="grid items-start gap-8 lg:grid-cols-[1.2fr_0.8fr]">
        <div>
          <p className="text-xs font-semibold uppercase tracking-[0.2em] text-accent-strong">
            Nur im Browser
          </p>
          <h1 className="mb-3 mt-2 text-[42px] font-semibold leading-tight">
            Studierendenliste filtern
          </h1>
          <p className="max-w-[520px] text-lg leading-relaxed text-muted">
            Excel-Datei hochladen, gewünschte Spalten behalten, nach Studium
            filtern und die bereinigte Datei herunterladen. Alles bleibt in
            deinem Browser.
          </p>
        </div>
        <div className="flex flex-col gap-3">
          <label className="flex cursor-pointer items-center gap-4 rounded-2xl border border-border bg-surface p-5 shadow-card transition hover:-translate-y-0.5 hover:shadow-[0_18px_50px_rgba(54,35,16,0.2)]">
            <input
              type="file"
              accept=".xlsx"
              onChange={handleFileChange}
              className="hidden"
            />
            <div>
              <span className="block font-semibold">Excel-Datei auswählen</span>
              <span className="text-sm text-muted">
                {fileName || "Keine Datei ausgewählt"}
              </span>
            </div>
          </label>
          {error && <p className="font-semibold text-accent-strong">{error}</p>}
        </div>
      </header>

      {headers.length > 0 && (
        <main className="grid gap-6">
          <section className="rounded-[20px] border border-border bg-surface p-6 shadow-panel">
            <div className="mb-4 flex items-center justify-between gap-4">
              <h2 className="text-[22px] font-semibold">Spalten</h2>
              <label className="flex items-center gap-2 text-sm text-muted">
                <input
                  type="checkbox"
                  checked={selectedColumns.length === headers.length}
                  onChange={(event) => selectAllColumns(event.target.checked)}
                  className="accent-accent"
                />
                <span>Alle Spalten</span>
              </label>
            </div>
            <div className="grid gap-3 [grid-template-columns:repeat(auto-fit,minmax(160px,1fr))]">
              {headers.map((header) => (
                <label
                  key={header}
                  className={`flex items-center gap-2.5 rounded-full border border-border bg-[#fff8f2] px-3.5 py-2.5 text-sm${
                    dimmedLookup.has(header.toLowerCase()) ? " opacity-55" : ""
                  }`}
                >
                  <input
                    type="checkbox"
                    checked={selectedColumns.includes(header)}
                    onChange={() => toggleColumn(header)}
                    className="accent-accent"
                  />
                  <span>{header}</span>
                </label>
              ))}
            </div>
          </section>

          <section className="rounded-[20px] border border-border bg-surface p-6 shadow-panel">
            <div className="mb-4 flex items-center justify-between gap-4">
              <h2 className="text-[22px] font-semibold">Studium-Filter</h2>
              <div className="flex items-center gap-3 text-sm text-muted">
                <span>Filterspalte</span>
                <select
                  value={filterColumn}
                  onChange={(event) => {
                    const value = event.target.value;
                    setFilterColumn(value);
                    setSelectedStudium([]);
                  }}
                  className="rounded-xl border border-border px-3 py-2 font-sans text-sm"
                >
                  <option value="">Kein Filter</option>
                  {headers
                    .filter((header) => FILTERABLE_COLUMNS.includes(header))
                    .map((header) => (
                      <option key={header} value={header}>
                        {header}
                      </option>
                    ))}
                </select>
              </div>
            </div>
            {filterColumn === "" ? (
              <p className="text-sm text-muted">
                Keine Filterspalte ausgewählt. Wähle "Studium" oder
                "Hörerstatus", um Zeilen zu filtern.
              </p>
            ) : (
              <div className="flex flex-col gap-2.5">
                {studiumOptions.map((value) => (
                  <label
                    key={String(value)}
                    className="grid grid-cols-[20px_1fr_auto] items-center gap-3 rounded-[14px] border border-border bg-[#fff8f2] px-3.5 py-2.5 text-sm"
                  >
                    <input
                      type="checkbox"
                      checked={selectedStudium.includes(value)}
                      onChange={() => toggleStudium(value)}
                      className="accent-accent"
                    />
                    <span className="truncate">
                      {String(value) || "(leer)"}
                    </span>
                    <span className="rounded-full bg-bgAccent px-2.5 py-1 font-mono text-xs text-accent-strong">
                      {studiumCounts.get(value) ?? 0}
                    </span>
                  </label>
                ))}
              </div>
            )}
          </section>

          <section className="flex flex-col gap-6 rounded-[20px] border border-border bg-surface p-6 shadow-panel lg:flex-row lg:items-center lg:justify-between">
            <div>
              <h2 className="text-[22px] font-semibold">Vorschau</h2>
              <p className="text-sm text-muted">
                {processedRows.length} Zeilen nach Filterung und Spaltenauswahl.
              </p>
            </div>
            <div className="flex w-full flex-col gap-4 md:w-[60%] lg:w-[45%]">
              <div className="flex flex-col gap-1 text-sm text-muted">
                <label className="flex items-center gap-2">
                  <span>Exportmodus</span>
                  <select
                    value={exportMode}
                    onChange={(event) =>
                      setExportMode(
                        event.target.value === "statistics"
                          ? "statistics"
                          : "student",
                      )
                    }
                    className="rounded-xl border border-border px-3 py-2 font-sans text-sm"
                  >
                    <option value="student">Studierendenzentriert</option>
                    <option value="statistics">Statistikzentriert</option>
                  </select>
                </label>
                <span className="text-xs text-muted">
                  {exportMode === "student"
                    ? "Studierende mit Mehrfachstudien werden zu einem Eintrag zusammengeführt und Studium-Werte kombiniert."
                    : "Studierende mit Mehrfachstudien bleiben als separate Einträge erhalten; Matrikelnummer wird durch eine Zufalls-ID ersetzt."}
                </span>
              </div>
              <button
                type="button"
                onClick={handleDownload}
                disabled={!canDownload}
                className="rounded-full bg-accent px-5 py-3 font-mono text-sm uppercase tracking-[0.12em] text-white transition hover:-translate-y-0.5 hover:bg-accent-strong disabled:cursor-not-allowed disabled:bg-[#d5b5ad] disabled:hover:translate-y-0"
              >
                Gefilterte Excel-Datei herunterladen
              </button>
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
