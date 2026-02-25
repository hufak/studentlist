import { useEffect, useMemo, useState } from "react";
import type { ChangeEvent } from "react";
import type { Worksheet } from "exceljs";

const DEFAULT_FILTER_COLUMN = "Studium";
const FILTERABLE_COLUMNS = ["Hörerstatus", "Studium"];
const PRESUMED_FEE_STATUS_HEADER =
  "presumed fee status (based on OECD 2025 list)";
const FEE_STATUS_LABELS: Record<string, string> = {
  exempt: "exempt (until exceeding tolerance semesters)",
  refund: "double fee, full refund",
  partial: "double fee, 50% refund",
  double: "double fee",
};
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

function shouldAnnotateCountryColumn(header: string) {
  const normalized = header.trim().toLowerCase();
  return normalized === "nationalität" || normalized.endsWith("land");
}

function isNationalityColumn(header: string) {
  return header.trim().toLowerCase() === "nationalität";
}

function normalizeCountryCode(value: Cell) {
  return String(value ?? "").trim().toUpperCase();
}

function toPrettyFeeStatus(value: string) {
  const normalized = value.trim().toLowerCase();
  return FEE_STATUS_LABELS[normalized] ?? value;
}

function getBaseUrl() {
  return import.meta.env.BASE_URL.endsWith("/")
    ? import.meta.env.BASE_URL
    : `${import.meta.env.BASE_URL}/`;
}

function formatStudyPrettyName(name: string, level: string) {
  const cleanName = name.trim();
  const cleanLevel = level.trim();
  if (!cleanName || !cleanLevel) return "";
  return `${cleanName} (${cleanLevel})`;
}

function mapStudiumText(value: Cell, lookup: Map<string, string>) {
  const raw = String(value ?? "");
  if (!raw.trim()) return raw;
  return raw
    .split(",")
    .map((part) => {
      const key = part.trim();
      if (!key) return "";
      return lookup.get(key) ?? key;
    })
    .filter((part) => part.length > 0)
    .join(", ");
}

function mapStudiumTextWithMeta(value: Cell, lookup: Map<string, string>) {
  const raw = String(value ?? "");
  if (!raw.trim()) {
    return { text: raw, usedLookup: false };
  }

  let usedLookup = false;
  const text = raw
    .split(",")
    .map((part) => {
      const key = part.trim();
      if (!key) return "";
      if (lookup.has(key)) {
        usedLookup = true;
      }
      return lookup.get(key) ?? key;
    })
    .filter((part) => part.length > 0)
    .join(", ");

  return { text, usedLookup };
}

function parseCsvRows(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let value = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          value += '"';
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        value += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }
    if (ch === ",") {
      row.push(value);
      value = "";
      continue;
    }
    if (ch === "\n") {
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }
    if (ch !== "\r") {
      value += ch;
    }
  }

  row.push(value);
  rows.push(row);
  return rows.filter((entry) => entry.some((cell) => String(cell).length > 0));
}

function toPrimitiveCellValue(value: unknown, fallbackText: string): Cell {
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return value;
  }
  if (value instanceof Date) return fallbackText || value.toISOString();
  if (value && typeof value === "object" && "result" in value) {
    const result = (value as { result?: unknown }).result;
    if (
      typeof result === "string" ||
      typeof result === "number" ||
      typeof result === "boolean"
    ) {
      return result;
    }
  }
  return fallbackText;
}

function setWorksheetCellItalic(
  worksheet: Worksheet,
  rowIndex1: number,
  columnIndex1: number,
) {
  const cell = worksheet.getCell(rowIndex1, columnIndex1);
  cell.font = {
    ...(cell.font ?? {}),
    italic: true,
  };
}

let excelJsPromise: Promise<typeof import("exceljs")> | null = null;

async function loadExcelJs() {
  if (!excelJsPromise) {
    excelJsPromise = import("exceljs");
  }
  return excelJsPromise;
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
  const [studyNameLookup, setStudyNameLookup] = useState<Map<string, string>>(
    () => new Map(),
  );

  useEffect(() => {
    let active = true;
    async function loadStudyNames() {
      try {
        const response = await fetch(`${getBaseUrl()}study_names.csv`, {
          cache: "no-store",
        });
        if (!response.ok) return;

        const csvText = await response.text();
        const csvRows = parseCsvRows(csvText);

        const [headerRow, ...dataRows] = csvRows;
        const csvHeader = (headerRow ?? []).map((value) =>
          String(value ?? "").trim().toLowerCase(),
        );
        const dnameIndex = csvHeader.indexOf("dname");
        const nameIndex = csvHeader.indexOf("name");
        const levelIndex = csvHeader.indexOf("level");
        if (dnameIndex < 0 || nameIndex < 0 || levelIndex < 0) return;

        const nextLookup = new Map<string, string>();
        dataRows.forEach((row) => {
          const dname = String(row[dnameIndex] ?? "").trim();
          const name = String(row[nameIndex] ?? "").trim();
          const level = String(row[levelIndex] ?? "").trim();
          const pretty = formatStudyPrettyName(name, level);
          if (dname && pretty) {
            nextLookup.set(dname, pretty);
          }
        });

        if (active) {
          setStudyNameLookup(nextLookup);
        }
      } catch {
        if (active) {
          setStudyNameLookup(new Map());
        }
      }
    }

    loadStudyNames();
    return () => {
      active = false;
    };
  }, []);

  const columnIndex = useMemo(() => {
    if (!filterColumn) return -1;
    return headers.indexOf(filterColumn);
  }, [headers, filterColumn]);

  const studiumOptions = useMemo(() => {
    if (columnIndex < 0) return [];
    const values = unique(rows.map((row) => String(row[columnIndex] ?? "")));
    if (filterColumn !== "Studium") {
      return values.sort();
    }
    return values.sort((left, right) => {
      const prettyLeft = mapStudiumText(left, studyNameLookup);
      const prettyRight = mapStudiumText(right, studyNameLookup);
      const byPretty = prettyLeft.localeCompare(prettyRight, "de", {
        sensitivity: "base",
      });
      if (byPretty !== 0) return byPretty;
      return left.localeCompare(right, "de", { sensitivity: "base" });
    });
  }, [rows, columnIndex, filterColumn, studyNameLookup]);

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
  const availableFilterColumns = useMemo(() => new Set(headers), [headers]);

  const trimmedRows = useMemo(() => {
    if (selectedColumnIndexes.length === 0) return [];
    return filteredRows.map((row) =>
      selectedColumnIndexes.map((item) => row[item.index] ?? ""),
    );
  }, [filteredRows, selectedColumnIndexes]);

  const uniqueTrimmedRows = useMemo(() => {
    if (selectedColumnIndexes.length === 0 || filteredRows.length === 0) return [];
    const studiumFullIndex = headers.findIndex((header) => header === "Studium");
    const merged = new Map<string, { row: Row; values: Set<string> }>();
    const order: string[] = [];
    filteredRows.forEach((row) => {
      const keyParts = headers.map((_, index) => {
        if (index === studiumFullIndex) return "";
        return String(row[index] ?? "");
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
        entry.values.add(
          studiumFullIndex >= 0 ? String(row[studiumFullIndex] ?? "").trim() : "",
        );
      }
    });

    return order.map((key) => {
      const entry = merged.get(key)!;
      const mergedStudium = Array.from(entry.values)
        .filter((value) => value.length)
        .join(", ");
      return selectedColumnIndexes.map((item) => {
        if (item.header === "Studium" && studiumFullIndex >= 0) {
          return mergedStudium;
        }
        return entry.row[item.index] ?? "";
      });
    });
  }, [selectedColumnIndexes, filteredRows, headers]);

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
      const { default: ExcelJS } = await loadExcelJs();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      const worksheet = workbook.worksheets[0];
      if (!worksheet) {
        throw new Error("The sheet is empty.");
      }
      const data: Cell[][] = [];
      worksheet.eachRow({ includeEmpty: false }, (excelRow) => {
        const rowValues: Cell[] = [];
        for (let colIndex = 1; colIndex <= excelRow.cellCount; colIndex += 1) {
          const cell = excelRow.getCell(colIndex);
          rowValues.push(toPrimitiveCellValue(cell.value, cell.text ?? ""));
        }
        data.push(rowValues);
      });

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

  async function handleDownload() {
    if (selectedColumnIndexes.length === 0) return;
    setError("");

    let countryLookup = new Map<string, string>();
    let feeLookup = new Map<string, string>();
    let shouldAugmentCountryData = true;
    try {
      const response = await fetch(`${getBaseUrl()}countries.csv`, {
        cache: "no-store",
      });
      if (!response.ok) {
        throw new Error(`countries.csv could not be loaded (${response.status}).`);
      }
      const csvText = await response.text();
      const csvRows = parseCsvRows(csvText);

      const [csvHeaderRow, ...csvDataRows] = csvRows;
      const csvHeader = (csvHeaderRow ?? []).map((value) =>
        String(value ?? "").trim().toLowerCase(),
      );
      const codeIndex = csvHeader.indexOf("code");
      const nameIndex = csvHeader.indexOf("name");
      const feeIndex = csvHeader.indexOf("fee");

      csvDataRows.forEach((row) => {
        const code = normalizeCountryCode(row[codeIndex >= 0 ? codeIndex : 0]);
        const name = String(row[nameIndex >= 0 ? nameIndex : 1] ?? "").trim();
        const feeRaw = String(row[feeIndex >= 0 ? feeIndex : 4] ?? "").trim();
        const fee = toPrettyFeeStatus(feeRaw);
        if (code) {
          countryLookup.set(code, name || "?");
          feeLookup.set(code, fee || "?");
        }
      });
    } catch (err: unknown) {
      shouldAugmentCountryData = false;
      setError(
        err instanceof Error
          ? `${err.message} Export wird mit Originalwerten fortgesetzt.`
          : "countries.csv konnte nicht geladen werden. Export wird mit Originalwerten fortgesetzt.",
      );
    }

    const outputHeaders: string[] = [];
    const italicHeaderColumns = new Set<number>();
    selectedColumnIndexes.forEach((item) => {
      outputHeaders.push(item.header);
      if (shouldAugmentCountryData && isNationalityColumn(item.header)) {
        outputHeaders.push(PRESUMED_FEE_STATUS_HEADER);
        italicHeaderColumns.add(outputHeaders.length - 1);
      }
    });

    const countryAnnotatedCells = new Map<
      string,
      { originalValue: string; lookedUpName: string }
    >();
    const outputCellItalics: boolean[][] = [];
    const outputRows: Row[] = processedRows.map((row, rowIndex) => {
      const next: Row = [];
      const italicFlags: boolean[] = [];
      selectedColumnIndexes.forEach((item, columnIndex) => {
        const originalCell = row[columnIndex] ?? "";
        let renderedCell: Cell = originalCell;
        let shouldItalicCell = false;

        if (item.header === "Studium") {
          const mappedStudium = mapStudiumTextWithMeta(
            originalCell,
            studyNameLookup,
          );
          renderedCell = mappedStudium.text;
          shouldItalicCell = mappedStudium.usedLookup;
        } else if (
          shouldAugmentCountryData &&
          shouldAnnotateCountryColumn(item.header)
        ) {
          const originalValue = String(originalCell);
          const lookupCode = normalizeCountryCode(originalCell);
          const lookedUpName = lookupCode
            ? countryLookup.get(lookupCode)
            : undefined;
          renderedCell = `${originalValue} (${lookedUpName ?? "?"})`;
          const outputColumnIndex = next.length;
          countryAnnotatedCells.set(`${rowIndex}:${outputColumnIndex}`, {
            originalValue,
            lookedUpName: lookedUpName ?? "?",
          });
        }

        if (exportMode === "statistics" && item.header === "Matrikelnummer") {
          shouldItalicCell = true;
        }
        next.push(renderedCell);
        italicFlags.push(shouldItalicCell);

        if (shouldAugmentCountryData && isNationalityColumn(item.header)) {
          const lookupCode = normalizeCountryCode(originalCell);
          const feeStatus = lookupCode ? feeLookup.get(lookupCode) : undefined;
          next.push(feeStatus ?? "?");
          italicFlags.push(true);
        }
      });
      outputCellItalics.push(italicFlags);
      return next;
    });

    const output: Row[] = [outputHeaders, ...outputRows];
    const { default: ExcelJS } = await loadExcelJs();
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Filtered");
    worksheet.addRow(outputHeaders);
    outputRows.forEach((row) => {
      worksheet.addRow(row.map((cell) => (cell === null ? "" : cell)));
    });
    italicHeaderColumns.forEach((columnIndex) => {
      setWorksheetCellItalic(worksheet, 1, columnIndex + 1);
    });
    outputCellItalics.forEach((rowItalics, rowIndex) => {
      rowItalics.forEach((shouldItalicCell, columnIndex) => {
        if (shouldItalicCell) {
          setWorksheetCellItalic(worksheet, rowIndex + 2, columnIndex + 1);
        }
      });
    });
    countryAnnotatedCells.forEach((parts, key) => {
      const [rowIndexText, columnIndexText] = key.split(":");
      const rowIndex = Number(rowIndexText);
      const columnIndex = Number(columnIndexText);
      const cell = worksheet.getCell(rowIndex + 2, columnIndex + 1);
      cell.value = {
        richText: [
          { text: `${parts.originalValue} (` },
          { text: parts.lookedUpName, font: { italic: true } },
          { text: ")" },
        ],
      };
    });
    const columnWidths = buildColumnWidths(output);
    columnWidths.forEach((entry, index) => {
      worksheet.getColumn(index + 1).width = entry.wch;
    });
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: output.length, column: outputHeaders.length },
    };

    const baseName = fileName ? fileName.replace(/\.xlsx$/i, "") : "output";
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `${baseName}-filtered.xlsx`;
    link.click();
    URL.revokeObjectURL(link.href);
  }

  const canDownload =
    headers.length > 0 &&
    selectedColumnIndexes.length > 0 &&
    (filterColumn === "" || selectedStudium.length > 0);
  const missingFilterSelection =
    headers.length > 0 && filterColumn !== "" && selectedStudium.length === 0;

  return (
    <div className="mx-auto flex w-full max-w-[1100px] flex-col gap-8 px-6 pb-20 pt-12">
      <header className="grid items-start gap-8 lg:grid-cols-[1.2fr_0.8fr]">
        <div>
          <p className="text-xs font-semibold uppercase tracking-[0.2em] text-accent-strong">
            Alles lokal im Browser
          </p>
          <h1 className="mb-3 mt-2 text-[42px] font-semibold leading-tight">
            Studierendenliste filtern
          </h1>
          <p className="max-w-[520px] text-lg leading-relaxed text-muted">
            Excel-Datei hochladen, gewünschte Spalten behalten, nach Studium
            filtern und die bereinigte Datei erzeugen. Alles bleibt auf dem
            lokalen Computer.
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
              <span className="block font-semibold">Excel-Datei laden</span>
              <span className="text-sm text-muted">
                {fileName || "Keine Datei geladen"}
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
                // Columns that are usually not needed stay muted, while the label text signals caution.
                <label
                  key={header}
                  className="flex items-center gap-2.5 rounded-full border border-border bg-[#fff8f2] px-3.5 py-2.5 text-sm"
                >
                  <input
                    type="checkbox"
                    checked={selectedColumns.includes(header)}
                    onChange={() => toggleColumn(header)}
                    className="accent-accent"
                  />
                  <span
                    className={
                      dimmedLookup.has(header.toLowerCase())
                        ? "text-warning-strong"
                        : "text-text font-semibold"
                    }
                  >
                    {header}
                  </span>
                </label>
              ))}
            </div>
          </section>

          <section className="rounded-[20px] border border-border bg-surface p-6 shadow-panel">
            <div className="mb-4 flex items-center justify-between gap-4">
              <h2 className="text-[22px] font-semibold">Studium-Filter</h2>
              <div className="flex items-center gap-2 text-sm">
                {["", "Studium", "Hörerstatus"].map((value) => {
                  const isActive = filterColumn === value;
                  const isAvailable =
                    value === "" || availableFilterColumns.has(value);
                  const label =
                    value === ""
                      ? "alle Studierenden"
                      : value === "Studium"
                        ? "nach Studium"
                        : "nach Hörer*status";
                  return (
                    <button
                      key={value || "none"}
                      type="button"
                      disabled={!isAvailable}
                      onClick={() => {
                        setFilterColumn(value);
                        setSelectedStudium([]);
                      }}
                      className={`rounded-full border px-3 py-1.5 font-mono text-xs uppercase tracking-[0.08em] transition ${
                        isActive
                          ? "border-accent bg-accent text-white"
                          : isAvailable
                            ? "border-border bg-[#fff8f2] text-muted hover:border-accent hover:text-accent-strong"
                            : "cursor-not-allowed border-border bg-[#f1e8de] text-[#9a8a78]"
                      }`}
                    >
                      {label}
                    </button>
                  );
                })}
              </div>
            </div>
            {filterColumn === "" ? (
              <p className="text-sm text-muted">
                Keine Filterspalte ausgewählt, somit werden die Daten aller Studierenden exportiert. Wähle "Studium" oder
                "Hörer*status", um Einträge zu filtern.
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
                      {filterColumn === "Studium"
                        ? mapStudiumText(value, studyNameLookup) || "(leer)"
                        : String(value) || "(leer)"}
                    </span>
                    <span className="rounded-full bg-bgAccent px-2.5 py-1 font-mono text-xs text-accent-strong">
                      {studiumCounts.get(value) ?? 0}
                    </span>
                  </label>
                ))}
              </div>
            )}
          </section>

          <section className="flex flex-col gap-6 rounded-[20px] border border-border bg-surface p-6 shadow-panel lg:flex-row lg:items-start lg:justify-between">
            <div>
              <h2 className="text-[22px] font-semibold">Export</h2>
              <p className="text-sm text-muted">
                {missingFilterSelection
                  ? "Bitte wähle mindestens eine Filterkategorie zum Exportieren"
                  : `${processedRows.length} Zeilen nach Filterung und Spaltenauswahl.`}
              </p>
            </div>
            <div className="flex w-full flex-col gap-4 md:w-[60%] lg:w-[45%]">
              <button
                type="button"
                onClick={handleDownload}
                disabled={!canDownload}
                className="rounded-full bg-accent px-5 py-3 font-mono text-sm uppercase tracking-[0.12em] text-white transition hover:-translate-y-0.5 hover:bg-accent-strong disabled:cursor-not-allowed disabled:bg-[#d5b5ad] disabled:hover:translate-y-0"
              >
                Gefilterte Excel-Datei herunterladen
              </button>
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
            </div>
          </section>
        </main>
      )}
    </div>
  );
}
