<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import type { Row as ExcelRow, Worksheet } from 'exceljs';

const DEFAULT_FILTER_COLUMN = 'Studium';
const FILTERABLE_COLUMNS = ['Hörerstatus', 'Studium'];
const PRESUMED_FEE_STATUS_HEADER = 'presumed fee status (based on OECD 2025 list)';
const FEE_STATUS_LABELS: Record<string, string> = {
	exempt: 'exempt (until exceeding tolerance semesters)',
	refund: 'double fee, full refund',
	partial: 'double fee, 50% refund',
	double: 'double fee',
};
const DIMMED_COLUMNS = [
	'Titel vor',
	'Titel nach',
	'Geburtsdatum',
	'Heimat_PLZ',
	'Heimat_Ort',
	'Heimat_Straße',
	'Studien_PLZ',
	'Studien_Ort',
	'Studien_Straße',
];
const FILTER_MODES = [
	{ value: '', label: 'alle Studierenden' },
	{ value: 'Studium', label: 'nach Studium' },
	{ value: 'Hörerstatus', label: 'nach Hörer*status' },
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
	return value.toString(36).padStart(7, '0').slice(0, 7);
}

function toCellText(cell: Cell) {
	if (cell === null || cell === undefined) return '';
	return String(cell);
}

function buildColumnWidths(rows: Row[]) {
	if (rows.length === 0) return [];
	const columnCount = rows[0]?.length ?? 0;
	return Array.from({ length: columnCount }, (_, columnIndex) => {
		const maxLength = rows.reduce((longest, row) => {
			const length = toCellText(row[columnIndex] ?? '').length;
			return Math.max(longest, length);
		}, 0);
		return { wch: Math.max(10, Math.min(maxLength + 2, 60)) };
	});
}

function shouldAnnotateCountryColumn(header: string) {
	const normalized = header.trim().toLowerCase();
	return normalized === 'nationalität' || normalized.endsWith('land');
}

function isNationalityColumn(header: string) {
	return header.trim().toLowerCase() === 'nationalität';
}

function normalizeCountryCode(value: Cell) {
	return String(value ?? '').trim().toUpperCase();
}

function toPrettyFeeStatus(value: string) {
	const normalized = value.trim().toLowerCase();
	return FEE_STATUS_LABELS[normalized] ?? value;
}

function getBaseUrl() {
	return import.meta.env.BASE_URL.endsWith('/')
		? import.meta.env.BASE_URL
		: `${import.meta.env.BASE_URL}/`;
}

function formatStudyPrettyName(name: string, level: string) {
	const cleanName = name.trim();
	const cleanLevel = level.trim();
	if (!cleanName || !cleanLevel) return '';
	return `${cleanName} (${cleanLevel})`;
}

function mapStudiumText(value: Cell, lookup: Map<string, string>) {
	const raw = String(value ?? '');
	if (!raw.trim()) return raw;
	return raw
		.split(',')
		.map((part) => {
			const key = part.trim();
			if (!key) return '';
			return lookup.get(key) ?? key;
		})
		.filter((part) => part.length > 0)
		.join(', ');
}

function mapStudiumTextWithMeta(value: Cell, lookup: Map<string, string>) {
	const raw = String(value ?? '');
	if (!raw.trim()) {
		return { text: raw, usedLookup: false };
	}

	let usedLookup = false;
	const text = raw
		.split(',')
		.map((part) => {
			const key = part.trim();
			if (!key) return '';
			if (lookup.has(key)) {
				usedLookup = true;
			}
			return lookup.get(key) ?? key;
		})
		.filter((part) => part.length > 0)
		.join(', ');

	return { text, usedLookup };
}

function parseCsvRows(text: string): string[][] {
	const rows: string[][] = [];
	let row: string[] = [];
	let value = '';
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
		if (ch === ',') {
			row.push(value);
			value = '';
			continue;
		}
		if (ch === '\n') {
			row.push(value);
			rows.push(row);
			row = [];
			value = '';
			continue;
		}
		if (ch !== '\r') {
			value += ch;
		}
	}

	row.push(value);
	rows.push(row);
	return rows.filter((entry) => entry.some((cell) => String(cell).length > 0));
}

function toPrimitiveCellValue(value: unknown, fallbackText: string): Cell {
	if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
		return value;
	}
	if (value instanceof Date) return fallbackText || value.toISOString();
	if (value && typeof value === 'object' && 'result' in value) {
		const result = (value as { result?: unknown }).result;
		if (typeof result === 'string' || typeof result === 'number' || typeof result === 'boolean') {
			return result;
		}
	}
	return fallbackText;
}

function setWorksheetCellItalic(worksheet: Worksheet, rowIndex1: number, columnIndex1: number) {
	const cell = worksheet.getCell(rowIndex1, columnIndex1);
	cell.font = { ...(cell.font ?? {}), italic: true };
}

let excelJsPromise: Promise<typeof import('exceljs')> | null = null;

/** exceljs ships as CommonJS, so the namespace object carries the library on
 * its default export at runtime while its types describe the namespace. */
async function loadExcelJs() {
	if (!excelJsPromise) {
		excelJsPromise = import('exceljs');
	}
	const module = await excelJsPromise;
	return ((module as unknown as { default?: typeof module }).default ?? module);
}

const fileName = ref('');
const headers = ref<string[]>([]);
const rows = ref<Row[]>([]);
const selectedColumns = ref<string[]>([]);
const filterColumn = ref('');
const selectedStudium = ref<string[]>([]);
const error = ref('');
const exportMode = ref<'student' | 'statistics'>('student');
const studyNameLookup = ref<Map<string, string>>(new Map());

const dimmedLookup = new Set(DIMMED_COLUMNS.map((name) => name.toLowerCase()));

onMounted(async () => {
	try {
		const response = await fetch(`${getBaseUrl()}study_names.csv`, { cache: 'no-store' });
		if (!response.ok) return;

		const csvRows = parseCsvRows(await response.text());
		const [headerRow, ...dataRows] = csvRows;
		const csvHeader = (headerRow ?? []).map((value) => String(value ?? '').trim().toLowerCase());
		const dnameIndex = csvHeader.indexOf('dname');
		const nameIndex = csvHeader.indexOf('name');
		const levelIndex = csvHeader.indexOf('level');
		if (dnameIndex < 0 || nameIndex < 0 || levelIndex < 0) return;

		const nextLookup = new Map<string, string>();
		dataRows.forEach((row) => {
			const dname = String(row[dnameIndex] ?? '').trim();
			const name = String(row[nameIndex] ?? '').trim();
			const level = String(row[levelIndex] ?? '').trim();
			const pretty = formatStudyPrettyName(name, level);
			if (dname && pretty) {
				nextLookup.set(dname, pretty);
			}
		});
		studyNameLookup.value = nextLookup;
	} catch {
		studyNameLookup.value = new Map();
	}
});

const columnIndex = computed(() =>
	filterColumn.value ? headers.value.indexOf(filterColumn.value) : -1,
);

const availableFilterColumns = computed(() => new Set(headers.value));

const studiumOptions = computed(() => {
	if (columnIndex.value < 0) return [];
	const values = unique(rows.value.map((row) => String(row[columnIndex.value] ?? '')));
	if (filterColumn.value !== 'Studium') {
		return values.sort();
	}
	return values.sort((left, right) => {
		const prettyLeft = mapStudiumText(left, studyNameLookup.value);
		const prettyRight = mapStudiumText(right, studyNameLookup.value);
		const byPretty = prettyLeft.localeCompare(prettyRight, 'de', { sensitivity: 'base' });
		if (byPretty !== 0) return byPretty;
		return left.localeCompare(right, 'de', { sensitivity: 'base' });
	});
});

const studiumCounts = computed(() => {
	const counts = new Map<string, number>();
	if (columnIndex.value < 0) return counts;
	rows.value.forEach((row) => {
		const value = String(row[columnIndex.value] ?? '');
		counts.set(value, (counts.get(value) ?? 0) + 1);
	});
	return counts;
});

const filteredRows = computed(() => {
	if (rows.value.length === 0) return [];
	if (columnIndex.value < 0 || selectedStudium.value.length === 0) return rows.value;
	const allowed = new Set(selectedStudium.value);
	return rows.value.filter((row) => allowed.has(String(row[columnIndex.value] ?? '')));
});

const selectedColumnIndexes = computed(() => {
	const set = new Set(selectedColumns.value);
	return headers.value
		.map((header, index) => ({ header, index }))
		.filter((item) => set.has(item.header));
});

const trimmedRows = computed(() => {
	if (selectedColumnIndexes.value.length === 0) return [];
	return filteredRows.value.map((row) =>
		selectedColumnIndexes.value.map((item) => row[item.index] ?? ''),
	);
});

const uniqueTrimmedRows = computed(() => {
	if (selectedColumnIndexes.value.length === 0 || filteredRows.value.length === 0) return [];
	const studiumFullIndex = headers.value.findIndex((header) => header === 'Studium');
	const merged = new Map<string, { row: Row; values: Set<string> }>();
	const order: string[] = [];
	filteredRows.value.forEach((row) => {
		const keyParts = headers.value.map((_, index) => {
			if (index === studiumFullIndex) return '';
			return String(row[index] ?? '');
		});
		const key = JSON.stringify(keyParts);
		if (!merged.has(key)) {
			merged.set(key, { row: row.map((cell) => cell ?? ''), values: new Set() });
			order.push(key);
		}
		const entry = merged.get(key);
		if (entry) {
			entry.values.add(
				studiumFullIndex >= 0 ? String(row[studiumFullIndex] ?? '').trim() : '',
			);
		}
	});

	return order.map((key) => {
		const entry = merged.get(key)!;
		const mergedStudium = Array.from(entry.values)
			.filter((value) => value.length)
			.join(', ');
		return selectedColumnIndexes.value.map((item) => {
			if (item.header === 'Studium' && studiumFullIndex >= 0) {
				return mergedStudium;
			}
			return entry.row[item.index] ?? '';
		});
	});
});

const processedRows = computed(() => {
	const baseRows = exportMode.value === 'student' ? uniqueTrimmedRows.value : trimmedRows.value;
	if (exportMode.value !== 'statistics') return baseRows;

	const matrikelnummerIndex = selectedColumnIndexes.value.findIndex(
		(item) => item.header === 'Matrikelnummer',
	);
	if (matrikelnummerIndex < 0) return baseRows;

	const ids = new Map<string, string>();
	const usedIds = new Set<string>();

	return baseRows.map((row) => {
		const original = String(row[matrikelnummerIndex] ?? '').trim();
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
});

const canDownload = computed(
	() =>
		headers.value.length > 0
		&& selectedColumnIndexes.value.length > 0
		&& (filterColumn.value === '' || selectedStudium.value.length > 0),
);
const missingFilterSelection = computed(
	() =>
		headers.value.length > 0
		&& filterColumn.value !== ''
		&& selectedStudium.value.length === 0,
);
const allColumnsSelected = computed(
	() => headers.value.length > 0 && selectedColumns.value.length === headers.value.length,
);

function resetState() {
	headers.value = [];
	rows.value = [];
	selectedColumns.value = [];
	filterColumn.value = '';
	selectedStudium.value = [];
}

async function handleFileChange(event: Event) {
	const file = (event.currentTarget as HTMLInputElement).files?.[0];
	if (!file) return;
	error.value = '';

	try {
		const buffer = await file.arrayBuffer();
		const ExcelJS = await loadExcelJs();
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.load(buffer);
		const worksheet = workbook.worksheets[0];
		if (!worksheet) {
			throw new Error('The sheet is empty.');
		}
		const data: Cell[][] = [];
		worksheet.eachRow({ includeEmpty: false }, (excelRow: ExcelRow) => {
			const rowValues: Cell[] = [];
			for (let colIndex = 1; colIndex <= excelRow.cellCount; colIndex += 1) {
				const cell = excelRow.getCell(colIndex);
				rowValues.push(toPrimitiveCellValue(cell.value, cell.text ?? ''));
			}
			data.push(rowValues);
		});

		if (!data || data.length === 0) {
			throw new Error('The sheet is empty.');
		}

		const nonEmptyCounts = data.map(
			(row) => row.filter((cell) => String(cell ?? '').trim()).length,
		);
		const maxNonEmpty = Math.max(...nonEmptyCounts);
		if (maxNonEmpty === 0) {
			throw new Error('No populated rows were found.');
		}

		const startIndex = nonEmptyCounts.findIndex((count) => count === maxNonEmpty);
		const trimmedData = data.slice(startIndex);
		const [headerRow, ...dataRows] = trimmedData;
		const maxColumns = headerRow.length;
		const cleanHeaders = headerRow.slice(0, maxColumns).map((value, index) => {
			const text = String(value ?? '').trim();
			return text.length > 0 ? text : `Column ${index + 1}`;
		});

		fileName.value = file.name;
		headers.value = cleanHeaders;
		rows.value = dataRows as Row[];
		const preferredColumns = ['Vorname', 'Zuname', 'Studium', 'Email'];
		const preferredLookup = new Set(preferredColumns.map((name) => name.toLowerCase()));
		const defaultSelected = cleanHeaders.filter((header) =>
			preferredLookup.has(header.toLowerCase()),
		);
		selectedColumns.value = defaultSelected.length > 0 ? defaultSelected : [];

		filterColumn.value = cleanHeaders.includes(DEFAULT_FILTER_COLUMN)
			? DEFAULT_FILTER_COLUMN
			: FILTERABLE_COLUMNS.find((column) => cleanHeaders.includes(column)) || '';
		selectedStudium.value = [];
	} catch (err: unknown) {
		resetState();
		fileName.value = '';
		error.value = err instanceof Error ? err.message : 'Failed to read file.';
	}
}

function toggleColumn(header: string) {
	const set = new Set(selectedColumns.value);
	if (set.has(header)) {
		set.delete(header);
	} else {
		set.add(header);
	}
	selectedColumns.value = Array.from(set);
}

function toggleStudium(value: string) {
	const set = new Set(selectedStudium.value);
	if (set.has(value)) {
		set.delete(value);
	} else {
		set.add(value);
	}
	selectedStudium.value = Array.from(set);
}

function selectAllColumns(event: Event) {
	selectedColumns.value = (event.target as HTMLInputElement).checked ? [...headers.value] : [];
}

function selectFilterColumn(value: string) {
	filterColumn.value = value;
	selectedStudium.value = [];
}

async function handleDownload() {
	if (selectedColumnIndexes.value.length === 0) return;
	error.value = '';

	const countryLookup = new Map<string, string>();
	const feeLookup = new Map<string, string>();
	let shouldAugmentCountryData = true;
	try {
		const response = await fetch(`${getBaseUrl()}countries.csv`, { cache: 'no-store' });
		if (!response.ok) {
			throw new Error(`countries.csv could not be loaded (${response.status}).`);
		}
		const csvRows = parseCsvRows(await response.text());
		const [csvHeaderRow, ...csvDataRows] = csvRows;
		const csvHeader = (csvHeaderRow ?? []).map((value) =>
			String(value ?? '').trim().toLowerCase(),
		);
		const codeIndex = csvHeader.indexOf('code');
		const nameIndex = csvHeader.indexOf('name');
		const feeIndex = csvHeader.indexOf('fee');

		csvDataRows.forEach((row) => {
			const code = normalizeCountryCode(row[codeIndex >= 0 ? codeIndex : 0]);
			const name = String(row[nameIndex >= 0 ? nameIndex : 1] ?? '').trim();
			const feeRaw = String(row[feeIndex >= 0 ? feeIndex : 4] ?? '').trim();
			const fee = toPrettyFeeStatus(feeRaw);
			if (code) {
				countryLookup.set(code, name || '?');
				feeLookup.set(code, fee || '?');
			}
		});
	} catch (err: unknown) {
		shouldAugmentCountryData = false;
		error.value =
			err instanceof Error
				? `${err.message} Export wird mit Originalwerten fortgesetzt.`
				: 'countries.csv konnte nicht geladen werden. Export wird mit Originalwerten fortgesetzt.';
	}

	const outputHeaders: string[] = [];
	const italicHeaderColumns = new Set<number>();
	selectedColumnIndexes.value.forEach((item) => {
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
	const outputRows: Row[] = processedRows.value.map((row, rowIndex) => {
		const next: Row = [];
		const italicFlags: boolean[] = [];
		selectedColumnIndexes.value.forEach((item, columnIndex) => {
			const originalCell = row[columnIndex] ?? '';
			let renderedCell: Cell = originalCell;
			let shouldItalicCell = false;

			if (item.header === 'Studium') {
				const mappedStudium = mapStudiumTextWithMeta(originalCell, studyNameLookup.value);
				renderedCell = mappedStudium.text;
				shouldItalicCell = mappedStudium.usedLookup;
			} else if (shouldAugmentCountryData && shouldAnnotateCountryColumn(item.header)) {
				const originalValue = String(originalCell);
				const lookupCode = normalizeCountryCode(originalCell);
				const lookedUpName = lookupCode ? countryLookup.get(lookupCode) : undefined;
				renderedCell = `${originalValue} (${lookedUpName ?? '?'})`;
				const outputColumnIndex = next.length;
				countryAnnotatedCells.set(`${rowIndex}:${outputColumnIndex}`, {
					originalValue,
					lookedUpName: lookedUpName ?? '?',
				});
			}

			if (exportMode.value === 'statistics' && item.header === 'Matrikelnummer') {
				shouldItalicCell = true;
			}
			next.push(renderedCell);
			italicFlags.push(shouldItalicCell);

			if (shouldAugmentCountryData && isNationalityColumn(item.header)) {
				const lookupCode = normalizeCountryCode(originalCell);
				const feeStatus = lookupCode ? feeLookup.get(lookupCode) : undefined;
				next.push(feeStatus ?? '?');
				italicFlags.push(true);
			}
		});
		outputCellItalics.push(italicFlags);
		return next;
	});

	const output: Row[] = [outputHeaders, ...outputRows];
	const ExcelJS = await loadExcelJs();
	const workbook = new ExcelJS.Workbook();
	const worksheet = workbook.addWorksheet('Filtered');
	worksheet.addRow(outputHeaders);
	outputRows.forEach((row) => {
		worksheet.addRow(row.map((cell) => (cell === null ? '' : cell)));
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
		const [rowIndexText, columnIndexText] = key.split(':');
		const cell = worksheet.getCell(Number(rowIndexText) + 2, Number(columnIndexText) + 1);
		cell.value = {
			richText: [
				{ text: `${parts.originalValue} (` },
				{ text: parts.lookedUpName, font: { italic: true } },
				{ text: ')' },
			],
		};
	});
	buildColumnWidths(output).forEach((entry, index) => {
		worksheet.getColumn(index + 1).width = entry.wch;
	});
	worksheet.autoFilter = {
		from: { row: 1, column: 1 },
		to: { row: output.length, column: outputHeaders.length },
	};

	const baseName = fileName.value ? fileName.value.replace(/\.xlsx$/i, '') : 'output';
	const buffer = await workbook.xlsx.writeBuffer();
	const blob = new Blob([buffer], {
		type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
	});
	const link = document.createElement('a');
	link.href = URL.createObjectURL(blob);
	link.download = `${baseName}-filtered.xlsx`;
	link.click();
	URL.revokeObjectURL(link.href);
}
</script>

<template>
	<div class="page">
		<header class="hero">
			<div class="hero-intro">
				<p class="eyebrow">Alles lokal im Browser</p>
				<h1>Studierendenliste filtern</h1>
				<p class="lead">
					Excel-Datei hochladen, gewünschte Spalten behalten, nach Studium filtern und die
					bereinigte Datei erzeugen. Alles bleibt auf dem lokalen Computer.
				</p>
			</div>
			<div class="hero-upload">
				<label class="upload">
					<input type="file" accept=".xlsx" @change="handleFileChange">
					<span class="upload-title">Excel-Datei laden</span>
					<span class="upload-file">{{ fileName || 'Keine Datei geladen' }}</span>
				</label>
				<p v-if="error" class="error">{{ error }}</p>
			</div>
		</header>

		<main v-if="headers.length > 0" class="panels">
			<section class="panel">
				<div class="panel-head">
					<h2>Spalten</h2>
					<label class="inline-check">
						<input type="checkbox" :checked="allColumnsSelected" @change="selectAllColumns">
						<span>Alle Spalten</span>
					</label>
				</div>
				<div class="column-grid">
					<label v-for="header in headers" :key="header" class="chip">
						<input
							type="checkbox"
							:checked="selectedColumns.includes(header)"
							@change="toggleColumn(header)">
						<span :class="dimmedLookup.has(header.toLowerCase()) ? 'chip-dimmed' : 'chip-label'">
							{{ header }}
						</span>
					</label>
				</div>
			</section>

			<section class="panel">
				<div class="panel-head">
					<h2>Studium-Filter</h2>
					<div class="mode-buttons">
						<button
							v-for="mode in FILTER_MODES"
							:key="mode.value || 'none'"
							type="button"
							:disabled="mode.value !== '' && !availableFilterColumns.has(mode.value)"
							:class="['mode-button', { 'mode-button--active': filterColumn === mode.value }]"
							@click="selectFilterColumn(mode.value)">
							{{ mode.label }}
						</button>
					</div>
				</div>
				<p v-if="filterColumn === ''" class="muted">
					Keine Filterspalte ausgewählt, somit werden die Daten aller Studierenden
					exportiert. Wähle "Studium" oder "Hörer*status", um Einträge zu filtern.
				</p>
				<div v-else class="option-list">
					<label v-for="value in studiumOptions" :key="String(value)" class="option">
						<input
							type="checkbox"
							:checked="selectedStudium.includes(value)"
							@change="toggleStudium(value)">
						<span class="option-label">
							{{ filterColumn === 'Studium'
								? mapStudiumText(value, studyNameLookup) || '(leer)'
								: String(value) || '(leer)' }}
						</span>
						<span class="count">{{ studiumCounts.get(value) ?? 0 }}</span>
					</label>
				</div>
			</section>

			<section class="panel panel--export">
				<div>
					<h2>Export</h2>
					<p class="muted">
						{{ missingFilterSelection
							? 'Bitte wähle mindestens eine Filterkategorie zum Exportieren'
							: `${processedRows.length} Zeilen nach Filterung und Spaltenauswahl.` }}
					</p>
				</div>
				<div class="export-actions">
					<button type="button" class="primary" :disabled="!canDownload" @click="handleDownload">
						Gefilterte Excel-Datei herunterladen
					</button>
					<label class="inline-check">
						<span>Exportmodus</span>
						<select v-model="exportMode">
							<option value="student">Studierendenzentriert</option>
							<option value="statistics">Statistikzentriert</option>
						</select>
					</label>
					<span class="hint">
						{{ exportMode === 'student'
							? 'Studierende mit Mehrfachstudien werden zu einem Eintrag zusammengeführt und Studium-Werte kombiniert.'
							: 'Studierende mit Mehrfachstudien bleiben als separate Einträge erhalten; Matrikelnummer wird durch eine Zufalls-ID ersetzt.' }}
					</span>
				</div>
			</section>
		</main>
	</div>
</template>

<style scoped>
.page {
	margin: 0 auto;
	display: flex;
	width: 100%;
	max-width: 1100px;
	flex-direction: column;
	gap: 24px;
	padding: 32px 24px 48px;
}

:root[data-embed='1'] .page {
	padding: 0 0 16px;
	max-width: none;
}

.hero {
	display: grid;
	gap: 24px;
	align-items: start;
}

@media (min-width: 900px) {
	.hero {
		grid-template-columns: 1.2fr 0.8fr;
	}
}

/* the host page provides its own heading and description */
:root[data-embed='1'] .hero-intro {
	display: none;
}

.eyebrow {
	margin: 0;
	font-size: 12px;
	font-weight: 600;
	text-transform: uppercase;
	letter-spacing: 0.2em;
	color: var(--sl-accent);
}

h1 {
	margin: 8px 0 12px;
	font-size: 36px;
	line-height: 1.15;
	font-weight: 600;
}

h2 {
	margin: 0;
	font-size: 20px;
	font-weight: 600;
}

.lead {
	margin: 0;
	max-width: 520px;
	font-size: 16px;
	line-height: 1.6;
	color: var(--sl-text-muted);
}

.upload {
	display: flex;
	flex-direction: column;
	gap: 4px;
	cursor: pointer;
	border: 1px solid var(--sl-border);
	border-radius: var(--sl-radius);
	background: var(--sl-background);
	padding: 16px;
}

.upload input {
	display: none;
}

.upload-title {
	font-weight: 600;
}

.upload-file,
.hint,
.muted {
	font-size: 13px;
	color: var(--sl-text-muted);
}

.error {
	margin: 8px 0 0;
	font-weight: 600;
	color: var(--sl-error);
}

.panels {
	display: grid;
	gap: 16px;
}

.panel {
	border: 1px solid var(--sl-border);
	border-radius: var(--sl-radius);
	background: var(--sl-background);
	padding: 20px;
}

.panel--export {
	display: flex;
	flex-direction: column;
	gap: 20px;
}

@media (min-width: 900px) {
	.panel--export {
		flex-direction: row;
		align-items: flex-start;
		justify-content: space-between;
	}
}

.panel-head {
	display: flex;
	align-items: center;
	justify-content: space-between;
	gap: 16px;
	margin-bottom: 16px;
	flex-wrap: wrap;
}

.inline-check {
	display: inline-flex;
	align-items: center;
	gap: 8px;
	font-size: 13px;
	color: var(--sl-text-muted);
}

.column-grid {
	display: grid;
	gap: 8px;
	grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
}

.chip,
.option {
	display: flex;
	align-items: center;
	gap: 10px;
	border: 1px solid var(--sl-border);
	border-radius: var(--sl-radius);
	background: var(--sl-hover);
	padding: 8px 12px;
	font-size: 13px;
	min-height: var(--sl-clickable);
	cursor: pointer;
}

.option {
	display: grid;
	grid-template-columns: auto 1fr auto;
}

.chip-label {
	font-weight: 600;
}

.chip-dimmed {
	color: var(--sl-warning);
}

.option-list {
	display: flex;
	flex-direction: column;
	gap: 8px;
}

.option-label {
	overflow: hidden;
	text-overflow: ellipsis;
	white-space: nowrap;
}

.count {
	border-radius: var(--sl-radius);
	background: var(--sl-hover);
	border: 1px solid var(--sl-border);
	padding: 2px 8px;
	font-size: 12px;
	color: var(--sl-text-muted);
}

.mode-buttons {
	display: flex;
	flex-wrap: wrap;
	gap: 8px;
}

.mode-button,
.primary,
select {
	font: inherit;
	border-radius: var(--sl-radius);
	border: 1px solid var(--sl-border);
	background: var(--sl-hover);
	color: var(--sl-text);
	padding: 8px 14px;
	min-height: var(--sl-clickable);
	cursor: pointer;
}

.mode-button--active {
	border-color: var(--sl-accent);
	background: var(--sl-accent);
	color: var(--sl-accent-text);
}

.mode-button:disabled,
.primary:disabled {
	cursor: not-allowed;
	opacity: 0.5;
}

.primary {
	border-color: var(--sl-accent);
	background: var(--sl-accent);
	color: var(--sl-accent-text);
	font-weight: 600;
}

.export-actions {
	display: flex;
	flex-direction: column;
	gap: 10px;
	width: 100%;
}

@media (min-width: 900px) {
	.export-actions {
		width: 45%;
	}
}
</style>
