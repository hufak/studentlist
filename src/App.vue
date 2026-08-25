<script setup lang="ts">
import { computed, onMounted, ref } from 'vue';
import type { Row as ExcelRow, Worksheet } from 'exceljs';

// Not a spreadsheet column: derived from the Studium column via the study details lookup.
const STV_FILTER = '__stv__';
const NO_STV = '';
// Drives the generated StV level suffixes, e.g. "Restaurierung+CHC (Diplom+MA)".
const STUDY_LEVEL_ORDER = ['Diplom', 'BA', 'MA', 'PhD'];
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
const WORKSPACE_TABS = [
	{ value: 'export', label: 'Filtern+Exportieren' },
	{ value: 'budget', label: 'StV-Budget' },
];
// Shared by the rendered budget table and its clipboard copy, so the two cannot drift.
const BUDGET_COLUMNS = ['StV', 'Studierende', 'Anteil', 'Sockel', 'Budget anteilig', 'Budget'];
const BUDGET_TOTAL_LABEL = 'Gesamt';
// The share of the student contribution revenue that is split between the StVs.
const BUDGET_SHARE = 0.3;
const BUDGET_SHARE_LABEL = `davon ${BUDGET_SHARE * 100}%`;
// Order matters twice: the buttons render in it, and the first available mode
// past "alle Studierenden" becomes the default once a sheet is loaded.
const FILTER_MODES = [
	{ value: '', label: 'alle Studierenden' },
	{ value: STV_FILTER, label: 'nach StV' },
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

const percentFormatter = new Intl.NumberFormat('de-AT', {
	style: 'percent',
	minimumFractionDigits: 1,
	maximumFractionDigits: 1,
});

const decimalFormatter = new Intl.NumberFormat('de-AT', {
	minimumFractionDigits: 2,
	maximumFractionDigits: 2,
});

function formatShare(share: number) {
	return percentFormatter.format(share);
}

const currencyFormatter = new Intl.NumberFormat('de-AT', {
	style: 'currency',
	currency: 'EUR',
});

/** Money is carried in whole cents throughout, so the split stays exact. */
function formatCurrency(cents: number) {
	return currencyFormatter.format(cents / 100);
}

function formatCurrencyInput(value: string | number) {
	const amount = parseCurrencyInput(value);
	return Number.isFinite(amount) && amount >= 0 ? decimalFormatter.format(amount) : String(value ?? '');
}

function parseCurrencyInput(value: string | number) {
	const normalized = String(value ?? '')
		.trim()
		.replace(/€/g, '')
		.replace(/\s/g, '')
		.replace(/\.(?=\d{3}(?:,|$))/g, '')
		.replace(',', '.');
	return Number.parseFloat(normalized);
}

function sortStudyLevels(levels: Iterable<string>) {
	return Array.from(levels).sort((left, right) => {
		const leftRank = STUDY_LEVEL_ORDER.indexOf(left);
		const rightRank = STUDY_LEVEL_ORDER.indexOf(right);
		// Levels the order does not know about trail the known ones, alphabetically.
		if (leftRank < 0 && rightRank < 0) return left.localeCompare(right, 'de');
		if (leftRank < 0) return 1;
		if (rightRank < 0) return -1;
		return leftRank - rightRank;
	});
}

function splitStudiumKeys(value: Cell) {
	return String(value ?? '')
		.split(',')
		.map((part) => part.trim())
		.filter((part) => part.length > 0);
}

function mapStudiumText(value: Cell, lookup: Map<string, string>) {
	const raw = String(value ?? '');
	if (!raw.trim()) return raw;
	return splitStudiumKeys(raw)
		.map((key) => lookup.get(key) ?? key)
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
const activeTab = ref('export');
const revenueInput = ref<string | number>('');
const baseBudgetInput = ref<string | number>(300);
const exportMode = ref<'student' | 'statistics'>('student');
const studyNameLookup = ref<Map<string, string>>(new Map());
const stvLookup = ref<Map<string, string>>(new Map());
const stvLabelLookup = ref<Map<string, string>>(new Map());

const dimmedLookup = new Set(DIMMED_COLUMNS.map((name) => name.toLowerCase()));

async function fetchCsvRows(fileName: string) {
	const response = await fetch(`${getBaseUrl()}${fileName}`, { cache: 'no-store' });
	if (!response.ok) return null;

	const csvRows = parseCsvRows(await response.text());
	const [headerRow, ...dataRows] = csvRows;
	const csvHeader = (headerRow ?? []).map((value) => String(value ?? '').trim().toLowerCase());
	return { csvHeader, dataRows };
}

// Maps the spreadsheet's Studium values onto pretty study names ("name (level)").
async function loadStudyNameLookup() {
	try {
		const csv = await fetchCsvRows('angewandte_evidenz_study_name_lookup.csv');
		if (!csv) return;

		const dnameIndex = csv.csvHeader.indexOf('dname');
		const nameIndex = csv.csvHeader.indexOf('name');
		const levelIndex = csv.csvHeader.indexOf('level');
		if (dnameIndex < 0 || nameIndex < 0 || levelIndex < 0) return;

		const nextLookup = new Map<string, string>();
		csv.dataRows.forEach((row) => {
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
}

// Maps pretty study names onto the StV responsible for them, and derives each StV's
// display label from the levels of the degrees it covers, e.g. "TransArts (BA+MA)".
async function loadStvLookup() {
	try {
		const csv = await fetchCsvRows('angewandte_study_details.csv');
		if (!csv) return;

		const nameIndex = csv.csvHeader.indexOf('name');
		const levelIndex = csv.csvHeader.indexOf('level');
		const stvIndex = csv.csvHeader.indexOf('stv');
		if (nameIndex < 0 || levelIndex < 0 || stvIndex < 0) return;

		const nextLookup = new Map<string, string>();
		const levelsByStv = new Map<string, Set<string>>();
		csv.dataRows.forEach((row) => {
			const name = String(row[nameIndex] ?? '').trim();
			const level = String(row[levelIndex] ?? '').trim();
			const stv = String(row[stvIndex] ?? '').trim();
			const pretty = formatStudyPrettyName(name, level);
			if (pretty && stv) {
				nextLookup.set(pretty, stv);
				const levels = levelsByStv.get(stv) ?? new Set<string>();
				levels.add(level);
				levelsByStv.set(stv, levels);
			}
		});

		const nextLabels = new Map<string, string>();
		levelsByStv.forEach((levels, stv) => {
			nextLabels.set(stv, `${stv} (${sortStudyLevels(levels).join('+')})`);
		});

		stvLookup.value = nextLookup;
		stvLabelLookup.value = nextLabels;
	} catch {
		stvLookup.value = new Map();
		stvLabelLookup.value = new Map();
	}
}

onMounted(() => {
	void loadStudyNameLookup();
	void loadStvLookup();
});

const isStvMode = computed(() => filterColumn.value === STV_FILTER);

const studiumColumnIndex = computed(() => headers.value.indexOf('Studium'));

const columnIndex = computed(() => {
	if (isStvMode.value) return studiumColumnIndex.value;
	return filterColumn.value ? headers.value.indexOf(filterColumn.value) : -1;
});

const availableFilterColumns = computed(() => new Set(headers.value));

/** Identifies the student behind an enrolment row: the sheet holds one row per
 * enrolment, so one student's several studies differ only in the Studium column.
 * The student-centred export merges on this, and the StV counts count it, so a
 * count always matches the number of entries that selecting it exports. */
function studentIdentity(row: Row) {
	return JSON.stringify(
		headers.value.map((_, index) =>
			index === studiumColumnIndex.value ? '' : String(row[index] ?? ''),
		),
	);
}

function isFilterModeAvailable(mode: string) {
	if (mode === '') return true;
	if (mode === STV_FILTER) return studiumColumnIndex.value >= 0 && stvLookup.value.size > 0;
	return availableFilterColumns.value.has(mode);
}

// The StVs each row belongs to; a student with multiple studies can belong to several.
const rowStvs = computed(() => {
	const studiumIndex = studiumColumnIndex.value;
	if (studiumIndex < 0) return [];
	return rows.value.map((row) => {
		const keys = splitStudiumKeys(row[studiumIndex]);
		if (keys.length === 0) return [NO_STV];
		return unique(
			keys.map((key) => {
				const pretty = studyNameLookup.value.get(key) ?? key;
				return stvLookup.value.get(pretty) ?? NO_STV;
			}),
		);
	});
});

/** Distinct students per StV across the whole sheet, independent of the current
 * filter: several studies of one student can share an StV, and one student can
 * fall under several StVs and is then counted under each. */
const stvStudentCounts = computed(() => {
	const studentsByStv = new Map<string, Set<string>>();
	rows.value.forEach((row, index) => {
		const student = studentIdentity(row);
		(rowStvs.value[index] ?? []).forEach((stv) => {
			const students = studentsByStv.get(stv) ?? new Set<string>();
			students.add(student);
			studentsByStv.set(stv, students);
		});
	});
	const counts = new Map<string, number>();
	studentsByStv.forEach((students, stv) => counts.set(stv, students.size));
	return counts;
});

const revenueCents = computed(() => {
	const value = parseCurrencyInput(revenueInput.value);
	return Number.isFinite(value) && value > 0 ? Math.round(value * 100) : 0;
});

const budgetPoolCents = computed(() => Math.round(revenueCents.value * BUDGET_SHARE));

const baseBudgetCents = computed(() => {
	const value = parseCurrencyInput(baseBudgetInput.value);
	return Number.isFinite(value) && value > 0 ? Math.round(value * 100) : 0;
});

/** Splits the pool across the StVs in proportion to their headcounts, by the
 * largest remainder method: hand every StV its whole cents, then give the odd
 * cents left over to the largest fractions. Rounding each share on its own
 * would leave the column a few cents off the pool it is meant to add up to. */
function apportionCents(pool: number, weights: number[]) {
	const total = weights.reduce((sum, weight) => sum + weight, 0);
	if (total <= 0 || pool <= 0) return weights.map(() => 0);

	const exact = weights.map((weight) => (pool * weight) / total);
	const amounts = exact.map((value) => Math.floor(value));
	let leftover = pool - amounts.reduce((sum, value) => sum + value, 0);

	exact
		.map((value, index) => ({ index, remainder: value - Math.floor(value) }))
		// ties resolved by position, so the same input always splits the same way
		.sort((left, right) => right.remainder - left.remainder || left.index - right.index)
		.forEach(({ index }) => {
			if (leftover > 0) {
				amounts[index] += 1;
				leftover -= 1;
			}
		});

	return amounts;
}

/** Every known StV with its headcount, share of the whole and slice of the
 * pool. The share is of the column total rather than of the student body: a
 * student under two StVs counts for both, so the shares still add up to 100%
 * while the total sits above the number of students. */
const stvBudget = computed(() => {
	const counts = stvStudentCounts.value;
	const names = unique([...stvLabelLookup.value.keys(), ...counts.keys()]).filter(
		(name) => name !== NO_STV,
	);
	const entries = names
		.map((name) => ({
			name,
			label: stvLabelLookup.value.get(name) ?? name,
			students: counts.get(name) ?? 0,
		}))
		.sort((left, right) => left.label.localeCompare(right.label, 'de', { sensitivity: 'base' }));
	const total = entries.reduce((sum, entry) => sum + entry.students, 0);
	const proportionalPoolCents = Math.max(
		0,
		budgetPoolCents.value - baseBudgetCents.value * entries.length,
	);
	const proportionalAmounts = apportionCents(
		proportionalPoolCents,
		entries.map((entry) => entry.students),
	);
	return {
		entries: entries.map((entry, index) => ({
			...entry,
			share: total > 0 ? entry.students / total : 0,
			baseCents: baseBudgetCents.value,
			proportionalCents: proportionalAmounts[index] ?? 0,
			amountCents: baseBudgetCents.value + (proportionalAmounts[index] ?? 0),
		})),
		total,
		baseAmountCents: baseBudgetCents.value * entries.length,
		proportionalAmountCents: proportionalAmounts.reduce((sum, value) => sum + value, 0),
		totalAmountCents:
			baseBudgetCents.value * entries.length +
			proportionalAmounts.reduce((sum, value) => sum + value, 0),
		withoutStv: counts.get(NO_STV) ?? 0,
	};
});

/** Tab separated rows, which is what spreadsheets read a plain text paste as:
 * one cell per tab, one row per newline. Values are the rendered ones, so the
 * pasted table matches the displayed one. */
const stvBudgetAsTsv = computed(() => {
	const budget = stvBudget.value;
	return [
		BUDGET_COLUMNS,
		...budget.entries.map((entry) => [
			entry.label,
			String(entry.students),
			formatShare(entry.share),
			formatCurrency(entry.baseCents),
			formatCurrency(entry.proportionalCents),
			formatCurrency(entry.amountCents),
		]),
		[
			BUDGET_TOTAL_LABEL,
			String(budget.total),
			formatShare(budget.total > 0 ? 1 : 0),
			formatCurrency(budget.baseAmountCents),
			formatCurrency(budget.proportionalAmountCents),
			formatCurrency(budget.totalAmountCents),
		],
	]
		.map((cells) => cells.join('\t'))
		.join('\n');
});

/** The async clipboard needs a secure context and, framed, a clipboard-write
 * permission policy from the host; fall back to the old selection copy, which
 * needs neither, rather than leaving the button dead when embedded. */
function writeToClipboard(text: string) {
	if (navigator.clipboard?.writeText) {
		return navigator.clipboard.writeText(text).catch(() => {
			if (!copyViaSelection(text)) throw new Error('clipboard unavailable');
		});
	}
	return copyViaSelection(text)
		? Promise.resolve()
		: Promise.reject(new Error('clipboard unavailable'));
}

function copyViaSelection(text: string) {
	const textarea = document.createElement('textarea');
	textarea.value = text;
	textarea.setAttribute('readonly', '');
	textarea.style.position = 'fixed';
	textarea.style.top = '0';
	textarea.style.opacity = '0';
	document.body.appendChild(textarea);
	textarea.select();
	let copied = false;
	try {
		copied = document.execCommand('copy');
	} catch {
		copied = false;
	}
	document.body.removeChild(textarea);
	return copied;
}

const copyState = ref<'idle' | 'copied' | 'failed'>('idle');
let copyStateTimer = 0;

async function copyBudgetTable() {
	try {
		await writeToClipboard(stvBudgetAsTsv.value);
		copyState.value = 'copied';
	} catch {
		copyState.value = 'failed';
	}
	window.clearTimeout(copyStateTimer);
	copyStateTimer = window.setTimeout(() => {
		copyState.value = 'idle';
	}, 2500);
}

const studiumOptions = computed(() => {
	if (columnIndex.value < 0) return [];
	if (isStvMode.value) {
		return unique(rowStvs.value.flat()).sort((left, right) => {
			if (left === NO_STV) return 1;
			if (right === NO_STV) return -1;
			return filterOptionLabel(left).localeCompare(filterOptionLabel(right), 'de', {
				sensitivity: 'base',
			});
		});
	}
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
	// Students, not enrolments: several studies of one student can share an StV.
	if (isStvMode.value) return stvStudentCounts.value;
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
	if (isStvMode.value) {
		return rows.value.filter((_, index) =>
			(rowStvs.value[index] ?? []).some((stv) => allowed.has(stv)),
		);
	}
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
	const studiumFullIndex = studiumColumnIndex.value;
	const merged = new Map<string, { row: Row; values: Set<string> }>();
	const order: string[] = [];
	filteredRows.value.forEach((row) => {
		const key = studentIdentity(row);
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

		filterColumn.value =
			FILTER_MODES.find((mode) => mode.value !== '' && isFilterModeAvailable(mode.value))
				?.value ?? '';
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

function filterOptionLabel(value: string) {
	if (isStvMode.value) {
		if (!value) return '(keine StV)';
		return stvLabelLookup.value.get(value) ?? value;
	}
	if (filterColumn.value === 'Studium') {
		return mapStudiumText(value, studyNameLookup.value) || '(leer)';
	}
	return value || '(leer)';
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
				<h1>Studierendenevidenz filtern</h1>
				<p class="lead">
					Excel-Datei hochladen, gewünschte Spalten behalten, nach Studien oder StV filtern und die
					bereinigte Datei erzeugen. Alle Daten bleiben lokal auf dem Computer.
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

		<main v-if="headers.length > 0" class="workspace">
			<div class="tabs" role="tablist">
				<button
					v-for="tab in WORKSPACE_TABS"
					:id="`tab-${tab.value}`"
					:key="tab.value"
					type="button"
					role="tab"
					:aria-selected="activeTab === tab.value"
					:aria-controls="`panel-${tab.value}`"
					:class="['tab', { 'tab--active': activeTab === tab.value }]"
					@click="activeTab = tab.value">
					{{ tab.label }}
				</button>
			</div>

			<div
				v-show="activeTab === 'export'"
				id="panel-export"
				class="panels"
				role="tabpanel"
				aria-labelledby="tab-export">
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
								:disabled="!isFilterModeAvailable(mode.value)"
								:class="['mode-button', { 'mode-button--active': filterColumn === mode.value }]"
								@click="selectFilterColumn(mode.value)">
								{{ mode.label }}
							</button>
						</div>
					</div>
					<p v-if="filterColumn === ''" class="muted">
						Kein Filter ausgewählt, somit werden die Daten aller Studierenden
						exportiert. Wähle "Studium", "StV" oder "Hörer*status", um Einträge zu filtern.
					</p>
					<div v-else class="option-list">
						<p v-if="isStvMode" class="muted">
							Exportiert alle Studierenden, deren Studium laut Studiendaten der gewählten StV
							zugeordnet ist. Mehrfachstudien zählen zu jeder betroffenen StV.
						</p>
						<label v-for="value in studiumOptions" :key="String(value)" class="option">
							<input
								type="checkbox"
								:checked="selectedStudium.includes(value)"
								@change="toggleStudium(value)">
							<span class="option-label">{{ filterOptionLabel(value) }}</span>
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
			</div>

			<section
				v-show="activeTab === 'budget'"
				id="panel-budget"
				class="panel"
				role="tabpanel"
				aria-labelledby="tab-budget">
				<div class="panel-head">
					<h2>StV-Budget</h2>
				</div>
				<p class="muted">
					Anzahl der Studierenden, für die jede StV zuständig ist. Studierende mit
					Mehrfachstudien zählen für jede betroffene StV, daher liegt die Summe über der
					Zahl der Studierenden und die Prozentanteile beziehen sich auf diese Summe.
				</p>
				<div class="budget-inputs">
					<label class="budget-field">
						<span class="budget-field-label">Erträge Studierendenbeiträge</span>
						<div class="currency-input">
							<span aria-hidden="true">€</span>
							<input
								v-model="revenueInput"
								class="budget-input"
								type="text"
								inputmode="decimal"
								lang="de-AT"
								placeholder="0,00"
								@blur="revenueInput = formatCurrencyInput(revenueInput)">
						</div>
					</label>
					<div class="budget-field">
						<span class="budget-field-label">{{ BUDGET_SHARE_LABEL }}</span>
						<output class="budget-output">{{ formatCurrency(budgetPoolCents) }}</output>
					</div>
					<label class="budget-field">
						<span class="budget-field-label">Sockel pro StV</span>
						<div class="currency-input">
							<span aria-hidden="true">€</span>
							<input
								v-model="baseBudgetInput"
								class="budget-input"
								type="text"
								inputmode="decimal"
								lang="de-AT"
								placeholder="0,00"
								@blur="baseBudgetInput = formatCurrencyInput(baseBudgetInput)">
						</div>
					</label>
				</div>
				<div class="table-actions">
					<span class="copy-status" role="status">
						{{ copyState === 'copied'
							? 'In die Zwischenablage kopiert'
							: copyState === 'failed'
								? 'Kopieren nicht möglich'
								: '' }}
					</span>
					<button
						type="button"
						class="copy-button"
						title="Tabelle in die Zwischenablage kopieren"
						aria-label="Tabelle in die Zwischenablage kopieren"
						@click="copyBudgetTable">
						<svg
							class="copy-icon"
							viewBox="0 0 24 24"
							fill="none"
							stroke="currentColor"
							stroke-width="2"
							stroke-linecap="round"
							stroke-linejoin="round"
							aria-hidden="true"
							focusable="false">
							<template v-if="copyState === 'copied'">
								<path d="M20 6 9 17l-5-5" />
							</template>
							<template v-else>
								<rect x="8" y="2" width="8" height="4" rx="1" />
								<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
							</template>
						</svg>
					</button>
				</div>
				<div class="table-scroll">
					<table class="budget-table">
						<thead>
							<tr>
								<th
									v-for="(column, index) in BUDGET_COLUMNS"
									:key="column"
									scope="col"
									:class="{ numeric: index > 0 }">
									{{ column }}
								</th>
							</tr>
						</thead>
						<tbody>
							<tr v-for="entry in stvBudget.entries" :key="entry.name">
								<td>{{ entry.label }}</td>
								<td class="numeric">{{ entry.students }}</td>
								<td class="numeric">{{ formatShare(entry.share) }}</td>
								<td class="numeric">{{ formatCurrency(entry.baseCents) }}</td>
								<td class="numeric">{{ formatCurrency(entry.proportionalCents) }}</td>
								<td class="numeric">{{ formatCurrency(entry.amountCents) }}</td>
							</tr>
						</tbody>
						<tfoot>
							<tr>
								<th scope="row">{{ BUDGET_TOTAL_LABEL }}</th>
								<td class="numeric">{{ stvBudget.total }}</td>
								<td class="numeric">{{ formatShare(stvBudget.total > 0 ? 1 : 0) }}</td>
								<td class="numeric">{{ formatCurrency(stvBudget.baseAmountCents) }}</td>
								<td class="numeric">{{ formatCurrency(stvBudget.proportionalAmountCents) }}</td>
								<td class="numeric">{{ formatCurrency(stvBudget.totalAmountCents) }}</td>
							</tr>
						</tfoot>
					</table>
				</div>
				<p v-if="stvBudget.withoutStv > 0" class="muted">
					{{ stvBudget.withoutStv }} Studierende sind keiner StV zugeordnet und bleiben in
					dieser Aufstellung unberücksichtigt.
				</p>
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

.workspace {
	display: grid;
	gap: 16px;
}

/* laid out like the studentstats2025 view tabs: equal columns across the full
 * width, underlined rather than boxed */
.tabs {
	display: grid;
	grid-template-columns: repeat(2, minmax(0, 1fr));
	gap: 4px;
	border-bottom: 1px solid var(--sl-border);
}

.tab {
	min-width: 0;
	font: inherit;
	font-weight: 600;
	border: none;
	border-bottom: 2px solid transparent;
	border-radius: var(--sl-radius) var(--sl-radius) 0 0;
	background: transparent;
	color: var(--sl-text-muted);
	padding: 12px;
	min-height: 56px;
	line-height: 1.2;
	text-align: center;
	text-wrap: balance;
	cursor: pointer;
}

.tab:hover {
	background: var(--sl-hover);
	color: var(--sl-text);
}

.tab--active {
	border-bottom-color: var(--sl-accent);
	color: var(--sl-text);
}

.panels {
	display: grid;
	gap: 16px;
}

/* sits directly on top of the table it copies */
.table-actions {
	display: flex;
	align-items: center;
	justify-content: flex-end;
	gap: 8px;
	margin-top: 16px;
}

.copy-button {
	display: inline-flex;
	align-items: center;
	justify-content: center;
	border-radius: var(--sl-radius);
	border: 1px solid var(--sl-border);
	background: var(--sl-hover);
	color: var(--sl-text);
	padding: 6px;
	cursor: pointer;
}

.copy-button:hover {
	border-color: var(--sl-accent);
	color: var(--sl-accent);
}

.copy-icon {
	width: 18px;
	height: 18px;
}

.copy-status {
	font-size: 13px;
	color: var(--sl-text-muted);
}

.budget-inputs {
	display: flex;
	flex-wrap: wrap;
	gap: 16px;
	margin-top: 16px;
}

.budget-field {
	display: flex;
	flex-direction: column;
	gap: 4px;
}

.budget-field-label {
	font-size: 13px;
	color: var(--sl-text-muted);
}

.budget-input,
.budget-output {
	font: inherit;
	border-radius: var(--sl-radius);
	border: 1px solid var(--sl-border);
	padding: 8px 12px;
	min-height: var(--sl-clickable);
	min-width: 12ch;
	text-align: right;
	font-variant-numeric: tabular-nums;
	color: var(--sl-text);
}

.budget-input {
	background: var(--sl-background);
}

.currency-input {
	display: flex;
	align-items: center;
	border: 1px solid var(--sl-border);
	border-radius: var(--sl-radius);
	background: var(--sl-background);
	padding-left: 12px;
	color: var(--sl-text-muted);
}

.currency-input:focus-within {
	border-color: var(--sl-accent);
}

.currency-input .budget-input {
	border: 0;
	outline: 0;
	min-width: 10ch;
}

/* computed, not editable: reads as a readout rather than as another field */
.budget-output {
	display: flex;
	align-items: center;
	justify-content: flex-end;
	background: var(--sl-hover);
	font-weight: 600;
}

.table-scroll {
	overflow-x: auto;
}

.budget-table {
	width: 100%;
	border-collapse: collapse;
	margin-top: 8px;
}

.budget-table th,
.budget-table td {
	text-align: left;
	padding: 8px 12px;
	border-bottom: 1px solid var(--sl-border);
	white-space: nowrap;
}

.budget-table thead th {
	font-size: 13px;
	color: var(--sl-text-muted);
	font-weight: 600;
}

.budget-table tbody tr:hover {
	background: var(--sl-hover);
}

.budget-table .numeric {
	text-align: right;
	font-variant-numeric: tabular-nums;
}

.budget-table tfoot th,
.budget-table tfoot td {
	border-bottom: none;
	border-top: 2px solid var(--sl-border);
	font-weight: 600;
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

/* Dispreferred, not dangerous: these columns are merely discouraged from export. */
.chip-dimmed {
	color: var(--sl-text-muted);
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
