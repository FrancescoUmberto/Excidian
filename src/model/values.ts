export interface ErrorValue {
	error: string;
}

/** A cell's value as read from the file. Dates are UTC `Date`s. */
export type CellValue = number | string | boolean | Date | ErrorValue | null;

/** What can be written into a cell. */
export type CellInput =
	| { kind: "empty" }
	| { kind: "number"; value: number }
	| { kind: "text"; value: string }
	| { kind: "date"; value: Date }
	| { kind: "formula"; formula: string };

const DAY_MS = 86_400_000;
const EPOCH_1900 = Date.UTC(1899, 11, 30);
const EPOCH_1904 = Date.UTC(1904, 0, 1);

export function serialToDate(serial: number, date1904: boolean): Date {
	const ms = (date1904 ? EPOCH_1904 : EPOCH_1900) + serial * DAY_MS;
	return new Date(Math.round(ms / 1000) * 1000);
}

export function dateToSerial(date: Date, date1904: boolean): number {
	return (date.getTime() - (date1904 ? EPOCH_1904 : EPOCH_1900)) / DAY_MS;
}

/** Turns typed text into a cell input: formula, number, date or plain text. */
export function parseInput(text: string): CellInput {
	const t = text.trim();
	if (t === "") return { kind: "empty" };
	if (t.startsWith("=") && t.length > 1) return { kind: "formula", formula: t.slice(1) };
	if (/^[+-]?\d+(?:[.,]\d+)?$/.test(t)) return { kind: "number", value: Number(t.replace(",", ".")) };

	let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
	if (m) return dateInput(Number(m[1]), Number(m[2]), Number(m[3])) ?? { kind: "text", value: t };
	m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); // dd/mm/yyyy
	if (m) return dateInput(Number(m[3]), Number(m[2]), Number(m[1])) ?? { kind: "text", value: t };

	return { kind: "text", value: t };
}

function dateInput(year: number, month: number, day: number): CellInput | undefined {
	const date = new Date(Date.UTC(year, month - 1, day));
	// Rejects impossible dates such as 31/02/2026 instead of rolling them over.
	if (date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return undefined;
	return { kind: "date", value: date };
}

/** The text put in the editor when a cell is edited. */
export function editText(value: CellValue, formula: string | undefined): string {
	if (formula) return `=${formula}`;
	if (value === null) return "";
	if (value instanceof Date) {
		const iso = value.toISOString();
		return iso.endsWith("T00:00:00.000Z") ? iso.slice(0, 10) : iso.slice(0, 16).replace("T", " ");
	}
	if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
	if (typeof value === "object") return value.error;
	return String(value);
}
