export interface CellRef {
	row: number;
	col: number;
}

export interface Range {
	top: number;
	left: number;
	bottom: number;
	right: number;
}

export const MAX_ROW = 1048576;
export const MAX_COL = 16384;

export function colName(n: number): string {
	let s = "";
	while (n > 0) {
		const m = (n - 1) % 26;
		s = String.fromCharCode(65 + m) + s;
		n = Math.floor((n - 1) / 26);
	}
	return s;
}

export function colNumber(letters: string): number {
	let n = 0;
	for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
	return n;
}

export function parseRef(ref: string): CellRef {
	const m = ref.replace(/\$/g, "").match(/^([A-Z]{1,3})(\d+)$/i);
	if (!m) throw new Error(`Invalid cell reference "${ref}".`);
	return { row: Number(m[2]), col: colNumber(m[1]) };
}

export function formatRef(row: number, col: number): string {
	return `${colName(col)}${row}`;
}

/** Parses "A1:F20" or a single cell "A1". */
export function parseRange(ref: string): Range {
	const [a, b = a] = ref.split(":");
	const start = parseRef(a);
	const end = parseRef(b);
	return {
		top: Math.min(start.row, end.row),
		left: Math.min(start.col, end.col),
		bottom: Math.max(start.row, end.row),
		right: Math.max(start.col, end.col),
	};
}

export function formatRange(r: Range): string {
	const start = formatRef(r.top, r.left);
	const end = formatRef(r.bottom, r.right);
	return start === end ? start : `${start}:${end}`;
}

export function inRange(r: Range, row: number, col: number): boolean {
	return row >= r.top && row <= r.bottom && col >= r.left && col <= r.right;
}
