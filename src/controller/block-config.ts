import { PathOptions, resolveWorkbookPath } from "../model/locations";
import { colNumber, MAX_COL } from "../model/ooxml/refs";

export interface AreaRange {
	top: number;
	left: number;
	right: number;
	/** Undefined means "down to the last used row". */
	bottom?: number;
}

export type BlockTarget =
	| { kind: "table"; table: string }
	| {
			kind: "sheet";
			sheet?: string;
			range?: AreaRange;
			/** Column numbers to show, in this order (overrides the range's columns). */
			columns?: number[];
	  };

export interface BlockConfig {
	file: string;
	target: BlockTarget;
	/** Show column letters and row numbers (default true). */
	headings: boolean;
}

/**
 * Reads the `key: value` (or `key=value`) lines of an excidian code block:
 *
 *     file: ~/Finance/bank.xlsx   (or onedrive:/Finance/bank.xlsx)
 *     table: Movimenti            (an Excel table, or:)
 *     sheet: 2026                 (a sheet, optionally with)
 *     range: A1:F
 *     cols: [B, C, D]             (only these columns, in this order)
 *     headings: false             (hide column letters and row numbers)
 */
export function parseBlockConfig(source: string, pathOptions: PathOptions = {}): BlockConfig {
	const cfg: Record<string, string> = {};
	for (const line of source.split("\n")) {
		const m = line.match(/^\s*(\w+)\s*[:=]\s*(.+?)\s*$/);
		if (m) cfg[m[1].toLowerCase()] = m[2].replace(/^["']|["']$/g, "");
	}
	if (!cfg.file) throw new Error("Missing `file:`, the path to your .xlsx file.");
	if (cfg.table && (cfg.sheet || cfg.range || cfg.cols)) {
		throw new Error("`range:` and `cols:` work with `sheet:`. For a table, use `table:` alone (table names are unique).");
	}
	return {
		file: resolveWorkbookPath(cfg.file, pathOptions),
		target: cfg.table
			? { kind: "table", table: cfg.table }
			: {
					kind: "sheet",
					sheet: cfg.sheet,
					range: cfg.range ? parseAreaRange(cfg.range) : undefined,
					columns: cfg.cols ? parseColumnList(cfg.cols) : undefined,
				},
		headings: !/^(false|no|off|0|hide|hidden)$/i.test(cfg.headings ?? ""),
	};
}

/** Parses "[B, C, D]", "b,c,d" or "[B:D, F]" into column numbers, keeping the order given. */
export function parseColumnList(text: string): number[] {
	const items = text
		.trim()
		.replace(/^\[|\]$/g, "")
		.split(",")
		.map((s) => s.trim().replace(/\$/g, ""))
		.filter(Boolean);
	if (!items.length) throw new Error(`No columns in "${text}". Use something like [B, C, D].`);

	const columns: number[] = [];
	for (const item of items) {
		const m = item.match(/^([A-Z]{1,3})(?::([A-Z]{1,3}))?$/i);
		if (!m) throw new Error(`"${item}" isn't a column. Use letters, like [B, C, D] or [B:D, F].`);
		const from = colNumber(m[1]);
		const to = m[2] ? colNumber(m[2]) : from;
		if (Math.max(from, to) > MAX_COL) throw new Error(`Column "${item}" is beyond the last column (XFD).`);
		const step = to >= from ? 1 : -1;
		for (let c = from; c !== to + step; c += step) if (!columns.includes(c)) columns.push(c);
	}
	return columns;
}

/** Parses "A1:F200", or "A1:F" / "A:F" for an open-ended range. */
export function parseAreaRange(text: string): AreaRange {
	const m = text.trim().match(/^\$?([A-Z]{1,3})\$?(\d+)?:\$?([A-Z]{1,3})\$?(\d+)?$/i);
	if (!m) throw new Error(`Invalid range "${text}". Use something like A1:F200.`);
	const left = colNumber(m[1]);
	const right = colNumber(m[3]);
	const top = m[2] ? Number(m[2]) : 1;
	const bottom = m[4] ? Number(m[4]) : undefined;
	if (right < left || (bottom !== undefined && bottom < top)) throw new Error(`Invalid range "${text}".`);
	return { top, left, right, bottom };
}
