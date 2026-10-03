import { shiftFormula } from "./ooxml/formula";
import type { Package } from "./ooxml/package";
import { formatRange, formatRef, inRange, MAX_COL, MAX_ROW, parseRange, parseRef, Range } from "./ooxml/refs";
import type { CellFormat, Styles } from "./ooxml/styles";
import { child, createChild, elements, removeChildren, richText, setText, textOf } from "./ooxml/xml";
import type { Table } from "./table";
import { readValidations, ValidationRule } from "./validation";
import { CellInput, CellValue, dateToSerial, serialToDate } from "./values";

export interface CellInfo {
	value: CellValue;
	formula?: string;
	/** Excel number format code, e.g. `#,##0.00 "€"`. */
	format: string;
	/** Font, fill and alignment. */
	look: CellFormat;
}

export interface ColumnInfo {
	/** In Excel's unit: roughly the number of characters that fit. */
	width: number;
	hidden: boolean;
	style?: number;
}

// Excel's width of a column nobody resized (Calibri 11).
const DEFAULT_COLUMN_WIDTH = 8.43;

/** Workbook-wide data a sheet needs to read and write cells. */
export interface SheetContext {
	pkg: Package;
	styles: Styles;
	sharedStrings: string[];
	date1904: boolean;
}

interface CellEntry {
	row: number;
	col: number;
	el: Element;
}

interface SharedMaster {
	formula: string;
	row: number;
	col: number;
}

const key = (row: number, col: number) => `${row}:${col}`;

/** A worksheet (xl/worksheets/sheetN.xml). Cell XML is indexed on first use. */
export class Sheet {
	private rows = new Map<number, Element>();
	private cells = new Map<string, CellEntry>();
	private sharedMasters = new Map<string, SharedMaster>();
	private indexed = false;
	private rules?: ValidationRule[];
	private columnList?: { min: number; max: number; info: ColumnInfo }[];
	private defaultWidth = DEFAULT_COLUMN_WIDTH;

	constructor(
		readonly name: string,
		private path: string,
		private ctx: SheetContext,
		readonly tables: Table[],
	) {}

	private get root(): Element {
		return this.ctx.pkg.xml(this.path)!.documentElement;
	}

	private get sheetData(): Element {
		return child(this.root, "sheetData") ?? createChild(this.root, "sheetData");
	}

	private index() {
		if (this.indexed) return;
		this.indexed = true;
		let rowNum = 0;
		for (const rowEl of elements(this.sheetData, "row")) {
			const r = rowEl.getAttribute("r");
			rowNum = r ? Number(r) : rowNum + 1;
			rowEl.setAttribute("r", String(rowNum));
			this.rows.set(rowNum, rowEl);

			let colNum = 0;
			for (const el of elements(rowEl, "c")) {
				const ref = el.getAttribute("r");
				colNum = ref ? parseRef(ref).col : colNum + 1;
				el.setAttribute("r", formatRef(rowNum, colNum));
				this.cells.set(key(rowNum, colNum), { row: rowNum, col: colNum, el });

				const f = child(el, "f");
				const si = f?.getAttribute("si");
				if (f && si !== null && si !== undefined && f.getAttribute("t") === "shared" && textOf(f)) {
					this.sharedMasters.set(si, { formula: textOf(f), row: rowNum, col: colNum });
				}
			}
		}
	}

	cell(row: number, col: number): CellInfo {
		this.index();
		const el = this.cells.get(key(row, col))?.el;
		const style = this.styleOf(el, row, col);
		return {
			value: el ? this.readValue(el, style) : null,
			formula: el ? this.readFormula(el, row, col) : undefined,
			format: this.ctx.styles.formatCode(style),
			look: this.ctx.styles.cellFormat(style),
		};
	}

	/** A cell's own style, else its row's, else its column's (how empty cells get formatted). */
	private styleOf(el: Element | undefined, row: number, col: number): number {
		const own = el?.getAttribute("s");
		if (own) return Number(own);
		const rowEl = this.rows.get(row);
		if (rowEl?.getAttribute("customFormat") === "1" && rowEl.getAttribute("s")) return Number(rowEl.getAttribute("s"));
		return this.column(col).style ?? 0;
	}

	column(col: number): ColumnInfo {
		if (!this.columnList) {
			const format = child(this.root, "sheetFormatPr");
			const base = format?.getAttribute("baseColWidth");
			this.defaultWidth = Number(format?.getAttribute("defaultColWidth") ?? (base ? Number(base) + 0.71 : DEFAULT_COLUMN_WIDTH));
			const cols = child(this.root, "cols");
			this.columnList = (cols ? elements(cols, "col") : []).map((c) => ({
				min: Number(c.getAttribute("min")),
				max: Number(c.getAttribute("max")),
				info: {
					width: c.getAttribute("width") ? Number(c.getAttribute("width")) : this.defaultWidth,
					hidden: c.getAttribute("hidden") === "1" || c.getAttribute("hidden") === "true",
					style: c.getAttribute("style") ? Number(c.getAttribute("style")) : undefined,
				},
			}));
		}
		return this.columnList.find((c) => col >= c.min && col <= c.max)?.info ?? { width: this.defaultWidth, hidden: false };
	}

	isRowHidden(row: number): boolean {
		this.index();
		const hidden = this.rows.get(row)?.getAttribute("hidden");
		return hidden === "1" || hidden === "true";
	}

	/** Whether the cell holds anything (a value or a formula), ignoring formatting. */
	hasContent(row: number, col: number): boolean {
		this.index();
		const el = this.cells.get(key(row, col))?.el;
		return !!el && (!!child(el, "v") || !!child(el, "f") || !!child(el, "is"));
	}

	/** From A1 to the furthest cell with content, or null for an empty sheet. */
	usedRange(): Range | null {
		this.index();
		let bottom = 0;
		let right = 0;
		for (const { row, col } of this.cells.values()) {
			if (!this.hasContent(row, col)) continue;
			bottom = Math.max(bottom, row);
			right = Math.max(right, col);
		}
		return bottom ? { top: 1, left: 1, bottom, right } : null;
	}

	mergedRanges(): Range[] {
		const list = child(this.root, "mergeCells");
		return (list ? elements(list, "mergeCell") : []).map((m) => parseRange(m.getAttribute("ref") ?? "A1"));
	}

	private readValue(el: Element, style: number): CellValue {
		const v = child(el, "v");
		const raw = v ? textOf(v) : undefined;
		switch (el.getAttribute("t")) {
			case "s":
				return raw === undefined ? null : (this.ctx.sharedStrings[Number(raw)] ?? "");
			case "inlineStr": {
				const is = child(el, "is");
				return is ? richText(is) : "";
			}
			case "str":
				return raw ?? "";
			case "b":
				return raw === undefined ? null : raw === "1";
			case "e":
				return raw ? { error: raw } : null;
			case "d":
				return raw ? new Date(raw) : null;
			default: {
				if (raw === undefined || raw === "") return null;
				const n = Number(raw);
				return this.ctx.styles.isDate(style) ? serialToDate(n, this.ctx.date1904) : n;
			}
		}
	}

	private readFormula(el: Element, row: number, col: number): string | undefined {
		const f = child(el, "f");
		if (!f) return undefined;
		const own = textOf(f);
		if (own) return own;
		// A shared-formula child cell: derive its formula from the group's first cell.
		const master = this.sharedMasters.get(f.getAttribute("si") ?? "");
		return master ? shiftFormula(master.formula, row - master.row, col - master.col) : undefined;
	}

	validations(): ValidationRule[] {
		return (this.rules ??= readValidations(this.root));
	}

	validationAt(row: number, col: number): ValidationRule | undefined {
		return this.validations().find((rule) => rule.ranges.some((r) => inRange(r, row, col)));
	}

	/** Grows rules that end on the row above down to `row`, as Excel does when a table grows. */
	extendValidations(row: number, left: number, right: number) {
		for (const rule of this.validations()) {
			let changed = false;
			const ranges = rule.ranges.map((r) => {
				if (r.bottom !== row - 1 || r.left > right || r.right < left) return r;
				changed = true;
				return { ...r, bottom: row };
			});
			if (changed) {
				rule.setRanges(ranges);
				this.markDirty();
			}
		}
	}

	isTableHeader(row: number, col: number): boolean {
		return this.tables.some((t) => t.headerRows > 0 && row === t.ref.top && inRange(t.ref, row, col));
	}

	setCell(row: number, col: number, input: CellInput) {
		if (row < 1 || row > MAX_ROW || col < 1 || col > MAX_COL) throw new Error(`Cell ${formatRef(row, col)} is outside the sheet.`);
		this.index();
		const k = key(row, col);
		let el = this.cells.get(k)?.el;

		const f = el && child(el, "f");
		if (f?.getAttribute("t") === "shared" && f.getAttribute("ref")) this.unshare(f.getAttribute("si") ?? "");

		if (input.kind === "empty") {
			if (!el) return;
			clearContent(el);
			if (!el.hasAttribute("s")) {
				el.parentNode?.removeChild(el);
				this.cells.delete(k);
			}
			this.markDirty();
			return;
		}

		el ??= this.createCell(row, col);
		clearContent(el);
		switch (input.kind) {
			case "number":
				setText(createChild(el, "v"), String(input.value));
				break;
			case "text": {
				el.setAttribute("t", "inlineStr");
				const t = createChild(createChild(el, "is"), "t");
				setText(t, input.value);
				if (/^\s|\s$/.test(input.value)) t.setAttribute("xml:space", "preserve");
				break;
			}
			case "date":
				el.setAttribute("s", String(this.ctx.styles.dateStyle(Number(el.getAttribute("s") ?? 0))));
				setText(createChild(el, "v"), String(dateToSerial(input.value, this.ctx.date1904)));
				break;
			case "formula":
				setText(createChild(el, "f"), input.formula);
				break;
		}
		this.growDimension(row, col);
		this.markDirty();
	}

	/** Creates empty, formatted cells across a row so it looks like the row above. */
	prepareRow(row: number, left: number, right: number) {
		this.index();
		for (let col = left; col <= right; col++) {
			if (!this.cells.has(key(row, col))) this.createCell(row, col);
		}
		this.growDimension(row, right);
		this.markDirty();
	}

	private createCell(row: number, col: number): Element {
		let rowEl = this.rows.get(row);
		if (!rowEl) {
			const next = elements(this.sheetData, "row").find((r) => Number(r.getAttribute("r")) > row) ?? null;
			rowEl = createChild(this.sheetData, "row", next);
			rowEl.setAttribute("r", String(row));
			this.rows.set(row, rowEl);
		}
		// "spans" is only a loading hint and would be wrong once the row gains cells.
		rowEl.removeAttribute("spans");

		const nextCell = elements(rowEl, "c").find((c) => parseRef(c.getAttribute("r")!).col > col) ?? null;
		const el = createChild(rowEl, "c", nextCell);
		el.setAttribute("r", formatRef(row, col));

		// New rows keep the formatting of the row above (dates, currency, ...), except below a header.
		const above = this.cells.get(key(row - 1, col))?.el.getAttribute("s");
		if (above && above !== "0" && !this.isTableHeader(row - 1, col)) el.setAttribute("s", above);

		this.cells.set(key(row, col), { row, col, el });
		return el;
	}

	/** Gives every cell of a shared-formula group its own formula, so its first cell can change. */
	private unshare(si: string) {
		const master = this.sharedMasters.get(si);
		if (!master) return;
		for (const { row, col, el } of this.cells.values()) {
			const f = child(el, "f");
			if (!f || f.getAttribute("t") !== "shared" || f.getAttribute("si") !== si) continue;
			setText(f, shiftFormula(master.formula, row - master.row, col - master.col));
			f.removeAttribute("t");
			f.removeAttribute("si");
			f.removeAttribute("ref");
		}
		this.sharedMasters.delete(si);
	}

	private growDimension(row: number, col: number) {
		const dim = child(this.root, "dimension");
		const ref = dim?.getAttribute("ref");
		if (!dim || !ref) return;
		const r = parseRange(ref);
		const grown = {
			top: Math.min(r.top, row),
			left: Math.min(r.left, col),
			bottom: Math.max(r.bottom, row),
			right: Math.max(r.right, col),
		};
		dim.setAttribute("ref", formatRange(grown));
	}

	private markDirty() {
		this.ctx.pkg.markDirty(this.path);
	}
}

function clearContent(el: Element) {
	for (const attr of ["t", "cm", "vm"]) el.removeAttribute(attr);
	removeChildren(el);
}
