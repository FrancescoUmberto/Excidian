import { formatNumber, formatValue } from "./number-format";
import { readThemeColors } from "./ooxml/colors";
import { shiftFormula } from "./ooxml/formula";
import { Package } from "./ooxml/package";
import { colNumber, formatRef, MAX_ROW, parseRange, Range } from "./ooxml/refs";
import { Styles } from "./ooxml/styles";
import { child, createChild, elements, NS_RELATIONSHIPS, richText, textOf } from "./ooxml/xml";
import { Sheet, SheetContext } from "./sheet";
import { Table } from "./table";
import { checkInput, describeRule, ListOption, ValidationIssue, ValidationRule } from "./validation";
import { CellInput, dateToSerial, editText, serialToDate } from "./values";

interface DefinedName {
	name: string;
	/** Set for names that only apply to one sheet. */
	sheet?: string;
	formula: string;
}

const MAX_LIST_CELLS = 2000;

const REL_TYPE = (name: string) => (r: { type: string }) => r.type.endsWith(`/${name}`);

// Elements that come before <calcPr> in workbook.xml, in schema order.
const BEFORE_CALC_PR = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets", "functionGroups", "externalReferences", "definedNames"];

export interface TableLocation {
	table: Table;
	sheet: Sheet;
}

export class Workbook {
	private constructor(
		private pkg: Package,
		private path: string,
		readonly sheets: Sheet[],
		private definedNames: DefinedName[],
		private date1904: boolean,
	) {}

	static async load(bytes: Uint8Array | ArrayBuffer): Promise<Workbook> {
		const pkg = await Package.load(bytes);
		const path = pkg.rels("").find(REL_TYPE("officeDocument"))?.target ?? "xl/workbook.xml";
		const doc = pkg.xml(path);
		if (!doc) throw new Error("The file is not an Excel workbook (.xlsx).");

		const rels = pkg.rels(path);
		const workbookPr = child(doc.documentElement, "workbookPr");
		const themePath = rels.find(REL_TYPE("theme"))?.target;
		const ctx: SheetContext = {
			pkg,
			styles: new Styles(pkg, rels.find(REL_TYPE("styles"))?.target, readThemeColors(themePath ? pkg.xml(themePath)?.documentElement : undefined)),
			sharedStrings: readSharedStrings(pkg, rels.find(REL_TYPE("sharedStrings"))?.target),
			date1904: ["1", "true"].includes(workbookPr?.getAttribute("date1904") ?? ""),
		};

		const sheetList = child(doc.documentElement, "sheets");
		const sheetElements = sheetList ? elements(sheetList, "sheet") : [];
		const nameList = child(doc.documentElement, "definedNames");
		const definedNames = (nameList ? elements(nameList, "definedName") : []).map((n) => {
			const local = n.getAttribute("localSheetId");
			return {
				name: n.getAttribute("name") ?? "",
				sheet: local !== null ? (sheetElements[Number(local)]?.getAttribute("name") ?? undefined) : undefined,
				formula: textOf(n),
			};
		});

		const sheets: Sheet[] = [];
		for (const s of sheetElements) {
			const name = s.getAttribute("name") ?? "";
			const id = s.getAttributeNS(NS_RELATIONSHIPS, "id") || s.getAttribute("r:id");
			const rel = rels.find((r) => r.id === id);
			if (!rel || !REL_TYPE("worksheet")(rel) || !pkg.xml(rel.target)) continue; // chart sheets etc.
			const tables = pkg
				.rels(rel.target)
				.filter(REL_TYPE("table"))
				.filter((r) => pkg.xml(r.target))
				.map((r) => new Table(pkg, r.target, name));
			sheets.push(new Sheet(name, rel.target, ctx, tables));
		}
		return new Workbook(pkg, path, sheets, definedNames, ctx.date1904);
	}

	/** A sheet by name or 1-based position; the first sheet when omitted. */
	sheet(nameOrPosition?: string): Sheet {
		if (!nameOrPosition) {
			if (!this.sheets[0]) throw new Error("The workbook has no sheets.");
			return this.sheets[0];
		}
		const byName = this.sheets.find((s) => s.name === nameOrPosition);
		if (byName) return byName;
		const position = Number(nameOrPosition);
		if (Number.isInteger(position) && this.sheets[position - 1]) return this.sheets[position - 1];
		throw new Error(`Sheet "${nameOrPosition}" not found. Sheets: ${this.sheets.map((s) => s.name).join(", ")}`);
	}

	tables(): TableLocation[] {
		return this.sheets.flatMap((sheet) => sheet.tables.map((table) => ({ table, sheet })));
	}

	/** Finds a table by name (case-insensitive, like Excel). */
	table(name: string): TableLocation {
		const found = this.tables().find(({ table }) => table.matches(name));
		if (found) return found;
		const names = this.tables().map(({ table, sheet }) => `${table.name} (${sheet.name})`);
		throw new Error(`Table "${name}" not found. ${names.length ? `Tables: ${names.join(", ")}` : "The workbook has no tables."}`);
	}

	/** The table a sheet cell belongs to, counting the row just below a table that can grow. */
	tableAt(sheet: Sheet, row: number, col: number): TableLocation | undefined {
		for (const table of sheet.tables) {
			const ref = table.ref;
			if (col < ref.left || col > ref.right) continue;
			if (row >= ref.top && row <= ref.bottom) return { table, sheet };
			if (row === ref.bottom + 1 && !this.appendBlocker({ table, sheet })) return { table, sheet };
		}
		return undefined;
	}

	/** Why a row can't be added below the table, or null if it can. */
	appendBlocker({ table, sheet }: TableLocation): string | null {
		if (table.totalsRows > 0) return "The table has a totals row. Turn it off in Excel to add rows here.";
		const ref = table.ref;
		const row = ref.bottom + 1;
		if (row > MAX_ROW) return "The table already reaches the last row of the sheet.";
		for (let col = ref.left; col <= ref.right; col++) {
			if (sheet.hasContent(row, col)) return `Cell ${formatRef(row, col)} below the table isn't empty.`;
		}
		const other = sheet.tables.find((t) => t !== table && t.ref.top <= row && t.ref.bottom >= row && t.ref.left <= ref.right && t.ref.right >= ref.left);
		if (other) return `Table ${other.name} is right below this one.`;
		return null;
	}

	/**
	 * Writes a cell of a table. Writing to the row just below the table first grows
	 * the table by one row, like typing below a table in Excel does.
	 */
	writeTableCell(location: TableLocation, row: number, col: number, input: CellInput) {
		const { table, sheet } = location;
		const ref = table.ref;
		if (col < ref.left || col > ref.right) throw new Error("That cell is outside the table.");
		if (row === ref.bottom + 1) {
			const blocker = this.appendBlocker(location);
			if (blocker) throw new Error(blocker);
			this.appendTableRow(location);
		} else if (row < table.dataRange.top || row > ref.bottom) {
			throw new Error("That cell is outside the table's rows.");
		}
		sheet.setCell(row, col, input);
	}

	private appendTableRow({ table, sheet }: TableLocation) {
		const ref = table.ref;
		const row = ref.bottom + 1;
		table.growTo(row);
		sheet.prepareRow(row, ref.left, ref.right);
		sheet.extendValidations(row, ref.left, ref.right);

		// Fill calculated columns, preferring the formula of the row above.
		table.columns.forEach((column, i) => {
			const col = ref.left + i;
			const above = sheet.cell(row - 1, col).formula;
			let formula: string | undefined;
			if (above && row - 1 >= table.dataRange.top) formula = shiftFormula(above, 1, 0);
			else if (column.calculatedFormula) formula = shiftFormula(column.calculatedFormula, row - table.dataRange.top, 0);
			if (formula) sheet.setCell(row, col, { kind: "formula", formula });
		});
	}

	/** The validation rule of a cell. A new table row uses the rules of the row above, which grow with it. */
	validationAt(sheet: Sheet, row: number, col: number, table?: Table): ValidationRule | undefined {
		const rule = sheet.validationAt(row, col);
		if (rule || !table || row !== table.ref.bottom + 1) return rule;
		return sheet.validationAt(row - 1, col);
	}

	/** The values of a list rule, whatever they come from: typed list, range, name or table column. */
	listOptions(sheet: Sheet, rule: ValidationRule): ListOption[] {
		return rule.type === "list" && rule.formula1 ? this.resolveList(sheet, rule.formula1, 0) : [];
	}

	/** E.g. "Decimal between -10000 and 10000". */
	describeValidation(sheet: Sheet, rule: ValidationRule): string {
		const show = (formula?: string) => {
			if (formula === undefined) return undefined;
			const n = this.resolveNumber(sheet, formula);
			if (n === undefined) return formula;
			if (rule.type === "date") return serialToDate(n, this.date1904).toLocaleDateString(undefined, { timeZone: "UTC" });
			if (rule.type === "time") return formatValue(serialToDate(n, this.date1904), "h:mm");
			return formatNumber(n, "General");
		};
		return describeRule(rule, show(rule.formula1), show(rule.formula2));
	}

	/** What to show while editing: the rule's input message if it has one, else its description. */
	validationHelp(sheet: Sheet, rule: ValidationRule): string {
		if (rule.showPrompt && (rule.prompt || rule.promptTitle)) return [rule.promptTitle, rule.prompt].filter(Boolean).join(": ");
		// A list explains itself through its dropdown.
		return rule.type === "list" ? "" : this.describeValidation(sheet, rule);
	}

	/** Checks a value against the cell's validation rule; null if it's fine. */
	checkInput(sheet: Sheet, row: number, col: number, input: CellInput, table?: Table): ValidationIssue | null {
		const rule = this.validationAt(sheet, row, col, table);
		if (!rule) return null;
		return checkInput(rule, input, {
			options: this.listOptions(sheet, rule),
			resolveNumber: (formula) => this.resolveNumber(sheet, formula),
			date1904: this.date1904,
			description: this.describeValidation(sheet, rule),
		});
	}

	private resolveList(sheet: Sheet, formula: string, depth: number): ListOption[] {
		if (depth > 5) return [];
		const f = formula.trim().replace(/^=/, "");

		// A typed list: "Grocery,Health,Sport"
		if (f.startsWith('"')) {
			const items = f.slice(1, f.endsWith('"') ? -1 : undefined).split(",").map((s) => s.trim()).filter(Boolean);
			return [...new Set(items)].map((s) => ({ label: s, value: s }));
		}
		// INDIRECT("Table[Column]") or INDIRECT("Sheet!A1:A9"): the usual way to point at a table column.
		const indirect = f.match(/^INDIRECT\(\s*"((?:[^"]|"")*)"\s*\)$/i);
		if (indirect) return this.resolveList(sheet, indirect[1].replace(/""/g, '"'), depth + 1);

		const structured = f.match(/^([A-Za-z_\\][\w.]*)\[\[?([^[\]]+)\]?\]$/);
		if (structured) {
			const location = this.tables().find(({ table }) => table.matches(structured[1]));
			const index = location?.table.columns.findIndex((c) => c.name.toLowerCase() === structured[2].toLowerCase()) ?? -1;
			if (!location || index < 0) return [];
			const data = location.table.dataRange;
			return this.readList(location.sheet, { ...data, left: data.left + index, right: data.left + index });
		}

		const area = this.resolveArea(sheet, f);
		if (area) return this.readList(area.sheet, area.range);

		const wanted = f.toLowerCase();
		const name =
			this.definedNames.find((n) => n.name.toLowerCase() === wanted && n.sheet === sheet.name) ??
			this.definedNames.find((n) => n.name.toLowerCase() === wanted && !n.sheet);
		return name ? this.resolveList(sheet, name.formula, depth + 1) : [];
	}

	private readList(sheet: Sheet, range: Range): ListOption[] {
		const used = sheet.usedRange();
		if (!used) return [];
		const options: ListOption[] = [];
		const seen = new Set<string>();
		let scanned = 0;
		for (let r = range.top; r <= Math.min(range.bottom, used.bottom); r++) {
			for (let c = range.left; c <= Math.min(range.right, used.right); c++) {
				if (++scanned > MAX_LIST_CELLS) return options;
				const info = sheet.cell(r, c);
				if (info.value === null || info.value === "") continue;
				const label = formatValue(info.value, info.format);
				if (seen.has(label)) continue;
				seen.add(label);
				options.push({ label, value: editText(info.value, undefined) });
			}
		}
		return options;
	}

	/** A reference like `$A$2:$A$9`, `'My sheet'!B3` or `H:H`, relative to `sheet`. */
	private resolveArea(sheet: Sheet, ref: string): { sheet: Sheet; range: Range } | undefined {
		let target = sheet;
		let area = ref;
		const qualified = ref.match(/^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/);
		if (qualified) {
			const name = (qualified[1] ?? qualified[2]).replace(/''/g, "'");
			const found = this.sheets.find((s) => s.name === name);
			if (!found) return undefined;
			target = found;
			area = qualified[3];
		}
		area = area.replace(/\$/g, "");
		const columns = area.match(/^([A-Z]{1,3}):([A-Z]{1,3})$/i);
		if (columns) return { sheet: target, range: { top: 1, bottom: MAX_ROW, left: colNumber(columns[1]), right: colNumber(columns[2]) } };
		if (/^[A-Z]{1,3}\d+(:[A-Z]{1,3}\d+)?$/i.test(area)) return { sheet: target, range: parseRange(area) };
		return undefined;
	}

	/** Number value of a rule bound: a constant, DATE(y,m,d), or a cell reference. */
	private resolveNumber(sheet: Sheet, formula: string): number | undefined {
		const f = formula.trim().replace(/^=/, "");
		if (f !== "" && Number.isFinite(Number(f))) return Number(f);
		const date = f.match(/^DATE\((\d+),(\d+),(\d+)\)$/i);
		if (date) return dateToSerial(new Date(Date.UTC(Number(date[1]), Number(date[2]) - 1, Number(date[3]))), this.date1904);
		const area = this.resolveArea(sheet, f);
		if (!area) return undefined;
		const value = area.sheet.cell(area.range.top, area.range.left).value;
		if (typeof value === "number") return value;
		if (value instanceof Date) return dateToSerial(value, this.date1904);
		return undefined;
	}

	/** Serializes the workbook. Excel recalculates formulas the next time it opens the file. */
	async save(): Promise<Uint8Array> {
		const calcChain = this.pkg.rels(this.path).find(REL_TYPE("calcChain"));
		// The calculation chain lists formula cells; edits can make it stale, and Excel rebuilds it.
		if (calcChain) this.pkg.removePart(calcChain.target, this.path);

		const root = this.pkg.xml(this.path)!.documentElement;
		let calcPr = child(root, "calcPr");
		if (!calcPr) {
			const after = elements(root).filter((el) => BEFORE_CALC_PR.includes(el.localName)).pop();
			calcPr = createChild(root, "calcPr", after ? after.nextSibling : root.firstChild);
		}
		calcPr.setAttribute("fullCalcOnLoad", "1");
		this.pkg.markDirty(this.path);

		return this.pkg.toBytes();
	}
}

function readSharedStrings(pkg: Package, path: string | undefined): string[] {
	const root = path ? pkg.xml(path)?.documentElement : undefined;
	return root ? elements(root, "si").map(richText) : [];
}
