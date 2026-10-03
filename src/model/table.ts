import type { Package } from "./ooxml/package";
import { formatRange, parseRange, Range } from "./ooxml/refs";
import { child, elements, textOf } from "./ooxml/xml";

export interface TableColumn {
	name: string;
	/** Formula Excel fills into every row of a calculated column. */
	calculatedFormula?: string;
}

/** An Excel table (ListObject) defined in xl/tables/tableN.xml. */
export class Table {
	constructor(
		private pkg: Package,
		readonly path: string,
		readonly sheetName: string,
	) {}

	private get root(): Element {
		return this.pkg.xml(this.path)!.documentElement;
	}

	/** The name used in formulas, e.g. Movimenti[Importo]. */
	get name(): string {
		return this.root.getAttribute("displayName") || this.root.getAttribute("name") || "";
	}

	matches(name: string): boolean {
		const wanted = name.toLowerCase();
		return [this.root.getAttribute("displayName"), this.root.getAttribute("name")].some((n) => n?.toLowerCase() === wanted);
	}

	get ref(): Range {
		return parseRange(this.root.getAttribute("ref") ?? "A1");
	}

	get headerRows(): number {
		const v = this.root.getAttribute("headerRowCount");
		return v === null ? 1 : Number(v);
	}

	get totalsRows(): number {
		return Number(this.root.getAttribute("totalsRowCount") ?? 0);
	}

	/** The rows between the header and the totals row. */
	get dataRange(): Range {
		const ref = this.ref;
		return { ...ref, top: ref.top + this.headerRows, bottom: ref.bottom - this.totalsRows };
	}

	get columns(): TableColumn[] {
		const list = child(this.root, "tableColumns");
		return (list ? elements(list, "tableColumn") : []).map((col) => {
			const formula = child(col, "calculatedColumnFormula");
			return {
				name: decodeEscapes(col.getAttribute("name") ?? ""),
				calculatedFormula: formula ? textOf(formula) : undefined,
			};
		});
	}

	/** Moves the bottom edge of the table (and its filter and sort ranges) to `bottom`. */
	growTo(bottom: number) {
		const ref = { ...this.ref, bottom };
		this.root.setAttribute("ref", formatRange(ref));
		const dataBottom = bottom - this.totalsRows;
		for (const el of [child(this.root, "autoFilter"), child(this.root, "sortState")]) {
			const r = el?.getAttribute("ref");
			if (el && r) el.setAttribute("ref", formatRange({ ...parseRange(r), bottom: dataBottom }));
		}
		const nestedSort = child(this.root, "autoFilter") && child(child(this.root, "autoFilter")!, "sortState");
		const nestedRef = nestedSort?.getAttribute("ref");
		if (nestedSort && nestedRef) nestedSort.setAttribute("ref", formatRange({ ...parseRange(nestedRef), bottom: dataBottom }));
		this.pkg.markDirty(this.path);
	}
}

/** Column names store characters like line breaks as _x000A_. */
function decodeEscapes(name: string): string {
	return name.replace(/_x([0-9a-fA-F]{4})_/g, (_, hex: string) => {
		const ch = String.fromCharCode(parseInt(hex, 16));
		return /[\r\n]/.test(ch) ? " " : ch;
	});
}
