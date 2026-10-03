import { contrastText, resolveColor } from "./colors";
import type { Package } from "./package";
import { child, elements } from "./xml";

/** How a cell looks. Colours are "#RRGGBB"; unset means "use the Obsidian theme". */
export interface CellFormat {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	color?: string;
	fill?: string;
	align?: "left" | "center" | "right" | "justify";
	wrap?: boolean;
}

const ALIGNMENTS: Record<string, CellFormat["align"]> = {
	left: "left",
	fill: "left",
	center: "center",
	centerContinuous: "center",
	distributed: "center",
	right: "right",
	justify: "justify",
};

/** <b/>, <i/>… count as on unless val="0"/"false". */
function flag(el: Element | undefined): boolean {
	if (!el) return false;
	const val = el.getAttribute("val");
	return val !== "0" && val !== "false" && val !== "none";
}

/** Excel's short-date format: shown in the user's locale, so it gets special handling. */
export const LOCALE_SHORT_DATE = "mm-dd-yy";

const BUILTIN_FORMATS: Record<number, string> = {
	0: "General",
	1: "0",
	2: "0.00",
	3: "#,##0",
	4: "#,##0.00",
	9: "0%",
	10: "0.00%",
	11: "0.00E+00",
	14: LOCALE_SHORT_DATE,
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mm:ss.0",
	48: "##0.0E+0",
	49: "@",
};

function isBuiltinDate(id: number): boolean {
	return (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
}

/** True when a format code shows dates or times (has d, m, y, h or s outside literals). */
export function isDateFormat(code: string): boolean {
	const bare = code
		.split(";")[0]
		.replace(/"[^"]*"/g, "")
		.replace(/\\.|[_*]./g, "")
		.replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, "");
	return /[dmyhs]/i.test(bare) && !/^general$/i.test(bare.trim());
}

/** Read access to the workbook's cell formats (cellXfs), plus adding a date format. */
export class Styles {
	private xfs: Element[] = [];
	private fonts: Element[] = [];
	private fills: Element[] = [];
	private customFormats = new Map<number, string>();
	private dateVariants = new Map<number, number>();
	private formats = new Map<number, CellFormat>();

	constructor(
		private pkg: Package,
		private path: string | undefined,
		private theme: string[],
	) {
		const root = path ? pkg.xml(path)?.documentElement : undefined;
		if (!root) return;
		const numFmts = child(root, "numFmts");
		for (const f of numFmts ? elements(numFmts, "numFmt") : []) {
			this.customFormats.set(Number(f.getAttribute("numFmtId")), f.getAttribute("formatCode") ?? "General");
		}
		const list = (name: string, item: string) => {
			const el = child(root, name);
			return el ? elements(el, item) : [];
		};
		this.xfs = list("cellXfs", "xf");
		this.fonts = list("fonts", "font");
		this.fills = list("fills", "fill");
	}

	/** Font, fill and alignment of a style, ready to display. */
	cellFormat(styleIndex: number): CellFormat {
		let format = this.formats.get(styleIndex);
		if (format) return format;
		format = {};
		const xf = this.xfs[styleIndex];
		if (xf) {
			const font = this.fonts[Number(xf.getAttribute("fontId") ?? 0)];
			if (font) {
				if (flag(child(font, "b"))) format.bold = true;
				if (flag(child(font, "i"))) format.italic = true;
				if (flag(child(font, "u"))) format.underline = true;
				if (flag(child(font, "strike"))) format.strike = true;
				const color = resolveColor(child(font, "color"), this.theme, "font");
				if (color && !color.isDefault) format.color = color.hex;
			}

			const pattern = this.fills[Number(xf.getAttribute("fillId") ?? 0)];
			const patternFill = pattern && child(pattern, "patternFill");
			if (patternFill?.getAttribute("patternType") === "solid") {
				const fill = resolveColor(child(patternFill, "fgColor"), this.theme, "fill");
				if (fill && !fill.isDefault) {
					format.fill = fill.hex;
					// Excel's default black text would vanish on a dark fill, and theme text on a light one.
					format.color ??= contrastText(fill.hex);
				}
			}

			const alignment = child(xf, "alignment");
			const align = ALIGNMENTS[alignment?.getAttribute("horizontal") ?? ""];
			if (align) format.align = align;
			const wrap = alignment?.getAttribute("wrapText");
			if (wrap === "1" || wrap === "true") format.wrap = true;
		}
		this.formats.set(styleIndex, format);
		return format;
	}

	private formatId(styleIndex: number): number {
		return Number(this.xfs[styleIndex]?.getAttribute("numFmtId") ?? 0);
	}

	formatCode(styleIndex: number): string {
		const id = this.formatId(styleIndex);
		return this.customFormats.get(id) ?? BUILTIN_FORMATS[id] ?? "General";
	}

	isDate(styleIndex: number): boolean {
		const id = this.formatId(styleIndex);
		const custom = this.customFormats.get(id);
		return custom !== undefined ? isDateFormat(custom) : isBuiltinDate(id);
	}

	/** A style like `styleIndex` but showing a date; appended to the workbook if needed. */
	dateStyle(styleIndex: number): number {
		if (this.isDate(styleIndex)) return styleIndex;
		const known = this.dateVariants.get(styleIndex);
		if (known !== undefined) return known;

		const root = this.path ? this.pkg.xml(this.path)?.documentElement : undefined;
		const cellXfs = root && child(root, "cellXfs");
		const source = this.xfs[styleIndex] ?? this.xfs[0];
		if (!cellXfs || !source) return styleIndex;

		const xf = source.cloneNode(true) as Element;
		xf.setAttribute("numFmtId", "14");
		xf.setAttribute("applyNumberFormat", "1");
		cellXfs.appendChild(xf);
		this.xfs.push(xf);
		cellXfs.setAttribute("count", String(this.xfs.length));
		this.pkg.markDirty(this.path!);

		const index = this.xfs.length - 1;
		this.dateVariants.set(styleIndex, index);
		return index;
	}
}
