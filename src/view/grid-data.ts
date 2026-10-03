/** Everything the grid view needs to draw. Built by the controller from the model. */
export interface GridData {
	title: string;
	tooltip: string;
	columns: string[];
	/** Column widths in Excel's unit (about one character each). */
	columnWidths?: number[];
	/**
	 * Fit the page width, splitting it in the proportions of `columnWidths`;
	 * otherwise columns get their real widths and the grid scrolls sideways.
	 */
	fit?: boolean;
	/** Show the column header row (default true). */
	columnHeader?: boolean;
	/** Show row numbers (default true). */
	rowNumbers?: boolean;
	rows: RowData[];
	note?: string;
	warnings?: WarningData[];
}

/** A banner above the grid, with buttons reported back through onWarningAction. */
export interface WarningData {
	id: string;
	text: string;
	actions: { id: string; label: string }[];
}

export type RowKind = "data" | "totals" | "new";

export interface RowData {
	label: string;
	kind: RowKind;
	cells: CellData[];
}

export interface CellData {
	/** Sheet coordinates, handed back to the controller on edit. */
	row: number;
	col: number;
	text: string;
	editText: string;
	editable: boolean;
	numeric?: boolean;
	formula?: boolean;
	/** A formula with no calculated result yet. */
	pending?: boolean;
	/** Tooltip. */
	hint?: string;
	/** Choices offered in a dropdown while editing (from a list validation rule). */
	options?: { label: string; value: string }[];
	/** Shown under the cell while editing, e.g. the rule's input message. */
	help?: string;
	/** A merged cell's size, in grid columns and rows. */
	span?: { cols: number; rows: number };
	/** Covered by a merged cell to its left or above: not drawn. */
	covered?: boolean;
	/** The header row of an Excel table shown inside a sheet. */
	tableHeader?: boolean;
	style?: CellStyle;
}

/** CSS-ready formatting. */
export interface CellStyle {
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	color?: string;
	background?: string;
	align?: "left" | "center" | "right" | "justify";
	wrap?: boolean;
}
