import { setIcon } from "obsidian";
import { CellPopup } from "./cell-popup";
import type { CellData, CellStyle, GridData, RowData, WarningData } from "./grid-data";

const ROW_NUMBER_PX = 44;

export interface GridCallbacks {
	onCommit(cell: CellData, text: string): void;
	onRefresh(): void;
	onWarningAction(warningId: string, actionId: string): void;
}

interface ActiveEdit {
	r: number;
	c: number;
	input: HTMLInputElement;
	finish(commit: boolean, move?: [number, number]): void;
}

/**
 * Draws a grid and handles cell editing. Knows nothing about files or Excel:
 * it renders GridData and reports edits through the callbacks.
 */
export class GridView {
	private titleEl?: HTMLElement;
	private statusEl?: HTMLElement;
	private warningsEl?: HTMLElement;
	private tableEl?: HTMLTableElement;
	private tbody?: HTMLTableSectionElement;
	private noteEl?: HTMLElement;
	private columnsKey = "";
	private rowNumbers = true;
	private colEls: HTMLTableColElement[] = [];
	private rowEls: HTMLTableRowElement[] = [];
	private data: CellData[][] = [];
	private editing: ActiveEdit | null = null;
	private pendingFocus: { r: number; c: number } | null = null;
	private statusTimer?: number;

	constructor(
		private root: HTMLElement,
		private callbacks: GridCallbacks,
	) {
		root.addClass("obxcel");
	}

	/** Draws the grid, updating the existing one in place so an open editor survives refreshes. */
	render(grid: GridData) {
		const rowNumbers = grid.rowNumbers !== false;
		const columnHeader = grid.columnHeader !== false;
		const columnsKey = [rowNumbers, columnHeader, ...grid.columns].join("\u0000");
		if (!this.tbody || columnsKey !== this.columnsKey) this.build(grid.columns, rowNumbers, columnHeader);
		this.columnsKey = columnsKey;

		this.titleEl!.setText(grid.title);
		this.titleEl!.title = grid.tooltip;
		this.renderWarnings(grid.warnings ?? []);
		this.sizeColumns(grid);
		this.data = grid.rows.map((row) => row.cells);
		grid.rows.forEach((row, i) => this.renderRow(row, i));
		while (this.rowEls.length > grid.rows.length) {
			if (this.editing && this.editing.r === this.rowEls.length - 1) this.editing.finish(false);
			this.rowEls.pop()!.remove();
		}

		this.noteEl!.setText(grid.note ?? "");
		this.noteEl!.toggle(!!grid.note);

		const focus = this.pendingFocus;
		this.pendingFocus = null;
		if (focus && !this.editing) this.startEdit(focus.r, focus.c);
	}

	showError(message: string) {
		this.reset();
		this.root.createDiv({ cls: "obxcel-error", text: `Obxcel: ${message}` });
	}

	setStatus(text: string) {
		if (!this.statusEl) return;
		this.statusEl.setText(text);
		window.clearTimeout(this.statusTimer);
		this.statusTimer = window.setTimeout(() => this.statusEl?.setText(""), 2000);
	}

	private reset() {
		this.editing?.finish(false);
		this.root.empty();
		this.tbody = undefined;
		this.columnsKey = "";
		this.rowEls = [];
		this.editing = null;
	}

	private build(columns: string[], rowNumbers: boolean, columnHeader: boolean) {
		this.reset();
		this.rowNumbers = rowNumbers;
		const bar = this.root.createDiv("obxcel-toolbar");
		this.titleEl = bar.createSpan("obxcel-title");
		this.statusEl = bar.createSpan("obxcel-status");
		const refresh = bar.createEl("button", { cls: "clickable-icon", attr: { "aria-label": "Reload from disk" } });
		setIcon(refresh, "refresh-cw");
		refresh.onclick = () => this.callbacks.onRefresh();
		this.warningsEl = this.root.createDiv("obxcel-warnings");

		const wrap = this.root.createDiv("obxcel-wrap");
		// Keep clicks from moving the editor cursor into the code block source.
		wrap.addEventListener("mousedown", (e) => e.stopPropagation());
		const table = wrap.createEl("table", { cls: "obxcel-table" });
		this.tableEl = table;
		const colgroup = table.createEl("colgroup");
		if (rowNumbers) colgroup.createEl("col", { cls: "obxcel-col-rownum" });
		this.colEls = columns.map(() => colgroup.createEl("col"));
		if (columnHeader) {
			const head = table.createEl("thead").createEl("tr");
			if (rowNumbers) head.createEl("th", { cls: "obxcel-corner" });
			for (const label of columns) head.createEl("th", { text: label, attr: { title: label } });
		}

		this.tbody = table.createEl("tbody");
		// Columns share the page width, so long values are cut off: show them in full on hover.
		this.tbody.addEventListener("mouseover", (e) => {
			const td = (e.target as HTMLElement).closest("td");
			if (td && !td.hasAttribute("title") && td.scrollWidth > td.clientWidth) td.title = td.textContent ?? "";
		});
		this.tbody.addEventListener("click", (e) => {
			const td = (e.target as HTMLElement).closest("td");
			const tr = td?.parentElement as HTMLTableRowElement | null;
			if (!td || !tr) return;
			const r = this.rowEls.indexOf(tr);
			if (r !== -1) this.startEdit(r, Array.prototype.indexOf.call(tr.children, td) - this.offset);
		});
		this.noteEl = this.root.createDiv("obxcel-note");
	}

	/**
	 * Sizes the columns from the given widths (Excel's), or else from how long their
	 * contents are. A fitted grid splits the page width in those proportions; an
	 * unfitted one gives each column its real width and scrolls sideways if needed.
	 */
	private sizeColumns(grid: GridData) {
		const widths =
			grid.columnWidths ??
			grid.columns.map((label, j) => {
				let longest = label.length;
				for (const row of grid.rows.slice(0, 200)) longest = Math.max(longest, row.cells[j]?.text.length ?? 0);
				return Math.min(Math.max(longest, 4), 40);
			});

		if (grid.fit !== false) {
			const total = widths.reduce((a, b) => a + b, 0);
			this.tableEl!.setCssStyles({ width: "100%" });
			widths.forEach((w, j) => this.colEls[j]?.setCssStyles({ width: `${((w / total) * 100).toFixed(2)}%` }));
			return;
		}

		// Excel's pixel width of a column for a width in characters (Calibri 11: 7px per character + padding).
		const pixels = widths.map((w) => Math.round(w * 7 + 5));
		pixels.forEach((px, j) => this.colEls[j]?.setCssStyles({ width: `${px}px` }));
		const rowNumberWidth = this.rowNumbers ? ROW_NUMBER_PX : 0;
		this.tableEl!.setCssStyles({ width: `${pixels.reduce((a, b) => a + b, rowNumberWidth)}px` });
	}

	private renderWarnings(warnings: WarningData[]) {
		const el = this.warningsEl!;
		el.empty();
		for (const warning of warnings) {
			const box = el.createDiv("obxcel-warning");
			setIcon(box.createSpan("obxcel-warning-icon"), "alert-triangle");
			box.createSpan({ cls: "obxcel-warning-text", text: warning.text });
			for (const action of warning.actions) {
				const button = box.createEl("button", { text: action.label });
				button.onclick = () => this.callbacks.onWarningAction(warning.id, action.id);
			}
		}
	}

	private renderRow(row: RowData, i: number) {
		let tr = this.rowEls[i];
		if (!tr) {
			tr = this.tbody!.createEl("tr");
			if (this.rowNumbers) tr.createEl("th", { cls: "obxcel-rownum" });
			row.cells.forEach(() => tr.createEl("td"));
			this.rowEls.push(tr);
		}
		tr.className = `obxcel-row-${row.kind}`;
		if (this.rowNumbers) (tr.firstElementChild as HTMLElement).setText(row.label);
		row.cells.forEach((cell, j) => {
			const td = tr.children[j + this.offset] as HTMLTableCellElement | undefined;
			if (td) this.paint(td, cell);
		});
	}

	/** Children of a row before its first cell (the row number). */
	private get offset(): number {
		return this.rowNumbers ? 1 : 0;
	}

	private cellEl(r: number, c: number): HTMLTableCellElement | undefined {
		return this.rowEls[r]?.children[c + this.offset] as HTMLTableCellElement | undefined;
	}

	private paint(td: HTMLTableCellElement, cell: CellData) {
		if (this.editing && this.cellEl(this.editing.r, this.editing.c) === td) return;
		td.className = "";
		td.removeAttribute("style");
		// Every grid position keeps a <td> (so indexes stay simple); merged-over ones are hidden.
		td.colSpan = cell.span?.cols ?? 1;
		td.rowSpan = cell.span?.rows ?? 1;
		if (cell.covered) {
			td.empty();
			td.addClass("obxcel-covered");
			return;
		}
		td.setText(cell.pending ? cell.editText : cell.text);
		td.toggleClass("obxcel-num", !!cell.numeric);
		td.toggleClass("obxcel-formula", !!cell.formula);
		td.toggleClass("obxcel-pending", !!cell.pending);
		td.toggleClass("obxcel-readonly", !cell.editable);
		td.toggleClass("obxcel-has-list", !!cell.options?.length && cell.editable);
		td.toggleClass("obxcel-table-header", !!cell.tableHeader);
		td.toggleClass("obxcel-wrap-text", !!cell.style?.wrap);
		if (cell.style) applyStyle(td, cell.style);
		if (cell.hint) td.title = cell.hint;
		else td.removeAttribute("title");
	}

	private startEdit(r: number, c: number) {
		if (this.editing?.r === r && this.editing.c === c) return;
		this.editing?.finish(true);

		const cell = this.data[r]?.[c];
		const td = this.cellEl(r, c);
		if (!cell || !td || !cell.editable) return;

		const original = cell.editText;
		td.empty();
		td.addClass("obxcel-editing");
		const input = td.createEl("input", { type: "text", value: original });
		const popup =
			cell.options?.length || cell.help
				? new CellPopup(td, input, cell.options ?? [], cell.help, (value) => {
						input.value = value;
						finish(true);
					})
				: null;

		const finish = (commit: boolean, move?: [number, number]) => {
			if (this.editing?.input !== input) return;
			this.editing = null;
			popup?.destroy();
			const text = input.value;
			// The grid may have been refreshed while editing; use the latest data for this cell.
			const current = this.data[r]?.[c] ?? cell;
			td.empty();
			if (commit && text !== original) {
				td.className = "obxcel-saving";
				td.setText(text);
				this.callbacks.onCommit(current, text);
			} else {
				this.paint(td, current);
			}

			if (move) {
				let [nr, nc] = [r + move[0], c + move[1]];
				// Step over cells hidden under a merged cell.
				while (this.data[nr]?.[nc]?.covered) [nr, nc] = [nr + move[0], nc + move[1]];
				if (this.cellEl(nr, nc)) this.startEdit(nr, nc);
				// The next row may appear once the save finishes (e.g. a new table row).
				else if (move[0] > 0) this.pendingFocus = { r: nr, c: nc };
			}
		};
		this.editing = { r, c, input, finish };

		// Enter/Tab take the highlighted choice of the dropdown, if there is one.
		const takeChoice = () => {
			const choice = popup?.activeValue;
			if (choice !== undefined) input.value = choice;
		};

		input.addEventListener("keydown", (e) => {
			e.stopPropagation();
			if (popup?.hasOptions && (e.key === "ArrowDown" || e.key === "ArrowUp")) {
				e.preventDefault();
				popup.move(e.key === "ArrowDown" ? 1 : -1);
			} else if (e.key === "Enter") {
				e.preventDefault();
				takeChoice();
				finish(true, [e.shiftKey ? -1 : 1, 0]);
			} else if (e.key === "Tab") {
				e.preventDefault();
				takeChoice();
				finish(true, [0, e.shiftKey ? -1 : 1]);
			} else if (e.key === "Escape") {
				e.preventDefault();
				finish(false);
			}
		});
		input.addEventListener("blur", () => finish(true));
		input.focus();
		input.select();
	}
}

function applyStyle(td: HTMLElement, style: CellStyle) {
	td.toggleClass("obxcel-bold", !!style.bold);
	td.toggleClass("obxcel-italic", !!style.italic);
	td.toggleClass("obxcel-underline", !!style.underline);
	td.toggleClass("obxcel-strike", !!style.strike);
	if (style.align) td.addClass(`obxcel-align-${style.align}`);
	// Colours come from the workbook, so they can only be set per cell.
	if (style.color) td.setCssStyles({ color: style.color });
	if (style.background) td.setCssStyles({ backgroundColor: style.background });
}
