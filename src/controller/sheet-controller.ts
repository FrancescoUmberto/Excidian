import { dirname, join } from "path";
import { MarkdownRenderChild, Notice } from "obsidian";
import type { DeviceSettings } from "../model/device-settings";
import { ConflictCopy, findConflictCopies } from "../model/locations";
import { formatRef } from "../model/ooxml/refs";
import type { ValidationIssue } from "../model/validation";
import { CellInput, parseInput } from "../model/values";
import type { Sheet } from "../model/sheet";
import type { TableLocation, Workbook } from "../model/workbook";
import { isOpenInExcel, readWorkbook, updateWorkbook, watchFile } from "../model/workbook-file";
import type { CellData, GridData, WarningData } from "../view/grid-data";
import { GridView } from "../view/grid-view";
import { BlockConfig, parseBlockConfig } from "./block-config";
import { buildGrid } from "./grid-builder";

/** Services the plugin shares with every code block. */
export interface ControllerContext {
	settings: DeviceSettings;
	/** Absolute path of the vault folder, for relative `file:` paths. */
	vaultRoot?: string;
	/** Controllers currently on screen, so settings changes can restart them. */
	active: Set<SheetController>;
}

/** Connects one obxcel code block to its workbook: loads it, shows it, saves edits. */
export class SheetController extends MarkdownRenderChild {
	private view: GridView;
	private cfg?: BlockConfig;
	private grid?: GridData;
	private conflicts: ConflictCopy[] = [];
	private stopWatching?: () => void;
	private ignoreChangesUntil = 0;
	private warnedOpenInExcel = false;

	constructor(
		containerEl: HTMLElement,
		private source: string,
		/** Vault-relative path of the note containing the block. */
		private notePath: string,
		private context: ControllerContext,
	) {
		super(containerEl);
		this.view = new GridView(containerEl, {
			onCommit: (cell, text) => void this.commit(cell, text),
			onRefresh: () => void this.reload(),
			onWarningAction: (warningId, actionId) => this.onWarningAction(warningId, actionId),
		});
	}

	onload() {
		this.context.active.add(this);
		this.start();
	}

	onunload() {
		this.context.active.delete(this);
		this.stop();
	}

	/** Re-reads the block settings, e.g. after the OneDrive folder changed. */
	restart() {
		this.stop();
		this.start();
	}

	private start() {
		try {
			const { vaultRoot, settings } = this.context;
			this.cfg = parseBlockConfig(this.source, {
				oneDriveRoot: settings.oneDriveRoot,
				vaultRoot,
				noteDir: vaultRoot ? dirname(join(vaultRoot, this.notePath)) : undefined,
			});
		} catch (e) {
			this.cfg = undefined;
			this.view.showError(errorMessage(e));
			return;
		}
		this.stopWatching = watchFile(this.cfg.file, () => {
			// Our own saves also change the file; those are already on screen.
			if (Date.now() >= this.ignoreChangesUntil) void this.reload();
		});
		void this.reload();
	}

	private stop() {
		this.stopWatching?.();
		this.stopWatching = undefined;
	}

	private async reload() {
		if (!this.cfg) return;
		try {
			this.show(await readWorkbook(this.cfg.file));
		} catch (e) {
			this.view.showError(errorMessage(e));
		}
	}

	private show(wb: Workbook) {
		this.conflicts = findConflictCopies(this.cfg!.file);
		this.grid = { ...buildGrid(wb, this.cfg!), warnings: this.conflictWarnings() };
		this.view.render(this.grid);
	}

	private conflictWarnings(): WarningData[] {
		return this.conflicts
			.filter((copy) => !this.context.settings.isDismissed(conflictId(copy)))
			.map((copy) => ({
				id: conflictId(copy),
				text:
					`Possible sync conflict copy: ${copy.name} (changed ${copy.modified.toLocaleString()}). ` +
					"Some entries may be in that file instead of this one. Merge them in Excel, then delete the copy.",
				actions: [
					{ id: "reveal", label: "Show file" },
					{ id: "dismiss", label: "Dismiss" },
				],
			}));
	}

	private onWarningAction(warningId: string, actionId: string) {
		const copy = this.conflicts.find((c) => conflictId(c) === warningId);
		if (!copy) return;
		if (actionId === "reveal") {
			revealInFileManager(copy.path);
		} else if (actionId === "dismiss" && this.grid) {
			this.context.settings.dismiss(warningId);
			this.grid = { ...this.grid, warnings: this.conflictWarnings() };
			this.view.render(this.grid);
		}
	}

	private async commit(cell: CellData, text: string) {
		const cfg = this.cfg!;
		const input = parseInput(text);

		if (!this.warnedOpenInExcel && isOpenInExcel(cfg.file)) {
			this.warnedOpenInExcel = true;
			new Notice("Obxcel: this file is open in Excel. Saving there will overwrite edits made here; close it or reload it in Excel.", 8000);
		}

		this.ignoreChangesUntil = Date.now() + 10_000;
		let warning = null as ValidationIssue | null;
		try {
			const wb = await updateWorkbook(cfg.file, (wb) => {
				// Checked against the fresh file, right before writing.
				const issue = checkEdit(wb, cfg, cell, input);
				if (issue?.blocking) throw new Error(issueText(issue));
				warning = issue;
				applyEdit(wb, cfg, cell, input);
			});
			this.show(wb);
			this.view.setStatus(`Saved ${formatRef(cell.row, cell.col)}`);
			if (warning) new Notice(`Obxcel: saved, but ${issueText(warning)}`, 8000);
		} catch (e) {
			new Notice(`Obxcel: could not save ${formatRef(cell.row, cell.col)}: ${errorMessage(e)}`);
			if (this.grid) this.view.render(this.grid);
		} finally {
			this.ignoreChangesUntil = Date.now() + 1000;
		}
	}
}

/**
 * The table an edit goes through: the block's table, or in sheet mode the table the
 * cell is in (or just below), so typing under a table grows it like in Excel.
 */
function editTarget(wb: Workbook, cfg: BlockConfig, cell: CellData): { sheet: Sheet; location?: TableLocation } {
	if (cfg.target.kind === "table") {
		const location = wb.table(cfg.target.table);
		return { sheet: location.sheet, location };
	}
	const sheet = wb.sheet(cfg.target.sheet);
	const location = wb.tableAt(sheet, cell.row, cell.col);
	return { sheet, location: location && cell.row >= location.table.dataRange.top ? location : undefined };
}

function checkEdit(wb: Workbook, cfg: BlockConfig, cell: CellData, input: CellInput): ValidationIssue | null {
	const { sheet, location } = editTarget(wb, cfg, cell);
	return wb.checkInput(sheet, cell.row, cell.col, input, location?.table);
}

function issueText(issue: ValidationIssue): string {
	return issue.title ? `${issue.title}: ${issue.message}` : issue.message;
}

function applyEdit(wb: Workbook, cfg: BlockConfig, cell: CellData, input: CellInput) {
	const { sheet, location } = editTarget(wb, cfg, cell);
	if (location) wb.writeTableCell(location, cell.row, cell.col, input);
	else sheet.setCell(cell.row, cell.col, input);
}

/** Includes the modification time, so a new conflict with the same name warns again. */
function conflictId(copy: ConflictCopy): string {
	return `${copy.path}|${copy.modified.getTime()}`;
}

function revealInFileManager(path: string) {
	try {
		// Obsidian desktop runs on Electron, whose shell can open Finder / Explorer.
		type Electron = { shell: { showItemInFolder(path: string): void } };
		const { shell } = (window as unknown as { require(module: string): Electron }).require("electron");
		shell.showItemInFolder(path);
	} catch {
		new Notice(`Obxcel: ${path}`);
	}
}

function errorMessage(e: unknown): string {
	return e instanceof Error ? e.message : String(e);
}
