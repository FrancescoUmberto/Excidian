import * as fs from "fs";
import * as path from "path";
import { isConflictCopyName } from "./locations";
import { Workbook } from "./workbook";

// Writes to the same file are chained so two quick edits never race each other.
const writeQueues = new Map<string, Promise<unknown>>();
// Files already backed up during this Obsidian session.
const backedUp = new Set<string>();

export function backupPath(file: string): string {
	const ext = path.extname(file);
	return path.join(path.dirname(file), `${path.basename(file, ext)}.obxcel-backup${ext}`);
}

export async function readWorkbook(file: string): Promise<Workbook> {
	let bytes: Buffer;
	try {
		bytes = await fs.promises.readFile(file);
	} catch (e) {
		const code = (e as NodeJS.ErrnoException).code;
		if (code === "ENOENT") throw new Error(`File not found: ${file}`);
		if (code === "EACCES" || code === "EPERM") throw new Error(`No permission to read ${file}`);
		throw e;
	}
	return Workbook.load(bytes);
}

/**
 * Loads the file fresh from disk, applies `change`, and saves it. Re-reading right
 * before writing keeps changes made in Excel in the meantime. Resolves to the
 * updated workbook.
 */
export function updateWorkbook(file: string, change: (wb: Workbook) => void): Promise<Workbook> {
	const previous = writeQueues.get(file) ?? Promise.resolve();
	const next = previous
		.catch(() => undefined)
		.then(async () => {
			const wb = await readWorkbook(file);
			change(wb);
			const bytes = await wb.save();

			if (!backedUp.has(file)) {
				await fs.promises.copyFile(file, backupPath(file));
				backedUp.add(file);
			}
			// Write next to the file, then swap it in, so a crash never leaves half a workbook.
			const tmp = path.join(path.dirname(file), `.${path.basename(file)}.obxcel-tmp`);
			await fs.promises.writeFile(tmp, bytes);
			await fs.promises.rename(tmp, file);
			return wb;
		});
	writeQueues.set(file, next);
	return next;
}

/** Excel keeps a "~$name.xlsx" owner file next to a workbook while it is open. */
export function isOpenInExcel(file: string): boolean {
	const dir = path.dirname(file);
	const base = path.basename(file);
	return [`~$${base}`, `~$${base.slice(2)}`].some((name) => fs.existsSync(path.join(dir, name)));
}

/**
 * Calls `onChange` (debounced) when the file, or a sync-conflict copy of it,
 * changes on disk. Returns a function that stops watching.
 */
export function watchFile(file: string, onChange: () => void): () => void {
	// Watch the folder, not the file: Excel and our own saves replace the file,
	// which would silently end a watcher attached to the old one.
	const base = path.basename(file);
	let timer: ReturnType<typeof setTimeout> | undefined;
	let watcher: fs.FSWatcher | undefined;
	try {
		watcher = fs.watch(path.dirname(file), (_event, name) => {
			if (name && name.toString() !== base && !isConflictCopyName(file, name.toString())) return;
			clearTimeout(timer);
			timer = setTimeout(onChange, 300);
		});
	} catch {
		// The folder can't be watched; manual refresh still works.
	}
	return () => {
		clearTimeout(timer);
		watcher?.close();
	};
}
