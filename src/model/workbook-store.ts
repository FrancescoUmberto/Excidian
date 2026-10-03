import { backupName, isConflictCopyName, SiblingFile, StoredFile } from "./storage";
import { Workbook } from "./workbook";

// Writes to the same file are chained so two quick edits never race each other.
const writeQueues = new Map<string, Promise<unknown>>();
// Files already backed up during this Obsidian session.
const backedUp = new Set<string>();

export async function readWorkbook(file: StoredFile): Promise<Workbook> {
	return Workbook.load(await file.read());
}

/**
 * Loads the file fresh, applies `change`, and saves it. Re-reading right before
 * writing keeps changes made in Excel (or synced from another device) meanwhile.
 * The first save of a session keeps a backup of the original next to the file.
 * Resolves to the updated workbook.
 */
export function updateWorkbook(file: StoredFile, change: (wb: Workbook) => void): Promise<Workbook> {
	const previous = writeQueues.get(file.id) ?? Promise.resolve();
	const next = previous
		.catch(() => undefined)
		.then(async () => {
			const original = await file.read();
			const wb = await Workbook.load(original);
			change(wb);
			const bytes = await wb.save();
			if (!backedUp.has(file.id)) {
				await file.writeSibling(backupName(file.name), original);
				backedUp.add(file.id);
			}
			await file.write(bytes);
			return wb;
		});
	writeQueues.set(file.id, next);
	return next;
}

/** Sync-conflict copies of the file in its folder, newest first. */
export async function findConflictCopies(file: StoredFile): Promise<SiblingFile[]> {
	const siblings = await file.siblings();
	return siblings.filter((s) => isConflictCopyName(file.name, s.name)).sort((a, b) => b.modified.getTime() - a.modified.getTime());
}

/** Excel keeps a "~$name.xlsx" owner file next to a workbook while it is open. */
export async function isOpenInExcel(file: StoredFile): Promise<boolean> {
	const owners = [`~$${file.name}`, `~$${file.name.slice(2)}`];
	return (await file.siblings()).some((s) => owners.includes(s.name));
}
