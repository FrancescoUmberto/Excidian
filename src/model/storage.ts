// Where workbooks are read from and written to, independent of the platform.
// Desktop implements this with Node's file system (any file on the computer);
// mobile with Obsidian's vault API (files inside the vault only).
// Nothing here may import Node modules: this file also runs on mobile.

/** Another file in the workbook's folder. */
export interface SiblingFile {
	name: string;
	modified: Date;
	/** Shows the file in the system file manager, where the platform can. */
	reveal?: () => void;
}

/** A workbook file on some storage. */
export interface StoredFile {
	/** Unique key for this file (its absolute or vault path). */
	readonly id: string;
	/** File name with extension, e.g. "bank.xlsx". */
	readonly name: string;
	/** Where the file is, for tooltips and messages. */
	readonly displayPath: string;
	read(): Promise<Uint8Array>;
	/** Replaces the file's content, as safely as the platform allows. */
	write(bytes: Uint8Array): Promise<void>;
	/** Creates or replaces another file in the same folder. */
	writeSibling(name: string, bytes: Uint8Array): Promise<void>;
	/** The other files in the same folder. */
	siblings(): Promise<SiblingFile[]>;
	/** Calls `onChange` (debounced) when this file or a conflict copy of it changes. Returns a stop function. */
	watch(onChange: () => void): () => void;
}

/** Turns the `file:` value of a code block into a file, or throws a message for the user. */
export interface FileResolver {
	resolve(spec: string, notePath: string): StoredFile;
}

/** Splits "bank.xlsx" into ["bank", ".xlsx"]. */
export function splitName(name: string): [string, string] {
	const dot = name.lastIndexOf(".");
	return dot > 0 ? [name.slice(0, dot), name.slice(dot)] : [name, ""];
}

/** "bank.xlsx" → "bank.excidian-backup.xlsx" */
export function backupName(name: string): string {
	const [stem, ext] = splitName(name);
	return `${stem}.excidian-backup${ext}`;
}

/**
 * Whether `candidate` looks like a sync-conflict copy of `name`: OneDrive appends the
 * device name ("bank-MacBook-Pro.xlsx"), other clients a number ("bank (1).xlsx").
 */
export function isConflictCopyName(name: string, candidate: string): boolean {
	const [stem, ext] = splitName(name);
	if (candidate.startsWith("~$") || !candidate.startsWith(stem) || !candidate.endsWith(ext)) return false;
	if (candidate.length <= stem.length + ext.length) return false;
	const middle = candidate.slice(stem.length, candidate.length - ext.length);
	return /^(-.+| \(.+\))$/.test(middle);
}

/**
 * Resolves the `file:` value of a block to a path inside the vault, the only place
 * mobile apps can reach: `./x.xlsx` and `../x.xlsx` from the note's folder, other
 * relative paths from the vault root.
 */
export function resolveInVault(spec: string, notePath: string): string {
	const s = spec.trim();
	if (/^onedrive:/i.test(s) || s.startsWith("/") || s.startsWith("~") || s.startsWith("\\") || /^[A-Za-z]:[\\/]/.test(s)) {
		throw new Error(
			`"${s}" is outside the vault, which Obsidian mobile can't reach. Keep the workbook in the vault and use a path ` +
				"relative to the note (./bank.xlsx) or to the vault (Finance/bank.xlsx).",
		);
	}
	const fromNote = /^\.\.?([\\/]|$)/.test(s);
	const parts = fromNote ? notePath.split("/").slice(0, -1) : [];
	for (const segment of s.replace(/\\/g, "/").split("/")) {
		if (segment === "..") {
			if (!parts.length) throw new Error(`"${s}" points outside the vault.`);
			parts.pop();
		} else if (segment && segment !== ".") {
			parts.push(segment);
		}
	}
	if (!parts.length) throw new Error(`"${s}" isn't a file.`);
	return parts.join("/");
}
