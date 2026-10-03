// Desktop only: a workbook anywhere on the computer, read and written with Node's fs.
import * as fs from "fs";
import * as path from "path";
import { isConflictCopyName, SiblingFile, StoredFile } from "../../model/storage";

export class NodeFile implements StoredFile {
	readonly id: string;
	readonly name: string;
	readonly displayPath: string;
	private dir: string;

	constructor(absolutePath: string) {
		this.id = absolutePath;
		this.displayPath = absolutePath;
		this.name = path.basename(absolutePath);
		this.dir = path.dirname(absolutePath);
	}

	async read(): Promise<Uint8Array> {
		try {
			return await fs.promises.readFile(this.id);
		} catch (e) {
			const code = (e as NodeJS.ErrnoException).code;
			if (code === "ENOENT") throw new Error(`File not found: ${this.id}`);
			if (code === "EACCES" || code === "EPERM") throw new Error(`No permission to read ${this.id}`);
			throw e;
		}
	}

	/** Writes next to the file, then swaps it in, so a crash never leaves half a workbook. */
	async write(bytes: Uint8Array) {
		const tmp = path.join(this.dir, `.${this.name}.excidian-tmp`);
		await fs.promises.writeFile(tmp, bytes);
		await fs.promises.rename(tmp, this.id);
	}

	async writeSibling(name: string, bytes: Uint8Array) {
		await fs.promises.writeFile(path.join(this.dir, name), bytes);
	}

	async siblings(): Promise<SiblingFile[]> {
		let names: string[];
		try {
			names = await fs.promises.readdir(this.dir);
		} catch {
			return [];
		}
		const out: SiblingFile[] = [];
		for (const name of names) {
			if (name === this.name) continue;
			const full = path.join(this.dir, name);
			try {
				out.push({ name, modified: (await fs.promises.stat(full)).mtime, reveal: () => revealInFileManager(full) });
			} catch {
				// vanished meanwhile
			}
		}
		return out;
	}

	watch(onChange: () => void): () => void {
		// Watch the folder, not the file: Excel and our own saves replace the file,
		// which would silently end a watcher attached to the old one.
		let timer: number | undefined;
		let watcher: fs.FSWatcher | undefined;
		try {
			watcher = fs.watch(this.dir, (_event, changed) => {
				const name = changed?.toString();
				if (name && name !== this.name && !isConflictCopyName(this.name, name)) return;
				window.clearTimeout(timer);
				timer = window.setTimeout(onChange, 300);
			});
		} catch {
			// The folder can't be watched; manual refresh still works.
		}
		return () => {
			window.clearTimeout(timer);
			watcher?.close();
		};
	}
}

/** Opens Finder / Explorer at the file, through Electron's shell. */
function revealInFileManager(file: string) {
	type Electron = { shell: { showItemInFolder(path: string): void } };
	const { shell } = (window as unknown as { require(module: string): Electron }).require("electron");
	shell.showItemInFolder(file);
}
