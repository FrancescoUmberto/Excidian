// Workbooks inside the vault, through Obsidian's vault API. Works on every platform;
// used on mobile, where the vault is the only place an app can reach.
// Only types are imported from "obsidian", so this can be tested with a fake vault.
import type { App, EventRef, TAbstractFile, TFile } from "obsidian";
import { FileResolver, isConflictCopyName, resolveInVault, SiblingFile, StoredFile } from "../model/storage";

function isFile(f: TAbstractFile | null): f is TFile {
	return !!f && "stat" in f;
}

/** Copies the bytes into an ArrayBuffer of their own, as the vault API expects. */
function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

export class VaultFile implements StoredFile {
	readonly id: string;
	readonly name: string;
	readonly displayPath: string;
	private folder: string;

	constructor(
		private app: App,
		vaultPath: string,
	) {
		this.id = vaultPath;
		this.displayPath = vaultPath;
		const slash = vaultPath.lastIndexOf("/");
		this.name = vaultPath.slice(slash + 1);
		this.folder = slash === -1 ? "" : vaultPath.slice(0, slash);
	}

	async read(): Promise<Uint8Array> {
		const file = this.app.vault.getFileByPath(this.id);
		if (!isFile(file)) throw new Error(`File not found in the vault: ${this.id}`);
		return new Uint8Array(await this.app.vault.readBinary(file));
	}

	async write(bytes: Uint8Array) {
		await this.writePath(this.id, bytes);
	}

	async writeSibling(name: string, bytes: Uint8Array) {
		await this.writePath(this.folder ? `${this.folder}/${name}` : name, bytes);
	}

	private async writePath(vaultPath: string, bytes: Uint8Array) {
		const existing = this.app.vault.getFileByPath(vaultPath);
		if (isFile(existing)) await this.app.vault.modifyBinary(existing, toArrayBuffer(bytes));
		else await this.app.vault.createBinary(vaultPath, toArrayBuffer(bytes));
	}

	async siblings(): Promise<SiblingFile[]> {
		const folder = this.folder ? this.app.vault.getFolderByPath(this.folder) : this.app.vault.getRoot();
		return (folder?.children ?? [])
			.filter(isFile)
			.filter((f) => f.name !== this.name)
			.map((f) => ({ name: f.name, modified: new Date(f.stat.mtime) }));
	}

	watch(onChange: () => void): () => void {
		let timer: number | undefined;
		const relevant = (vaultPath: string | undefined) => {
			if (!vaultPath) return false;
			const slash = vaultPath.lastIndexOf("/");
			const folder = slash === -1 ? "" : vaultPath.slice(0, slash);
			const name = vaultPath.slice(slash + 1);
			return folder === this.folder && (name === this.name || isConflictCopyName(this.name, name));
		};
		const handler = (file: TAbstractFile, oldPath?: string) => {
			if (!relevant(file.path) && !relevant(oldPath)) return;
			window.clearTimeout(timer);
			timer = window.setTimeout(onChange, 300);
		};
		const refs: EventRef[] = [
			this.app.vault.on("modify", handler),
			this.app.vault.on("create", handler),
			this.app.vault.on("delete", handler),
			this.app.vault.on("rename", handler),
		];
		return () => {
			window.clearTimeout(timer);
			refs.forEach((ref) => this.app.vault.offref(ref));
		};
	}
}

/** Resolves block paths inside the vault; paths outside it get a clear message. */
export class VaultResolver implements FileResolver {
	constructor(private app: App) {}

	resolve(spec: string, notePath: string): StoredFile {
		return new VaultFile(this.app, resolveInVault(spec, notePath));
	}
}
