import assert from "node:assert/strict";
import { describe, it } from "node:test";
import type { App } from "obsidian";
import { resolveInVault } from "../src/model/storage";
import { parseInput } from "../src/model/values";
import { findConflictCopies, isOpenInExcel, readWorkbook, updateWorkbook } from "../src/model/workbook-store";
import { VaultFile, VaultResolver } from "../src/platform/vault-file";
import { financeWorkbook } from "./fixture";

// The vault watcher debounces with window.setTimeout, as in Obsidian.
(globalThis as unknown as { window: typeof globalThis }).window ??= globalThis;

type Handler = (file: { path: string }, oldPath?: string) => void;

/** Just enough of Obsidian's vault API, in memory: files, folders and change events. */
function fakeVault(initial: Record<string, Uint8Array>) {
	const files = new Map(Object.entries(initial).map(([path, data]) => [path, { data, mtime: 1000 }]));
	const handlers: Handler[] = [];
	let clock = 2000;
	const fileObj = (path: string) => {
		const entry = files.get(path);
		return entry ? { path, name: path.split("/").pop()!, stat: { mtime: entry.mtime } } : null;
	};
	const emit = (path: string) => handlers.forEach((h) => h({ path }));
	const vault = {
		getFileByPath: fileObj,
		getFolderByPath: (folder: string) => ({
			children: [...files.keys()].filter((p) => p.slice(0, p.lastIndexOf("/")) === folder).map(fileObj),
		}),
		getRoot: () => ({ children: [...files.keys()].filter((p) => !p.includes("/")).map(fileObj) }),
		readBinary: async (f: { path: string }) => files.get(f.path)!.data.slice().buffer,
		modifyBinary: async (f: { path: string }, data: ArrayBuffer) => {
			files.set(f.path, { data: new Uint8Array(data), mtime: clock++ });
			emit(f.path);
		},
		createBinary: async (path: string, data: ArrayBuffer) => {
			files.set(path, { data: new Uint8Array(data), mtime: clock++ });
			emit(path);
		},
		on: (_event: string, handler: Handler) => (handlers.push(handler), handler),
		offref: (ref: Handler) => handlers.splice(handlers.indexOf(ref), 1),
	};
	return { app: { vault } as unknown as App, files, emit, handlers };
}

describe("paths on mobile", () => {
	it("resolves relative paths inside the vault", () => {
		assert.equal(resolveInVault("./Sheets/bank.xlsx", "Finance/October.md"), "Finance/Sheets/bank.xlsx");
		assert.equal(resolveInVault("../bank.xlsx", "Finance/2026/October.md"), "Finance/bank.xlsx");
		assert.equal(resolveInVault("Sheets/bank.xlsx", "Finance/October.md"), "Sheets/bank.xlsx");
		assert.equal(resolveInVault(".\\Sheets\\bank.xlsx", "October.md"), "Sheets/bank.xlsx");
	});

	it("explains that files outside the vault can't be reached", () => {
		for (const spec of ["onedrive:/Finance/bank.xlsx", "/Users/me/bank.xlsx", "~/bank.xlsx", "C:\\Users\\me\\bank.xlsx"]) {
			assert.throws(() => resolveInVault(spec, "October.md"), /outside the vault, which Obsidian mobile can't reach/, spec);
		}
		assert.throws(() => resolveInVault("../../bank.xlsx", "Finance/October.md"), /points outside the vault/);
	});
});

describe("workbooks in the vault", () => {
	it("reads, saves through the vault API and keeps a backup next to the file", async () => {
		const original = await financeWorkbook();
		const { app, files } = fakeVault({ "Finance/bank.xlsx": original });
		const file = new VaultResolver(app).resolve("./bank.xlsx", "Finance/October.md");
		assert.equal(file.displayPath, "Finance/bank.xlsx");

		await updateWorkbook(file, (wb) => wb.writeTableCell(wb.table("Movimenti"), 4, 4, parseInput("Sport")));
		const reread = await readWorkbook(file);
		assert.equal(reread.table("Movimenti").table.ref.bottom, 4);
		assert.equal(reread.sheet("2026").cell(4, 4).value, "Sport");
		assert.deepEqual(files.get("Finance/bank.excidian-backup.xlsx")?.data, original);
	});

	it("says clearly when the file isn't in the vault", async () => {
		const { app } = fakeVault({});
		await assert.rejects(readWorkbook(new VaultFile(app, "Finance/missing.xlsx")), /File not found in the vault: Finance\/missing.xlsx/);
	});

	it("finds conflict copies and Excel's owner file in the same folder", async () => {
		const bytes = new Uint8Array([1]);
		const { app } = fakeVault({
			"bank.xlsx": bytes,
			"bank-iPhone.xlsx": bytes,
			"~$bank.xlsx": bytes,
			"Other/bank (1).xlsx": bytes,
		});
		const file = new VaultFile(app, "bank.xlsx");
		assert.deepEqual((await findConflictCopies(file)).map((c) => c.name), ["bank-iPhone.xlsx"]);
		assert.equal((await findConflictCopies(file))[0].reveal, undefined, "no file manager on mobile");
		assert.equal(await isOpenInExcel(file), true);
	});

	it("reloads when the file or a conflict copy changes, and stops when asked", async () => {
		const { app, emit, handlers } = fakeVault({ "Finance/bank.xlsx": new Uint8Array([1]) });
		let calls = 0;
		const stop = new VaultFile(app, "Finance/bank.xlsx").watch(() => calls++);
		emit("Finance/bank.xlsx");
		emit("Finance/bank-iPhone.xlsx");
		emit("Finance/notes.md");
		emit("Other/bank.xlsx");
		await new Promise((r) => setTimeout(r, 350));
		assert.equal(calls, 1, "changes are debounced into one reload");
		stop();
		assert.equal(handlers.length, 0);
	});
});
