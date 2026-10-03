import assert from "node:assert/strict";
import * as fs from "node:fs";
import * as os from "node:os";
import * as path from "node:path";
import { afterEach, beforeEach, describe, it } from "node:test";
import { DeviceSettings } from "../src/model/device-settings";
import { detectOneDriveRoots, findConflictCopies, isConflictCopyName, Machine, resolveWorkbookPath } from "../src/model/locations";

let home: string;
const mac = (): Machine => ({ platform: "darwin", home, env: {} });

function touch(...parts: string[]): string {
	const file = path.join(home, ...parts);
	fs.mkdirSync(path.dirname(file), { recursive: true });
	fs.writeFileSync(file, "");
	return file;
}

beforeEach(() => {
	home = fs.realpathSync(fs.mkdtempSync(path.join(os.tmpdir(), "excidian-home-")));
});
afterEach(() => fs.rmSync(home, { recursive: true, force: true }));

describe("OneDrive paths", () => {
	it("detects macOS CloudStorage folders, modern names first, without duplicates", () => {
		fs.mkdirSync(path.join(home, "Library/CloudStorage/OneDrive - Politecnico di Bari"), { recursive: true });
		fs.mkdirSync(path.join(home, "Library/CloudStorage/OneDrive-PolitecnicodiBari"));
		fs.mkdirSync(path.join(home, "Library/CloudStorage/GoogleDrive-me"));
		fs.symlinkSync(path.join(home, "Library/CloudStorage/OneDrive-PolitecnicodiBari"), path.join(home, "OneDrive"));
		assert.deepEqual(detectOneDriveRoots(mac()), [
			path.join(home, "Library/CloudStorage/OneDrive-PolitecnicodiBari"),
			path.join(home, "Library/CloudStorage/OneDrive - Politecnico di Bari"),
		]);
	});

	it("uses the OneDrive environment variables on Windows", () => {
		const machine: Machine = { platform: "win32", home, env: { OneDriveCommercial: "C:\\Users\\me\\OneDrive - Uni", OneDrive: "C:\\Users\\me\\OneDrive - Uni" } };
		assert.deepEqual(detectOneDriveRoots(machine), ["C:\\Users\\me\\OneDrive - Uni"]);
	});

	it("resolves onedrive: paths to the OneDrive folder that has the file", () => {
		fs.mkdirSync(path.join(home, "Library/CloudStorage/OneDrive-Personal"), { recursive: true });
		const file = touch("Library/CloudStorage/OneDrive-Uni/Finance/bank.xlsx");
		assert.equal(resolveWorkbookPath("onedrive:/Finance/bank.xlsx", {}, mac()), file);
		assert.equal(resolveWorkbookPath("OneDrive:Finance/bank.xlsx", {}, mac()), file);
	});

	it("prefers the folder set in settings", () => {
		touch("Library/CloudStorage/OneDrive-Uni/Finance/bank.xlsx");
		const custom = touch("Elsewhere/Finance/bank.xlsx");
		assert.equal(resolveWorkbookPath("onedrive:/Finance/bank.xlsx", { oneDriveRoot: "~/Elsewhere" }, mac()), custom);
	});

	it("explains what went wrong", () => {
		assert.throws(() => resolveWorkbookPath("onedrive:/bank.xlsx", {}, mac()), /No OneDrive folder found/);
		fs.mkdirSync(path.join(home, "Library/CloudStorage/OneDrive-Uni"), { recursive: true });
		assert.throws(() => resolveWorkbookPath("onedrive:/bank.xlsx", {}, mac()), /"bank.xlsx" not found in OneDrive. Looked in: .*OneDrive-Uni/);
	});

	it("leaves ordinary paths alone apart from ~", () => {
		assert.equal(resolveWorkbookPath("~/Documents/bank.xlsx", {}, mac()), path.join(home, "Documents/bank.xlsx"));
		assert.equal(resolveWorkbookPath("/data/bank.xlsx", {}, mac()), "/data/bank.xlsx");
	});
});

describe("vault-relative paths", () => {
	const vault = { vaultRoot: "/Vault", noteDir: "/Vault/Finance/2026" };

	it("resolves ./ and ../ from the note's folder", () => {
		assert.equal(resolveWorkbookPath("./Sheets/countability--01-10-2026.xlsx", vault, mac()), "/Vault/Finance/2026/Sheets/countability--01-10-2026.xlsx");
		assert.equal(resolveWorkbookPath("../bank.xlsx", vault, mac()), "/Vault/Finance/bank.xlsx");
	});

	it("resolves other relative paths from the vault root", () => {
		assert.equal(resolveWorkbookPath("Sheets/bank.xlsx", vault, mac()), "/Vault/Sheets/bank.xlsx");
	});

	it("never falls back to the process working directory", () => {
		assert.throws(() => resolveWorkbookPath("./bank.xlsx", {}, mac()), /vault folder is unknown/);
	});
});

describe("conflict copies", () => {
	it("recognises OneDrive and numbered copies only", () => {
		const file = "/x/bank.xlsx";
		assert.ok(isConflictCopyName(file, "bank-MacBook-Pro-di-Umberto.xlsx"));
		assert.ok(isConflictCopyName(file, "bank-DESKTOP-4F2K9.xlsx"));
		assert.ok(isConflictCopyName(file, "bank (1).xlsx"));
		for (const name of ["bank.xlsx", "bank.excidian-backup.xlsx", "~$bank.xlsx", "bankrupt.xlsx", "bank-copy.csv", "bank-.xlsx"]) {
			assert.equal(isConflictCopyName(file, name), false, name);
		}
	});

	it("lists copies next to the file, newest first", () => {
		const file = touch("OneDrive/bank.xlsx");
		const older = touch("OneDrive/bank-DESKTOP-1.xlsx");
		touch("OneDrive/bank-MacBook.xlsx");
		touch("OneDrive/bank.excidian-backup.xlsx");
		fs.utimesSync(older, new Date(2026, 0, 1), new Date(2026, 0, 1));
		assert.deepEqual(findConflictCopies(file).map((c) => c.name), ["bank-MacBook.xlsx", "bank-DESKTOP-1.xlsx"]);
	});
});

describe("device settings", () => {
	it("stores values and dismissed warnings in the device store", () => {
		const data = new Map<string, unknown>();
		const settings = new DeviceSettings({ load: (k) => data.get(k) ?? null, save: (k, v) => data.set(k, v) });
		assert.equal(settings.oneDriveRoot, "");
		settings.oneDriveRoot = "  ~/OneDrive  ";
		settings.dismiss("a");
		settings.dismiss("b");

		const again = new DeviceSettings({ load: (k) => data.get(k) ?? null, save: (k, v) => data.set(k, v) });
		assert.equal(again.oneDriveRoot, "~/OneDrive");
		assert.ok(again.isDismissed("a"));
		assert.equal(again.dismissedCount, 2);
		again.clearDismissed();
		assert.equal(again.isDismissed("a"), false);
		assert.equal(again.oneDriveRoot, "~/OneDrive");
	});
});
