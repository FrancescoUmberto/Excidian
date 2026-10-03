// Desktop only: resolves block paths to files anywhere on the computer, using Node.
import * as fs from "fs";
import * as os from "os";
import * as path from "path";

export interface PathOptions {
	/** This device's OneDrive folder, overriding detection. */
	oneDriveRoot?: string;
	/** Absolute path of the vault; relative paths like "Sheets/bank.xlsx" start here. */
	vaultRoot?: string;
	/** Absolute path of the note's folder; "./" and "../" paths start here. */
	noteDir?: string;
}

/** The parts of the machine that path resolution depends on (replaceable in tests). */
export interface Machine {
	platform: NodeJS.Platform;
	home: string;
	env: Record<string, string | undefined>;
}

const thisMachine = (): Machine => ({ platform: process.platform, home: os.homedir(), env: process.env });

const ONEDRIVE_PREFIX = /^onedrive:[\\/]*/i;

export function expandHome(p: string, home: string): string {
	return p === "~" || p.startsWith("~/") || p.startsWith("~\\") ? path.join(home, p.slice(1)) : p;
}

function isDirectory(p: string): boolean {
	try {
		return fs.statSync(p).isDirectory();
	} catch {
		return false;
	}
}

/** OneDrive folders on this device, most likely first. */
export function detectOneDriveRoots(machine: Machine = thisMachine()): string[] {
	const roots: string[] = [];
	if (machine.platform === "win32") {
		for (const name of ["OneDrive", "OneDriveCommercial", "OneDriveConsumer"]) {
			const value = machine.env[name];
			if (value) roots.push(value);
		}
	}
	// macOS keeps OneDrive in ~/Library/CloudStorage/OneDrive-<account>; older
	// installs and Linux clients use ~/OneDrive or "~/OneDrive - <organisation>".
	for (const dir of [path.join(machine.home, "Library", "CloudStorage"), machine.home]) {
		let names: string[];
		try {
			names = fs.readdirSync(dir);
		} catch {
			continue;
		}
		names
			.filter((n) => /^onedrive/i.test(n))
			.sort((a, b) => Number(a.includes(" ")) - Number(b.includes(" ")) || a.localeCompare(b))
			.map((n) => path.join(dir, n))
			.filter(isDirectory)
			.forEach((p) => roots.push(p));
	}

	// ~/OneDrive is often a link to the CloudStorage folder: keep each real folder once.
	const seen = new Set<string>();
	return roots.filter((root) => {
		let real = root;
		try {
			real = fs.realpathSync(root);
		} catch {
			// keep as is
		}
		if (seen.has(real)) return false;
		seen.add(real);
		return true;
	});
}

/**
 * Turns the `file:` value of a code block into an absolute path:
 * - `/abs/path` or `~/path` as is;
 * - `./x.xlsx`, `../x.xlsx` from the note's folder, like a relative link;
 * - `Sheets/x.xlsx` from the vault root;
 * - `onedrive:/path/in/OneDrive` from this device's OneDrive folder, so the same
 *   note works on every computer.
 */
export function resolveWorkbookPath(file: string, options: PathOptions = {}, machine: Machine = thisMachine()): string {
	if (!ONEDRIVE_PREFIX.test(file)) {
		const expanded = expandHome(file, machine.home);
		if (path.isAbsolute(expanded)) return path.resolve(expanded);
		const fromNote = /^\.\.?([\\/]|$)/.test(expanded);
		const base = fromNote ? (options.noteDir ?? options.vaultRoot) : options.vaultRoot;
		if (!base) throw new Error(`Can't resolve the relative path "${file}": the vault folder is unknown. Use an absolute path.`);
		return path.resolve(base, expanded);
	}

	const relative = file.replace(ONEDRIVE_PREFIX, "");
	const override = options.oneDriveRoot?.trim();
	const roots = override ? [expandHome(override, machine.home)] : detectOneDriveRoots(machine);
	if (!roots.length) throw new Error("No OneDrive folder found on this device. Set it in Settings → Excidian.");

	for (const root of roots) {
		const full = path.resolve(root, relative);
		if (fs.existsSync(full)) return full;
	}
	const where = override ? `the OneDrive folder set in Settings → Excidian (${roots[0]})` : `OneDrive. Looked in: ${roots.join(", ")}`;
	throw new Error(`"${relative}" not found in ${where}. If it's online only, make it available offline.`);
}
