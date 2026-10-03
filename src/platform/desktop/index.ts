// Desktop only. main.ts loads this module with a dynamic import() so that its Node
// imports (fs, os, path, electron) are never evaluated on mobile, where they don't exist.
import * as path from "path";
import { App, FileSystemAdapter } from "obsidian";
import type { DeviceSettings } from "../../model/device-settings";
import type { FileResolver, StoredFile } from "../../model/storage";
import { NodeFile } from "./node-file";
import { detectOneDriveRoots, resolveWorkbookPath } from "./paths";

/** Resolves block paths anywhere on the computer: vault-relative, ~, absolute and onedrive:. */
export class DesktopResolver implements FileResolver {
	private vaultRoot?: string;

	constructor(
		app: App,
		private settings: DeviceSettings,
	) {
		this.vaultRoot = app.vault.adapter instanceof FileSystemAdapter ? app.vault.adapter.getBasePath() : undefined;
	}

	resolve(spec: string, notePath: string): StoredFile {
		const absolute = resolveWorkbookPath(spec, {
			oneDriveRoot: this.settings.oneDriveRoot,
			vaultRoot: this.vaultRoot,
			noteDir: this.vaultRoot ? path.dirname(path.join(this.vaultRoot, notePath)) : undefined,
		});
		return new NodeFile(absolute);
	}
}

export { detectOneDriveRoots };
