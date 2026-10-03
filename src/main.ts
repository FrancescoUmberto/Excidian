import { FileSystemAdapter, Plugin } from "obsidian";
import { ControllerContext, SheetController } from "./controller/sheet-controller";
import { DeviceSettings } from "./model/device-settings";
import { detectOneDriveRoots } from "./model/locations";
import { ExcidianSettingTab } from "./view/settings-tab";

export default class ExcidianPlugin extends Plugin {
	async onload() {
		const context: ControllerContext = {
			settings: new DeviceSettings({
				load: (key): unknown => {
					const value: unknown = this.app.loadLocalStorage(key);
					return value;
				},
				save: (key, value) => this.app.saveLocalStorage(key, value),
			}),
			active: new Set(),
			vaultRoot: this.app.vault.adapter instanceof FileSystemAdapter ? this.app.vault.adapter.getBasePath() : undefined,
		};

		this.addSettingTab(
			new ExcidianSettingTab(this.app, this, {
				settings: context.settings,
				detectOneDriveRoots: () => detectOneDriveRoots(),
				onChange: () => context.active.forEach((controller) => controller.restart()),
			}),
		);

		this.registerMarkdownCodeBlockProcessor("excidian", (source, el, ctx) => {
			ctx.addChild(new SheetController(el, source, ctx.sourcePath, context));
		});
	}
}
