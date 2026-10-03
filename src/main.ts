import { Platform, Plugin } from "obsidian";
import { ControllerContext, SheetController } from "./controller/sheet-controller";
import { DeviceSettings } from "./model/device-settings";
import type { FileResolver } from "./model/storage";
import { VaultResolver } from "./platform/vault-file";
import { ExcidianSettingTab } from "./view/settings-tab";

export default class ExcidianPlugin extends Plugin {
	async onload() {
		const settings = new DeviceSettings({
			load: (key): unknown => {
				const value: unknown = this.app.loadLocalStorage(key);
				return value;
			},
			save: (key, value) => this.app.saveLocalStorage(key, value),
		});

		// Desktop reaches any file on the computer through Node; mobile has no Node,
		// so it works with files inside the vault through Obsidian's vault API.
		// The desktop module is imported dynamically so mobile never loads Node code.
		let resolver: FileResolver;
		let detectOneDriveRoots = (): string[] => [];
		if (Platform.isDesktopApp) {
			const desktop = await import("./platform/desktop");
			resolver = new desktop.DesktopResolver(this.app, settings);
			detectOneDriveRoots = () => desktop.detectOneDriveRoots();
		} else {
			resolver = new VaultResolver(this.app);
		}

		const context: ControllerContext = { settings, resolver, active: new Set() };

		this.addSettingTab(
			new ExcidianSettingTab(this.app, this, {
				settings,
				showOneDrive: Platform.isDesktopApp,
				detectOneDriveRoots,
				onChange: () => context.active.forEach((controller) => controller.restart()),
			}),
		);

		this.registerMarkdownCodeBlockProcessor("excidian", (source, el, ctx) => {
			ctx.addChild(new SheetController(el, source, ctx.sourcePath, context));
		});
	}
}
