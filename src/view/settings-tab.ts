import { App, Plugin, PluginSettingTab, SettingDefinitionItem } from "obsidian";
import type { DeviceSettings } from "../model/device-settings";

export interface SettingsTabOptions {
	settings: DeviceSettings;
	detectOneDriveRoots(): string[];
	/** Called when the tab closes after something changed. */
	onChange(): void;
}

const ONEDRIVE_ROOT = "oneDriveRoot";

/**
 * Settings → Excidian, declared with Obsidian's settings API so the options also
 * show up in settings search. Values live in this device's storage, not the vault.
 */
export class ExcidianSettingTab extends PluginSettingTab {
	private changed = false;

	constructor(
		app: App,
		plugin: Plugin,
		private options: SettingsTabOptions,
	) {
		super(app, plugin);
	}

	getSettingDefinitions(): SettingDefinitionItem[] {
		const { settings } = this.options;
		const detected = this.options.detectOneDriveRoots();
		const dismissed = settings.dismissedCount;
		return [
			{
				name: "OneDrive folder",
				desc: createFragment((desc) => {
					desc.appendText("Used by ");
					desc.createEl("code", { text: "file: onedrive:/…" });
					desc.appendText(" paths. Saved on this device only, so each computer can point to its own. Leave empty to detect it automatically.");
					desc.createEl("br");
					desc.appendText(detected.length ? `Detected: ${detected.join(", ")}` : "No OneDrive folder detected on this device.");
				}),
				aliases: ["onedrive", "cloud", "sync"],
				control: { type: "text", key: ONEDRIVE_ROOT, placeholder: detected[0] ?? "/path/to/OneDrive" },
			},
			{
				name: "Show dismissed conflict warnings again",
				desc: dismissed ? `${dismissed} dismissed on this device.` : "None dismissed.",
				disabled: () => this.options.settings.dismissedCount === 0,
				action: () => {
					settings.clearDismissed();
					this.changed = true;
					this.update();
				},
			},
		];
	}

	getControlValue(key: string): unknown {
		return key === ONEDRIVE_ROOT ? this.options.settings.oneDriveRoot : undefined;
	}

	setControlValue(key: string, value: unknown) {
		if (key === ONEDRIVE_ROOT && typeof value === "string") {
			this.options.settings.oneDriveRoot = value;
			this.changed = true;
		}
	}

	hide() {
		super.hide();
		if (this.changed) {
			this.changed = false;
			this.options.onChange();
		}
	}
}
