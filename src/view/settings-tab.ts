import { App, Plugin, PluginSettingTab, Setting } from "obsidian";
import type { DeviceSettings } from "../model/device-settings";

export interface SettingsTabOptions {
	settings: DeviceSettings;
	detectOneDriveRoots(): string[];
	/** Called when the tab closes after something changed. */
	onChange(): void;
}

export class ObxcelSettingTab extends PluginSettingTab {
	private changed = false;

	constructor(
		app: App,
		plugin: Plugin,
		private options: SettingsTabOptions,
	) {
		super(app, plugin);
	}

	display() {
		const { containerEl } = this;
		const { settings } = this.options;
		containerEl.empty();

		const detected = this.options.detectOneDriveRoots();
		new Setting(containerEl)
			.setName("OneDrive folder")
			.setDesc(
				createFragment((desc) => {
					desc.appendText("Used by ");
					desc.createEl("code", { text: "file: onedrive:/…" });
					desc.appendText(" paths. Saved on this device only, so each computer can point to its own. Leave empty to detect it automatically.");
					desc.createEl("br");
					desc.appendText(detected.length ? `Detected: ${detected.join(", ")}` : "No OneDrive folder detected on this device.");
				}),
			)
			.addText((text) =>
				text
					.setPlaceholder(detected[0] ?? "/path/to/OneDrive")
					.setValue(settings.oneDriveRoot)
					.onChange((value) => {
						settings.oneDriveRoot = value;
						this.changed = true;
					}),
			);

		const dismissed = settings.dismissedCount;
		new Setting(containerEl)
			.setName("Dismissed conflict warnings")
			.setDesc(dismissed ? `${dismissed} dismissed on this device.` : "None dismissed.")
			.addButton((button) =>
				button
					.setButtonText("Show again")
					.setDisabled(dismissed === 0)
					.onClick(() => {
						settings.clearDismissed();
						this.changed = true;
						this.display();
					}),
			);
	}

	hide() {
		if (this.changed) {
			this.changed = false;
			this.options.onChange();
		}
	}
}
