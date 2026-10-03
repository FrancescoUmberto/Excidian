/** Storage that stays on this device (in Obsidian: app.loadLocalStorage / saveLocalStorage). */
export interface DeviceStore {
	load(key: string): unknown;
	save(key: string, value: unknown): void;
}

interface Stored {
	oneDriveRoot?: string;
	dismissedWarnings?: string[];
}

const KEY = "obxcel-device-settings";
const MAX_DISMISSED = 50;

/**
 * Settings that differ per computer, such as where OneDrive lives. They are not
 * stored in the vault, so syncing the vault doesn't carry them to other devices.
 */
export class DeviceSettings {
	constructor(private store: DeviceStore) {}

	private read(): Stored {
		const raw = this.store.load(KEY);
		if (typeof raw !== "string") return {};
		try {
			const parsed: unknown = JSON.parse(raw);
			return parsed && typeof parsed === "object" ? (parsed as Stored) : {};
		} catch {
			return {};
		}
	}

	private write(change: Partial<Stored>) {
		this.store.save(KEY, JSON.stringify({ ...this.read(), ...change }));
	}

	get oneDriveRoot(): string {
		return this.read().oneDriveRoot ?? "";
	}

	set oneDriveRoot(value: string) {
		this.write({ oneDriveRoot: value.trim() || undefined });
	}

	isDismissed(warningId: string): boolean {
		return this.read().dismissedWarnings?.includes(warningId) ?? false;
	}

	dismiss(warningId: string) {
		const list = (this.read().dismissedWarnings ?? []).filter((id) => id !== warningId);
		this.write({ dismissedWarnings: [...list, warningId].slice(-MAX_DISMISSED) });
	}

	get dismissedCount(): number {
		return this.read().dismissedWarnings?.length ?? 0;
	}

	clearDismissed() {
		this.write({ dismissedWarnings: undefined });
	}
}
