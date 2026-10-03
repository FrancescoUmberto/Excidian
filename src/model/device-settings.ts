/** Storage that stays on this device (in Obsidian: app.loadLocalStorage / saveLocalStorage). */
export interface DeviceStore {
	load(key: string): unknown;
	save(key: string, value: unknown): void;
}

interface Stored {
	oneDriveRoot?: string;
	dismissedWarnings?: string[];
}

const KEY = "excidian-device-settings";
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
		let parsed: unknown;
		try {
			parsed = JSON.parse(raw);
		} catch {
			return {};
		}
		if (!parsed || typeof parsed !== "object") return {};
		// Keep only well-formed fields, in case the stored value was damaged.
		const root = "oneDriveRoot" in parsed ? parsed.oneDriveRoot : undefined;
		const dismissed = "dismissedWarnings" in parsed ? parsed.dismissedWarnings : undefined;
		return {
			oneDriveRoot: typeof root === "string" ? root : undefined,
			dismissedWarnings: Array.isArray(dismissed) ? dismissed.filter((id): id is string => typeof id === "string") : undefined,
		};
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
