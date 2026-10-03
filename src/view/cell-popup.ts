export interface PopupOption {
	label: string;
	value: string;
}

/**
 * The panel under a cell being edited: an optional help line and an optional list
 * of choices. Typing filters the list; arrows move the highlight. It lives in
 * <body> so the grid's scroll container doesn't clip it.
 */
export class CellPopup {
	private el: HTMLElement;
	private listEl?: HTMLElement;
	private shown: PopupOption[] = [];
	private active = -1;
	private reposition = () => this.position();

	constructor(
		private anchor: HTMLElement,
		private input: HTMLInputElement,
		private options: PopupOption[],
		help: string | undefined,
		private onPick: (value: string) => void,
	) {
		const doc = anchor.ownerDocument;
		this.el = doc.body.createDiv("obxcel-popup");
		// Clicks inside must not blur the input (which would save and close the editor).
		this.el.addEventListener("mousedown", (e) => e.preventDefault());
		if (help) this.el.createDiv({ cls: "obxcel-popup-help", text: help });

		if (options.length) {
			this.listEl = this.el.createDiv("obxcel-popup-list");
			// Start with every choice visible and the current value highlighted.
			this.show(options, options.findIndex((o) => o.value === input.value || o.label === input.value));
			input.addEventListener("input", () => this.filter(input.value));
		}

		this.position();
		doc.addEventListener("scroll", this.reposition, true);
		doc.defaultView?.addEventListener("resize", this.reposition);
	}

	/** The highlighted choice's value, if any. */
	get activeValue(): string | undefined {
		return this.shown[this.active]?.value;
	}

	get hasOptions(): boolean {
		return this.options.length > 0;
	}

	move(delta: number) {
		if (!this.shown.length) return;
		this.setActive(this.active === -1 && delta < 0 ? this.shown.length - 1 : (this.active + delta + this.shown.length) % this.shown.length);
	}

	destroy() {
		const doc = this.anchor.ownerDocument;
		doc.removeEventListener("scroll", this.reposition, true);
		doc.defaultView?.removeEventListener("resize", this.reposition);
		this.el.remove();
	}

	private filter(text: string) {
		const q = text.trim().toLowerCase();
		if (!q) return this.show(this.options, -1);
		const matches = this.options.filter((o) => o.label.toLowerCase().includes(q));
		// Prefer an option that starts with what was typed.
		const best = matches.findIndex((o) => o.label.toLowerCase().startsWith(q));
		this.show(matches, best === -1 ? 0 : best);
	}

	private show(options: PopupOption[], active: number) {
		this.shown = options;
		this.listEl!.empty();
		if (!options.length) {
			this.listEl!.createDiv({ cls: "obxcel-popup-empty", text: "No matching choice" });
		}
		options.forEach((option) => {
			const item = this.listEl!.createDiv({ cls: "obxcel-popup-item", text: option.label });
			item.addEventListener("click", () => this.onPick(option.value));
		});
		this.setActive(active);
		this.position();
	}

	private setActive(index: number) {
		const items = this.listEl ? Array.from(this.listEl.querySelectorAll<HTMLElement>(".obxcel-popup-item")) : [];
		items[this.active]?.removeClass("is-selected");
		this.active = index;
		items[index]?.addClass("is-selected");
		items[index]?.scrollIntoView({ block: "nearest" });
	}

	private position() {
		const rect = this.anchor.getBoundingClientRect();
		const viewport = this.anchor.ownerDocument.defaultView?.innerHeight ?? 800;
		this.el.setCssStyles({ left: `${rect.left}px`, minWidth: `${rect.width}px` });
		// Open upwards when there isn't room below.
		const below = viewport - rect.bottom;
		const height = this.el.offsetHeight;
		this.el.setCssStyles({ top: below < height && rect.top > below ? `${rect.top - height}px` : `${rect.bottom}px` });
	}
}
