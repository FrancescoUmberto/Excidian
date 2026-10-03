import { elements } from "./xml";

/** Excel's legacy indexed palette (indexes 0–63). */
const INDEXED = [
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
	"800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
	"9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
	"000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
	"00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
	"3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
	"003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
];

/** Office's default theme, used when the workbook has none. */
const DEFAULT_THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];

/**
 * Reads the theme colours in the order cell styles refer to them. The theme file
 * lists dk1, lt1, dk2, lt2, but styles number them lt1, dk1, lt2, dk2.
 */
export function readThemeColors(themeRoot: Element | undefined): string[] {
	const scheme = themeRoot && findDescendant(themeRoot, "clrScheme");
	if (!scheme) return DEFAULT_THEME;
	const colors = elements(scheme).map((slot) => {
		const value = elements(slot)[0];
		// <a:srgbClr val="…"/> or a system colour <a:sysClr lastClr="…"/>
		return (value?.localName === "srgbClr" ? value.getAttribute("val") : value?.getAttribute("lastClr")) ?? "000000";
	});
	if (colors.length < 12) return DEFAULT_THEME;
	[colors[0], colors[1]] = [colors[1], colors[0]];
	[colors[2], colors[3]] = [colors[3], colors[2]];
	return colors;
}

function findDescendant(el: Element, localName: string): Element | undefined {
	for (const c of elements(el)) {
		if (c.localName === localName) return c;
		const found = findDescendant(c, localName);
		if (found) return found;
	}
	return undefined;
}

export interface ResolvedColor {
	/** "#RRGGBB" */
	hex: string;
	/**
	 * Plain black/white text or white background: Excel's defaults. Left out so
	 * the Obsidian theme (light or dark) decides instead.
	 */
	isDefault: boolean;
}

/** Resolves a <color>, <fgColor> or <bgColor> element. */
export function resolveColor(el: Element | undefined, theme: string[], kind: "font" | "fill"): ResolvedColor | undefined {
	if (!el || el.getAttribute("auto") === "1") return undefined;
	let rgb: string | undefined;
	let isDefault = false;

	const themeIndex = el.getAttribute("theme");
	const indexed = el.getAttribute("indexed");
	const argb = el.getAttribute("rgb");
	const tint = Number(el.getAttribute("tint") ?? 0);
	if (themeIndex !== null) {
		const i = Number(themeIndex);
		rgb = theme[i];
		// Untinted theme "Text 1" for fonts and "Background 1" for fills are the defaults.
		isDefault = !tint && (kind === "font" ? i === 1 : i === 0);
	} else if (indexed !== null) {
		const i = Number(indexed);
		if (i >= 64) return undefined; // system foreground / background
		rgb = INDEXED[i];
	} else if (argb) {
		rgb = argb.length === 8 ? argb.slice(2) : argb;
	}
	if (!rgb) return undefined;

	if (tint) rgb = applyTint(rgb, tint);
	rgb = rgb.toUpperCase();
	if (!tint) isDefault ||= kind === "font" ? rgb === "000000" : rgb === "FFFFFF";
	return { hex: `#${rgb}`, isDefault };
}

/** Lightens (tint > 0) or darkens (tint < 0) a colour the way Excel does, in HSL space. */
function applyTint(rgb: string, tint: number): string {
	const [h, s, l] = toHsl(rgb);
	const lightness = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
	return fromHsl(h, s, Math.min(1, Math.max(0, lightness)));
}

function toHsl(rgb: string): [number, number, number] {
	const [r, g, b] = [0, 2, 4].map((i) => parseInt(rgb.slice(i, i + 2), 16) / 255);
	const max = Math.max(r, g, b);
	const min = Math.min(r, g, b);
	const l = (max + min) / 2;
	if (max === min) return [0, 0, l];
	const d = max - min;
	const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
	const h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
	return [h / 6, s, l];
}

function fromHsl(h: number, s: number, l: number): string {
	const hue = (p: number, q: number, t: number) => {
		if (t < 0) t += 1;
		if (t > 1) t -= 1;
		if (t < 1 / 6) return p + (q - p) * 6 * t;
		if (t < 1 / 2) return q;
		if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
		return p;
	};
	let r = l;
	let g = l;
	let b = l;
	if (s) {
		const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
		const p = 2 * l - q;
		r = hue(p, q, h + 1 / 3);
		g = hue(p, q, h);
		b = hue(p, q, h - 1 / 3);
	}
	return [r, g, b].map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("");
}

/** Black or white, whichever reads better on the given background. */
export function contrastText(hex: string): string {
	const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
	const luminance = 0.2126 * r + 0.7152 * g + 0.0722 * b;
	return luminance > 0.55 ? "#000000" : "#FFFFFF";
}
