import { LOCALE_SHORT_DATE } from "./ooxml/styles";
import type { CellValue } from "./values";

/** Shows a cell value the way its Excel format code asks for, in the user's locale. */
export function formatValue(value: CellValue, code: string): string {
	if (value === null) return "";
	if (value instanceof Date) return formatDate(value, code);
	if (typeof value === "number") return formatNumber(value, code);
	if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
	if (typeof value === "object") return value.error;
	return value;
}

function general(n: number): string {
	if (Number.isInteger(n) && Math.abs(n) < 1e11) return String(n);
	const rounded = Number(n.toPrecision(10));
	if (Math.abs(rounded) >= 1e11 || (rounded !== 0 && Math.abs(rounded) < 1e-9)) return rounded.toExponential(5).toUpperCase();
	return rounded.toLocaleString(undefined, { maximumFractionDigits: 10, useGrouping: false });
}

/** Splits a format code on ";" outside quotes and escapes. */
function sections(code: string): string[] {
	const out: string[] = [];
	let current = "";
	let quoted = false;
	for (let i = 0; i < code.length; i++) {
		const ch = code[i];
		if (ch === '"') quoted = !quoted;
		if (ch === "\\" && !quoted) {
			current += ch + (code[++i] ?? "");
			continue;
		}
		if (ch === ";" && !quoted) {
			out.push(current);
			current = "";
		} else current += ch;
	}
	out.push(current);
	return out;
}

/** The colour a format code gives a number, e.g. "red" for -5 in `#,##0;[Red]-#,##0`. */
export function numberColor(n: number, code: string): string | undefined {
	if (!code || /^general$/i.test(code)) return undefined;
	const parts = sections(code);
	let section = parts[0];
	if (n < 0 && parts[1]) section = parts[1];
	else if (n === 0 && parts[2]) section = parts[2];
	return section.match(/\[(red|green|blue|magenta|cyan|yellow|black|white)\]/i)?.[1].toLowerCase();
}

export function formatNumber(n: number, code: string): string {
	if (!code || /^general$/i.test(code) || code === "@") return general(n);
	const parts = sections(code);
	let section = parts[0];
	let value = n;
	let minus = n < 0;
	if (n < 0 && parts[1] !== undefined && parts[1] !== "") {
		section = parts[1];
		value = -n;
		minus = false;
	} else if (n === 0 && parts[2] !== undefined && parts[2] !== "") {
		section = parts[2];
	}
	if (minus) value = -value;

	// Currency blocks like [$€-410] become literal text; colours and conditions are dropped.
	section = section.replace(/\[\$([^\]-]*)(?:-[^\]]*)?\]/g, (_, symbol: string) => `"${symbol}"`);
	section = section.replace(/\[[^\]]*\]/g, "");

	let out = "";
	let pattern = "";
	let numberAt = -1;
	let percent = false;
	for (let i = 0; i < section.length; i++) {
		const ch = section[i];
		if (ch === '"') {
			const end = section.indexOf('"', i + 1);
			out += section.slice(i + 1, end === -1 ? undefined : end);
			i = end === -1 ? section.length : end;
		} else if (ch === "\\") {
			out += section[++i] ?? "";
		} else if (ch === "_") {
			out += " ";
			i++;
		} else if (ch === "*") {
			i++;
		} else if (ch === "%") {
			percent = true;
			out += "%";
		} else if (/[0#?]/.test(ch) || (/[.,]/.test(ch) && /[0#?]/.test(section[i + 1] ?? ""))) {
			if (numberAt === -1) numberAt = out.length;
			if (numberAt === out.length) {
				// Collect the whole number pattern, including an exponent part.
				let j = i;
				while (j < section.length && /[0#?.,]/.test(section[j])) j++;
				if (/[eE]/.test(section[j] ?? "") && /[+-]/.test(section[j + 1] ?? "")) {
					j += 2;
					while (j < section.length && /[0#?]/.test(section[j])) j++;
				}
				pattern += section.slice(i, j);
				i = j - 1;
			}
		} else {
			out += ch;
		}
	}
	if (numberAt === -1) return (minus ? "-" : "") + out;

	if (percent) value *= 100;
	const decimals = pattern.match(/\.([0#?]+)/)?.[1].length ?? 0;
	let digits: string;
	if (/[eE]/.test(pattern)) {
		const [mantissa, exp] = value.toExponential(decimals).split("e");
		const e = Number(exp);
		digits = `${mantissa}E${e < 0 ? "-" : "+"}${String(Math.abs(e)).padStart(2, "0")}`;
	} else {
		digits = value.toLocaleString(undefined, {
			minimumFractionDigits: decimals,
			maximumFractionDigits: decimals,
			useGrouping: /[0#?],[0#?]/.test(pattern),
		});
	}
	return (minus ? "-" : "") + out.slice(0, numberAt) + digits + out.slice(numberAt);
}

const DATE_TOKEN = /"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|./gi;

export function formatDate(d: Date, code: string): string {
	const section = sections(code)[0].replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, "").replace(/\[([hms])\]/gi, "$1");
	if (section === LOCALE_SHORT_DATE) return d.toLocaleDateString(undefined, { timeZone: "UTC" });

	const tokens = section.match(DATE_TOKEN) ?? [];
	const twelveHour = tokens.some((t) => /^(AM\/PM|A\/P)$/i.test(t));
	const kinds = tokens.map((t) => t.toLowerCase());
	const pad = (n: number) => String(n).padStart(2, "0");
	const hours = d.getUTCHours();

	return tokens
		.map((token, i) => {
			const lower = kinds[i];
			if (token.startsWith('"')) return token.slice(1, -1);
			if (token.startsWith("\\")) return token.slice(1);
			if (lower.startsWith("m") && lower.length <= 2) {
				// "m" means minutes right after hours or right before seconds.
				const prev = kinds.slice(0, i).reverse().find((k) => /^[ymdhs]/.test(k));
				const next = kinds.slice(i + 1).find((k) => /^[ymdhs]/.test(k));
				const minutes = prev?.startsWith("h") || next?.startsWith("s");
				const value = minutes ? d.getUTCMinutes() : d.getUTCMonth() + 1;
				return lower === "mm" ? pad(value) : String(value);
			}
			switch (lower) {
				case "yyyy":
					return String(d.getUTCFullYear());
				case "yy":
					return pad(d.getUTCFullYear() % 100);
				case "mmmmm":
					return d.toLocaleString(undefined, { month: "narrow", timeZone: "UTC" });
				case "mmmm":
					return d.toLocaleString(undefined, { month: "long", timeZone: "UTC" });
				case "mmm":
					return d.toLocaleString(undefined, { month: "short", timeZone: "UTC" });
				case "dddd":
					return d.toLocaleString(undefined, { weekday: "long", timeZone: "UTC" });
				case "ddd":
					return d.toLocaleString(undefined, { weekday: "short", timeZone: "UTC" });
				case "dd":
					return pad(d.getUTCDate());
				case "d":
					return String(d.getUTCDate());
				case "hh":
					return pad(twelveHour ? hours % 12 || 12 : hours);
				case "h":
					return String(twelveHour ? hours % 12 || 12 : hours);
				case "ss":
					return pad(d.getUTCSeconds());
				case "s":
					return String(d.getUTCSeconds());
				case "am/pm":
					return hours < 12 ? "AM" : "PM";
				case "a/p":
					return hours < 12 ? "A" : "P";
				default:
					return token;
			}
		})
		.join("");
}
