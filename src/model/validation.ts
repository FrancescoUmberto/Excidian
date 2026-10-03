import { formatRange, parseRange, Range } from "./ooxml/refs";
import { child, elements, setText, textOf } from "./ooxml/xml";
import { CellInput, dateToSerial } from "./values";

export type ValidationType = "none" | "list" | "whole" | "decimal" | "date" | "time" | "textLength" | "custom";

/** A data validation rule (Data → Data Validation in Excel). */
export interface ValidationRule {
	type: ValidationType;
	operator: string;
	formula1?: string;
	formula2?: string;
	showDropdown: boolean;
	/** Whether Excel rejects or questions invalid values at all. */
	showError: boolean;
	errorStyle: "stop" | "warning" | "information";
	errorTitle?: string;
	error?: string;
	showPrompt: boolean;
	promptTitle?: string;
	prompt?: string;
	/** The cells the rule applies to. Change through `setRanges`. */
	ranges: Range[];
	setRanges(ranges: Range[]): void;
}

export interface ListOption {
	label: string;
	/** What gets written when the option is picked. */
	value: string;
}

export interface ValidationIssue {
	title?: string;
	message: string;
	/** "Stop" rules reject the value; warnings and information only report it. */
	blocking: boolean;
}

const truthy = (v: string | null) => v === "1" || v === "true";

/**
 * Reads the rules of a worksheet. Excel keeps most rules in <dataValidations>, but
 * rules whose list lives on another sheet go to an <x14:dataValidations> extension.
 */
export function readValidations(root: Element): ValidationRule[] {
	const rules: ValidationRule[] = [];
	const main = child(root, "dataValidations");
	for (const dv of main ? elements(main, "dataValidation") : []) {
		rules.push(
			fromElement(
				dv,
				dv.getAttribute("sqref") ?? "",
				(sqref) => dv.setAttribute("sqref", sqref),
				(name) => textOf(child(dv, name)),
			),
		);
	}

	const extLst = child(root, "extLst");
	for (const ext of extLst ? elements(extLst, "ext") : []) {
		const list = child(ext, "dataValidations");
		for (const dv of list ? elements(list, "dataValidation") : []) {
			const sqref = child(dv, "sqref");
			rules.push(
				fromElement(
					dv,
					textOf(sqref),
					(value) => sqref && setText(sqref, value),
					(name) => {
						const formula = child(dv, name);
						return formula ? textOf(child(formula, "f")) : "";
					},
				),
			);
		}
	}
	return rules;
}

function fromElement(dv: Element, sqref: string, writeSqref: (sqref: string) => void, formula: (name: string) => string): ValidationRule {
	const attr = (name: string) => dv.getAttribute(name) || undefined;
	const rule: ValidationRule = {
		type: (attr("type") ?? "none") as ValidationType,
		operator: attr("operator") ?? "between",
		formula1: formula("formula1") || undefined,
		formula2: formula("formula2") || undefined,
		// The file format stores the opposite: showDropDown="1" hides the arrow.
		showDropdown: !truthy(dv.getAttribute("showDropDown")),
		showError: truthy(dv.getAttribute("showErrorMessage")),
		errorStyle: (attr("errorStyle") ?? "stop") as ValidationRule["errorStyle"],
		errorTitle: attr("errorTitle"),
		error: attr("error"),
		showPrompt: truthy(dv.getAttribute("showInputMessage")),
		promptTitle: attr("promptTitle"),
		prompt: attr("prompt"),
		ranges: sqref
			.trim()
			.split(/\s+/)
			.filter(Boolean)
			.flatMap((ref) => {
				try {
					return [parseRange(ref)];
				} catch {
					return [];
				}
			}),
		setRanges(ranges) {
			rule.ranges = ranges;
			writeSqref(ranges.map(formatRange).join(" "));
		},
	};
	return rule;
}

const TYPE_NAMES: Record<string, string> = {
	whole: "Whole number",
	decimal: "Decimal",
	date: "Date",
	time: "Time",
	textLength: "Text length",
};

const OPERATOR_NAMES: Record<string, string> = {
	equal: "equal to",
	notEqual: "not equal to",
	greaterThan: "greater than",
	lessThan: "less than",
	greaterThanOrEqual: "at least",
	lessThanOrEqual: "at most",
};

/** A short description such as "Decimal between -100 and 100". `a`/`b` are the bounds as display text. */
export function describeRule(rule: ValidationRule, a?: string, b?: string): string {
	if (rule.type === "list") return "Pick a value from the list";
	if (rule.type === "custom") return "Custom rule (checked by Excel)";
	const what = TYPE_NAMES[rule.type];
	if (!what) return "";
	if (rule.operator === "between" || rule.operator === "notBetween") {
		return `${what} ${rule.operator === "between" ? "between" : "not between"} ${a ?? "?"} and ${b ?? "?"}`;
	}
	return `${what} ${OPERATOR_NAMES[rule.operator] ?? rule.operator} ${a ?? "?"}`;
}

function compare(operator: string, n: number, a: number, b: number | undefined): boolean {
	const lo = Math.min(a, b ?? a);
	const hi = Math.max(a, b ?? a);
	switch (operator) {
		case "notBetween":
			return n < lo || n > hi;
		case "equal":
			return n === a;
		case "notEqual":
			return n !== a;
		case "greaterThan":
			return n > a;
		case "lessThan":
			return n < a;
		case "greaterThanOrEqual":
			return n >= a;
		case "lessThanOrEqual":
			return n <= a;
		default:
			return n >= lo && n <= hi;
	}
}

export interface CheckContext {
	/** The list values, for list rules. */
	options: ListOption[];
	/** Number value of a rule formula (constant or cell reference), if it can be worked out. */
	resolveNumber(formula: string): number | undefined;
	date1904: boolean;
	/** Display text for the rule, used in the default error message. */
	description: string;
}

/** Checks a value against a rule the way Excel would. Null when the value is fine (or can't be judged). */
export function checkInput(rule: ValidationRule, input: CellInput, ctx: CheckContext): ValidationIssue | null {
	// Clearing a cell is always allowed; formulas are only known once Excel calculates them.
	if (!rule.showError || input.kind === "empty" || input.kind === "formula") return null;

	let ok: boolean;
	if (rule.type === "list") {
		if (!ctx.options.length) return null;
		const text = (input.kind === "number" ? String(input.value) : input.kind === "text" ? input.value : "").toLowerCase();
		ok = ctx.options.some((o) => o.value.toLowerCase() === text || o.label.toLowerCase() === text);
	} else if (rule.type in TYPE_NAMES) {
		const a = rule.formula1 !== undefined ? ctx.resolveNumber(rule.formula1) : undefined;
		const b = rule.formula2 !== undefined ? ctx.resolveNumber(rule.formula2) : undefined;
		if (a === undefined) return null;
		let n: number | undefined;
		if (rule.type === "textLength") n = (input.kind === "text" ? input.value : String(input.kind === "number" ? input.value : "")).length;
		else if (input.kind === "number") n = input.value;
		else if (input.kind === "date") n = dateToSerial(input.value, ctx.date1904);
		ok = n !== undefined && compare(rule.operator, n, a, b) && (rule.type !== "whole" || Number.isInteger(n));
	} else {
		return null; // custom formulas can't be evaluated here
	}
	if (ok) return null;

	const fallback =
		rule.type === "list"
			? `Pick one of: ${ctx.options.slice(0, 8).map((o) => o.label).join(", ")}${ctx.options.length > 8 ? ", …" : ""}`
			: `Allowed: ${ctx.description.charAt(0).toLowerCase()}${ctx.description.slice(1)}.`;
	return { title: rule.errorTitle, message: rule.error || fallback, blocking: rule.errorStyle === "stop" };
}
