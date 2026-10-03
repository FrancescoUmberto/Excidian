import { colName, colNumber } from "./refs";

// An A1 reference not glued to a name, a function call or a structured reference.
const A1_REF = /(^|[^A-Za-z0-9_.[\]])(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?![A-Za-z0-9_([])/g;

/**
 * Moves the relative references of a formula by the given offset, the way Excel
 * does when a formula is copied (used for shared formulas and new table rows).
 * Text inside double quotes is left alone.
 */
export function shiftFormula(formula: string, dRow: number, dCol: number): string {
	if (dRow === 0 && dCol === 0) return formula;
	return formula
		.split('"')
		.map((part, i) =>
			i % 2 === 1
				? part
				: part.replace(A1_REF, (match, pre: string, colAbs: string, col: string, rowAbs: string, row: string) => {
						const c = colAbs ? colNumber(col) : colNumber(col) + dCol;
						const r = rowAbs ? Number(row) : Number(row) + dRow;
						if (c < 1 || r < 1) return match;
						return `${pre}${colAbs}${colName(c)}${rowAbs}${r}`;
					}),
		)
		.join('"');
}
