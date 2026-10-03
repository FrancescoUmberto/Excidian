import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { parseBlockConfig } from "../src/controller/block-config";
import { buildGrid } from "../src/controller/grid-builder";
import { parseInput } from "../src/model/values";
import { Workbook } from "../src/model/workbook";
import { financeWorkbook, readEntry } from "./fixture";

const CATEGORIES = ["Stipendio", "Spesa", "Salute", "Sport", "Abbonamenti"];
const config = (lines: string) => parseBlockConfig(`file: /tmp/finance.xlsx\n${lines}`);

describe("data validation", () => {
	it("offers the list of a rule stored in Excel's cross-sheet extension, also on the new row", async () => {
		const grid = buildGrid(await Workbook.load(await financeWorkbook()), config("table: Movimenti"));
		const category = 3;
		for (const row of grid.rows) {
			assert.deepEqual(row.cells[category].options?.map((o) => o.label), CATEGORIES, `row ${row.label}`);
		}
		assert.equal(grid.rows[0].cells[0].options, undefined, "no rule on Mese");
		assert.equal(grid.rows[0].cells[2].help, "Decimal between -10000 and 10000");
	});

	it("resolves typed lists, named ranges and INDIRECT table columns, and shows input messages", async () => {
		const grid = buildGrid(await Workbook.load(await financeWorkbook()), config("sheet: Riepilogo\nrange: A1:B5"));
		const b = (row: number) => grid.rows[row - 1].cells[1];
		assert.deepEqual(b(2).options?.map((o) => o.label), ["Sì", "No"]);
		assert.equal(b(2).help, "Chiuso: Il mese è chiuso?");
		assert.deepEqual(b(4).options?.map((o) => o.label), CATEGORIES, "defined name → TabListe[Tipo]");
		assert.deepEqual(b(5).options?.map((o) => o.label), CATEGORIES, 'INDIRECT("TabListe[Tipo]")');
	});

	it("rejects values that break a stop rule and accepts the rest", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const { table, sheet } = wb.table("Movimenti");
		const check = (row: number, col: number, text: string) => wb.checkInput(sheet, row, col, parseInput(text), table);

		assert.deepEqual(check(2, 4, "Cinema"), { title: "Categoria", message: "Scegli una categoria dalla lista", blocking: true });
		assert.equal(check(2, 4, "sport"), null, "case-insensitive like Excel");
		assert.equal(check(4, 4, "Salute"), null, "new row uses the rule of the row above");
		assert.equal(check(4, 4, "Cinema")?.blocking, true);
		assert.equal(check(2, 3, "20000")?.message, "L'importo deve essere tra -10000 e 10000");
		assert.equal(check(2, 3, "abc")?.blocking, true, "text in a number rule");
		assert.equal(check(2, 3, "-45,50"), null);
		assert.equal(check(2, 3, ""), null, "clearing is always allowed");
		assert.equal(check(2, 3, "=1/0"), null, "formulas are left to Excel");
	});

	it("builds a default message when the rule has none, and lets non-stop rules through", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const summary = wb.sheet("Riepilogo");
		assert.match(wb.checkInput(summary, 4, 2, parseInput("Cinema"))!.message, /^Pick one of: Stipendio, Spesa, Salute, Sport, Abbonamenti$/);
		assert.equal(wb.checkInput(summary, 5, 2, parseInput("Cinema")), null, "rule without showErrorMessage");
	});

	it("grows the rules with the table, in both places Excel stores them", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		wb.writeTableCell(wb.table("Movimenti"), 4, 4, parseInput("Sport"));
		const saved = await wb.save();
		const xml = readEntry(saved, "xl/worksheets/sheet2.xml")!;
		assert.match(xml, /<xm:sqref>D2:D4<\/xm:sqref>/);
		assert.match(xml, /sqref="C2:C4"/);

		const reread = await Workbook.load(saved);
		const { table, sheet } = reread.table("Movimenti");
		assert.equal(reread.checkInput(sheet, 4, 4, parseInput("Cinema"), table)?.blocking, true);
		assert.equal(reread.checkInput(sheet, 5, 4, parseInput("Salute"), table), null);
	});
});
