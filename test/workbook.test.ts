import assert from "node:assert/strict";
import { describe, it } from "node:test";
import JSZip from "jszip";
import { parseBlockConfig } from "../src/controller/block-config";
import { buildGrid } from "../src/controller/grid-builder";
import { formatDate, formatNumber } from "../src/model/number-format";
import { shiftFormula } from "../src/model/ooxml/formula";
import { parseInput } from "../src/model/values";
import { Workbook } from "../src/model/workbook";
import { financeWorkbook } from "./fixture";

const tableConfig = parseBlockConfig("file: /tmp/finance.xlsx\ntable: movimenti");

async function entry(bytes: Uint8Array, path: string): Promise<string | undefined> {
	return (await JSZip.loadAsync(bytes)).file(path)?.async("string");
}

describe("reading", () => {
	it("finds sheets and tables", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		assert.deepEqual(wb.sheets.map((s) => s.name), ["Riepilogo", "2026", "Liste"]);
		const { table, sheet } = wb.table("MOVIMENTI");
		assert.equal(table.name, "Movimenti");
		assert.equal(sheet.name, "2026");
		assert.deepEqual(table.columns.map((c) => c.name), ["Mese", "Data", "Importo", "Categoria", "Saldo"]);
		assert.throws(() => wb.table("Nope"), /Tables: Movimenti \(2026\), TabListe \(Liste\)/);
	});

	it("builds a table grid with formatted values and a row for new entries", async () => {
		const grid = buildGrid(await Workbook.load(await financeWorkbook()), tableConfig);
		assert.equal(grid.title, "Movimenti · 2026");
		assert.deepEqual(grid.columns, ["Mese", "Data", "Importo", "Categoria", "Saldo"]);
		assert.deepEqual(grid.rows.map((r) => r.kind), ["data", "data", "new"]);

		const [first, second] = grid.rows;
		assert.equal(first.cells[0].text, "ottobre");
		assert.equal(first.cells[1].text, "01/10/2026");
		assert.equal(first.cells[1].editText, "2026-10-01");
		assert.match(first.cells[2].text, /^1.500.00 €$/);
		assert.equal(first.cells[3].text, "Stipendio");
		assert.equal(second.cells[4].editText, "=SUM($C$2:C3)", "shared formula child");
		assert.ok(second.cells[4].formula);
	});

	it("protects table headers and merged cells in sheet mode", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const sheet2 = buildGrid(wb, parseBlockConfig("file: /tmp/x.xlsx\nsheet: 2026"));
		assert.equal(sheet2.rows[0].cells[0].editable, false);
		assert.equal(sheet2.rows[1].cells[0].editable, true);
		const summary = buildGrid(wb, parseBlockConfig("file: /tmp/x.xlsx\nsheet: Riepilogo"));
		assert.equal(summary.rows[2].cells[0].text, "Riepilogo annuale");
		assert.equal(summary.rows[2].cells[1].editable, false, "merged cell");
	});
});

describe("writing", () => {
	it("appends a row to a table, filling calculated columns and keeping formats", async () => {
		const original = await financeWorkbook();
		const wb = await Workbook.load(original);
		const location = wb.table("Movimenti");
		wb.writeTableCell(location, 4, 2, parseInput("05/10/2026"));
		wb.writeTableCell(location, 4, 3, parseInput("-12,50"));
		wb.writeTableCell(location, 4, 4, parseInput("Caffè"));
		const saved = await wb.save();

		const reread = await Workbook.load(saved);
		const { table, sheet } = reread.table("Movimenti");
		assert.equal(table.ref.bottom, 4);
		assert.match((await entry(saved, "xl/tables/table1.xml"))!, /<autoFilter ref="A1:E4">/);
		assert.equal(sheet.cell(4, 1).formula, 'TEXT(Movimenti[[#This Row],[Data]],"mmmm")');
		assert.equal(sheet.cell(4, 5).formula, "SUM($C$2:C4)");
		assert.deepEqual(sheet.cell(4, 2).value, new Date(Date.UTC(2026, 9, 5)));
		assert.equal(sheet.cell(4, 2).format, "dd/mm/yyyy");
		assert.equal(sheet.cell(4, 3).value, -12.5);
		assert.equal(sheet.cell(4, 3).format, '#,##0.00\\ "€"');
		assert.equal(sheet.cell(4, 4).value, "Caffè");

		const grid = buildGrid(reread, tableConfig);
		assert.deepEqual(grid.rows.map((r) => r.kind), ["data", "data", "data", "new"]);
		assert.ok(grid.rows[2].cells[4].pending, "new formula waits for Excel to calculate it");
	});

	it("leaves untouched parts byte for byte and resets the calculation chain", async () => {
		const original = await financeWorkbook();
		const wb = await Workbook.load(original);
		wb.sheet("2026").setCell(2, 4, parseInput("Bonus"));
		const saved = await wb.save();

		for (const path of ["xl/worksheets/sheet1.xml", "xl/sharedStrings.xml", "xl/styles.xml", "xl/tables/table1.xml"]) {
			assert.equal(await entry(saved, path), await entry(original, path), path);
		}
		assert.equal(await entry(saved, "xl/calcChain.xml"), undefined);
		assert.doesNotMatch((await entry(saved, "xl/_rels/workbook.xml.rels"))!, /calcChain/);
		assert.doesNotMatch((await entry(saved, "[Content_Types].xml"))!, /calcChain/);
		assert.match((await entry(saved, "xl/workbook.xml"))!, /<calcPr calcId="191029" fullCalcOnLoad="1"\/>/);

		const sheetXml = (await entry(saved, "xl/worksheets/sheet2.xml"))!;
		assert.match(sheetXml, /^<\?xml version="1.0" encoding="UTF-8" standalone="yes"\?>/);
		assert.match(sheetXml, /<c r="D2" t="inlineStr"><is><t>Bonus<\/t><\/is><\/c>/);
		assert.doesNotMatch(sheetXml, /xmlns=""/);
	});

	it("keeps the rest of a shared formula group working when its first cell changes", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		const sheet = wb.sheet("2026");
		sheet.setCell(2, 5, parseInput("0"));
		const reread = (await Workbook.load(await wb.save())).sheet("2026");
		assert.equal(reread.cell(2, 5).value, 0);
		assert.equal(reread.cell(2, 5).formula, undefined);
		assert.equal(reread.cell(3, 5).formula, "SUM($C$2:C3)");
	});

	it("clears a cell but keeps its formatting", async () => {
		const wb = await Workbook.load(await financeWorkbook());
		wb.sheet("2026").setCell(3, 3, parseInput(""));
		const xml = (await entry(await wb.save(), "xl/worksheets/sheet2.xml"))!;
		assert.match(xml, /<c r="C3" s="3"\/>/);
	});

	it("refuses to grow a table that has a totals row", async () => {
		const wb = await Workbook.load(await financeWorkbook({ totalsRow: true }));
		const location = wb.table("Movimenti");
		assert.match(wb.appendBlocker(location)!, /totals row/);
		assert.throws(() => wb.writeTableCell(location, 5, 2, parseInput("1")), /totals row/);
		assert.throws(() => wb.writeTableCell(location, 2, 6, parseInput("1")), /outside the table/);
		const grid = buildGrid(wb, tableConfig);
		assert.deepEqual(grid.rows.map((r) => r.kind), ["data", "data", "totals"]);
		assert.match(grid.note!, /totals row/);
	});
});

describe("helpers", () => {
	it("shifts relative references only", () => {
		assert.equal(shiftFormula("SUM($C$2:C2)", 1, 0), "SUM($C$2:C3)");
		assert.equal(shiftFormula('A1&"B2"', 1, 1), 'B2&"B2"');
		assert.equal(shiftFormula("LOG10(A1)+Sheet1!B$3", 2, 0), "LOG10(A3)+Sheet1!B$3");
		assert.equal(shiftFormula("Tab1[Col2]+Tab1[[#This Row],[AB2]]", 1, 0), "Tab1[Col2]+Tab1[[#This Row],[AB2]]");
	});

	it("parses typed input", () => {
		assert.deepEqual(parseInput("12,50"), { kind: "number", value: 12.5 });
		assert.deepEqual(parseInput("=SUM(A1:A3)"), { kind: "formula", formula: "SUM(A1:A3)" });
		assert.deepEqual(parseInput("31/02/2026"), { kind: "text", value: "31/02/2026" });
		assert.deepEqual(parseInput("  "), { kind: "empty" });
	});

	it("formats numbers and dates", () => {
		const d = new Date(Date.UTC(2026, 9, 3, 14, 5));
		assert.equal(formatDate(d, "dd/mm/yyyy hh:mm"), "03/10/2026 14:05");
		assert.equal(formatDate(d, "yyyy-mm-dd"), "2026-10-03");
		assert.equal(formatNumber(0.256, "0.0%"), "25.6%".replace(".", (0.5).toLocaleString().charAt(1)));
		assert.match(formatNumber(-45, '#,##0.00 "€";[Red]-#,##0.00 "€"'), /^-45.00 €$/);
		assert.match(formatNumber(-45, "#,##0.00;(#,##0.00)"), /^\(45.00\)$/);
		assert.equal(formatNumber(12345, "0.00E+00"), "1.23E+04");
		assert.equal(formatNumber(0.1 + 0.2, "General"), (0.3).toLocaleString());
	});

	it("validates block settings", () => {
		assert.throws(() => parseBlockConfig("table: X"), /Missing `file:`/);
		assert.throws(() => parseBlockConfig("file: /tmp/a.xlsx\ntable: X\nsheet: Y"), /work with `sheet:`/);
		assert.deepEqual(parseBlockConfig("file: /tmp/a.xlsx\nsheet: 2026\nrange: B2:D").target, {
			kind: "sheet",
			sheet: "2026",
			range: { top: 2, left: 2, right: 4, bottom: undefined },
			columns: undefined,
		});
	});
});
