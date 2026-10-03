import assert from "node:assert/strict";
import { describe, it } from "node:test";
import { parseBlockConfig, parseColumnList } from "../src/controller/block-config";
import { buildGrid } from "../src/controller/grid-builder";
import { contrastText, readThemeColors, resolveColor } from "../src/model/ooxml/colors";
import { parseXml } from "../src/model/ooxml/xml";
import { Workbook } from "../src/model/workbook";
import { financeWorkbook } from "./fixture";

const config = (lines: string) => parseBlockConfig(`file: /tmp/finance.xlsx\n${lines}`);
const load = async () => Workbook.load(await financeWorkbook());

describe("sheet layout", () => {
	it("uses Excel's column widths and skips hidden rows and columns", async () => {
		const grid = buildGrid(await load(), config("sheet: Riepilogo"));
		assert.deepEqual(grid.columns, ["A", "B", "D"]);
		assert.deepEqual(grid.columnWidths, [20, 8.43, 8.43]);
		assert.deepEqual(
			grid.rows.map((r) => `${r.label}:${r.kind}`),
			["1:data", "2:data", "3:data", "4:data", "5:data", "7:new"],
		);
	});

	it("spans merged cells and carries their formatting", async () => {
		const grid = buildGrid(await load(), config("sheet: Riepilogo"));
		const [title, covered] = grid.rows[2].cells;
		assert.equal(title.text, "Riepilogo annuale");
		assert.deepEqual(title.span, { cols: 2, rows: 1 });
		assert.deepEqual(title.style, { bold: true, color: "#000000", background: "#DAE3F3", align: "center", wrap: true });
		assert.equal(covered.covered, true);
	});

	it("colours numbers by their format and leaves default colours to the theme", async () => {
		const grid = buildGrid(await load(), config("sheet: Riepilogo"));
		const [total, , negative] = grid.rows[0].cells;
		assert.deepEqual(total.style, { bold: true }, "no colour: Excel's black text is left to the theme");
		assert.equal(negative.style?.color, "var(--text-error)");
		assert.match(negative.text, /^-250.00$/);
	});

	it("marks table headers inside a sheet and offers table rules below the table", async () => {
		const wb = await load();
		const grid = buildGrid(wb, config("sheet: 2026"));
		assert.ok(grid.rows[0].cells.every((c) => c.tableHeader && !c.editable));
		const newRow = grid.rows[grid.rows.length - 1];
		assert.equal(newRow.kind, "new");
		assert.equal(newRow.cells[3].options?.length, 5, "category dropdown on the row below the table");
		const sheet = wb.sheet("2026");
		assert.equal(wb.tableAt(sheet, 4, 4)?.table.name, "Movimenti");
		assert.equal(wb.tableAt(sheet, 4, 6), undefined);
	});

	it("hides headings on request, keeping a table's column names", async () => {
		const wb = await load();
		const sheet = buildGrid(wb, config("sheet: Riepilogo\nheadings: false"));
		assert.equal(sheet.rowNumbers, false);
		assert.equal(sheet.columnHeader, false);
		const table = buildGrid(wb, config("table: Movimenti\nheadings: no"));
		assert.equal(table.rowNumbers, false);
		assert.equal(table.columnHeader, true);
		assert.equal(buildGrid(wb, config("sheet: Riepilogo")).rowNumbers, true);
	});

	it("fits tables to the page and keeps sheets at Excel's widths", async () => {
		const wb = await load();
		assert.equal(buildGrid(wb, config("table: Movimenti")).fit, true);
		assert.equal(buildGrid(wb, config("sheet: Riepilogo")).fit, false);
	});
});

describe("choosing columns", () => {
	it("parses column lists in either key style, keeping the order given", () => {
		assert.deepEqual(parseColumnList("[b,c,d]"), [2, 3, 4]);
		assert.deepEqual(parseColumnList("B:D, F"), [2, 3, 4, 6]);
		assert.deepEqual(parseColumnList("[D, B, d]"), [4, 2], "duplicates dropped");
		assert.deepEqual(parseColumnList("[D:B]"), [4, 3, 2]);
		assert.deepEqual(parseColumnList("[$AA]"), [27]);
		assert.throws(() => parseColumnList("[1, B]"), /"1" isn't a column/);
		assert.throws(() => parseColumnList("[]"), /No columns/);

		assert.deepEqual(config("sheet: 2026\ncols=[b,c,d]").target, { kind: "sheet", sheet: "2026", range: undefined, columns: [2, 3, 4] });
		assert.deepEqual(config("sheet= 2026\ncols: [B]").target, { kind: "sheet", sheet: "2026", range: undefined, columns: [2] });
		assert.throws(() => config("table: Movimenti\ncols=[b]"), /work with `sheet:`/);
	});

	it("shows only the listed columns", async () => {
		const grid = buildGrid(await load(), config("sheet: 2026\ncols=[b,c,d]"));
		assert.deepEqual(grid.columns, ["B", "C", "D"]);
		const [date, amount, category] = grid.rows[1].cells;
		assert.equal(date.text, "01/10/2026");
		assert.match(amount.text, /^1.500.00 €$/);
		assert.equal(category.text, "Stipendio");
		assert.ok(grid.rows[0].cells.every((c) => c.tableHeader), "table header still recognised");
		assert.equal(grid.rows[grid.rows.length - 1].kind, "new");
	});

	it("shows listed columns even if hidden, and only spans merges when in order", async () => {
		const wb = await load();
		const reordered = buildGrid(wb, config("sheet: Riepilogo\ncols=[C,A]"));
		assert.deepEqual(reordered.columns, ["C", "A"]);
		assert.deepEqual(reordered.columnWidths, [5, 20]);
		assert.equal(reordered.rows[0].cells[0].text, "nascosto");
		assert.equal(reordered.rows[2].cells[1].text, "Riepilogo annuale");
		assert.equal(reordered.rows[2].cells[1].span, undefined);

		const ordered = buildGrid(wb, config("sheet: Riepilogo\ncols=[A,B]"));
		assert.deepEqual(ordered.rows[2].cells[0].span, { cols: 2, rows: 1 });
		assert.equal(ordered.rows[2].cells[1].covered, true);
	});
});

describe("colours", () => {
	const theme = readThemeColors(undefined);
	const color = (xml: string, kind: "font" | "fill") => resolveColor(parseXml(xml).documentElement, theme, kind);

	it("resolves rgb, theme + tint and indexed colours", () => {
		assert.deepEqual(color('<color rgb="FFC00000"/>', "font"), { hex: "#C00000", isDefault: false });
		assert.equal(color('<fgColor theme="4" tint="-0.249977111117893"/>', "fill")?.hex, "#2F5597");
		assert.deepEqual(color('<color indexed="10"/>', "font"), { hex: "#FF0000", isDefault: false });
		assert.equal(color('<color indexed="64"/>', "font"), undefined, "system colour");
		assert.equal(color('<color auto="1"/>', "font"), undefined);
	});

	it("treats Excel's default black text and white fill as theme defaults", () => {
		assert.equal(color('<color theme="1"/>', "font")?.isDefault, true);
		assert.equal(color('<color rgb="FF000000"/>', "font")?.isDefault, true);
		assert.equal(color('<fgColor theme="0"/>', "fill")?.isDefault, true);
		assert.equal(color('<fgColor rgb="FFFFFFFF"/>', "fill")?.isDefault, true);
		assert.equal(color('<fgColor theme="0" tint="-0.15"/>', "fill")?.isDefault, false, "a tinted white is a real grey");
	});

	it("reads the workbook theme in style order", () => {
		const themeXml = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="X">
			<a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FEFEFE"/></a:lt1>
			<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
			${[1, 2, 3, 4, 5, 6].map((i) => `<a:accent${i}><a:srgbClr val="00000${i}"/></a:accent${i}>`).join("")}
			<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`;
		assert.deepEqual(readThemeColors(parseXml(themeXml).documentElement).slice(0, 5), ["FEFEFE", "111111", "EEEEEE", "222222", "000001"]);
	});

	it("picks readable text for a fill", () => {
		assert.equal(contrastText("#DAE3F3"), "#000000");
		assert.equal(contrastText("#2F5597"), "#FFFFFF");
	});
});
