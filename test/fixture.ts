import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

/** The text of one part of a saved workbook, or undefined if it isn't there. */
export function readEntry(bytes: Uint8Array, path: string): string | undefined {
	const data = unzipSync(bytes)[path];
	return data && strFromU8(data);
}

/**
 * Builds a workbook laid out the way Excel writes one: shared strings, a calculation
 * chain, and on sheet "2026" a table "Movimenti" with a calculated column, a
 * running-balance shared formula, active filters and data validation rules; sheet
 * "Liste" holds the categories in table "TabListe".
 */
export async function financeWorkbook(options: { totalsRow?: boolean } = {}): Promise<Uint8Array> {
	const serial = (y: number, m: number, d: number) => (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86_400_000;
	const strings = ["Mese", "Data", "Importo", "Categoria", "Saldo", "Stipendio", "Spesa", "Totale", "Riepilogo annuale", "Tipo", "Salute", "Sport", "Abbonamenti", "nascosto", "riga nascosta"];
	const s = (text: string) => strings.indexOf(text);
	const ref = options.totalsRow ? "A1:E4" : "A1:E3";

	const files: Record<string, Uint8Array> = {};
	const zip = { file: (path: string, text: string) => (files[path] = strToU8(text)) };
	zip.file(
		"[Content_Types].xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
	);
	zip.file(
		"_rels/.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
	);
	zip.file(
		"xl/workbook.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><workbookPr defaultThemeVersion="166925"/><bookViews><workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="12300"/></bookViews><sheets><sheet name="Riepilogo" sheetId="1" r:id="rId1"/><sheet name="2026" sheetId="2" r:id="rId2"/><sheet name="Liste" sheetId="3" r:id="rId6"/></sheets><definedNames><definedName name="Categorie">TabListe[Tipo]</definedName></definedNames><calcPr calcId="191029"/></workbook>`,
	);
	zip.file(
		"xl/_rels/workbook.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>`,
	);
	zip.file(
		"xl/styles.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;€&quot;"/><numFmt numFmtId="165" formatCode="dd/mm/yyyy"/><numFmt numFmtId="166" formatCode="#,##0.00;[Red]\\-#,##0.00"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor theme="4" tint="0.79998168889431442"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
	);
	zip.file(
		"xl/sharedStrings.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((t) => `<si><t>${t}</t></si>`).join("")}</sst>`,
	);
	zip.file(
		"xl/calcChain.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A2" i="2"/><c r="E2"/><c r="A3"/><c r="E3"/><c r="B1" i="1"/></calcChain>`,
	);
	zip.file(
		"xl/worksheets/sheet1.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:D6"/><cols><col min="1" max="1" width="20" customWidth="1"/><col min="3" max="3" width="5" hidden="1" customWidth="1"/></cols><sheetData><row r="1" spans="1:4"><c r="A1" s="1" t="s"><v>${s("Totale")}</v></c><c r="B1" s="3"><f>SUM(Movimenti[Importo])</f><v>1455</v></c><c r="C1" t="s"><v>${s("nascosto")}</v></c><c r="D1" s="4"><v>-250</v></c></row><row r="3" spans="1:2"><c r="A3" s="5" t="s"><v>${s("Riepilogo annuale")}</v></c><c r="B3" s="5"/></row><row r="6" hidden="1"><c r="A6" t="s"><v>${s("riga nascosta")}</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A3:B3"/></mergeCells><dataValidations count="3"><dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" promptTitle="Chiuso" prompt="Il mese è chiuso?" sqref="B2"><formula1>"Sì,No"</formula1></dataValidation><dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="B4"><formula1>Categorie</formula1></dataValidation><dataValidation type="list" allowBlank="1" sqref="B5"><formula1>INDIRECT("TabListe[Tipo]")</formula1></dataValidation></dataValidations><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`,
	);

	const totals = options.totalsRow
		? `<row r="4"><c r="A4" t="s"><v>${s("Totale")}</v></c><c r="C4" s="3"><f>SUBTOTAL(109,Movimenti[Importo])</f><v>1455</v></c></row>`
		: "";
	zip.file(
		"xl/worksheets/sheet2.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="${ref}"/><sheetData><row r="1" spans="1:5"><c r="A1" s="1" t="s"><v>${s("Mese")}</v></c><c r="B1" s="1" t="s"><v>${s("Data")}</v></c><c r="C1" s="1" t="s"><v>${s("Importo")}</v></c><c r="D1" s="1" t="s"><v>${s("Categoria")}</v></c><c r="E1" s="1" t="s"><v>${s("Saldo")}</v></c></row><row r="2" spans="1:5"><c r="A2" t="str"><f>TEXT(Movimenti[[#This Row],[Data]],"mmmm")</f><v>ottobre</v></c><c r="B2" s="2"><v>${serial(2026, 10, 1)}</v></c><c r="C2" s="3"><v>1500</v></c><c r="D2" t="s"><v>${s("Stipendio")}</v></c><c r="E2" s="3"><f t="shared" ref="E2:E3" si="0">SUM($C$2:C2)</f><v>1500</v></c></row><row r="3" spans="1:5"><c r="A3" t="str"><f>TEXT(Movimenti[[#This Row],[Data]],"mmmm")</f><v>ottobre</v></c><c r="B3" s="2"><v>${serial(2026, 10, 3)}</v></c><c r="C3" s="3"><v>-45</v></c><c r="D3" t="s"><v>${s("Spesa")}</v></c><c r="E3" s="3"><f t="shared" si="0"/><v>1455</v></c></row>${totals}</sheetData><dataValidations count="1"><dataValidation type="decimal" allowBlank="1" showErrorMessage="1" errorTitle="Importo" error="L'importo deve essere tra -10000 e 10000" sqref="C2:C3"><formula1>-10000</formula1><formula2>10000</formula2></dataValidation></dataValidations><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts><extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Categoria" error="Scegli una categoria dalla lista"><x14:formula1><xm:f>Liste!$A$2:$A$6</xm:f></x14:formula1><xm:sqref>D2:D3</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst></worksheet>`,
	);
	zip.file(
		"xl/worksheets/_rels/sheet2.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>`,
	);
	const totalsAttrs = options.totalsRow ? ' totalsRowCount="1"' : "";
	const filterRef = options.totalsRow ? "A1:E3" : ref;
	zip.file(
		"xl/tables/table1.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Movimenti" displayName="Movimenti" ref="${ref}"${totalsAttrs}><autoFilter ref="${filterRef}"><filterColumn colId="2"><customFilters><customFilter operator="notEqual" val=" "/></customFilters></filterColumn><filterColumn colId="3"><filters><filter val="Spesa"/><filter val="Stipendio"/></filters></filterColumn></autoFilter><tableColumns count="5"><tableColumn id="1" name="Mese"><calculatedColumnFormula>TEXT(Movimenti[[#This Row],[Data]],"mmmm")</calculatedColumnFormula></tableColumn><tableColumn id="2" name="Data"/><tableColumn id="3" name="Importo"/><tableColumn id="4" name="Categoria"/><tableColumn id="5" name="Saldo"><calculatedColumnFormula>SUM($C$2:C2)</calculatedColumnFormula></tableColumn></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>`,
	);
	const categories = ["Stipendio", "Spesa", "Salute", "Sport", "Abbonamenti"];
	zip.file(
		"xl/worksheets/sheet3.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><dimension ref="A1:A6"/><sheetData><row r="1"><c r="A1" s="1" t="s"><v>${s("Tipo")}</v></c></row>${categories.map((c, i) => `<row r="${i + 2}"><c r="A${i + 2}" t="s"><v>${s(c)}</v></c></row>`).join("")}</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>`,
	);
	zip.file(
		"xl/worksheets/_rels/sheet3.xml.rels",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table2.xml"/></Relationships>`,
	);
	zip.file(
		"xl/tables/table2.xml",
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="2" name="TabListe" displayName="TabListe" ref="A1:A6"><autoFilter ref="A1:A6"/><tableColumns count="1"><tableColumn id="1" name="Tipo"/></tableColumns><tableStyleInfo name="TableStyleLight1" showRowStripes="1"/></table>`,
	);
	return zipSync(files);
}
