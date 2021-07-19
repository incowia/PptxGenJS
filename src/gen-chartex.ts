/**
 * PptxGenJS: Extended Chart Generation
 */

import {ChartExType, DEF_SHAPE_LINE_COLOR, LETTERS,} from './core-enums'
import {ISlideRelChartEx, SunburstChartExData} from './core-interfaces'
import JSZip from 'jszip'
import {createColorElement, getUuid} from "./gen-utils";

/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a extended chart.
 * @param {ISlideRelChartEx} chartExObject - extended chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
export function createExcelWorksheet(chartExObject: ISlideRelChartEx, zip: JSZip): Promise<any> {
	let data = chartExObject.data;

	return new Promise((resolve, reject) => {
		let zipExcel = new JSZip()

		// A: Add folders
		zipExcel.folder('_rels')
		zipExcel.folder('docProps')
		zipExcel.folder('xl/_rels')
		zipExcel.folder('xl/tables')
		zipExcel.folder('xl/theme')
		zipExcel.folder('xl/worksheets')
		zipExcel.folder('xl/worksheets/_rels')

		// B: Add core contents
		{
			zipExcel.file(
				'[Content_Types].xml',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
					'  <Default Extension="xml" ContentType="application/xml"/>' +
					'  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
					'  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
					'  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
					'  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
					'  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
					'  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
					'  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
					'  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
					'  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
					'</Types>\n'
			)
			zipExcel.file(
				'_rels/.rels',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
					'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
					'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
					'</Relationships>\n'
			)
			zipExcel.file(
				'docProps/app.xml',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
					'<Application>Microsoft Excel</Application>' +
					'<DocSecurity>0</DocSecurity>' +
					'<ScaleCrop>false</ScaleCrop>' +
					'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
					'</Properties>\n'
			)
			zipExcel.file(
				'docProps/core.xml',
				'<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
					'<dc:creator>PptxGenJS</dc:creator>' +
					'<cp:lastModifiedBy>Ely, Brent</cp:lastModifiedBy>' +
					'<dcterms:created xsi:type="dcterms:W3CDTF">' +
					new Date().toISOString() +
					'</dcterms:created>' +
					'<dcterms:modified xsi:type="dcterms:W3CDTF">' +
					new Date().toISOString() +
					'</dcterms:modified>' +
					'</cp:coreProperties>\n'
			)
			zipExcel.file(
				'xl/_rels/workbook.xml.rels',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
					'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
					'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
					'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
					'</Relationships>\n'
			)
			zipExcel.file(
				'xl/styles.xml',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
					'<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
					'<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
			)
			zipExcel.file(
				'xl/theme/theme1.xml',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
			)
			zipExcel.file(
				'xl/workbook.xml',
				'<?xml version="1.0" encoding="UTF-8"?>' +
					'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
					'<fileVersion appName="xl" lastEdited="6" lowestEdited="6" rupBuild="14420"/>' +
					'<workbookPr />' +
					'<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="15960" windowHeight="18080"/></bookViews>' +
					'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1" /></sheets>' +
					'<calcPr calcId="171026" concurrentCalc="0"/>' +
					'</workbook>\n'
			)
			zipExcel.file(
				'xl/worksheets/_rels/sheet1.xml.rels',
				'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
					'</Relationships>\n'
			)
		}

		// sharedStrings.xml
		{
			// A: Start XML
			let strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			if (chartExObject.opts.type === ChartExType.sunburst) {
				const {rowsData, colsData} = getRowsAndColumnsData(data)
				const {labels, uniqueLabels} = getLabels(rowsData)
				strSharedStrings +=
					'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + (labels.length + 1) + '" uniqueCount="' + (uniqueLabels.length + 1) + '">'
				// C: Add `name`/Series
				strSharedStrings += '<si><t>' + encodeXmlEntities((data.name || ' ').replace('X-Axis', 'X-Values')) + '</t></si>'
				// D: Add `labels`/Categories
				uniqueLabels.forEach(label => {
					strSharedStrings += '<si><t>' + encodeXmlEntities(label.toString()) + '</t></si>'
				})
				strSharedStrings += '</sst>\n'
			}
			zipExcel.file('xl/sharedStrings.xml', strSharedStrings)
		}

		// tables/table1.xml
		{
			let strTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			if (chartExObject.opts.type === ChartExType.sunburst) {
				let {rowsData, colsData} = getRowsAndColumnsData(data)
				strTableXml +=
					'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:' +
					LETTERS[colsData.length - 1] +
					(rowsData.length + 1) +
					'" totalsRowShown="0">'
				strTableXml += '<tableColumns count="' + (colsData.length) + '">'
				colsData.slice(0, colsData.length - 1).forEach((col, idx) => {
					strTableXml += `<tableColumn id="${idx + 1}" name=" " />`
				})
				strTableXml += '<tableColumn id="' + (colsData.length) + '" name="' + encodeXmlEntities(data.name) + '" />'
			}
			strTableXml += '</tableColumns>'
			strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0" />'
			strTableXml += '</table>'
			zipExcel.file('xl/tables/table1.xml', strTableXml)
		}

		// worksheets/sheet1.xml
		{
			let strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
			strSheetXml +=
				'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
			if (chartExObject.opts.type === ChartExType.sunburst) {
				let {rowsData, colsData} = getRowsAndColumnsData(data)
				strSheetXml += '<dimension ref="A1:' + LETTERS[colsData.length - 1] + (Math.max(rowsData.length + 1, 17)) + '" />'
			}
			strSheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1" /></sheetView></sheetViews>'
			strSheetXml += '<sheetFormatPr baseColWidth="10" defaultColWidth="11.5" defaultRowHeight="12" />'

			if(chartExObject.opts.type === ChartExType.sunburst) {
				/* EX: INPUT: `data`
				{
					name: 'Data series 1',
					values: [
						// first tree
						[ 'Branch 1',  'Root 1',  'Leaf 1', 22 ],
						[ 'Branch 1',  'Root 1',  'Leaf 2', 12 ],
						[ 'Branch 1',  'Root 1',  'Leaf 3', 18 ],
						[ 'Branch 1',  'Root 2',  'Leaf 4', 87 ],
						[ 'Branch 1',  'Root 2',  'Leaf 5', 88 ],
						[ 'Branch 1',  'Leaf 6',        '', 17 ],
						[ 'Branch 1',  'Leaf 7',        '', 14 ],
						// second tree
						[ 'Branch 2',  'Root 3',  'Leaf 8', 25 ],
						[ 'Branch 2',  'Leaf 9',        '', 16 ],
						[ 'Branch 2',  'Root 4', 'Leaf 10', 24 ],
						[ 'Branch 2',  'Root 4', 'Leaf 11', 89 ],
						// third tree
						[ 'Branch 3',  'Root 5', 'Leaf 12', 16 ],
						[ 'Branch 3',  'Root 5', 'Leaf 13', 19 ],
						[ 'Branch 3',  'Root 6', 'Leaf 14', 86 ],
						[ 'Branch 3',  'Root 6', 'Leaf 15', 23 ],
						[ 'Branch 3', 'Leaf 16',        '', 21 ]
					]
				}
				*/
				/* EX: OUTPUT: scatterChart Worksheet:
					-|------------|----------|----------|----------| Datenreihe1
					1| 'Branch 1' | 'Root 1' | 'Leaf 1' | Blatt 1  | 22
					1| 'Branch 1' | 'Root 1' | 'Leaf 2' | Blatt 2  | 12
					1| 'Branch 1' | 'Root 1' | 'Leaf 3' | Blatt 3  | 18
					1| 'Branch 1' | 'Root 2' | 'Leaf 4' |          | 87
					1| 'Branch 1' | 'Root 2' | 'Leaf 5' | Blatt 5  | 88
					1| 'Branch 1' | 'Leaf 6' |          |          | 17
					-|------------|----------|----------|----------|--------------
				*/
				// generate rows
				const {rowsData, colsData} = getRowsAndColumnsData(data)
				const {labels, uniqueLabels} = getLabels(rowsData)
				const colCount = colsData.length
				const values = colsData[colsData.length - 1]
				strSheetXml += '<sheetData>'
				strSheetXml += `<row r="1" spans="1:${colCount}" x14ac:dyDescent="0.25"><c r="${LETTERS[colCount - 1]}1" t="s"><v>0</v></c></row>`
				// TODO uniqueLabel-Index ermitteln als v
				rowsData.forEach((row, i) => {
					strSheetXml += `<row r="${i + 2}" spans="1:${colCount}" x14ac:dyDescent="0.25">`
					colsData.slice(0, colsData.length - 1).forEach((col, colIdx) => {
						strSheetXml += `<c r="${LETTERS[colIdx]}${i + 2}" t="s">`
						let uniqueLabelIdx = -1
						uniqueLabels.find((l, j) => {
							if (col[i] === l) {
								uniqueLabelIdx = j + 1 // name "Data series 1" has index 0
								return true
							}
							return false
						})
						strSheetXml += col[i] ? ` <v>${uniqueLabelIdx !== -1 ? uniqueLabelIdx : ''}</v>` : ''
						strSheetXml += `</c>`
					})
					strSheetXml += `<c r="${LETTERS[colCount - 1]}${i + 2}"><v>${values[i]}</v></c>`
					strSheetXml += `</row>`
				})
				let countEmptyRows = 17 - rowsData.length - 1
				if (countEmptyRows > 0) {
					for (let i = rowsData.length + 2; i <= 17; i++) {
						strSheetXml += `<row r="${i}" spans="1:${colCount}"/>`
					}
				}
			}
			strSheetXml += '</sheetData>'
			strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />'
			// Link the `table1.xml` file to define an actual Table in Excel
			strSheetXml += '</worksheet>\n'
			zipExcel.file('xl/worksheets/sheet1.xml', strSheetXml)
		}

		// C: Add XLSX to PPTX export
		zipExcel
			.generateAsync({ type: 'base64' })
			.then(content => {
				// 1: Create the embedded Excel worksheet with labels and data
				zip.file('ppt/embeddings/Microsoft_Excel_Worksheet' + chartExObject.globalId + '.xlsx', content, { base64: true })

				// 2: Create the chart.xml and rel files
				if (chartExObject.opts.type === ChartExType.sunburst) {
					zip.file(
						'ppt/charts/_rels/' + chartExObject.fileName + '.rels',
						'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
						'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
						'<Relationship Id="rId3" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors' +
						chartExObject.globalId +
						'.xml"/>' +
						'<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style' +
						chartExObject.globalId +
						'.xml"/>' +
						'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet' +
						chartExObject.globalId +
						'.xlsx"/>' +
						'</Relationships>'
					)
					zip.file(`ppt/charts/colors${chartExObject.globalId}.xml`, makeXmlColors(chartExObject))
					zip.file(`ppt/charts/style${chartExObject.globalId}.xml`, makeXmlStyle(chartExObject))
					zip.file('ppt/charts/' + chartExObject.fileName, makeXmlChartEx(chartExObject))
				}

				// 3: Done
				resolve(null)
			})
			.catch(strErr => {
				reject(strErr)
			})
	})
}

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChartEx} chartExObject - extended chart object
 * @return {string} XML
 */
export function makeXmlChartEx(chartExObject: ISlideRelChartEx): string {
	let data = chartExObject.data
	let {rowsData, colsData} = getRowsAndColumnsData(data);
	const values = colsData[colsData.length - 1]
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	strXml += '<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex">'
	strXml += ' <cx:chartData>'
	strXml += `  <cx:externalData r:id="rId1" cx:autoUpdate="0"/>`
	strXml += '  <cx:data id="0">'
	strXml += '   <cx:strDim type="cat">'
	strXml += `    <cx:f>Sheet1!$A$2:$${LETTERS[colsData.length - 2]}$${1 + rowsData.length}</cx:f>`
	colsData.slice(0, colsData.length - 1).reverse().forEach((col) => {
		strXml += `    <cx:lvl ptCount="${rowsData.length}">`
		col.forEach((label, iLabel) => {
			strXml += `     <cx:pt idx="${iLabel}">${encodeXmlEntities(label)}</cx:pt>`
		})
		strXml += '    </cx:lvl>'
	})
	strXml += '   </cx:strDim>'
	strXml += '   <cx:numDim type="size">'
	strXml += `    <cx:f>Sheet1!$${LETTERS[colsData.length - 1]}$2:$${LETTERS[colsData.length - 1]}$${1 + rowsData.length}</cx:f>`
	strXml += `    <cx:lvl ptCount="${rowsData.length}" formatCode="Standard">`
	values.forEach((value, iV) => {
		strXml += `     <cx:pt idx="${iV}">${value}</cx:pt>`
	})
	strXml += '    </cx:lvl>'
	strXml += '   </cx:numDim>'
	strXml += '  </cx:data>'
	strXml += ' </cx:chartData>'
	strXml += ' <cx:chart>'
	if (chartExObject.opts.title) {
		strXml += '<cx:title pos="t" align="ctr" overlay="0">'
		strXml += '	<cx:tx>'
		strXml += '		<cx:rich>'
		strXml += '			<a:bodyPr spcFirstLastPara="1" vertOverflow="ellipsis" horzOverflow="overflow" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/>'
		strXml += '			<a:lstStyle/>'
		strXml += '			<a:p>'
		strXml += '				<a:pPr algn="ctr" rtl="0">'
		strXml += '					<a:defRPr/>'
		strXml += '				</a:pPr>'
		strXml += '				<a:r>'
		strXml += '					<a:rPr lang="" sz="1862" b="0" i="0" u="none" strike="noStrike" baseline="0" dirty="0">'
		strXml += '						<a:solidFill>'
		strXml += '							<a:prstClr val="black">'
		strXml += '								<a:lumMod val="65000"/>'
		strXml += '								<a:lumOff val="35000"/>'
		strXml += '							</a:prstClr>'
		strXml += '						</a:solidFill>'
		strXml += '						<a:latin typeface="Calibri" panose="020F0502020204030204"/>'
		strXml += '					</a:rPr>'
		strXml += `					<a:t>${encodeXmlEntities(data.name)}</a:t>`
		strXml += '				</a:r>'
		strXml += '			</a:p>'
		strXml += '		</cx:rich>'
		strXml += '	</cx:tx>'
		strXml += '</cx:title>'
	}
	strXml += '   <cx:plotArea>'
	strXml += '    <cx:plotAreaRegion>'
	strXml += `     <cx:series layoutId="sunburst" uniqueId="{` + '00000000'.substring(0, 8 - (chartExObject.globalId + 1).toString().length).toString() + (chartExObject.globalId + 1) + getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}">'
	strXml += '      <cx:tx>'
	strXml += '       <cx:txData>'
	strXml += `        <cx:f>Sheet1!$${LETTERS[colsData.length - 1]}$1</cx:f>`
	strXml += `        <cx:v>${encodeXmlEntities(data.name)}</cx:v>`
	strXml += '      </cx:txData>'
	strXml += '     </cx:tx>'
	// colors for data points
	if (chartExObject.opts.sunburst && chartExObject.opts.sunburst.segments) {
		chartExObject.opts.sunburst.segments.forEach((segment, iColor) => {
			if (segment.fill) {
				strXml += `<cx:dataPt idx="${iColor}">`
				strXml += '	<cx:spPr>'
				strXml += '		<a:solidFill>'
				strXml += `			${createColorElement(segment.fill.color)}`
				strXml += '		</a:solidFill>'
				strXml += '     <a:ln w="6350">'
				strXml += '	     <a:solidFill>'
				strXml += `			${createColorElement(
					segment.line ? (segment.line.color || DEF_SHAPE_LINE_COLOR) : DEF_SHAPE_LINE_COLOR,
					'     <a:lumMod val="20000"/>' +
					'     <a:lumOff val="80000"/>'
				)}`
				strXml += '	     </a:solidFill>'
				strXml += '     </a:ln>'
				strXml += '	</cx:spPr>'
				strXml += '</cx:dataPt>'
			}
		})
	}
	console.log('chartExObject', chartExObject)
	strXml += '     <cx:dataLabels pos="ctr">'
	strXml += `      <cx:visibility seriesName="0" categoryName="${(chartExObject.opts.sunburst && chartExObject.opts.sunburst.dataLabel.visibility.category) ? 1 : 0}" value="${(chartExObject.opts.sunburst && chartExObject.opts.sunburst.dataLabel.visibility.value) ? 1 : 0}"/>` // show only value with categoryName="0" value="1"
	strXml += '     </cx:dataLabels>'
	strXml += '     <cx:dataId val="0"/>'
	strXml += '    </cx:series>'
	strXml += '   </cx:plotAreaRegion>'
	strXml += '  </cx:plotArea>'
	if (chartExObject.opts.legend) {
		strXml += `  <cx:legend pos="r" align="ctr" overlay="0"/>`
	}
	strXml += ' </cx:chart>'
	strXml += '</cx:chartSpace>'
	return strXml
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities(xml: string): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

function getRowsAndColumnsData(data: SunburstChartExData) {
	const valuesCopy = data.values.slice()
	let colCount = valuesCopy[0].length
	let rowsData = valuesCopy
	let colsData = []
	for (let i = 0; i < colCount; i++) {
		let col = []
		rowsData.forEach((row) => {
			col.push(row[i] ? row[i] : '')
		})
		colsData.push(col)
	}
	return {rowsData, colsData};
}

function getLabels(rowsData) {
	const labels = rowsData.slice().reduce((acc, row) => {
		row.slice(0, row.length - 1).forEach((label) => {
			acc.push(label)
		})
		return acc
	}, [])
	const uniqueLabels = labels.slice().reduce((acc, el) => {
		if (el !== '' && !acc.includes(el)) {
			acc.push(el.toString())
		}
		return acc
	}, [])
	return {labels, uniqueLabels}
}

function getLabelsVertical(colsData) {
	let countEmptyEntries = 0;
	const allLabels = colsData.slice(0, colsData.length - 1).reduce((acc, col) => {
		col.slice(0, col.length - 1).forEach((label) => {
			acc.push(label)
		})
		return acc
	}, [])
	const labels = allLabels.filter((l, i) => {
		if (l === '') {
			countEmptyEntries = countEmptyEntries + 1
			return false
		}
		return true
	})
	return {labels, countEmptyEntries}
}

export function makeXmlColors(chartObject) {
	// TODO HPE dynamisch generieren
	let strXml = '<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" meth="cycle" id="10">'
	strXml += '	<a:schemeClr val="accent1"/>'
	strXml += '	<a:schemeClr val="accent2"/>'
	strXml += '	<a:schemeClr val="accent3"/>'
	strXml += '	<a:schemeClr val="accent4"/>'
	strXml += '	<a:schemeClr val="accent5"/>'
	strXml += '	<a:schemeClr val="accent6"/>'
	strXml += '	<cs:variation/>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="60000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="80000"/>'
	strXml += '		<a:lumOff val="20000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="80000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="60000"/>'
	strXml += '		<a:lumOff val="40000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="50000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="70000"/>'
	strXml += '		<a:lumOff val="30000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="70000"/>'
	strXml += '	</cs:variation>'
	strXml += '	<cs:variation>'
	strXml += '		<a:lumMod val="50000"/>'
	strXml += '		<a:lumOff val="50000"/>'
	strXml += '	</cs:variation>'
	strXml += '</cs:colorStyle>'
	return strXml
}

export function makeXmlStyle(chartObject) {
	// TODO HPE dynamisch
	let strXml = '<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="381">'
	strXml += '	<cs:axisTitle>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:axisTitle>'
	strXml += '	<cs:categoryAxis>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:categoryAxis>'
	strXml += '	<cs:chartArea mods="allowNoFillOverride allowNoLineOverride">'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="bg1"/>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '		<cs:defRPr sz="1330"/>'
	strXml += '	</cs:chartArea>'
	strXml += '	<cs:dataLabel>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="lt1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:dataLabel>'
	strXml += '	<cs:dataLabelCallout>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="dk1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="lt1"/>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln>'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="dk1">'
	strXml += '						<a:lumMod val="25000"/>'
	strXml += '						<a:lumOff val="75000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '		<cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" horzOverflow="clip" vert="horz" wrap="square" lIns="36576" tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">'
	strXml += '			<a:spAutoFit/>'
	strXml += '		</cs:bodyPr>'
	strXml += '	</cs:dataLabelCallout>'
	strXml += '	<cs:dataPoint>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:fillRef>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="phClr"/>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln w="19050">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="lt1"/>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dataPoint>'
	strXml += '	<cs:dataPoint3D>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:fillRef>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="phClr"/>'
	strXml += '			</a:solidFill>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dataPoint3D>'
	strXml += '	<cs:dataPointLine>'
	strXml += '		<cs:lnRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:lnRef>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="28575" cap="rnd">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="phClr"/>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dataPointLine>'
	strXml += '	<cs:dataPointMarker>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:fillRef>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="phClr"/>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln w="9525">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="lt1"/>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dataPointMarker>'
	strXml += '	<cs:dataPointMarkerLayout symbol="circle" size="5"/>'
	strXml += '	<cs:dataPointWireframe>'
	strXml += '		<cs:lnRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:lnRef>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="28575" cap="rnd">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="phClr"/>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dataPointWireframe>'
	strXml += '	<cs:dataTable>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:dataTable>'
	strXml += '	<cs:downBar>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="dk1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="dk1">'
	strXml += '					<a:lumMod val="65000"/>'
	strXml += '					<a:lumOff val="35000"/>'
	strXml += '				</a:schemeClr>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln w="9525">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="65000"/>'
	strXml += '						<a:lumOff val="35000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:downBar>'
	strXml += '	<cs:dropLine>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="35000"/>'
	strXml += '						<a:lumOff val="65000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:dropLine>'
	strXml += '	<cs:errorBar>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="65000"/>'
	strXml += '						<a:lumOff val="35000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:errorBar>'
	strXml += '	<cs:floor>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '	</cs:floor>'
	strXml += '	<cs:gridlineMajor>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:gridlineMajor>'
	strXml += '	<cs:gridlineMinor>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:gridlineMinor>'
	strXml += '	<cs:hiLoLine>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="75000"/>'
	strXml += '						<a:lumOff val="25000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:hiLoLine>'
	strXml += '	<cs:leaderLine>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="35000"/>'
	strXml += '						<a:lumOff val="65000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:leaderLine>'
	strXml += '	<cs:legend>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:legend>'
	strXml += '	<cs:plotArea mods="allowNoFillOverride allowNoLineOverride">'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '	</cs:plotArea>'
	strXml += '	<cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '	</cs:plotArea3D>'
	strXml += '	<cs:seriesAxis>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:seriesAxis>'
	strXml += '	<cs:seriesLine>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="9525" cap="flat">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:srgbClr val="D9D9D9"/>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:round/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:seriesLine>'
	strXml += '	<cs:title>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1862"/>'
	strXml += '	</cs:title>'
	strXml += '	<cs:trendline>'
	strXml += '		<cs:lnRef idx="0">'
	strXml += '			<cs:styleClr val="auto"/>'
	strXml += '		</cs:lnRef>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:ln w="19050" cap="rnd">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="phClr"/>'
	strXml += '				</a:solidFill>'
	strXml += '				<a:prstDash val="sysDash"/>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:trendline>'
	strXml += '	<cs:trendlineLabel>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:trendlineLabel>'
	strXml += '	<cs:upBar>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="dk1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:spPr>'
	strXml += '			<a:solidFill>'
	strXml += '				<a:schemeClr val="lt1"/>'
	strXml += '			</a:solidFill>'
	strXml += '			<a:ln w="9525">'
	strXml += '				<a:solidFill>'
	strXml += '					<a:schemeClr val="tx1">'
	strXml += '						<a:lumMod val="15000"/>'
	strXml += '						<a:lumOff val="85000"/>'
	strXml += '					</a:schemeClr>'
	strXml += '				</a:solidFill>'
	strXml += '			</a:ln>'
	strXml += '		</cs:spPr>'
	strXml += '	</cs:upBar>'
	strXml += '	<cs:valueAxis>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1">'
	strXml += '				<a:lumMod val="65000"/>'
	strXml += '				<a:lumOff val="35000"/>'
	strXml += '			</a:schemeClr>'
	strXml += '		</cs:fontRef>'
	strXml += '		<cs:defRPr sz="1197"/>'
	strXml += '	</cs:valueAxis>'
	strXml += '	<cs:wall>'
	strXml += '		<cs:lnRef idx="0"/>'
	strXml += '		<cs:fillRef idx="0"/>'
	strXml += '		<cs:effectRef idx="0"/>'
	strXml += '		<cs:fontRef idx="minor">'
	strXml += '			<a:schemeClr val="tx1"/>'
	strXml += '		</cs:fontRef>'
	strXml += '	</cs:wall>'
	strXml += '</cs:chartStyle>'
	return strXml
}
