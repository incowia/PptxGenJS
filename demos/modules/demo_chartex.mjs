/**
 * NAME: demo_chartex.mjs
 * AUTH: incowia GmbH (https://github.com/incowia/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.7.0-beta
 * BLD.:
 */

import {
	BASE_TABLE_OPTS,
	BASE_TEXT_OPTS_L,
	BASE_TEXT_OPTS_R,
	COLOR_RED,
	COLOR_AMB,
	COLOR_GRN,
	COLOR_UNK,
	TESTMODE,
	BASE_CODE_OPTS, IMAGE_PATHS
} from "./enums.mjs";

// test data
const data2 = {
	name: 'Data series 2',
	values: [
		// first tree
		// segment size is sum of values from children (except leafs)
		[ 'Branch 1',        '',        '',    0 ],
		[ 'Branch 1',  'Root 1',        '',    0 ],
		[ 'Branch 1',  'Root 1',  'Leaf 1',   22 ],
		[ 'Branch 1',  'Root 1',  'Leaf 2',   12 ],
		[ 'Branch 1',  'Root 1',  'Leaf 3',   18 ],
		[ 'Branch 1',  'Root 2',        '',    0 ],
		[ 'Branch 1',  'Root 2',  'Leaf 4',   87 ],
		[ 'Branch 1',  'Root 2',  'Leaf 5',   88 ],
		[ 'Branch 1',  'Leaf 6',        '',   17 ],
		[ 'Branch 1',  'Leaf 7',        '',   14 ],
		// second tree
		// segment size is value plus values of children (except leafs)
		[ 'Branch 2',        '',        '', -138 ],
		[ 'Branch 2',  'Root 3',        '',  -25 ],
		[ 'Branch 2',  'Root 3',  'Leaf 8',   25 ],
		[ 'Branch 2',  'Leaf 9',        '',   16 ],
		[ 'Branch 2',  'Root 4',        '', -113 ],
		[ 'Branch 2',  'Root 4', 'Leaf 10',   24 ],
		[ 'Branch 2',  'Root 4', 'Leaf 11',   89 ],
		// third tree
		// adds 30 to segment plus values of children (except leafs)
		[ 'Branch 3',        '',        '',   30 ],
		[ 'Branch 3',  'Root 5',        '',   30 ],
		[ 'Branch 3',  'Root 5', 'Leaf 12',   16 ],
		[ 'Branch 3',  'Root 5', 'Leaf 13',   19 ],
		[ 'Branch 3',  'Root 6',        '',   30 ],
		[ 'Branch 3',  'Root 6', 'Leaf 14',   86 ],
		[ 'Branch 3',  'Root 6', 'Leaf 15',   23 ],
		[ 'Branch 3', 'Leaf 16',        '',   21 ]
	]
};

export function genSlides_ChartEx(pptx) {
	initTestData();

	pptx.addSection({ title: "Extended charts" });

	genSlide01(pptx);
	genSlide02(pptx);
	genSlide03(pptx);
}

function initTestData() {
	// TODO wenn nötig
}

// SLIDE 1: Default example in MS PowerPoint (only leafs)
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Extended Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-chartexs.html"); // TODO API-Doku
	slide.addTable([[{ text: "Extended Chart Examples: Sunburst Chart - Default example in MS PowerPoint", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const data1 = {
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
	};
	const chartEx1Options = {
		type: pptx.ChartExType.sunburst,
		x: 0.5,
		y: 0.65,
		w: 6.0,
		h: 6.0,
		sunburst: {
			dataLabel: {
				visibility: {
					category: true,
					value: false
				}
			}
		}
	};
	slide.addText([{ text: `Only leafs` }], { ...{ fontSize: 13 }, ...{ x: 0.5, y: 0.5, h: 0.25, w: 6.0 } });
	slide.addChartEx(data1, chartEx1Options);

	const chartEx2Options = {
		type: pptx.ChartExType.sunburst,
		x: 6.7,
		y: 0.65,
		w: 6.0,
		h: 6.0,
		sunburst: {
			dataLabel: {
				visibility: {
					category: true,
					value: false
				}
			}
		}
	};
	slide.addText([{ text: `Entire trees` }], { ...{ fontSize: 13 }, ...{ x: 6.7, y: 0.5, h: 0.25, w: 6.0 } });
	slide.addChartEx(data2, chartEx2Options);
	let text1 = `+30\nRoot 5`
	let text2 = `+30\nBranch 3`
	let text3 = `+30\nRoot 6`
	slide.addText([{ text: text1 }], { ...{color: 'FF0000'}, ...{ x: 6.9, y: 4.7, w: 1, h: 0.5, align: 'right' } })
	slide.addText([{ text: text2 }], { ...{color: 'FF0000'}, ...{ x: 10.5, y: 5.1, w: 1.2, h: 0.5 } })
	slide.addText([{ text: text3 }], { ...{color: 'FF0000'}, ...{ x: 10.1, y: 6.0, w: 1, h: 0.5 } })
}

// SLIDE 2: Show values instead of categories
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Extended Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-chartexs.html");
	slide.addTable([[{ text: "Extended Chart Examples: Sunburst Chart - Show values instead of categories", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const chartEx3Options = {
		type: pptx.ChartExType.sunburst,
		x: 0.5,
		y: 1.0,
		w: 6.0,
		h: 6.0,
		sunburst: {
			dataLabel: {
				visibility: {
					category: false,
					value: true
				}
			}
		}
	};
	slide.addText([{ text: `Default value format` }], { ...{ fontSize: 13 }, ...{ x: 0.5, y: 0.5, h: 0.25, w: 6.0 } });
	slide.addChartEx(data2, chartEx3Options);

	const chartEx4Options = {
		type: pptx.ChartExType.sunburst,
		x: 6.7,
		y: 1.0,
		w: 6.0,
		h: 6.0,
		sunburst: {
			dataLabel: {
				numFmt: '0;0', // shows absolute number only
				visibility: {
					category: false,
					value: true
				}
			}
		}
	};
	slide.addText([{ text: `Format values` }], { ...{ fontSize: 13 }, ...{ x: 6.7, y: 0.5, h: 0.25, w: 6.0 } });
	slide.addText([{ text: `example: numFmt: '0;0' -> shows absolute number only` }],
		{ ...BASE_CODE_OPTS, ...{ x: 6.7, y: 0.8, h: 0.25, w: 6.0 } });
	slide.addChartEx(data2, chartEx4Options);
}

// Slide 3: Format segments individually
function genSlide03(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Extended Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-chartexs.html");
	slide.addTable([[{ text: "Extended Chart Examples: Sunburst Chart - Format segments individually", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const chartEx5Options = {
		type: pptx.ChartExType.sunburst,
		x: 0.5,
		y: 1.0,
		w: 6.0,
		h: 6.0,
		sunburst: {
			dataLabel: {
				numFmt: '0;0', // shows absolute number only
				visibility: {
					category: true,
					value: true
				}
			},
			segments: [
				{
					path: ['Branch 3'],
					dataLabel: { numFmt: '"foo";"foo"' /* static string */ }
				}, {
					path: ['Branch 2'],
					dataLabel: { visibility: { category: false } }
				}, {
					path: ['Branch 1', 'Root 1'],
					dataLabel: { visibility: { category: false, value: false } },
					fill: { type: 'solid', color: '00FF00' }
				}, {
					path: ['Branch 2', 'Root 4', 'Leaf 11'],
					fill: { type: 'solid', color: 'FFFFFF' },
					line: { width: 1, color: '000000' },
					text: { fill: { type: 'solid', color: '000000' } }
				}
			]
		}
	};
	slide.addChartEx(data2, chartEx5Options);
}
