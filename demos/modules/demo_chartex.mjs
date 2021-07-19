/**
 * NAME: demo_chartex.mjs
 * AUTH: incowia GmbH (https://github.com/incowia/)
 * DESC: Common test/demo slides for all library features
 * DEPS: Used by various demos (./demos/browser, ./demos/node, etc.)
 * VER.: 3.7.0-beta
 * BLD.:
 */

import { BASE_TABLE_OPTS, BASE_TEXT_OPTS_L, BASE_TEXT_OPTS_R, COLOR_RED, COLOR_AMB, COLOR_GRN, COLOR_UNK, TESTMODE } from "./enums.mjs";

export function genSlides_ChartEx(pptx) {
	initTestData();

	pptx.addSection({ title: "Extended charts" });

	genSlide01(pptx);
	genSlide02(pptx);
}

function initTestData() {
	// TODO wenn nötig
}

// SLIDE 1: Default example in MS PowerPoint (only leafs)
function genSlide01(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Extended Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-chartexs.html"); // TODO API-Doku
	slide.addTable([[{ text: "Extended Chart Examples: Sunburst Chart - Default example in MS PowerPoint (only leafs)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const data = {
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
//	slide.addTable(data.values);
	const options = {
		type: pptx.ChartExType.sunburst,
		x: 0.5,
		y: 0.6,
		w: 6.0,
		h: 5.0,
		sunburst: {
			dataLabel: {
				visibility: {
					category: true,
					value: false
				}
			}
		}
	};
	slide.addChartEx(data, options);
}

// SLIDE 2: Default example in MS PowerPoint (entire trees)
function genSlide02(pptx) {
	let slide = pptx.addSlide({ sectionTitle: "Extended Charts" });
	slide.addNotes("API Docs: https://gitbrent.github.io/PptxGenJS/docs/api-chartexs.html");
	slide.addTable([[{ text: "Extended Chart Examples: Sunburst Chart - Default example in MS PowerPoint (entire trees)", options: BASE_TEXT_OPTS_L }, BASE_TEXT_OPTS_R]], BASE_TABLE_OPTS);

	const data = {
		name: 'Data series 1',
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
	slide.addTable(data.values);

	/*	const options = {
            type: pptx.ChartExType.sunburst,
            x: 0.5,
            y: 0.6,
            w: 6.0,
            h: 5.0,
        };
        slide.addChartEx(data, options);*/
}
