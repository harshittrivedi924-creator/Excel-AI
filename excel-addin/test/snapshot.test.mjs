import { test } from "node:test";
import assert from "node:assert/strict";

import { blankValue, isEmptyValues, trimBounds, buildSnapshot } from "../src/taskpane/snapshot.mjs";

test("blankValue", () => {
  assert.equal(blankValue(null), true);
  assert.equal(blankValue(undefined), true);
  assert.equal(blankValue(""), true);
  assert.equal(blankValue(0), false);
  assert.equal(blankValue("0"), false);
});

test("isEmptyValues rejects genuinely empty 2-D arrays", () => {
  assert.equal(isEmptyValues(null), true);
  assert.equal(isEmptyValues([]), true);
  assert.equal(isEmptyValues([[null]]), true);
  assert.equal(isEmptyValues([[""]]), true);
  assert.equal(isEmptyValues([[null, ""], ["", null]]), true);
  assert.equal(isEmptyValues([["a"]]), false);
  assert.equal(isEmptyValues([[0]]), false);
  assert.equal(isEmptyValues([[2555, 4544], [511, 3535]]), false);
});

test("trimBounds detects data offset inside an A1-window scan", () => {
  // What the A1:CV3000 region scan returns for the real workbook: headers on
  // row 2 (index 1), five data rows under them.
  const values = [];
  for (let r = 0; r < 3000; r++) values.push(new Array(100).fill(null));
  values[1][0] = "a";
  values[1][1] = "b";
  values[1][2] = "d";
  values[2][0] = 2555;
  values[2][1] = 4544;
  values[3][0] = 511;
  values[3][1] = 3535;
  values[4][0] = 214;
  values[4][1] = 355;
  values[5][0] = 3212135;
  values[5][1] = 546;
  values[6][0] = 35;
  values[6][1] = 3545;

  const bounds = trimBounds(values);
  assert.deepEqual(bounds, { startRow: 1, startCol: 0, endRow: 6, endCol: 2 });

  const snap = buildSnapshot(values, { maxRows: 3000, maxCols: 100 });
  assert.equal(snap.rows, 6);
  assert.equal(snap.cols, 3);
  assert.equal(snap.originRow, 2);
  assert.equal(snap.originCol, 1);
  assert.deepEqual(snap.values[0], ["a", "b", "d"]);
  assert.deepEqual(snap.values[1], [2555, 4544, null]);
});

test("buildSnapshot returns null for genuinely blank worksheets", () => {
  assert.equal(buildSnapshot([]), null);
  assert.equal(buildSnapshot([[null, ""], [""]]), null);
  assert.equal(buildSnapshot(new Array(50).fill(new Array(50).fill(null))), null);
});

test("buildSnapshot trims trailing blank rows/columns", () => {
  const values = [
    [null, null, null],
    ["a", 1, null],
    ["b", 2, null],
    ["c", 3, null],
    [null, null, null],
  ];
  const snap = buildSnapshot(values, { baseRow: 1, baseCol: 1 });
  assert.equal(snap.rows, 3);
  assert.equal(snap.cols, 2);
  assert.equal(snap.originRow, 2);
  assert.equal(snap.originCol, 1);
});

test("buildSnapshot caps oversized snapshots and flags truncated", () => {
  const values = new Array(20).fill(new Array(4).fill(5));
  const snap = buildSnapshot(values, { maxRows: 10, maxCols: 2 });
  assert.equal(snap.rows, 10);
  assert.equal(snap.cols, 2);
  assert.equal(snap.truncated, true);
});

test("buildSnapshot honors baseRow/baseCol from used-range origin", () => {
  // used range A2:C4 -> rowIndex=1, columnIndex=0 -> baseRow=2, baseCol=1
  const values = [
    ["a", "b", "d"],
    [2555, 4544, null],
    [511, 3535, null],
  ];
  const snap = buildSnapshot(values, { baseRow: 2, baseCol: 1 });
  assert.equal(snap.originRow, 2);
  assert.equal(snap.originCol, 1);
  assert.equal(snap.rows, 3);
  assert.equal(snap.cols, 3);
});

test("b ka total snapshot produces the expected numeric column", () => {
  // Regression guard for the reported worksheet (Sheet1, A2:C7).
  const values = [
    ["a", "b", "d"],
    [2555, 4544, null],
    [511, 3535, null],
    [214, 355, null],
    [3212135, 546, null],
    [35, 3545, null],
  ];
  const snap = buildSnapshot(values, { baseRow: 2, baseCol: 1 });
  const b = snap.values.slice(1).map((row) => row[1]);
  assert.deepEqual(b, [4544, 3535, 355, 546, 3545]);
  assert.equal(b.reduce((a, v) => a + v, 0), 12525);
});