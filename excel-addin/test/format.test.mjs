import { test } from "node:test";
import assert from "node:assert/strict";
import {
  groupThousands,
  successIcon,
  humanizeSuccess,
  humanizeError,
} from "../src/taskpane/format.mjs";

test("groupThousands formats standalone integers", () => {
  assert.equal(groupThousands("The result is 18000."), "The result is 18,000.");
  assert.equal(groupThousands("Sum of B2:B7 is 145577390."), "Sum of B2:B7 is 145,577,390.");
  assert.equal(groupThousands("total comes to 7500"), "total comes to 7,500");
});

test("groupThousands keeps decimals", () => {
  assert.equal(groupThousands("average is 12345.5"), "average is 12,345.5");
  assert.equal(groupThousands("percent 0.05"), "percent 0.05");
});

test("groupThousands leaves short numbers and cell references untouched", () => {
  assert.equal(groupThousands("I wrote 6 formula(s) to column D."), "I wrote 6 formula(s) to column D.");
  assert.equal(groupThousands("B7 already contains the value B7"), "B7 already contains the value B7");
  assert.equal(groupThousands("=SUM(A2:A4)"), "=SUM(A2:A4)");
  assert.equal(groupThousands("Revenue"), "Revenue");
});

test("successIcon reflects charts, sheets and dedupe", () => {
  assert.equal(successIcon({}), "✅");
  assert.equal(successIcon({ operation: "DEDUPE" }), "🧹");
  assert.equal(successIcon({ writes: { charts: [{ type: "ColumnClustered" }] } }), "📊");
  assert.equal(successIcon({ writes: { sheets: { Filtered: {} } } }), "📄");
  assert.equal(successIcon({ writes: { cells: [] } }), "✅");
});

test("humanizeSuccess wraps the readable backend message", () => {
  const text = humanizeSuccess({
    message: "Done. The result of Sum of A2:A4 is 18000.",
    operation: "SUM",
    writes: { cells: [], sheets: {}, charts: [] },
  });
  assert.ok(text.includes("✅"));
  assert.ok(text.includes("18,000"));
  assert.ok(!text.includes("{"));
});

test("humanizeError keeps the message readable and marks errors", () => {
  assert.equal(humanizeError("I couldn't find a column named 'Xyz'."), "❌ I couldn't find a column named 'Xyz'.");
  assert.equal(humanizeError("❌ already prefixed."), "❌ already prefixed.");
  assert.equal(humanizeError(""), "❌ Something went wrong. Try again.");
});