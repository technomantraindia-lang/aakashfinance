import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";
import { parse } from "@babel/parser";

const source = readFileSync(new URL("./src/App.jsx", import.meta.url), "utf8");
const ast = parse(source, { sourceType: "module", plugins: ["jsx"] });
const context = vm.createContext({ console });
vm.runInContext(ast.program.body
  .filter((node) => node.type === "FunctionDeclaration" && node.start >= source.indexOf("function getDataClient("))
  .map((node) => source.slice(node.start, node.end)).join("\n"), context);

const summary = `TVSCREDIT TVS CREDIT SERVICES LIMITED
MOHAMAD JAHID SHEKH 3026UC0342554
GJ3026UC0324485 TATA LPT 3718
2065000.00 48 APR 13.74% 05-AUG-26 05-JUL-30`;
const rows = [
  "1 05-08-2026 61907 32916 28991 2052611 0 61907",
  "2 05-09-2026 55960 33279 22681 2019332 0 55960",
  "3 05-10-2026 55960 34367 21593 1984965 0 55960",
  "4 05-11-2026 55960 34027 21933 1950938 0 55960"
];

test("TVS repayment columns and Gujarati-summary OCR fallback are parsed", () => {
  const parsed = context.parseBankPdfText(`${summary}\nRepayment Schedule:\n${rows.join("\n")}`, "tvc.pdf");
  assert.equal(parsed.owner, "MOHAMAD JAHID SHEKH");
  assert.equal(parsed.loanAccount, "3026UC0342554");
  assert.equal(parsed.regNo, "GJ3026UC0324485");
  assert.equal(parsed.loanAmount, "2065000");
  assert.equal(parsed.emiAmount, "55960");
  assert.equal(parsed.tenure, "48");
  assert.equal(parsed.interestRate, "13.74");
  assert.equal(parsed.emiStart, "05-08-2026");
  assert.equal(parsed.emiEnd, "05-07-2030");
  assert.equal(parsed.emiSchedule.length, 4);
  assert.equal(parsed.emiSchedule[0].amount, 61907);
  assert.equal(parsed.emiSchedule[0].principal, 32916);
  assert.equal(parsed.emiSchedule[0].interest, 28991);
  assert.equal(parsed.bankClosingPrincipal, "1984965");
});

test("TVS date-only rows recover when OCR drops S.No", () => {
  const text = `${summary}\nRepayment Schedule:\n${rows.map((row) => row.replace(/^\d+\s+/, "")).join("\n")}`;
  const parsed = context.parseTvsScheduleRows(text);
  assert.equal(parsed.length, 4);
  assert.equal(parsed[0].installment, 1);
  assert.equal(parsed[3].installment, 4);
});

test("unrelated text does not produce a TVS row", () => {
  assert.equal(context.parseTvsScheduleRows("TVS CREDIT 01-01-2024 4-01-2024").length, 0);
});
