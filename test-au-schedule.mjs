import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { parse } from "@babel/parser";

const source = readFileSync(new URL("./src/App.jsx", import.meta.url), "utf8");
class ImportDate extends Date {
  constructor(...args) {
    super(...(args.length ? args : [2026, 9, 7, 12]));
  }
}
const context = vm.createContext({ console, Date: ImportDate });
const ast = parse(source, { sourceType: "module", plugins: ["jsx"] });
vm.runInContext(
  ast.program.body
    .filter((node) => node.type === "FunctionDeclaration" && node.start >= source.indexOf("function toNumber("))
    .map((node) => source.slice(node.start, node.end))
    .join("\n"),
  context
);

const text = `
A U SMALL FINANCE BANK LIMITED
Customer: CHETAN SUTHAR
Loan Account Number: 9001011069137729
Vehicle No.: GJ06BT7868
Loan Amount: INR 2,600,000
Rate of Interest: 12.50
Tenure: 36 Months
EMI Amount: INR 88,244
EMI Period: 10-09-2026 to 10-08-2029
EMI No. Due Date EMI Amount Principal Interest Closing Principal
1 10-09-2026 88,244 61,161 27,083 2,538,839
2 10-10-2026 88,244 61,798 26,446 2,477,041
3 10-11-2026 88,244 62,442 25,802 2,414,599
`;

const parsed = context.normalizePdfFinanceRow(context.parseBankPdfText(text, "au-repayment-schedule.pdf"));
assert.equal(parsed.financier, "AU SMALL FINANCE");
assert.equal(parsed.loanAccount, "9001011069137729");
assert.equal(parsed.regNo, "GJ06BT7868");
assert.equal(Number(parsed.loanAmount), 2600000);
assert.equal(Number(parsed.emiAmount), 88244);
assert.equal(Number(parsed.interestRate), 12.5);
assert.equal(Number(parsed.tenure), 36);
assert.equal(parsed.emiStart, "10-09-2026");
assert.equal(parsed.emiEnd, "10-08-2029");
assert.equal(parsed.emiSchedule.length, 3);
assert.equal(parsed.scheduleParsed, "yes");

const badClosingText = text.replace("2,538,839", "25,653,712");
const sanitized = context.normalizePdfFinanceRow(context.parseBankPdfText(badClosingText, "au-repayment-schedule.pdf"));
assert.equal(sanitized.bankClosingPrincipal, "");
const sanitizedForMerge = context.sanitizePdfRowForMerge(sanitized);
assert.ok(Object.prototype.hasOwnProperty.call(sanitizedForMerge, "bankClosingPrincipal"));
assert.equal(sanitizedForMerge.bankClosingPrincipal, "");

console.log("AU Small Finance repayment schedule parser passed");
