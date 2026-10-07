import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import test from "node:test";
import vm from "node:vm";
import { parse } from "@babel/parser";

const source = readFileSync(new URL("./src/App.jsx", import.meta.url), "utf8");
class ImportDate extends Date {
  constructor(...args) { super(...(args.length ? args : [2026, 9, 6, 12])); }
}
const context = vm.createContext({ console, Date: ImportDate });
const ast = parse(source, { sourceType: "module", plugins: ["jsx"] });
vm.runInContext(ast.program.body
  .filter((node) => node.type === "FunctionDeclaration" && node.start >= source.indexOf("function toNumber("))
  .map((node) => source.slice(node.start, node.end)).join("\n"), context);

const summary = `IndoStar Capital Finance Ltd. — Commercial Vehicle Loan
Customer SHEKH SAMIMBANU Agreement No CVVADOD0348953
Loan Amount 12,50,000 Rate of Interest 15.0244 % p.a.
Vehicle No. GJ23AT-3640 Tenure / EMIs 35 EMIs
EMI Amount 44,550 EMI Period 10/06/2026 to 10/04/2029`;
const header = "EMI No. Due Date EMI Amount Principal Interest Balance Outstanding";
const rows = [
  "1 10/06/2026 44,550 22,639 21,911 12,27,361",
  "2 10/07/2026 44,550 29,183 15,367 11,98,178",
  "3 10/08/2026 44,550 29,548 15,002 11,68,630",
  "4 10/09/2026 44,550 29,918 14,632 11,38,712",
  "5 10/10/2026 44,550 30,293 14,257 11,08,419",
  "6 10/11/2026 44,550 30,672 13,878 10,77,747"
];
const screenshotText = [summary, header, ...rows].join("\n");

function checkSummary(result, count) {
  assert.equal(result.owner, "SHEKH SAMIMBANU");
  assert.equal(result.loanAccount, "CVVADOD0348953");
  assert.equal(result.regNo.replace(/[^A-Z0-9]/g, ""), "GJ23AT3640");
  assert.equal(result.financier, "INDOSTAR");
  assert.equal(Number(result.loanAmount), 1250000);
  assert.equal(Number(result.emiAmount), 44550);
  assert.equal(result.tenure, "35");
  assert.equal(result.interestRate, "15.0244");
  assert.equal(result.emiStart, "10-06-2026");
  assert.equal(result.emiEnd, "10-04-2029");
  assert.equal(Number(result.bankClosingPrincipal), 1138712);
  assert.equal(result.emiSchedule.length, count);
  assert.equal(result.emiSchedule[0].principal, 22639);
  assert.equal(result.emiSchedule[0].interest, 21911);
}

test("screenshot summary and six columns, including four-decimal contractual rate", () => {
  checkSummary(context.parseBankPdfText(screenshotText, "renamed.pdf"), 6);
  const parsed = context.parseIndostarScheduleRows(screenshotText);
  assert.equal(parsed[0].rate, "15.0244");
  assert.equal(Number(parsed[0].closingPrincipal), 1227361);
});

test("flattened OCR, wrapped cells, table borders and repeated passes", () => {
  const flat = screenshotText.replaceAll("\n", " ");
  const wrapped = [summary, header, ...rows.map((row) => row.replaceAll(" ", "\n"))].join("\n");
  const bordered = screenshotText.replaceAll(" ", " | ");
  for (const text of [flat, wrapped, bordered, `${screenshotText}\n${wrapped}`]) {
    checkSummary(context.parseBankPdfText(text, "renamed.pdf"), 6);
  }
});

test("spaced numeric dates and split Indian comma amounts", () => {
  const text = screenshotText.replaceAll("10/06/2026", "10 / 06 / 2026").replace("12,27,361", "12, 27,361");
  const parsed = context.parseIndostarScheduleRows(text);
  assert.equal(parsed.length, 6);
  assert.equal(parsed[0].dueDate, "10-06-2026");
  assert.equal(Number(parsed[0].closingPrincipal), 1227361);
});

test("totals cannot become closing balance; invalid arithmetic and dates are rejected", () => {
  const good = context.parseIndostarScheduleRows(`${screenshotText}\nTotal 2,67,300 1,72,253 95,047`);
  assert.equal(good.length, 6);
  assert.equal(Number(good.at(-1).closingPrincipal), 1077747);
  for (const row of [rows[0].replace("22,639", "2263900"), rows[0].replace("10/06", "31/02")]) {
    assert.equal(context.parseIndostarScheduleRows(`${summary}\n${header}\n${row}`).length, 0);
  }
});

test("original bank PDF keeps MI separate from outstanding and excludes disbursal", async () => {
  const { getDocument } = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const task = getDocument({ data: new Uint8Array(readFileSync(new URL("../bank/indostar.pdf", import.meta.url))) });
  try {
    const document = await task.promise;
    const text = await context.extractPdfNativeText(document);
    checkSummary(context.normalizePdfFinanceRow(context.parseBankPdfText(text, "renamed.pdf")), 35);
    assert.equal(context.parseTvsScheduleRows(text).length, 0);
    assert.equal(Object.keys(context.findTvsCreditSummaryValues(text)).length, 0);
    const parsed = context.parseIndostarScheduleRows(text);
    assert.equal(parsed[0].installment, 1);
    assert.equal(Number(parsed[0].openingPrincipal), 1250000);
    assert.equal(Number(parsed[0].serviceTax), 0);
    assert.equal(Number(parsed[0].closingPrincipal), 1227361);
    assert.equal(Number(parsed.at(-1).closingPrincipal), 0);
    assert.equal(Number(parsed.at(-1).installmentAmount), 44469);
    assert.equal(parsed.reduce((sum, row) => sum + Number(row.principalPaid), 0), 1250000);
  } finally {
    await task.destroy();
  }
});

test("TVS layout continues to use its own column order", () => {
  const text = "TVS CREDIT Repayment Schedule\nS.No Due Date EMI Principal Interest Balance Principle Vehicle Insurance\n1 10/06/2026 44550 22639 21911 1227361 0";
  assert.equal(context.isIndostarRepaymentLayout(text), false);
  const result = context.findScheduleTableValues(text);
  assert.equal(Number(result.loanAmount), 1250000);
  assert.equal(Number(result.bankClosingPrincipal), 1227361);
});

test("image upload reads the rendered table and resolves agreement against client records", {
  skip: !process.argv.includes("--ocr"), timeout: 180000
}, async () => {
  // Recreate the supplied table layout for a repeatable OCR test. This is a
  // generated fixture, not the original attached screenshot.
  const { createCanvas } = await import("@napi-rs/canvas");
  const { createWorker } = await import("tesseract.js");
  const canvas = createCanvas(2288, 1030);
  const ctx = canvas.getContext("2d");
  ctx.scale(2, 2);
  ctx.fillStyle = "white"; ctx.fillRect(0, 0, 1144, 515);
  ctx.font = "19px Arial"; ctx.fillStyle = "#000060";
  ctx.fillText("IndoStar Capital Finance Ltd. - Commercial Vehicle Loan", 318, 37);
  const cells = [
    ["Customer", "SHEKH SAMIMBANU", "Agreement No", "CVVADOD0348953"],
    ["Loan Amount", "12,50,000", "Rate of Interest", "15.0244 % p.a."],
    ["Vehicle No.", "GJ23AT-3640", "Tenure / EMIs", "35 EMIs"],
    ["EMI Amount", "44,550", "EMI Period", "10/06/2026 to 10/04/2029"]
  ];
  const summaryX = [30, 207, 561, 739, 1093];
  for (const [r, row] of cells.entries()) for (const [c, text] of row.entries()) {
    ctx.fillStyle = c % 2 === 0 ? "#ebf2f6" : "white";
    ctx.fillRect(summaryX[c], 56 + r * 42, summaryX[c + 1] - summaryX[c], 42);
    ctx.strokeStyle = "#cccccc"; ctx.strokeRect(summaryX[c], 56 + r * 42, summaryX[c + 1] - summaryX[c], 42);
    ctx.fillStyle = "#111"; ctx.font = `${c % 2 === 0 ? "bold " : ""}19px Arial`;
    ctx.fillText(text, summaryX[c] + 12, 83 + r * 42);
  }
  const columns = [30, 136, 290, 479, 668, 857, 1093];
  const table = [["EMI No.", "Due Date", "EMI Amount", "Principal", "Interest", "Balance Outstanding"], ...rows.map((row) => row.split(" "))];
  for (const [r, row] of table.entries()) for (const [c, text] of row.entries()) {
    ctx.fillStyle = r === 0 ? "#244e77" : r % 2 ? "white" : "#f6f6f6";
    ctx.fillRect(columns[c], 244 + r * 38, columns[c + 1] - columns[c], 38);
    ctx.strokeStyle = "#cccccc"; ctx.strokeRect(columns[c], 244 + r * 38, columns[c + 1] - columns[c], 38);
    ctx.fillStyle = r === 0 ? "white" : "#111"; ctx.font = `${r === 0 ? "bold " : ""}19px Arial`;
    ctx.fillText(text, columns[c] + 10, 270 + r * 38);
  }
  context.createWorker = () => createWorker("eng", 1, { langPath: fileURLToPath(new URL(".", import.meta.url)), gzip: false });
  const bytes = canvas.toBuffer("image/png");
  bytes.name = "indostar.png";
  bytes.type = "image/png";
  const text = await context.extractPdfTextWithOcr(bytes);
  try {
    const result = context.parseBankPdfText(text, bytes.name);
    // OCR can insert O near a zero in an identifier. Exercise the same
    // existing-record reconciliation used by the admin import, not a guessed
    // string replacement or a hardcoded correction inside the parser.
    const assets = [
      { loanAccount: "CVVADOD0348953", owner: "SHEKH SAMIMBANU", regNo: "GJ23AT-3640" },
      { loanAccount: "CVVADOD9999999", owner: "OTHER CUSTOMER", regNo: "GJ23AT-9999" }
    ];
    const agreement = context.findProfileAgreementInText(text, assets) || context.inferProfileAgreementFromPdf(result, assets);
    assert.equal(agreement, "CVVADOD0348953");
    if (agreement) result.loanAccount = agreement;
    checkSummary(result, 6);
  } catch (error) {
    throw new Error(`${error.message}\nOCR fixture output:\n${text}`, { cause: error });
  }
});
