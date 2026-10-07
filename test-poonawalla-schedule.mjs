import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";
import { fileURLToPath } from "node:url";
import { parse } from "@babel/parser";

// Exercise the application's actual helpers, without loading the React UI.
const source = readFileSync(new URL("./src/App.jsx", import.meta.url), "utf8");
const ast = parse(source, { sourceType: "module", plugins: ["jsx"] });
const context = vm.createContext({ console });
vm.runInContext(ast.program.body
  .filter((node) => node.type === "FunctionDeclaration" && node.start >= source.indexOf("function toNumber("))
  .map((node) => source.slice(node.start, node.end)).join("\n"), context);

const header = `Installment Schedule Detail as on 26-Feb-2026
Instl. # Due Date Amount Effective Rate(%) / Days
Opening Balance Installment Principal Interest`;
const rows = [
  "1 20-Mar-2026 32,00,000.00 1,06,724.00 74,992.00 31,732.00 12.31/29",
  "2 20-Apr-2026 31,25,008.00 1,06,724.00 74,667.00 32,057.00 12.31/30",
  "3 20-May-2026 30,50,341.00 1,06,724.00 75,433.00 31,291.00 12.31/30",
  "4 20-Jun-2026 29,74,908.00 1,06,724.00 76,206.00 30,518.00 12.31/30",
  "5 20-Jul-2026 28,98,702.00 1,06,724.00 76,988.00 29,736.00 12.31/30"
];

test("screenshot columns, month names and effective rate map correctly", () => {
  const text = [header, ...rows].join("\n");
  const parsed = context.parseOpeningBalanceScheduleRows(text);
  assert.equal(parsed.length, 5);
  assert.equal(parsed[0].dueDate, "20-03-2026");
  assert.equal(Number(parsed[0].openingPrincipal), 3200000);
  assert.equal(Number(parsed[0].installmentAmount), 106724);
  assert.equal(Number(parsed[0].principalPaid), 74992);
  assert.equal(Number(parsed[0].interest), 31732);
  assert.equal(Number(parsed[0].closingPrincipal), 3125008);
  assert.equal(parsed[0].rate, "12.31");
  const result = context.parseBankPdfText(text, "poonawalla.pdf");
  assert.equal(result.interestRate, "12.31");
  assert.equal(result.emiSchedule.length, 5);
});

test("lost shaded headers still identify Poonawalla and its borrower", () => {
  const text = `Repayment Schedule For UCV123456789
Example Industries Limited
Note: Broken Period Interest forms part of the first installment
EDO asin main peed esa
${rows.join("\n")}
Poonawalla Fincorp Limited`;
  const result = context.parseBankPdfText(text, "renamed-scan.pdf");
  assert.equal(result.owner, "Example Industries Limited");
  assert.equal(result.loanAccount, "UCV123456789");
  assert.equal(result.financier, "POONAWALLA FINCORP LIMITED");
  assert.equal(Number(result.loanAmount), 3200000);
  assert.equal(result.emiAmount, "106724");
  assert.equal(result.interestRate, "12.31");
});

test("duplicate OCR passes and totals do not change closing balances", () => {
  const last = "36 20-Feb-2029 1,05,641.00 1,06,725.00 1,05,641.00 1,084.00 12.31/30";
  const block = [header, ...rows, last, "Total 38,42,065.00 32,00,000.00 6,42,065.00"].join("\n");
  const parsed = context.parseOpeningBalanceScheduleRows(`${block}\n${block}`);
  assert.equal(parsed.length, 6);
  assert.equal(Number(parsed.at(-1).closingPrincipal), 0);
  assert.equal(Number(parsed[4].closingPrincipal), 2821714);
});

test("wrapped cells, Indian commas and spaced date separators are accepted", () => {
  const wrapped = rows[0].replace("20-Mar-2026", "20 - Mar - 2026").replaceAll(" ", "\n");
  const parsed = context.parseOpeningBalanceScheduleRows(`${header}\n${wrapped}`);
  assert.equal(parsed.length, 1);
  assert.equal(Number(parsed[0].openingPrincipal), 3200000);
});

test("reject wrong arithmetic and impossible dates instead of fabricating amounts", () => {
  for (const row of [rows[0].replace("74,992.00", "74,99200"), rows[0].replace("20-Mar", "31-Feb")]) {
    assert.equal(context.parseOpeningBalanceScheduleRows(`${header}\n${row}`).length, 0);
  }
  assert.equal(context.parseOpeningBalanceScheduleRows(`OTHER BANK\n${rows[0]}`).length, 0);
});

test("rate is read from Rate / Days even when interest is less than 40", () => {
  const parsed = context.parseOpeningBalanceScheduleRows(`${header}\n1 20-Mar-2026 100.00 105.00 100.00 5.00 12.31/29`);
  assert.equal(parsed[0].rate, "12.31");
});

test("actual scanned PDF through production rendering, all OCR passes and import parser", {
  skip: !process.argv.includes("--ocr"), timeout: 240000
}, async () => {
  const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
  const { createCanvas } = await import("@napi-rs/canvas");
  const { createWorker } = await import("tesseract.js");
  const tasks = [];
  context.pdfjsLib = { getDocument(options) {
    const task = pdfjs.getDocument(options);
    tasks.push(task);
    return task;
  } };
  context.window = { document: { createElement: () => createCanvas(1, 1) } };
  context.createWorker = async () => {
    const worker = await createWorker("eng", 1, { langPath: fileURLToPath(new URL(".", import.meta.url)), gzip: false });
    const recognize = worker.recognize.bind(worker);
    worker.recognize = (canvas) => recognize(canvas.toBuffer("image/png"));
    return worker;
  };
  try {
    const bytes = readFileSync(new URL("../bank/poonawalla.pdf", import.meta.url));
    const text = await context.extractPdfTextWithOcr({
      name: "poonawalla.pdf", type: "application/pdf",
      arrayBuffer: async () => new Uint8Array(bytes).buffer
    });
    const result = context.normalizePdfFinanceRow(context.parseBankPdfText(text, "poonawalla.pdf"));
    assert.equal(result.owner, "Kotyark Industries Limited");
    assert.equal(result.loanAccount, "UCV0222UCV000019733958");
    assert.equal(result.financier, "POONAWALLA FINCORP LIMITED");
    assert.equal(Number(result.loanAmount), 3200000);
    assert.equal(result.emiAmount, "106724");
    assert.equal(result.tenure, "36");
    assert.equal(result.interestRate, "12.31");
    assert.equal(result.emiStart, "20-03-2026");
    assert.equal(result.emiEnd, "20-02-2029");
    assert.equal(result.emiSchedule.length, 36);
    const schedule = result.emiSchedule;
    assert.equal(schedule.reduce((sum, row) => sum + row.principal, 0), 3200000);
    assert.equal(schedule.reduce((sum, row) => sum + row.interest, 0), 642065);
    assert.equal(schedule.reduce((sum, row) => sum + row.amount, 0), 3842065);
    for (const [index, row] of schedule.entries()) {
      assert.equal(row.installment, index + 1);
      assert.equal(row.amount, row.principal + row.interest);
    }
    assert.equal(schedule[0].principal, 74992);
    assert.equal(schedule[0].interest, 31732);
    assert.equal(schedule.at(-1).amount, 106725);
  } finally {
    await Promise.all(tasks.map((task) => task.destroy()));
  }
});
