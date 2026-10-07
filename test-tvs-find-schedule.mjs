import { readFileSync } from "fs";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

const pdfPath = "c:/Users/techn/Downloads/final code for kuber finance/finacel1/finacel/bank/tvc.pdf";

const data = new Uint8Array(readFileSync(pdfPath));
const pdfDocument = await getDocument({ data }).promise;

async function extractPdfNativeText(pdfDocument) {
  const pageCount = pdfDocument.numPages;
  const pageTexts = [];
  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    const page = await pdfDocument.getPage(pageNumber);
    const textContent = await page.getTextContent();

    const uniqueItems = [];
    for (const item of textContent.items) {
      if (!item.str) continue;
      const x = item.transform[4];
      const y = item.transform[5];
      const text = item.str.trim();
      if (!text) continue;

      const isDuplicate = uniqueItems.some(
        (existing) =>
          Math.abs(existing.y - y) <= 1.5 &&
          Math.abs(existing.x - x) <= 2.5 &&
          existing.text.trim() === text
      );

      if (!isDuplicate) {
        uniqueItems.push({ x, y, text: item.str, width: item.width });
      }
    }

    const linesBucket = [];
    for (const item of uniqueItems) {
      let bucket = linesBucket.find((b) => Math.abs(b.y - item.y) <= 3.5);
      if (!bucket) {
        bucket = { y: item.y, items: [] };
        linesBucket.push(bucket);
      }
      bucket.items.push(item);
    }
    linesBucket.sort((a, b) => b.y - a.y);
    const lines = linesBucket.map((bucket) => {
      const items = bucket.items.sort((a, b) => a.x - b.x);
      let line = "";
      for (let i = 0; i < items.length; i++) {
        if (i > 0) {
          const gap = items[i].x - (items[i - 1].x + (items[i - 1].width || 0));
          line += gap > 5 ? "  " : gap > 1 ? " " : "";
        }
        line += items[i].text;
      }
      return line.trim();
    }).filter(Boolean);
    pageTexts.push(lines.join("\n"));
  }
  return pageTexts.join("\n");
}

const text = await extractPdfNativeText(pdfDocument);

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function looseLabelPattern(label) {
  return label
    .trim()
    .split(/\s+/)
    .map((word) => word.split("").map((char) => escapeRegExp(char)).join("\\s*"))
    .join("\\s*");
}

function toNumber(value) {
  if (typeof value === "number") return isNaN(value) ? 0 : value;
  const cleaned = String(value ?? "").replace(/,/g, "").trim();
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function extractSignedMoneyLikeNumbers(valueText) {
  const matches = [...String(valueText ?? "").matchAll(/(-?[\d,]+(?:\.\d{1,2})?)/g)].map((m) => m[1].replace(/,/g, ""));
  return matches.filter((v) => !isNaN(parseFloat(v)));
}

function makeScheduleRow({ installment, dueDate, openingPrincipal, installmentAmount, principalPaid, interest, closingPrincipal, rate }) {
  return { installment, dueDate, openingPrincipal, installmentAmount, principalPaid, interest, closingPrincipal, rate };
}

function deriveAnnualRate(interest, openingPrincipal) {
  const i = toNumber(interest);
  const p = toNumber(openingPrincipal);
  if (i <= 0 || p <= 0) return "";
  const monthlyRate = i / p;
  const annualRate = monthlyRate * 12 * 100;
  return annualRate > 0 && annualRate <= 40 ? annualRate.toFixed(2) : "";
}

function parseTvsScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!/TVS\s+CREDIT/i.test(normalized)) return [];

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const values = extractSignedMoneyLikeNumbers(match[3]);
      if (values.length < 4) return null;
      const installmentAmount = values[0];
      const principalPaid = values[1];
      const interest = values[2];
      const closingPrincipal = values[3];
      const serviceTax = values.length >= 5 ? values[4] : "";
      if (toNumber(installmentAmount) <= 0 || toNumber(principalPaid) < 0) return null;
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({
        installment: match[1],
        dueDate: match[2],
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        serviceTax,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    })
    .filter(Boolean);

  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const values = extractSignedMoneyLikeNumbers(normalized.slice(rowStart, rowEnd));
    if (values.length < 4) return null;
    const installmentAmount = values[0];
    const principalPaid = values[1];
    const interest = values[2];
    const closingPrincipal = values[3];
    const serviceTax = values.length >= 5 ? values[4] : "";
    if (toNumber(installmentAmount) <= 0 || toNumber(principalPaid) < 0) return null;
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  }).filter(Boolean);
}

const rows = parseTvsScheduleRows(text);
console.log("=== TVS ROWS PARSED ===");
console.log("Total rows:", rows.length);
console.log("Row 1:", rows[0]);
console.log("Row 2:", rows[1]);
