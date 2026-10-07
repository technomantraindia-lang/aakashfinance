import { createWorker } from "tesseract.js";

const imagePath = "C:/Users/techn/.gemini/antigravity-ide/brain/cc25b798-c4ea-4a11-a6e4-9111b91b6d07/.user_uploaded/media_1791199782055.png";
const worker = await createWorker("eng");

await worker.setParameters({
  tessedit_pageseg_mode: "4"
});
const { data: psm4 } = await worker.recognize(imagePath);
await worker.terminate();

function formatOcrDate(rawDate) {
  let cleaned = String(rawDate ?? "").trim();
  const matchNoDash = cleaned.match(/^(\d{1,2})[./-]?(\d{1,2})[./-]?(\d{4})$/);
  if (matchNoDash) {
    const day = matchNoDash[1].padStart(2, "0");
    const month = matchNoDash[2].padStart(2, "0");
    const year = matchNoDash[3];
    return `${day}-${month}-${year}`;
  }
  return cleaned;
}

console.log("Original date '05-102026' formatted:", formatOcrDate("05-102026"));
