import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";
import vm from "node:vm";
import { parse } from "@babel/parser";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

const source = readFileSync(new URL("./src/App.jsx", import.meta.url), "utf8");
const ast = parse(source, { sourceType: "module", plugins: ["jsx"] });
const functions = ast.program.body.filter((node) => node.type === "FunctionDeclaration" &&
  (node.start >= source.indexOf("function getDataClient(") || node.id.name === "makeVehicleDetail"));
function findUpload(node) {
  if (!node || typeof node !== "object") return null;
  if (node.type === "VariableDeclarator" && node.id.name === "importClientPdf") return node.init;
  for (const value of Object.values(node)) {
    for (const child of Array.isArray(value) ? value : [value]) {
      const found = findUpload(child);
      if (found) return found;
    }
  }
  return null;
}
const uploadNode = findUpload(ast);
const bytes = new Uint8Array(readFileSync(new URL("../bank/Indostar_CVVADOD0348953_Repayment_Schedule.pdf", import.meta.url)));
const file = {
  name: "Indostar_CVVADOD0348953_Repayment_Schedule.pdf", type: "application/pdf",
  arrayBuffer: async () => bytes.slice().buffer
};
class FixedDate extends Date {
  constructor(...args) { super(...(args.length ? args : [2026, 9, 6, 12])); }
}

function setup(existingImports = []) {
  const tasks = [];
  const context = vm.createContext({
    console, Date: FixedDate,
    data: { clients: [{ id: "c1", name: "Saved Client Name" }], vehicles: [], dueTasks: [], clientImports: existingImports },
    pdfjsLib: { getDocument(options) { const task = getDocument(options); tasks.push(task); return task; } },
    setSaveStatus() {}, setToast(message) { context.toast = message; },
    notify: (value) => value, withAudit: (value) => value,
    persist(value) { context.data = value; context.saved = true; }
  });
  vm.runInContext(functions.map((node) => source.slice(node.start, node.end)).join("\n"), context);
  const releaseRate = ast.program.body.flatMap((node) => node.declarations || [])
    .find((node) => node.id.name === "BANK_RELEASE_RATE_PERCENT");
  vm.runInContext(`const ${source.slice(releaseRate.start, releaseRate.end)};`, context);
  const upload = vm.runInContext(`(${source.slice(uploadNode.start, uploadNode.end)})`, context);
  return { context, upload, cleanup: () => Promise.all(tasks.map((task) => task.destroy())) };
}

function checkDetail(context, detail) {
  assert.equal(Number(detail.loanAmount), 1250000);
  assert.equal(Number(detail.emiAmount), 44550);
  assert.equal(Number(detail.tenure), 35);
  assert.equal(Number(detail.interestRate), 15.0244);
  assert.equal(context.formatDisplayDate(detail.emiStart), "10-06-2026");
  assert.equal(context.formatDisplayDate(detail.emiEnd), "10-04-2029");
  assert.equal(detail.schedule.length, 35);
}

test("exact PDF through admin upload persists standalone row and displays actual loan fields", async () => {
  const { context, upload, cleanup } = setup();
  try {
    await upload({ target: { files: [file], value: "selected" } }, "c1");
    assert.equal(context.saved, true, context.toast);
    assert.equal(context.data.clientImports.length, 1);
    const asset = context.data.clientImports[0].rows[0];
    const vehicle = context.data.vehicles[0];
    assert.equal(vehicle.principal, 1138712);
    assert.equal(asset.owner, "SHEKH SAMIMBANU");
    assert.equal(context.data.clients[0].name, "Saved Client Name");
    checkDetail(context, context.makeVehicleDetail(vehicle, asset, "Saved Client Name"));
    // Reproduce the old screen path with no Excel/PDF row to fall back to.
    checkDetail(context, context.makeVehicleDetail(vehicle, undefined, "Saved Client Name"));
    checkDetail(context, context.makeVehicleDetail(vehicle, { loanAmount: "-", emiAmount: "-", tenure: "-" }, "Saved Client Name"));
    await upload({ target: { files: [file], value: "selected" } }, "c1");
    assert.equal(context.data.vehicles.length, 1);
    assert.equal(context.data.clientImports.length, 1);
  } finally { await cleanup(); }
});

test("reimport updates finance without changing an existing owner's name", async () => {
  const { context, upload, cleanup } = setup([{ id: "old-import", clientId: "c1", rows: [{
    owner: "Saved Owner Name", regNo: "GJ23AT-3640", loanAccount: "CVVADOD0348953",
    loanAmount: "1138712", emiAmount: "-", tenure: "-"
  }] }]);
  try {
    await upload({ target: { files: [file], value: "selected" } }, "c1");
    assert.equal(context.saved, true, context.toast);
    const asset = context.data.clientImports[0].rows[0];
    assert.equal(asset.owner, "Saved Owner Name");
    assert.equal(context.data.clients[0].name, "Saved Client Name");
    checkDetail(context, context.makeVehicleDetail(context.data.vehicles[0], asset, "Saved Client Name"));
  } finally { await cleanup(); }
});

test("backend stores calendar dates and configures MySQL DATE reads as strings", () => {
  const backendSource = readFileSync(new URL("../backend/server.js", import.meta.url), "utf8");
  const backendAst = parse(backendSource, { sourceType: "script" });
  const selected = backendAst.program.body.filter((node) => node.type === "FunctionDeclaration" && ["toMysqlDate", "normalizeVehicle"].includes(node.id.name));
  const backend = vm.createContext({ Date: FixedDate });
  vm.runInContext(selected.map((node) => backendSource.slice(node.start, node.end)).join("\n"), backend);
  const normalized = backend.normalizeVehicle({ emiStart: "10-06-2026", emiEnd: "2029-04-10", loanAmount: 1250000, principal: 1138712, emiAmount: 44550, tenure: 35 });
  assert.equal(normalized.emi_start, "2026-06-10");
  assert.equal(normalized.emi_end, "2029-04-10");
  assert.equal(normalized.loan_amount, 1250000);
  assert.equal(normalized.emi_amount, 44550);
  const pool = backendAst.program.body.flatMap((node) => node.declarations || []).find((node) => node.id.name === "pool");
  const option = pool.init.arguments[0].properties.find((node) => node.key.name === "dateStrings");
  assert.deepEqual(option.value.elements.map((node) => node.value), ["DATE"]);
});

test("PDF upload and reupload preserve manually saved registration exactly", async () => {
  for (const savedRegNo of ["GJ 23 AT 3640", "GJ23AT-9999"]) {
    const { context, upload, cleanup } = setup();
    context.data.vehicles.push({
      id: "saved-vehicle", clientId: "c1", regNo: savedRegNo,
      loanAccount: "CVVADOD0348953", principal: 0
    });
    try {
      for (let attempt = 0; attempt < 2; attempt += 1) {
        await upload({ target: { files: [file], value: "selected" } }, "c1");
        assert.equal(context.saved, true, context.toast);
        assert.equal(context.data.vehicles.length, 1);
        assert.equal(context.data.vehicles[0].regNo, savedRegNo);
        assert.equal(context.data.clientImports[0].rows[0].regNo, savedRegNo);
        const detail = context.makeVehicleDetail(context.data.vehicles[0], context.data.clientImports[0].rows[0], "Saved Client Name");
        assert.equal(detail.regNo, savedRegNo);
        checkDetail(context, detail);
        // The persistence payload also retains exactly what the user entered.
        context.data = JSON.parse(JSON.stringify(context.data));
      }
    } finally { await cleanup(); }
  }
});

test("an imported registration wins over the PDF when creating its vehicle", async () => {
  const savedRegNo = "GJ23AT-9999";
  const { context, upload, cleanup } = setup([{ id: "saved-import", clientId: "c1", rows: [{
    regNo: savedRegNo, owner: "Saved Owner", loanAccount: "CVVADOD0348953"
  }] }]);
  try {
    await upload({ target: { files: [file], value: "selected" } }, "c1");
    assert.equal(context.saved, true, context.toast);
    assert.equal(context.data.vehicles[0].regNo, savedRegNo);
    assert.equal(context.data.clientImports[0].rows[0].regNo, savedRegNo);
  } finally { await cleanup(); }
});

test("vehicle merge only fills a missing registration", async () => {
  const { context, cleanup } = setup();
  try {
    const row = { regNo: "GJ23AT-3640", bankClosingPrincipal: "1138712" };
    assert.equal(context.mergePdfIntoVehicle({ regNo: "GJ 23 AT 9999" }, row).regNo, "GJ 23 AT 9999");
    assert.equal(context.mergePdfIntoVehicle({ regNo: "" }, row).regNo, "GJ23AT-3640");
    assert.equal(context.mergePdfIntoVehicle({ regNo: "-" }, row).regNo, "GJ23AT-3640");
  } finally { await cleanup(); }
});
