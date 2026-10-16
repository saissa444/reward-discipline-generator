"use strict";
// Pass private workbooks from OUTSIDE the repository. No workbook, names, IDs or
// screenshots are written to the repository or included in console output.
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const http = require("node:http");
const {webkit, chromium, devices} = require("playwright");
const xlsx = require("../vendor/xlsx.full.min.js");
const core = require("../app-core.js");
let baseURL = process.env.TEST_URL;
const rosterPath = process.env.ROSTER_PATH;
const templatePath = process.env.TEMPLATE_PATH;
const engine = process.env.TEST_BROWSER || "webkit";

async function main() {
  assert(rosterPath && templatePath, "Set ROSTER_PATH and TEMPLATE_PATH to private files outside the repository.");
  const {students} = core.extractRoster(xlsx.read(fs.readFileSync(rosterPath)), xlsx);
  const template = core.extractTemplate(xlsx.read(fs.readFileSync(templatePath)), xlsx);
  const browser = await ({webkit, chromium}[engine]).launch({headless: true});
  let server;
  if (!baseURL) {
    const root = path.resolve(__dirname, "..");
    server = http.createServer((request, response) => {
      const url = new URL(request.url, "http://localhost");
      const file = path.resolve(root, "." + decodeURIComponent(url.pathname === "/" ? "/index.html" : url.pathname));
      if (!file.startsWith(root + path.sep) || !fs.existsSync(file) || !fs.statSync(file).isFile()) { response.writeHead(404); response.end(); return; }
      response.setHeader("Content-Type", file.endsWith(".js") ? "application/javascript" : "text/html; charset=utf-8");
      response.end(fs.readFileSync(file));
    });
    await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
    baseURL = "http://127.0.0.1:" + server.address().port + "/";
  }
  const context = await browser.newContext({...devices["iPhone 13"], acceptDownloads: true});
  // Exercise Safari's FileReader path, even without the newer Blob API.
  await context.addInitScript(() => { Blob.prototype.arrayBuffer = undefined; });
  const page = await context.newPage();
  const errors = [], nonLocalRequests = [], writes = [];
  page.on("pageerror", error => errors.push(error.message));
  page.on("console", message => { if (message.type() === "error") errors.push(message.text()); });
  page.on("request", request => {
    const url = new URL(request.url());
    if (["http:", "https:"].includes(url.protocol) && url.origin !== new URL(baseURL).origin) nonLocalRequests.push(url.origin);
    if (request.method() !== "GET") writes.push(request.method());
  });
  const $ = id => page.locator("#" + id);
  const text = async (id, expression) => assert(expression.test(await $(id).innerText()), "Status assertion failed for #" + id);
  const disabled = async id => assert.equal(await $(id).isDisabled(), true);
  const load = async (id, file, statusId) => {
    await $(id).setInputFiles(file);
    await page.waitForFunction(id => document.getElementById(id).classList.contains("ok"), statusId);
  };
  const choose = async (article, clause, item) => {
    await $("article").selectOption(article);
    await $("clause").selectOption(clause);
    if (item) await $("item").selectOption(item);
  };
  let downloadsChecked = 0;
  const checkDownload = async expected => {
    const downloaded = page.waitForEvent("download");
    await $("exportBtn").click();
    const download = await downloaded;
    assert.match(download.suggestedFilename(), /^獎懲紀錄_\d{8}\.xlsx$/);
    const temporary = path.join(fs.mkdtempSync(path.join(os.tmpdir(), "reward-export-")), "result.xlsx");
    try {
      await download.saveAs(temporary);
      const workbook = xlsx.read(fs.readFileSync(temporary));
      assert.equal(workbook.SheetNames[0], template.sheetName);
      const data = xlsx.utils.sheet_to_json(workbook.Sheets[template.sheetName], {header: 1, defval: "", raw: false});
      assert.deepEqual(data[0], template.headers);
      assert.equal(data.length - 1, expected.count);
      data.slice(1).forEach(row => assert.equal(row[template.fields.indexOf("amount")],
        template.defaultAmounts[row[template.fields.indexOf("article")]], "Exported base amount must match its article"));
      assert.equal(data[1][0], expected.rows?.[0]?.school || students[0].school);
      assert.equal(data[1][1], expected.rows?.[0]?.name || students[0].name);
      for (const [field, value] of Object.entries(expected.fields || {})) assert.equal(data[1][template.fields.indexOf(field)], value);
      (expected.rows || []).forEach((row, index) => {
        for (const [field, value] of Object.entries(row)) assert.equal(data[index + 1][template.fields.indexOf(field)], value);
      });
      downloadsChecked++;
    } finally { fs.rmSync(path.dirname(temporary), {recursive: true, force: true}); }
  };

  try {
    await page.goto(baseURL);
    await page.waitForFunction(() => document.documentElement.dataset.appReady === "true");
    await text("startupStatus", /已就緒/);
    await disabled("previewBtn");
    await disabled("clearRecordsBtn");
    assert.equal(await page.locator("#editApplyAmount").count(), 0);
    await load("rosterFile", rosterPath, "rosterStatus");
    await text("rosterStatus", new RegExp("姓名查詢.*" + students.length));
    await load("templateFile", templatePath, "templateStatus");
    await text("templateStatus", new RegExp(String(template.rules.length)));

    // Every actual roster row must match by each supported identifier.
    for (const field of ["female", "team", "name"]) {
      const counts = new Map();
      students.forEach(student => {
        const value = core.normalize(student[field]);
        counts.set(value, (counts.get(value) || 0) + 1);
      });
      const ambiguous = students.filter(student => counts.get(core.normalize(student[field])) > 1).length;
      await $("studentInput").fill(students.map(student => student[field]).filter(Boolean).join("、"));
      await text("studentCheckPanel", new RegExp("辨識 " + (students.length - ambiguous) + " 位"));
      assert.equal(await page.locator("#studentCheckPanel .error").count(), ambiguous);
    }
    await $("studentInput").fill([Number(students[0].female), students[0].team, students[0].name, students[1].female].join("、"));
    await text("studentCheckPanel", /共 3 筆（重複 2 次/);
    assert.equal(await page.locator("#studentCheckPanel .error").count(), 1);

    // Walk every real rule through the actual cascading DOM controls.
    for (const rule of template.rules) {
      await choose(rule.article, rule.clause, rule.item);
      assert.equal(await $("reason").inputValue(), rule.reason);
      assert.equal(await $("amount").inputValue(), template.defaultAmounts[rule.article]);
      assert.equal(await $("amount").isEditable(), false);
      if (!rule.item) await disabled("item");
    }
    const withItem = template.rules.find(rule => rule.item);
    await $("article").selectOption(withItem.article);
    await $("clause").selectOption(withItem.clause);
    await disabled("amount");
    await $("previewBtn").click();
    await text("resultStatus", /選擇.*目/);
    await $("item").selectOption(withItem.item);
    await disabled("adjustAmount");
    await $("adjustReason").selectOption(template.adjustmentOptions[0]);
    await $("adjustAmount").selectOption(template.amountOptions[0]);
    await $("previewBtn").click();
    await text("resultStatus", /已新增 4 筆.*目前共 4 筆/);
    assert.equal(await page.locator("#previewBody tr").count(), 4);
    assert.equal(await page.locator("#previewBody tr:first-child td[data-field]").count(), 9);
    await checkDownload({count: 4, fields: {...withItem, amount: template.defaultAmounts[withItem.article], adjustReason: template.adjustmentOptions[0], adjustAmount: template.amountOptions[0]}});
    await text("duplicateStatus", /共 3 筆，重複 2 次/);
    assert.equal(await $("duplicateStatus").evaluate(element => getComputedStyle(element).color), "rgb(180, 35, 24)");

    // Draft changes and failed batches preserve the accumulated export.
    const firstBatchURL = await $("exportBtn").getAttribute("href");
    await $("studentInput").fill(students[0].female + ",不存在的測試學生");
    assert.equal(await $("exportBtn").getAttribute("href"), firstBatchURL);
    await $("previewBtn").click();
    await text("resultStatus", /已新增 1 筆.*目前共 5 筆/);
    await text("resultStatus", /已略過 1 筆/);
    assert.equal(await page.locator("#studentCheckPanel .error").count(), 1);
    assert.equal(await $("resultStatus").getAttribute("class"), "status warning");
    assert.equal(await $("exportBtn").getAttribute("aria-disabled"), "false");
    await checkDownload({count: 5, fields: withItem});
    await text("duplicateStatus", /共 4 筆，重複 3 次/);
    const partialBatchURL = await $("exportBtn").getAttribute("href");
    await $("studentInput").fill("不存在的測試學生");
    await $("previewBtn").click();
    await text("resultStatus", /沒有可匯出/);
    assert.equal(await $("exportBtn").getAttribute("aria-disabled"), "false");
    assert.equal(await $("exportBtn").getAttribute("href"), partialBatchURL);
    assert.equal(await page.locator("#previewBody tr").count(), 5);
    await $("clearRecordsBtn").click();
    await $("studentInput").fill(students[0].female);
    const withoutItem = template.rules.find(rule => !rule.item);
    await choose(withoutItem.article, withoutItem.clause);
    await $("previewBtn").click();
    await checkDownload({count: 1, fields: {...withoutItem, adjustReason: "", adjustAmount: ""}});
    const withoutItemURL = await $("exportBtn").getAttribute("href");
    await choose(withItem.article, withItem.clause, withItem.item);
    assert.equal(await $("exportBtn").getAttribute("href"), withoutItemURL);

    // Single and selected-row edits must update the actual downloaded workbook.
    await $("clearRecordsBtn").click();
    await $("studentInput").fill(students.slice(0, 3).map(student => student.female).concat("不存在的測試學生").join(","));
    await choose(withoutItem.article, withoutItem.clause);
    await $("previewBtn").click();
    const cell = (index, field) => page.locator("#previewBody tr").nth(index).locator('[data-field="' + field + '"]');
    const editRow = index => page.locator("#previewBody .edit-row").nth(index).click();
    const baseAmount = template.defaultAmounts[withoutItem.article];
    const ruleAmount = template.defaultAmounts[withItem.article];
    assert.notEqual(ruleAmount, baseAmount);
    const initialURL = await $("exportBtn").getAttribute("href");
    await disabled("editSelectedBtn");
    await editRow(0);
    for (const rule of template.rules) {
      await $("editArticle").selectOption(rule.article);
      await $("editClause").selectOption(rule.clause);
      if (rule.item) await $("editItem").selectOption(rule.item);
      assert.equal(await $("editReason").inputValue(), rule.reason);
      assert.equal(await $("editAmount").inputValue(), template.defaultAmounts[rule.article]);
      assert.equal(await $("editAmount").isEditable(), false);
    }
    await $("cancelEditBtn").click();
    assert.equal(await $("exportBtn").getAttribute("href"), initialURL);
    await editRow(0);
    assert.equal(await $("editArticle").inputValue(), withoutItem.article);
    assert.equal(await $("editAmount").inputValue(), baseAmount);
    await $("editArticle").selectOption(withItem.article);
    await $("editClause").selectOption(withItem.clause);
    await $("editItem").selectOption(withItem.item);
    await $("editAdjustReason").selectOption(template.adjustmentOptions[0]);
    await $("editAdjustAmount").selectOption(template.amountOptions[0]);
    await $("saveEditBtn").click();
    assert.equal(await $("recordEditor").isVisible(), false);
    await text("resultStatus", /已略過 1 筆/);
    assert.equal(await $("resultStatus").getAttribute("class"), "status warning");
    assert.equal(await cell(0, "amount").innerText(), ruleAmount);
    assert.equal(await cell(1, "amount").innerText(), baseAmount);
    assert.equal(await cell(2, "amount").innerText(), baseAmount);
    const firstAdjusted = {school: students[0].school, name: students[0].name, ...withItem, amount: ruleAmount, adjustReason: template.adjustmentOptions[0], adjustAmount: template.amountOptions[0]};
    const untouched = {school: students[1].school, name: students[1].name, ...withoutItem, amount: baseAmount, adjustReason: "", adjustAmount: ""};
    const third = {school: students[2].school, name: students[2].name, ...withoutItem, amount: baseAmount, adjustReason: "", adjustAmount: ""};
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    const savedURL = await $("exportBtn").getAttribute("href");
    await editRow(0);
    await $("editArticle").selectOption(withoutItem.article);
    await $("editClause").selectOption(withoutItem.clause);
    await $("cancelEditBtn").click();
    assert.equal(await $("exportBtn").getAttribute("href"), savedURL);
    assert.equal(await cell(0, "amount").innerText(), ruleAmount);

    await page.locator("#previewBody .row-select").nth(0).check();
    await page.locator("#previewBody .row-select").nth(2).check();
    await text("selectedCount", /2 \/ 3/);
    assert.equal(await $("selectAll").evaluate(element => element.indeterminate), true);
    await $("editSelectedBtn").click();
    for (const id of ["editApplyRule", "editApplyReason", "editApplyAdjust"]) assert.equal(await $(id).isChecked(), false);
    await $("saveEditBtn").click();
    await text("editStatus", /至少一組/);
    assert.equal(await $("recordEditor").isVisible(), true);
    await $("editApplyRule").check();
    await $("editArticle").selectOption(withoutItem.article);
    await $("editClause").selectOption(withoutItem.clause);
    await $("saveEditBtn").click();
    await text("resultStatus", /已更新 2 筆/);
    await text("resultStatus", /已略過 1 筆/);
    Object.assign(firstAdjusted, withoutItem, {amount: baseAmount});
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    assert.equal(await cell(1, "amount").innerText(), baseAmount);
    assert.equal(await cell(0, "adjustReason").innerText(), template.adjustmentOptions[0]);
    assert.equal(await cell(2, "adjustReason").innerText(), "");

    // Invalid batch edits and cancellation leave the previous export untouched.
    const beforeInvalid = await $("exportBtn").getAttribute("href");
    await $("editSelectedBtn").click();
    await $("editApplyRule").check();
    await $("editArticle").selectOption(withItem.article);
    await $("editClause").selectOption(withItem.clause);
    await $("saveEditBtn").click();
    await text("editStatus", /選擇.*目/);
    assert.equal(await $("recordEditor").isVisible(), true);
    assert.equal(await $("exportBtn").getAttribute("href"), beforeInvalid);
    await $("cancelEditBtn").click();
    await $("editSelectedBtn").click();
    await $("editApplyRule").check();
    await $("editArticle").selectOption(withItem.article);
    await $("editClause").selectOption(withItem.clause);
    await $("editItem").selectOption(withItem.item);
    await $("saveEditBtn").click();
    Object.assign(firstAdjusted, withItem, {amount: ruleAmount});
    Object.assign(third, withItem, {amount: ruleAmount});
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    await editRow(0);
    await $("editAdjustReason").selectOption("");
    await disabled("editAdjustAmount");
    await $("saveEditBtn").click();
    firstAdjusted.adjustReason = "";
    firstAdjusted.adjustAmount = "";
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});

    // Custom reasons survive other edits and round-trip through the download.
    const singleReason = "單筆自行補充的獎懲事由。\n保留第二行內容。";
    await editRow(0);
    await $("editReason").fill(singleReason);
    await $("editApplyRule").uncheck();
    await $("editApplyRule").check();
    await $("editApplyAdjust").uncheck();
    await $("editApplyAdjust").check();
    assert.equal(await $("editReason").inputValue(), singleReason);
    await $("saveEditBtn").click();
    firstAdjusted.reason = singleReason;
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    await editRow(0);
    assert.equal(await $("editReason").inputValue(), singleReason);
    await $("editApplyRule").uncheck();
    await $("editApplyReason").uncheck();
    await $("editAdjustReason").selectOption(template.adjustmentOptions[0]);
    await $("editAdjustAmount").selectOption(template.amountOptions[0]);
    await $("saveEditBtn").click();
    firstAdjusted.adjustReason = template.adjustmentOptions[0];
    firstAdjusted.adjustAmount = template.amountOptions[0];
    assert.equal(await cell(0, "reason").innerText(), singleReason);
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});

    const batchReason = "核取多筆後共同填寫的自訂事由。\n第二行補充。";
    await $("editSelectedBtn").click();
    await $("editApplyReason").check();
    await $("editReason").fill(batchReason);
    await $("editApplyAdjust").check();
    await $("editApplyAdjust").uncheck();
    assert.equal(await $("editReason").inputValue(), batchReason);
    await $("saveEditBtn").click();
    firstAdjusted.reason = batchReason;
    third.reason = batchReason;
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    const customURL = await $("exportBtn").getAttribute("href");
    await editRow(0);
    await $("editReason").fill(" \n ");
    await $("saveEditBtn").click();
    await text("editStatus", /填寫獎懲事由/);
    assert.equal(await $("recordEditor").isVisible(), true);
    assert.equal(await $("exportBtn").getAttribute("href"), customURL);
    await $("cancelEditBtn").click();
    await editRow(0);
    await $("editArticle").selectOption(withoutItem.article);
    await $("editClause").selectOption(withoutItem.clause);
    assert.equal(await $("editReason").inputValue(), withoutItem.reason);
    await $("cancelEditBtn").click();
    assert.equal(await $("exportBtn").getAttribute("href"), customURL);

    await $("selectAll").check();
    await text("selectedCount", /3 \/ 3/);
    await $("selectAll").uncheck();
    await text("selectedCount", /0 \/ 3/);
    await disabled("editSelectedBtn");

    // A new batch keeps earlier edits and selection, including repeat students.
    await page.locator("#previewBody .row-select").nth(1).check();
    const beforeAppend = await $("exportBtn").getAttribute("href");
    await $("studentInput").fill([students[0].team, students[0].name, students[1].female].join(","));
    await choose(withItem.article, withItem.clause, withItem.item);
    const secondBatchReason = "接續新增第二批，保留原有編輯內容。";
    await $("reason").fill(secondBatchReason);
    assert.equal(await $("exportBtn").getAttribute("href"), beforeAppend);
    assert.equal(await page.locator("#previewBody tr").count(), 3);
    await $("previewBtn").click();
    await text("resultStatus", /已新增 3 筆.*目前共 6 筆/);
    await text("selectedCount", /1 \/ 6/);
    assert.equal(await page.locator("#previewBody .row-select").nth(1).isChecked(), true);
    assert.equal(await page.locator("#previewBody .row-select").nth(3).isChecked(), false);
    const appended = [students[0], students[0], students[1]].map(student => ({school: student.school, name: student.name,
      ...withItem, reason: secondBatchReason, amount: ruleAmount, adjustReason: "", adjustAmount: ""}));
    await checkDownload({count: 6, rows: [firstAdjusted, untouched, third, ...appended]});
    await text("duplicateStatus", /共 3 筆，重複 2 次/);
    await text("duplicateStatus", /共 2 筆，重複 1 次/);
    await editRow(3);
    await $("editReason").fill("只修改其中一筆重複紀錄。");
    await $("saveEditBtn").click();
    appended[0].reason = "只修改其中一筆重複紀錄。";
    await checkDownload({count: 6, rows: [firstAdjusted, untouched, third, ...appended]});
    for (const index of [5, 4, 3]) await page.locator("#previewBody .delete-row").nth(index).click();
    assert.equal(await $("duplicateStatus").isVisible(), false);
    await text("selectedCount", /1 \/ 3/);
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    await $("selectAll").check();
    await $("selectAll").uncheck();
    await $("studentInput").fill(students.slice(0, 3).map(student => student.female).concat("不存在的測試學生").join(","));
    await choose(withoutItem.article, withoutItem.clause);
    // Only the upper draft changed; the previously edited preview remains.
    assert.equal(await cell(0, "reason").innerText(), firstAdjusted.reason);

    // Deleting a row remaps selection to the same students and rebuilds Excel.
    const formState = () => page.evaluate(() => Object.fromEntries(
      ["studentInput", "article", "clause", "item", "reason", "amount", "adjustReason", "adjustAmount"]
        .map(id => [id, document.getElementById(id).value])));
    const beforeRemoval = await formState();
    const importState = await $("importStatus").innerText();
    const removeRow = index => page.locator("#previewBody .delete-row").nth(index).click();
    const assertEmptyPreview = async () => {
      assert.equal(await page.locator("#previewBody tr").count(), 0);
      assert.equal(await $("exportBtn").getAttribute("href"), null);
      assert.equal(await $("exportBtn").getAttribute("aria-disabled"), "true");
      await text("selectedCount", /0 \/ 0/);
      for (const id of ["selectAll", "editSelectedBtn", "clearRecordsBtn"]) await disabled(id);
      assert.deepEqual(await formState(), beforeRemoval);
      assert.equal(await $("importStatus").innerText(), importState);
      assert.equal(await $("previewBtn").isEnabled(), true);
      assert.equal(await $("duplicateStatus").isVisible(), false);
    };
    assert.equal(await page.locator("#previewBody .edit-row").count(), 3);
    assert.equal(await page.locator("#previewBody .delete-row").count(), 3);
    await page.locator("#previewBody .row-select").nth(0).check();
    await page.locator("#previewBody .row-select").nth(2).check();
    await removeRow(1);
    await text("resultStatus", /已刪除 1 筆.*共 2 筆/);
    await text("selectedCount", /2 \/ 2/);
    assert.equal(await $("selectAll").isChecked(), true);
    assert.equal(await cell(1, "school").innerText(), students[2].school);
    await checkDownload({count: 2, rows: [firstAdjusted, third]});
    await page.locator("#previewBody .row-select").nth(0).uncheck();
    await $("editSelectedBtn").click();
    await $("editReason").fill("刪除其他列後，仍編輯原本勾選的學生。");
    await $("saveEditBtn").click();
    third.reason = "刪除其他列後，仍編輯原本勾選的學生。";
    assert.equal(await cell(0, "reason").innerText(), firstAdjusted.reason);
    await checkDownload({count: 2, rows: [firstAdjusted, third]});
    await removeRow(1);
    await text("selectedCount", /0 \/ 1/);
    await disabled("editSelectedBtn");
    await checkDownload({count: 1, rows: [firstAdjusted]});
    await $("clearRecordsBtn").click();
    await text("resultStatus", /已清空預覽紀錄/);
    await assertEmptyPreview();

    // Clearing all selected records keeps the imports and common form reusable.
    await $("previewBtn").click();
    const originalRecords = students.slice(0, 3).map(student => ({school: student.school, name: student.name,
      ...withoutItem, amount: baseAmount, adjustReason: "", adjustAmount: ""}));
    await checkDownload({count: 3, rows: originalRecords});
    await $("selectAll").check();
    await $("clearRecordsBtn").click();
    await assertEmptyPreview();
    await $("previewBtn").click();
    await page.locator("#previewBody .row-select").nth(1).check();
    await removeRow(0);
    await text("selectedCount", /1 \/ 2/);
    assert.equal(await page.locator("#previewBody .row-select").nth(0).isChecked(), true);
    assert.equal(await page.locator("#previewBody .row-select").nth(1).isChecked(), false);
    await checkDownload({count: 2, rows: originalRecords.slice(1)});
    await editRow(0);
    assert.equal(await $("editReason").inputValue(), withoutItem.reason);
    await $("cancelEditBtn").click();
    await removeRow(1);
    await text("selectedCount", /1 \/ 1/);
    await checkDownload({count: 1, rows: [originalRecords[1]]});
    await removeRow(0);
    await text("resultStatus", /已刪除最後一筆/);
    await assertEmptyPreview();
    await $("studentInput").fill(students[0].female);
    await text("selectedCount", /0 \/ 0/);
    await disabled("selectAll");

    // Same-file retries, invalid files, opposite import order, and state reset.
    await load("rosterFile", rosterPath, "rosterStatus");
    await $("rosterFile").setInputFiles({name: "empty.xlsx", mimeType: "application/octet-stream", buffer: Buffer.alloc(0)});
    await page.waitForFunction(() => document.getElementById("rosterStatus").classList.contains("error"));
    await disabled("previewBtn");
    await load("rosterFile", rosterPath, "rosterStatus");
    await $("templateFile").setInputFiles(rosterPath);
    await page.waitForFunction(() => document.getElementById("templateStatus").classList.contains("error"));
    await disabled("article");
    await disabled("previewBtn");
    await load("templateFile", templatePath, "templateStatus");
    await $("clearBtn").click();
    assert.equal(await $("studentInput").inputValue(), "");
    await disabled("clause");
    await disabled("item");
    assert.equal(await $("exportBtn").getAttribute("href"), null);
    await page.reload();
    await page.waitForFunction(() => document.documentElement.dataset.appReady === "true");
    await load("templateFile", templatePath, "templateStatus");
    await load("rosterFile", rosterPath, "rosterStatus");
    await $("studentInput").fill(students[0].female);
    await choose(withoutItem.article, withoutItem.clause);
    await $("previewBtn").click();
    await text("resultStatus", /已新增 1 筆/);
    const firstMain = {school: students[0].school, name: students[0].name, ...withoutItem,
      amount: template.defaultAmounts[withoutItem.article], adjustReason: "", adjustAmount: ""};
    const mainBeforeChange = await $("exportBtn").getAttribute("href");
    const mainReason = "共同獎懲事由帶入後自行修改。\n這一行也應保留。";
    await $("reason").fill(mainReason);
    assert.equal(await $("exportBtn").getAttribute("href"), mainBeforeChange);
    await $("adjustReason").selectOption(template.adjustmentOptions[0]);
    await $("adjustAmount").selectOption(template.amountOptions[0]);
    assert.equal(await $("reason").inputValue(), mainReason);
    await $("previewBtn").click();
    const secondMain = {...firstMain, reason: mainReason,
      adjustReason: template.adjustmentOptions[0], adjustAmount: template.amountOptions[0]};
    await checkDownload({count: 2, rows: [firstMain, secondMain]});
    const beforeBlank = await $("exportBtn").getAttribute("href");
    await $("reason").fill("  ");
    await $("previewBtn").click();
    await text("resultStatus", /填寫獎懲事由/);
    assert.equal(await $("exportBtn").getAttribute("href"), beforeBlank);
    assert.equal(await page.locator("#previewBody tr").count(), 2);
    await choose(withItem.article, withItem.clause, withItem.item);
    assert.equal(await $("reason").inputValue(), withItem.reason);
    await $("previewBtn").click();
    const thirdMain = {...firstMain, ...withItem, amount: template.defaultAmounts[withItem.article]};
    await checkDownload({count: 3, rows: [firstMain, secondMain, thirdMain]});
    await text("duplicateStatus", /共 3 筆，重複 2 次/);
    const beforeClearInput = await $("exportBtn").getAttribute("href");
    await $("clearBtn").click();
    assert.equal(await $("studentInput").inputValue(), "");
    assert.equal(await $("exportBtn").getAttribute("href"), beforeClearInput);
    assert.equal(await page.locator("#previewBody tr").count(), 3);
    await $("previewBtn").click();
    await text("resultStatus", /請輸入學生/);
    assert.equal(await $("exportBtn").getAttribute("href"), beforeClearInput);
    await checkDownload({count: 3, rows: [firstMain, secondMain, thirdMain]});
    await $("clearRecordsBtn").click();
    assert.equal(await $("duplicateStatus").isVisible(), false);
    assert.equal(await $("exportBtn").getAttribute("href"), null);
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
    assert.deepEqual(errors, [], "Browser Console must be clean");
    assert.deepEqual(nonLocalRequests, [], "No external CDN or data requests");
    assert.deepEqual(writes, [], "No student data may be uploaded");
    const brokenContext = await browser.newContext();
    try {
      await brokenContext.route("**/vendor/xlsx.full.min.js*", route => route.abort());
      const brokenPage = await brokenContext.newPage();
      await brokenPage.goto(baseURL);
      assert.match(await brokenPage.locator("#startupStatus").innerText(), /Excel 元件未完整載入/);
      assert.equal(await brokenPage.locator("#previewBtn").isDisabled(), true);
    } finally { await brokenContext.close(); }
    console.log(JSON.stringify({engine, mobile: "iPhone 13", rosterCount: students.length, rulesCheckedInDOM: template.rules.length, editorRulesChecked: template.rules.length, exportedDownloads: downloadsChecked, fixedBaseAmountsInEveryExport: true, manualAmountEditingRemoved: true, batchesAppendWithEditsPreserved: true, repeatsKeptInExports: true, duplicateCountsInRed: true, failedBatchPreservesExport: true, clearInputPreservesRecords: true, customReasonsPreserved: true, singleAndBatchEditsSynced: true, deletionsSynced: true, selectionPreservedAfterDeletion: true, clearKeepsImportsAndForm: true, partialMatchExport: true, consoleErrors: errors.length, externalRequests: nonLocalRequests.length, uploadRequests: writes.length, missingLibraryHandled: true, result: "PASS"}));
  } finally { await browser.close(); if (server) await new Promise(resolve => server.close(resolve)); }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
