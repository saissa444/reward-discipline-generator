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
      assert.equal(data[1][0], students[0].school);
      assert.equal(data[1][1], students[0].name);
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
    await text("studentCheckPanel", /重複 3 次/);
    assert.equal(await page.locator("#studentCheckPanel .error").count(), 1);

    // Walk every real rule through the actual cascading DOM controls.
    for (const rule of template.rules) {
      await choose(rule.article, rule.clause, rule.item);
      assert.equal(await $("reason").inputValue(), rule.reason);
      assert.equal(await $("amount").inputValue(), template.defaultAmounts[rule.article]);
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
    await text("resultStatus", /成功建立 2 筆/);
    assert.equal(await page.locator("#previewBody tr").count(), 2);
    assert.equal(await page.locator("#previewBody tr:first-child td[data-field]").count(), 9);
    await checkDownload({count: 2, fields: {...withItem, amount: template.defaultAmounts[withItem.article], adjustReason: template.adjustmentOptions[0], adjustAmount: template.amountOptions[0]}});

    // Any changed student or rule invalidates the old downloadable workbook.
    await $("studentInput").fill(students[0].female + ",不存在的測試學生");
    assert.equal(await $("exportBtn").getAttribute("href"), null);
    await $("previewBtn").click();
    await text("resultStatus", /成功建立 1 筆/);
    await text("resultStatus", /已略過 1 筆/);
    assert.equal(await page.locator("#studentCheckPanel .error").count(), 1);
    assert.equal(await $("resultStatus").getAttribute("class"), "status warning");
    assert.equal(await $("exportBtn").getAttribute("aria-disabled"), "false");
    await checkDownload({count: 1, fields: withItem});
    await $("studentInput").fill("不存在的測試學生");
    await $("previewBtn").click();
    await text("resultStatus", /沒有可匯出/);
    assert.equal(await $("exportBtn").getAttribute("aria-disabled"), "true");
    assert.equal(await $("exportBtn").getAttribute("href"), null);
    await $("studentInput").fill(students[0].female);
    const withoutItem = template.rules.find(rule => !rule.item);
    await choose(withoutItem.article, withoutItem.clause);
    await $("previewBtn").click();
    await checkDownload({count: 1, fields: {...withoutItem, adjustReason: "", adjustAmount: ""}});
    await $("amount").selectOption(template.amountOptions.find(amount => amount !== template.defaultAmounts[withoutItem.article]));
    assert.equal(await $("exportBtn").getAttribute("href"), null);

    // Single and selected-row edits must update the actual downloaded workbook.
    await $("studentInput").fill(students.slice(0, 3).map(student => student.female).concat("不存在的測試學生").join(","));
    await choose(withoutItem.article, withoutItem.clause);
    await $("previewBtn").click();
    const cell = (index, field) => page.locator("#previewBody tr").nth(index).locator('[data-field="' + field + '"]');
    const editRow = index => page.locator("#previewBody .edit-row").nth(index).click();
    const baseAmount = template.defaultAmounts[withoutItem.article];
    const singleAmount = template.amountOptions.find(amount => amount !== baseAmount);
    const batchAmount = template.amountOptions.find(amount => amount !== baseAmount && amount !== singleAmount);
    const initialURL = await $("exportBtn").getAttribute("href");
    await disabled("editSelectedBtn");
    await editRow(0);
    for (const rule of template.rules) {
      await $("editArticle").selectOption(rule.article);
      await $("editClause").selectOption(rule.clause);
      if (rule.item) await $("editItem").selectOption(rule.item);
      assert.equal(await $("editReason").inputValue(), rule.reason);
      assert.equal(await $("editAmount").inputValue(), template.defaultAmounts[rule.article]);
    }
    await $("cancelEditBtn").click();
    assert.equal(await $("exportBtn").getAttribute("href"), initialURL);
    await editRow(0);
    assert.equal(await $("editArticle").inputValue(), withoutItem.article);
    assert.equal(await $("editAmount").inputValue(), baseAmount);
    await $("editAmount").selectOption(singleAmount);
    await $("editAdjustReason").selectOption(template.adjustmentOptions[0]);
    await $("editAdjustAmount").selectOption(template.amountOptions[0]);
    await $("saveEditBtn").click();
    assert.equal(await $("recordEditor").isVisible(), false);
    await text("resultStatus", /已略過 1 筆/);
    assert.equal(await $("resultStatus").getAttribute("class"), "status warning");
    assert.equal(await cell(0, "amount").innerText(), singleAmount);
    assert.equal(await cell(1, "amount").innerText(), baseAmount);
    assert.equal(await cell(2, "amount").innerText(), baseAmount);
    const firstAdjusted = {school: students[0].school, name: students[0].name, ...withoutItem, amount: singleAmount, adjustReason: template.adjustmentOptions[0], adjustAmount: template.amountOptions[0]};
    const untouched = {school: students[1].school, name: students[1].name, ...withoutItem, amount: baseAmount, adjustReason: "", adjustAmount: ""};
    const third = {school: students[2].school, name: students[2].name, ...withoutItem, amount: baseAmount, adjustReason: "", adjustAmount: ""};
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    const savedURL = await $("exportBtn").getAttribute("href");
    await editRow(0);
    await $("editAmount").selectOption(batchAmount);
    await $("cancelEditBtn").click();
    assert.equal(await $("exportBtn").getAttribute("href"), savedURL);
    assert.equal(await cell(0, "amount").innerText(), singleAmount);

    await page.locator("#previewBody .row-select").nth(0).check();
    await page.locator("#previewBody .row-select").nth(2).check();
    await text("selectedCount", /2 \/ 3/);
    assert.equal(await $("selectAll").evaluate(element => element.indeterminate), true);
    await $("editSelectedBtn").click();
    for (const id of ["editApplyRule", "editApplyAmount", "editApplyAdjust"]) assert.equal(await $(id).isChecked(), false);
    await $("saveEditBtn").click();
    await text("editStatus", /至少一組/);
    assert.equal(await $("recordEditor").isVisible(), true);
    await $("editApplyAmount").check();
    await $("editAmount").selectOption(batchAmount);
    await $("saveEditBtn").click();
    await text("resultStatus", /已更新 2 筆/);
    await text("resultStatus", /已略過 1 筆/);
    firstAdjusted.amount = batchAmount;
    third.amount = batchAmount;
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
    Object.assign(firstAdjusted, withItem);
    Object.assign(third, withItem);
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    await editRow(0);
    await $("editAdjustReason").selectOption("");
    await disabled("editAdjustAmount");
    await $("saveEditBtn").click();
    firstAdjusted.adjustReason = "";
    firstAdjusted.adjustAmount = "";
    await checkDownload({count: 3, rows: [firstAdjusted, untouched, third]});
    await $("selectAll").check();
    await text("selectedCount", /3 \/ 3/);
    await $("selectAll").uncheck();
    await text("selectedCount", /0 \/ 3/);
    await disabled("editSelectedBtn");
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
    await text("resultStatus", /成功建立 1 筆/);
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
    console.log(JSON.stringify({engine, mobile: "iPhone 13", rosterCount: students.length, rulesCheckedInDOM: template.rules.length, editorRulesChecked: template.rules.length, exportedDownloads: downloadsChecked, singleAndBatchEditsSynced: true, partialMatchExport: true, consoleErrors: errors.length, externalRequests: nonLocalRequests.length, uploadRequests: writes.length, missingLibraryHandled: true, result: "PASS"}));
  } finally { await browser.close(); if (server) await new Promise(resolve => server.close(resolve)); }
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
