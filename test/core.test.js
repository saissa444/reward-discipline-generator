"use strict";
// All names and identifiers in this file are fictional. Never commit real rosters.
const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");
const xlsx = require("../vendor/xlsx.full.min.js");
const core = require("../app-core.js");

const headers = ["學號", "姓名", "獎懲依據(條)", "獎懲依據(款)", "獎懲依據(目)", "獎懲事由", "額度", "加重減輕事由 (若無可不選擇)", "加重減輕額度"];
const rosterData = [
  ["姓名", "建制隊號", "女生隊號", "隊別", "學籍號碼"],
  ["測試甲", 991001, 2, 1, 1999100001],
  ["測試乙", 991002, 11, 1, 1999100002],
  ["同名測試", 991003, 12, 1, 1999100003],
  ["同名測試", 991004, "", 1, 1999100004]
];
const ruleData = [
  ["", "", "", "", "", "", "第19條", "申誡一次", "", "加重減輕事由"],
  ["第19條", "第1款", "", "", "測試事由甲。", "", 1, "申誡一次", "", "行為之動機"],
  ["第20條", "第17款", "第1目", "", "測試事由乙。", "", "第20條", "小過一次"],
  ["第20條", "第17款", "第2目", "", "測試事由丙。", "", 2, "小過一次"],
  ["第24條", "第1款", "", "", "測試事由丁。", "", "第24條", "開除學籍"],
  ["", "", "", "", "", "", 13, "以下空白"]
];
function workbook(sheets) {
  const result = xlsx.utils.book_new();
  for (const [name, data] of Object.entries(sheets)) xlsx.utils.book_append_sheet(result, xlsx.utils.aoa_to_sheet(data), name);
  return result;
}
const roster = core.extractRoster(workbook({姓名查詢: rosterData}), xlsx).students;
const template = core.extractTemplate(workbook({工作表1: [headers], 獎懲項目: ruleData}), xlsx);
const selection = {article: "第19條", clause: "第1款", item: "", amount: "申誡一次"};

test("every inline script parses, every local script exists, and no remote runtime CDN is used", () => {
  const html = fs.readFileSync(path.join(__dirname, "../index.html"), "utf8");
  let scripts = 0;
  for (const match of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/g)) {
    new vm.Script(match[2]); scripts++;
    const src = match[1].match(/src="([^"?]+)(?:[^"]*)"/);
    if (src) {
      assert(!/^https?:/.test(src[1]));
      assert(fs.existsSync(path.join(__dirname, "..", src[1])));
    }
  }
  assert.equal(scripts, 4);
});

test("prefer 姓名查詢 even when a larger sheet comes first", () => {
  const alternate = rosterData.concat([["測試戊", 991005, 13, 1, 1999100005]]);
  const result = core.extractRoster(workbook({其他: alternate, 姓名查詢: rosterData}), xlsx);
  assert.equal(result.sheetName, "姓名查詢");
  assert.equal(result.students.length, 4);
  assert.equal(result.students[0].female, "002");
});

test("ignore adjacent lookup block and invalid school numbers; fallback when preferred sheet is absent", () => {
  const data = rosterData.map(row => row.concat(["", ...row]));
  data.push(["非資料列", 991999, 99, 1, "錯誤"]);
  const result = core.extractRoster(workbook({其他: data}), xlsx);
  assert.equal(result.students.length, 4);
});

test("mixed female numbers, team numbers, names, school IDs and full-width input share one matcher", () => {
  const result = core.resolveStudents("２、991001 測試甲\n1999100001；０１１", roster);
  assert.equal(result.errors.length, 0);
  assert.equal(result.matches.length, 2);
  assert.equal(result.matches[0].count, 4);
  assert.equal(result.matches[1].student.female, "011");
});

test("ranges expand, reverse ranges fail, blank female number never matches 000", () => {
  assert.deepEqual(core.parseTokens("011-012"), ["011", "012"]);
  assert.throws(() => core.parseTokens("012-011"), /範圍/);
  assert(core.matchStudent("000", roster).error);
});

test("unknown and ambiguous students are errors, and prevent preview/export", () => {
  assert.match(core.matchStudent("同名測試", roster).error, /多位/);
  assert.throws(() => core.buildRecords("測試甲,無此人", roster, template, selection), /無法確認/);
});

test("template keeps original sheet name, exact nine headers, defaults and optional adjustments", () => {
  assert.equal(template.sheetName, "工作表1");
  assert.deepEqual(template.headers, headers);
  assert.equal(template.rules.length, 4);
  assert(template.amountOptions.includes("開除學籍"));
  assert(!template.amountOptions.includes("以下空白"));
  assert.deepEqual(template.adjustmentOptions, ["行為之動機"]);
  assert.throws(() => core.extractTemplate(workbook({獎懲項目: ruleData}), xlsx), /九個/);
});

test("required item, rule existence and valid amount are checked against source data", () => {
  assert.throws(() => core.buildRecords("2", roster, template, {...selection, article: "第20條", clause: "第17款"}), /選擇.*目/);
  assert.throws(() => core.buildRecords("2", roster, template, {...selection, clause: "第999款"}), /有效/);
  assert.throws(() => core.buildRecords("2", roster, template, {...selection, amount: "無效"}), /額度/);
});

test("deduplicate by school number; reason always comes from source, not editable selection", () => {
  const result = core.buildRecords("2,991001,測試甲,11", roster, template, {...selection, reason: "偽造事由"});
  assert.equal(result.records.length, 2);
  assert.equal(result.resolved.matches[0].count, 3);
  assert.equal(result.records[0].reason, "測試事由甲。");
});

test("optional adjustments may be blank, but amount requires a selected adjustment reason", () => {
  assert.equal(core.buildRecords("2", roster, template, selection).records[0].adjustAmount, "");
  assert.throws(() => core.buildRecords("2", roster, template, {...selection, adjustAmount: "小過一次"}), /先選擇/);
  assert.equal(core.buildRecords("2", roster, template, {...selection, adjustReason: "行為之動機", adjustAmount: "小過一次"}).records[0].adjustAmount, "小過一次");
});

test("export round trip retains headers, text school IDs, item, reason and both adjustments", () => {
  const selected = {article: "第20條", clause: "第17款", item: "第2目", amount: "小過一次", adjustReason: "行為之動機", adjustAmount: "申誡一次"};
  const {records} = core.buildRecords("2,11", roster, template, selected);
  const bytes = xlsx.write(core.exportWorkbook(records, template, xlsx), {type: "buffer", bookType: "xlsx"});
  const result = xlsx.read(bytes);
  assert.deepEqual(result.SheetNames, ["工作表1"]);
  const data = xlsx.utils.sheet_to_json(result.Sheets.工作表1, {header: 1, defval: ""});
  assert.deepEqual(data[0], headers);
  assert.deepEqual(data[1], ["1999100001", "測試甲", "第20條", "第17款", "第2目", "測試事由丙。", "小過一次", "行為之動機", "申誡一次"]);
  assert.equal(data.length, 3);
});
