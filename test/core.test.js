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

test("unknown and ambiguous inputs are reported while other students can preview and export", () => {
  assert.match(core.matchStudent("同名測試", roster).error, /多位/);
  const result = core.buildRecords("測試甲,無此人,同名測試,11", roster, template, selection);
  assert.equal(result.records.length, 2);
  assert.deepEqual(result.records.map(record => record.name), ["測試甲", "測試乙"]);
  assert.equal(result.resolved.errors.length, 2);
  assert.deepEqual(result.resolved.errors.map(error => error.token), ["無此人", "同名測試"]);
  const workbook = xlsx.read(xlsx.write(core.exportWorkbook(result.records, template, xlsx), {type: "buffer", bookType: "xlsx"}));
  const data = xlsx.utils.sheet_to_json(workbook.Sheets.工作表1, {header: 1});
  assert.equal(data.length, 3);
  assert.deepEqual(data.slice(1).map(row => row[1]), ["測試甲", "測試乙"]);
});

test("no matched students prevents an empty export", () => {
  assert.throws(() => core.buildRecords("無此人,同名測試", roster, template, selection), /沒有可匯出/);
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

test("keep every repeated identifier in input order with a custom multiline reason", () => {
  const reason = "自行補充的獎懲事由。\n第二行說明。";
  const result = core.buildRecords("2,991001,測試甲,11", roster, template, {...selection, reason});
  assert.equal(result.records.length, 4);
  assert.equal(result.resolved.matches[0].count, 3);
  assert.deepEqual(result.records.map(record => record.school), [roster[0].school, roster[0].school, roster[0].school, roster[1].school]);
  assert(result.records.every(record => record.reason === reason));
  const edited = core.editRecords(result.records, [1], template, {reason: "只編輯這一筆重複紀錄。"});
  assert.equal(edited[0].reason, reason);
  assert.equal(edited[2].reason, reason);
  assert.equal(edited[1].reason, "只編輯這一筆重複紀錄。");
  const exported = xlsx.read(xlsx.write(core.exportWorkbook(edited, template, xlsx), {type: "buffer", bookType: "xlsx"}));
  const data = xlsx.utils.sheet_to_json(exported.Sheets.工作表1, {header: 1});
  assert.deepEqual(data.slice(1).map(row => row[0]), edited.map(record => record.school));
  assert.deepEqual(data.slice(1).map(row => row[5]), edited.map(record => record.reason));
});

test("overlapping ranges retain repeated occurrences and skip only unresolved input", () => {
  const result = core.buildRecords("11-12,無此人,11,測試甲,11", roster, template, selection);
  assert.deepEqual(result.records.map(record => record.school), [roster[1].school, roster[2].school, roster[1].school, roster[0].school, roster[1].school]);
  assert.equal(result.resolved.errors.length, 1);
});

test("duplicate counts use school identity and update when records are removed", () => {
  const {records} = core.buildRecords("2,11,2,12,2", roster, template, selection);
  const before = JSON.stringify(records);
  assert.deepEqual(core.duplicateStudents(records), [{school: roster[0].school, name: roster[0].name, count: 3}]);
  assert.equal(JSON.stringify(records), before);
  assert.equal(core.duplicateStudents(records.slice(0, -1))[0].count, 2);
  assert.deepEqual(core.duplicateStudents([]), []);
  assert.deepEqual(core.duplicateStudents([{school: "1999100003", name: "同名測試"}, {school: "1999100004", name: "同名測試"}]), []);
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

test("single record editing changes only its reward fields and updates the exported values", () => {
  const {records} = core.buildRecords("2,11,12", roster, template, selection);
  const snapshot = JSON.stringify(records);
  const edited = core.editRecords(records, [0], template, {amount: "小過一次"});
  assert.equal(edited[0].amount, "小過一次");
  assert.equal(edited[1], records[1]);
  assert.equal(edited[2], records[2]);
  assert.equal(JSON.stringify(records), snapshot);
  assert.deepEqual(edited.map(record => [record.school, record.name]), records.map(record => [record.school, record.name]));
  const workbook = xlsx.read(xlsx.write(core.exportWorkbook(edited, template, xlsx), {type: "buffer", bookType: "xlsx"}));
  const data = xlsx.utils.sheet_to_json(workbook.Sheets.工作表1, {header: 1});
  assert.equal(data[1][6], "小過一次");
  assert.equal(data[2][6], "申誡一次");
});

test("batch editing applies only the chosen fields to the selected rows", () => {
  const {records} = core.buildRecords("2,11,12", roster, template, selection);
  const edited = core.editRecords(records, [0, 2], template, {article: "第20條", clause: "第17款", item: "第2目"});
  for (const index of [0, 2]) {
    assert.equal(edited[index].reason, "測試事由丙。");
    assert.equal(edited[index].amount, records[index].amount);
    assert.equal(edited[index].school, records[index].school);
  }
  assert.equal(edited[1], records[1]);
  assert.equal(records[0].reason, "測試事由甲。");
});

test("edited adjustments can be set and cleared without changing other reward fields", () => {
  const {records} = core.buildRecords("2,11", roster, template, selection);
  const adjusted = core.editRecords(records, [0, 1], template, {adjustReason: "行為之動機", adjustAmount: "小過一次"});
  assert(adjusted.every(record => record.adjustAmount === "小過一次"));
  const cleared = core.editRecords(adjusted, [0], template, {adjustReason: "", adjustAmount: ""});
  assert.equal(cleared[0].adjustAmount, "");
  assert.equal(cleared[0].adjustReason, "");
  assert.equal(cleared[1].adjustAmount, "小過一次");
});

test("invalid edits reject the whole batch and preserve the previous records", () => {
  const {records} = core.buildRecords("2,11", roster, template, selection);
  records[0].adjustReason = "行為之動機";
  const snapshot = JSON.stringify(records);
  assert.throws(() => core.editRecords(records, [0, 1], template, {adjustAmount: "小過一次"}), /先選擇/);
  assert.equal(JSON.stringify(records), snapshot);
  assert.throws(() => core.editRecords(records, [0, 1], template, {article: "第20條", clause: "第17款", item: ""}), /選擇.*目/);
  assert.equal(JSON.stringify(records), snapshot);
});

test("editing requires a valid selection, a chosen field and unchanged roster identity", () => {
  const {records} = core.buildRecords("2", roster, template, selection);
  assert.throws(() => core.editRecords(records, [], template, {amount: "小過一次"}), /先選取/);
  assert.throws(() => core.editRecords(records, [9], template, {amount: "小過一次"}), /重新選取/);
  assert.throws(() => core.editRecords(records, [0], template, {}), /至少一組/);
  assert.throws(() => core.editRecords(records, [0], template, {name: "其他學生"}), /名冊/);
});

test("amount and adjustment edits preserve each student's custom reason in the export", () => {
  const {records} = core.buildRecords("2,11", roster, template, {...selection, reason: "自訂事由甲。\n第二行。"});
  const distinct = core.editRecords(records, [1], template, {reason: "自訂事由乙。"});
  const edited = core.editRecords(distinct, [0, 1], template, {amount: "小過一次", adjustReason: "行為之動機", adjustAmount: "申誡一次"});
  assert.deepEqual(edited.map(record => record.reason), ["自訂事由甲。\n第二行。", "自訂事由乙。"]);
  const workbook = xlsx.read(xlsx.write(core.exportWorkbook(edited, template, xlsx), {type: "buffer", bookType: "xlsx"}));
  const data = xlsx.utils.sheet_to_json(workbook.Sheets.工作表1, {header: 1});
  assert.deepEqual(data.slice(1).map(row => row[5]), edited.map(record => record.reason));
});

test("reason-only batch edits preserve different clauses and unselected rows", () => {
  const {records} = core.buildRecords("2,11,12", roster, template, selection);
  const mixed = core.editRecords(records, [2], template, {article: "第20條", clause: "第17款", item: "第1目"});
  const edited = core.editRecords(mixed, [0, 2], template, {reason: "共同自訂事由。"});
  for (const index of [0, 2]) {
    assert.equal(edited[index].reason, "共同自訂事由。");
    for (const key of ["school", "name", "article", "clause", "item", "amount"]) assert.equal(edited[index][key], mixed[index][key]);
  }
  assert.equal(edited[1], mixed[1]);
});

test("changing a rule reloads its default unless an explicit custom reason is supplied", () => {
  const {records} = core.buildRecords("2", roster, template, {...selection, reason: "原本自訂內容。"});
  assert.equal(core.editRecords(records, [0], template, {article: "第19條", clause: "第1款", item: ""})[0].reason, "原本自訂內容。");
  const patch = {article: "第20條", clause: "第17款", item: "第2目"};
  assert.equal(core.editRecords(records, [0], template, patch)[0].reason, "測試事由丙。");
  assert.equal(core.editRecords(records, [0], template, {...patch, reason: "換條款後另外補充。"})[0].reason, "換條款後另外補充。");
});

test("an explicitly empty reason is rejected instead of silently using the template text", () => {
  assert.throws(() => core.buildRecords("2", roster, template, {...selection, reason: " \n "}), /填寫獎懲事由/);
  const {records} = core.buildRecords("2", roster, template, selection);
  assert.throws(() => core.editRecords(records, [0], template, {reason: ""}), /填寫獎懲事由/);
  assert.equal(records[0].reason, "測試事由甲。");
});
