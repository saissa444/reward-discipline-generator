/* Shared parsing and validation. No files or student data are sent to a server. */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.RewardCore = factory();
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  const fields = ["school", "name", "article", "clause", "item", "reason", "amount", "adjustReason", "adjustAmount"];
  const clean = value => String(value == null ? "" : value).trim();
  const normalize = value => clean(value).normalize("NFKC").replace(/\s+/g, "").toLowerCase();
  const identifier = value => normalize(value).replace(/^'(?=\d)/, "").replace(/\.0+$/, "");
  const unique = values => Array.from(new Set(values));
  const rows = (xlsx, sheet) => xlsx.utils.sheet_to_json(sheet, {header: 1, defval: "", raw: false});

  function column(headers, names) {
    const exact = headers.findIndex(header => names.some(name => normalize(header) === normalize(name)));
    return exact >= 0 ? exact : headers.findIndex(header => names.some(name => normalize(header).includes(normalize(name))));
  }

  function extractRoster(workbook, xlsx) {
    const preferred = workbook.SheetNames.filter(name => normalize(name) === "姓名查詢");
    const names = preferred.concat(workbook.SheetNames.filter(name => !preferred.includes(name)));
    let best = null;
    for (const sheetName of names) {
      const data = rows(xlsx, workbook.Sheets[sheetName]);
      for (let index = 0; index < Math.min(data.length, 30); index++) {
        const headers = data[index];
        const cols = {
          name: column(headers, ["姓名"]),
          school: column(headers, ["學籍號碼", "學籍號", "學號"]),
          female: column(headers, ["女生隊號"]),
          team: column(headers, ["建制隊號"])
        };
        if (cols.name < 0 || cols.school < 0) continue;
        const students = new Map();
        for (const row of data.slice(index + 1)) {
          const student = {
            name: clean(row[cols.name]), school: identifier(row[cols.school]),
            female: cols.female < 0 ? "" : identifier(row[cols.female]),
            team: cols.team < 0 ? "" : identifier(row[cols.team])
          };
          if (!student.name || !/^\d{10}$/.test(student.school)) continue;
          if (/^\d{1,3}$/.test(student.female)) student.female = student.female.padStart(3, "0");
          const previous = students.get(student.school);
          if (previous && ["name", "female", "team"].some(key => normalize(previous[key]) !== normalize(student[key]))) {
            throw new Error("同一學籍號碼有不同名冊資料，請先確認 Excel。");
          }
          students.set(student.school, student);
        }
        const result = {students: Array.from(students.values()), sheetName};
        if (result.students.length && preferred.includes(sheetName)) return result;
        if (!best || result.students.length > best.students.length) best = result;
        break; // Use the first complete header block, not a second lookup table beside it.
      }
    }
    if (!best || !best.students.length) {
      throw new Error("無法辨識名冊：請確認「姓名查詢」包含姓名及 10 位數學籍號碼。");
    }
    return best;
  }

  function headerField(value) {
    const header = normalize(value);
    const mapping = {
      "學號": "school", "學籍號碼": "school", "姓名": "name",
      "獎懲依據(條)": "article", "獎懲依據(款)": "clause", "獎懲依據(目)": "item",
      "獎懲事由": "reason", "額度": "amount", "加重減輕額度": "adjustAmount"
    };
    return header.startsWith("加重減輕事由") ? "adjustReason" : mapping[header];
  }

  function extractTemplate(workbook, xlsx) {
    const ruleSheetName = workbook.SheetNames.find(name => normalize(name) === "獎懲項目");
    if (!ruleSheetName) throw new Error("找不到「獎懲項目」工作表。");
    const rules = [], amounts = [], adjustments = [], defaults = {};
    for (const row of rows(xlsx, workbook.Sheets[ruleSheetName])) {
      const [article, clause, item, , reason, , key, amount, , adjustment] = Array.from({length: 10}, (_, i) => clean(row[i]));
      if (/^第\d+條$/.test(key) && amount) defaults[key] = amount;
      if (/^\d+$/.test(key) && amount && amount !== "以下空白") amounts.push(amount);
      if (adjustment && adjustment !== "加重減輕事由") adjustments.push(adjustment);
      if (/^第\d+條$/.test(article) && /^第\d+款$/.test(clause) && reason) {
        rules.push({article, clause, item, reason});
      }
    }
    if (!rules.length) throw new Error("「獎懲項目」未找到有效的條款與事由。");
    let layout = null;
    for (const sheetName of workbook.SheetNames.filter(name => name !== ruleSheetName)) {
      for (const row of rows(xlsx, workbook.Sheets[sheetName]).slice(0, 30)) {
        const mapped = row.map(headerField);
        if (fields.every(field => mapped.includes(field))) {
          const last = mapped.reduce((end, field, index) => field ? index : end, 0);
          layout = {sheetName, headers: row.slice(0, last + 1).map(clean), fields: mapped.slice(0, last + 1)};
          break;
        }
      }
      if (layout) break;
    }
    if (!layout) throw new Error("找不到範本的九個獎懲匯入欄位，請確認所選檔案。");
    return Object.assign(layout, {
      rules, defaultAmounts: defaults,
      amountOptions: unique(amounts.concat(Object.values(defaults))),
      adjustmentOptions: unique(adjustments)
    });
  }

  function parseTokens(value) {
    const tokens = [];
    for (const token of clean(value).normalize("NFKC").split(/[\s,、;]+/).filter(Boolean)) {
      const range = token.match(/^(\d{1,3})[-–~～](\d{1,3})$/);
      if (!range) { tokens.push(token); continue; }
      const first = Number(range[1]), last = Number(range[2]);
      if (first < 1 || last < first || last > 999) throw new Error("女生隊號範圍不正確：" + token);
      for (let n = first; n <= last; n++) tokens.push(String(n).padStart(3, "0"));
    }
    return tokens;
  }

  function matchStudent(token, roster) {
    const value = identifier(token);
    const candidates = roster.filter(student =>
      normalize(student.name) === normalize(token) || student.school === value ||
      (student.team && identifier(student.team) === value) ||
      (/^\d{1,3}$/.test(value) && /^\d{1,3}$/.test(student.female) &&
        student.female.padStart(3, "0") === value.padStart(3, "0"))
    );
    const matches = Array.from(new Map(candidates.map(student => [student.school, student])).values());
    if (!matches.length) return {error: "查無此學生"};
    if (matches.length > 1) return {error: "符合多位學生，請改用女生隊號或建制隊號"};
    return {student: matches[0]};
  }

  function resolveStudents(value, roster) {
    const tokens = parseTokens(value), matches = new Map(), errors = [];
    for (const token of tokens) {
      const match = matchStudent(token, roster);
      if (match.error) { errors.push({token, message: match.error}); continue; }
      const key = match.student.school;
      const previous = matches.get(key);
      if (previous) previous.count++;
      else matches.set(key, {student: match.student, count: 1});
    }
    return {tokens, matches: Array.from(matches.values()), errors};
  }

  function buildRecords(value, roster, template, selection) {
    if (!roster.length) throw new Error("請先匯入學生名冊。");
    if (!template) throw new Error("請先匯入獎懲範本。");
    const resolved = resolveStudents(value, roster);
    if (!resolved.tokens.length) throw new Error("請輸入學生。");
    if (resolved.errors.length) throw new Error("以下學生無法確認，請修正後再匯出：\n" + resolved.errors.map(error => error.token + "：" + error.message).join("\n"));
    if (!selection.article || !selection.clause) throw new Error("請依序選擇獎懲依據（條）、（款）。");
    const relevant = template.rules.filter(rule => rule.article === selection.article && rule.clause === selection.clause);
    if (relevant.some(rule => rule.item) && !selection.item) throw new Error("請選擇獎懲依據（目）。");
    const rule = relevant.find(row => row.item === selection.item);
    if (!rule) throw new Error("尚未帶入有效的獎懲事由。");
    if (!template.amountOptions.includes(selection.amount)) throw new Error("請選擇有效的額度。");
    if (selection.adjustReason && !template.adjustmentOptions.includes(selection.adjustReason)) throw new Error("請選擇有效的加重減輕事由。");
    if (selection.adjustAmount && !selection.adjustReason) throw new Error("選擇加重減輕額度前，請先選擇加重減輕事由。");
    if (selection.adjustAmount && !template.amountOptions.includes(selection.adjustAmount)) throw new Error("請選擇有效的加重減輕額度。");
    const records = resolved.matches.map(({student}) => ({
      school: student.school, name: student.name, article: rule.article, clause: rule.clause,
      item: rule.item, reason: rule.reason, amount: selection.amount,
      adjustReason: selection.adjustReason || "", adjustAmount: selection.adjustAmount || ""
    }));
    return {records, resolved};
  }

  function exportWorkbook(records, template, xlsx) {
    if (!records.length || !template) throw new Error("請先產生紀錄預覽。");
    const data = [template.headers, ...records.map(record => template.fields.map(field => record[field] || ""))];
    const sheet = xlsx.utils.aoa_to_sheet(data);
    sheet["!cols"] = template.fields.map(field => ({wch: field === "reason" ? 55 : field === "adjustReason" ? 30 : 18}));
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, sheet, template.sheetName);
    return workbook;
  }

  return {fields, clean, normalize, extractRoster, extractTemplate, parseTokens, matchStudent, resolveStudents, buildRecords, exportWorkbook};
});
