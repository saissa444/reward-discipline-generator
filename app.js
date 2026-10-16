(function () {
  "use strict";
  const $ = id => document.getElementById(id);
  const core = window.RewardCore;
  const xlsx = window.XLSX;
  if (!core || !xlsx || !xlsx.read || !xlsx.utils) {
    window.reportStartupError("Excel 元件未完整載入，請重新整理網頁後再試。");
    return;
  }

  let roster = [], template = null, output = [], downloadURL = null;
  const selectedRows = new Set();
  let editingRows = [], previewSummary = {skipped: 0};
  const versions = {roster: 0, template: 0};
  const imports = {roster: "尚未匯入學生名冊", template: "尚未匯入獎懲範本"};
  const status = (id, message, kind) => {
    $(id).textContent = message;
    $(id).className = "status" + (kind ? " " + kind : "");
  };

  function fillOptions(id, values, placeholder) {
    const select = $(id);
    select.replaceChildren(new Option(placeholder, ""));
    values.forEach(value => select.add(new Option(value, value)));
  }

  function invalidatePreview(message) {
    output = [];
    selectedRows.clear();
    previewSummary = {skipped: 0};
    if ($("recordEditor").open) $("recordEditor").close();
    editingRows = [];
    $("previewBody").replaceChildren();
    updateDuplicates();
    updateSelection();
    $("exportBtn").removeAttribute("href");
    $("exportBtn").setAttribute("aria-disabled", "true");
    $("exportBtn").tabIndex = -1;
    if (downloadURL) URL.revokeObjectURL(downloadURL);
    downloadURL = null;
    status("resultStatus", message || "匯入資料已變更，請重新批次新增紀錄。");
  }

  function refreshImports() {
    status("importStatus", "學生名冊：" + (roster.length ? roster.length + " 位" : "尚未匯入") +
      "；獎懲範本：" + (template ? template.rules.length + " 筆項目" : "尚未匯入"),
      roster.length && template ? "ok" : "");
    $("previewBtn").disabled = !roster.length || !template;
  }

  function resetRuleControls() {
    fillOptions("article", template ? Array.from(new Set(template.rules.map(rule => rule.article))) : [], template ? "請選擇條" : "請先匯入範本");
    fillOptions("clause", [], "請先選擇條");
    fillOptions("item", [], "請先選擇款");
    $("amount").value = "";
    fillOptions("adjustReason", template ? template.adjustmentOptions : [], "無（選填）");
    fillOptions("adjustAmount", template ? template.amountOptions : [], "無（選填）");
    $("article").disabled = !template;
    ["clause", "item", "amount", "adjustReason", "adjustAmount"].forEach(id => { $(id).disabled = true; });
    $("reason").value = "";
    $("reason").disabled = true;
  }

  function syncRuleDetails() {
    const rule = template && template.rules.find(row => row.article === $("article").value &&
      row.clause === $("clause").value && row.item === $("item").value);
    $("reason").value = rule ? rule.reason : "";
    $("reason").disabled = !rule;
    $("amount").disabled = !rule;
    $("adjustReason").disabled = !rule;
    $("adjustAmount").disabled = !rule || !$("adjustReason").value;
    $("amount").value = rule ? template.defaultAmounts[rule.article] || "" : "";
    $("adjustReason").value = "";
    $("adjustAmount").value = "";
    $("adjustAmount").disabled = true;
  }

  function updateClause() {
    const article = $("article").value;
    fillOptions("clause", Array.from(new Set(template.rules.filter(rule => rule.article === article).map(rule => rule.clause))), "請選擇款");
    $("clause").disabled = !article;
    fillOptions("item", [], "請先選擇款");
    $("item").disabled = true;
    syncRuleDetails();
  }

  function updateItem() {
    const relevant = template.rules.filter(rule => rule.article === $("article").value && rule.clause === $("clause").value);
    const items = Array.from(new Set(relevant.map(rule => rule.item).filter(Boolean)));
    fillOptions("item", items, items.length ? "請選擇目" : relevant.length ? "無目（不需選擇）" : "請先選擇款");
    $("item").disabled = !items.length;
    syncRuleDetails();
  }

  function checkStudents() {
    const panel = $("studentCheckPanel");
    panel.replaceChildren();
    if (!core.clean($("studentInput").value)) { panel.textContent = "尚未輸入學生資料"; return; }
    if (!roster.length) { panel.textContent = "請先匯入學生名冊"; return; }
    try {
      const resolved = core.resolveStudents($("studentInput").value, roster);
      resolved.matches.forEach(({student, count}) => {
        const row = document.createElement("div");
        row.className = "student-row" + (count > 1 ? " error" : "");
        row.textContent = student.name + (count > 1 ? "｜本批共 " + count + " 筆（重複 " + (count - 1) + " 次，全部新增）" : "｜辨識成功") +
          "\n女生隊號：" + (student.female || "—") + "　建制隊號：" + (student.team || "—") + "　學籍號碼：" + student.school;
        panel.appendChild(row);
      });
      resolved.errors.forEach(error => {
        const row = document.createElement("div");
        row.className = "student-row error";
        row.textContent = error.token + "：" + error.message;
        panel.appendChild(row);
      });
      const summary = document.createElement("p");
      summary.textContent = "輸入 " + resolved.tokens.length + " 次；辨識 " + resolved.matches.length + " 位；重複 " +
        resolved.matches.filter(match => match.count > 1).length + " 位；無法辨識 " + resolved.errors.length + " 筆。";
      panel.appendChild(summary);
    } catch (error) {
      panel.textContent = error.message;
      panel.classList.add("error");
      return;
    }
    panel.classList.remove("error");
  }

  function readExcel(file) {
    return new Promise((resolve, reject) => {
      if (!file.size) { reject(new Error("檔案是空的；若來自 iCloud，請先下載至「檔案」後重選。")); return; }
      const reader = new FileReader();
      // FileReader + ArrayBuffer also works on Safari without Blob.arrayBuffer().
      const timer = setTimeout(() => {
        reader.onabort = null;
        reader.abort();
        reject(new Error("讀取逾時，請先將 Excel 下載至 iPhone「檔案」再重新選取。"));
      }, 30000);
      reader.onload = () => {
        clearTimeout(timer);
        try {
          const bytes = new Uint8Array(reader.result);
          const workbook = xlsx.read(bytes, {type: "array"});
          if (!workbook.SheetNames.length) throw new Error("沒有可讀取的工作表。");
          resolve(workbook);
        } catch (error) { reject(new Error("Excel 解析失敗：" + error.message)); }
      };
      reader.onerror = () => { clearTimeout(timer); reject(new Error("無法讀取檔案，請先下載至本機後重選。")); };
      reader.onabort = () => { clearTimeout(timer); reject(new Error("檔案讀取已中止，請重新選取。")); };
      try { reader.readAsArrayBuffer(file); }
      catch (error) { clearTimeout(timer); reject(new Error("無法開啟檔案：" + error.message)); }
    });
  }

  async function importFile(kind, input) {
    const file = input.files && input.files[0];
    if (!file) return;
    const version = ++versions[kind];
    const statusId = kind === "roster" ? "rosterStatus" : "templateStatus";
    if (kind === "roster") roster = [];
    else { template = null; resetRuleControls(); }
    invalidatePreview();
    refreshImports();
    checkStudents();
    status(statusId, "正在讀取：" + file.name + "…");
    try {
      const workbook = await readExcel(file);
      if (versions[kind] !== version) return;
      if (kind === "roster") {
        const result = core.extractRoster(workbook, xlsx);
        roster = result.students;
        imports.roster = file.name + "｜「" + result.sheetName + "」：" + roster.length + " 位學生";
      } else {
        template = core.extractTemplate(workbook, xlsx);
        resetRuleControls();
        imports.template = file.name + "｜獎懲項目：" + template.rules.length + " 筆";
      }
      status(statusId, "匯入成功：" + imports[kind], "ok");
    } catch (error) {
      if (versions[kind] !== version) return;
      status(statusId, "匯入失敗：" + error.message, "error");
    } finally {
      if (versions[kind] === version) {
        input.value = ""; // Safari must dispatch change when retrying the same file.
        refreshImports();
        checkStudents();
      }
    }
  }

  function selection() {
    return Object.fromEntries(["article", "clause", "item", "reason", "adjustReason", "adjustAmount"].map(id => [id, $(id).value]));
  }

  function updateSelection() {
    $("selectedCount").textContent = "已選取 " + selectedRows.size + " / " + output.length + " 筆";
    $("editSelectedBtn").disabled = !selectedRows.size;
    $("clearRecordsBtn").disabled = !output.length;
    $("selectAll").disabled = !output.length;
    $("selectAll").checked = output.length > 0 && selectedRows.size === output.length;
    $("selectAll").indeterminate = selectedRows.size > 0 && selectedRows.size < output.length;
  }

  function updateDuplicates() {
    const duplicates = core.duplicateStudents(output);
    $("duplicateStatus").hidden = !duplicates.length;
    $("duplicateStatus").textContent = duplicates.length ? "重複學生提醒（各筆均會保留及匯出）\n" +
      duplicates.map(student => student.name + "（學號 " + student.school + "）：共 " + student.count +
        " 筆，重複 " + (student.count - 1) + " 次。").join("\n") : "";
  }

  function renderOutput() {
    $("previewBody").replaceChildren();
    output.forEach((record, index) => {
      const row = document.createElement("tr");
      const selectedCell = document.createElement("td");
      const label = document.createElement("label");
      label.className = "check-label";
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = "row-select";
      checkbox.checked = selectedRows.has(index);
      checkbox.setAttribute("aria-label", "選取 " + record.name + "（" + record.school + "）");
      checkbox.addEventListener("change", () => {
        if (checkbox.checked) selectedRows.add(index); else selectedRows.delete(index);
        updateSelection();
      });
      label.appendChild(checkbox);
      selectedCell.appendChild(label);
      row.appendChild(selectedCell);
      const actionCell = document.createElement("td");
      const actions = document.createElement("div");
      actions.className = "row-actions";
      const edit = document.createElement("button");
      edit.type = "button";
      edit.className = "secondary edit-row";
      edit.textContent = "編輯";
      edit.setAttribute("aria-label", "編輯 " + record.name + "（" + record.school + "）");
      edit.addEventListener("click", () => openEditor([index]));
      const remove = document.createElement("button");
      remove.type = "button";
      remove.className = "danger delete-row";
      remove.textContent = "刪除";
      remove.setAttribute("aria-label", "刪除 " + record.name + "（" + record.school + "）");
      remove.addEventListener("click", () => deleteRecord(index));
      actions.append(edit, remove);
      actionCell.appendChild(actions);
      row.appendChild(actionCell);
      core.fields.forEach(field => {
        const cell = document.createElement("td");
        cell.dataset.field = field;
        cell.textContent = record[field];
        row.appendChild(cell);
      });
      $("previewBody").appendChild(row);
    });
    updateSelection();
    updateDuplicates();
  }

  function replaceOutput(records, selectionIndexes) {
    // Create the next download before changing the visible records or old URL.
    const workbook = core.exportWorkbook(records, template, xlsx);
    const bytes = xlsx.write(workbook, {bookType: "xlsx", type: "array"});
    const nextURL = URL.createObjectURL(new Blob([bytes], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));
    const previousURL = downloadURL;
    output = records;
    downloadURL = nextURL;
    if (selectionIndexes) {
      selectedRows.clear();
      selectionIndexes.forEach(index => selectedRows.add(index));
    }
    const now = new Date();
    const stamp = [now.getFullYear(), String(now.getMonth() + 1).padStart(2, "0"), String(now.getDate()).padStart(2, "0")].join("");
    $("exportBtn").href = downloadURL;
    $("exportBtn").download = "獎懲紀錄_" + stamp + ".xlsx";
    $("exportBtn").setAttribute("aria-disabled", "false");
    $("exportBtn").tabIndex = 0;
    renderOutput();
    if (previousURL) URL.revokeObjectURL(previousURL);
  }

  function deleteRecord(index) {
    if (!output[index]) return;
    const records = output.filter((record, rowIndex) => rowIndex !== index);
    if (!records.length) {
      invalidatePreview("已刪除最後一筆紀錄；可再次批次新增。");
      return;
    }
    const selectionIndexes = Array.from(selectedRows)
      .filter(rowIndex => rowIndex !== index)
      .map(rowIndex => rowIndex > index ? rowIndex - 1 : rowIndex);
    try {
      replaceOutput(records, selectionIndexes);
      reportOutput("已刪除 1 筆紀錄；目前共 " + output.length + " 筆，可直接匯出");
    } catch (error) {
      status("resultStatus", "刪除失敗：" + error.message, "error");
    }
  }

  function reportOutput(message) {
    const {skipped} = previewSummary;
    status("resultStatus", message + "。" +
      (skipped ? "\n上次新增時已略過 " + skipped + " 筆無法辨識的輸入；有效學生紀錄已保留，可正常匯出。" : ""), skipped ? "warning" : "ok");
  }

  function syncEditor() {
    const applyRule = $("editApplyRule").checked;
    const relevant = template.rules.filter(rule => rule.article === $("editArticle").value && rule.clause === $("editClause").value);
    $("editArticle").disabled = !applyRule;
    $("editClause").disabled = !applyRule || !$("editArticle").value;
    $("editItem").disabled = !applyRule || !relevant.some(rule => rule.item);
    $("editReason").disabled = !$("editApplyReason").checked;
    $("editAdjustReason").disabled = !$("editApplyAdjust").checked;
    $("editAdjustAmount").disabled = !$("editApplyAdjust").checked || !$("editAdjustReason").value;
  }

  function fillEditorReason() {
    const rule = template.rules.find(rule => rule.article === $("editArticle").value &&
      rule.clause === $("editClause").value && rule.item === $("editItem").value);
    $("editReason").value = rule ? rule.reason : "";
    $("editAmount").value = rule ? template.defaultAmounts[rule.article] || "" : "";
  }

  function fillEditItems(value) {
    const relevant = template.rules.filter(rule => rule.article === $("editArticle").value && rule.clause === $("editClause").value);
    const items = Array.from(new Set(relevant.map(rule => rule.item).filter(Boolean)));
    fillOptions("editItem", items, items.length ? "請選擇目" : relevant.length ? "無目（不需選擇）" : "請先選擇款");
    $("editItem").value = value || "";
    fillEditorReason();
    syncEditor();
  }

  function fillEditClauses(clause, item) {
    const clauses = Array.from(new Set(template.rules.filter(rule => rule.article === $("editArticle").value).map(rule => rule.clause)));
    fillOptions("editClause", clauses, "請選擇款");
    $("editClause").value = clause || "";
    fillEditItems(item);
  }

  function openEditor(indexes) {
    if (!output.length || !indexes.length) return;
    editingRows = indexes.slice();
    const records = indexes.map(index => output[index]);
    const common = field => records.every(record => record[field] === records[0][field]) ? records[0][field] : "";
    const single = indexes.length === 1;
    $("editTitle").textContent = "編輯 " + indexes.length + " 筆紀錄";
    $("editStudentSummary").textContent = single ? records[0].name + "（學號 " + records[0].school + "）" :
      records.slice(0, 5).map(record => record.name).join("、") + (records.length > 5 ? "等 " + records.length + " 位學生" : "");
    ["editApplyRule", "editApplyReason", "editApplyAdjust"].forEach(id => { $(id).checked = single; });
    fillOptions("editArticle", Array.from(new Set(template.rules.map(rule => rule.article))), "請選擇條");
    $("editArticle").value = common("article");
    fillOptions("editAdjustReason", template.adjustmentOptions, "無（清除）");
    $("editAdjustReason").value = common("adjustReason");
    fillOptions("editAdjustAmount", template.amountOptions, "無（選填）");
    $("editAdjustAmount").value = common("adjustAmount");
    fillEditClauses(common("clause"), common("item"));
    $("editReason").value = common("reason");
    $("editAmount").value = common("amount");
    $("editStatus").textContent = "";
    $("editStatus").className = "";
    $("recordEditor").showModal();
  }

  function saveEdits(event) {
    event.preventDefault();
    try {
      const patch = {};
      if ($("editApplyRule").checked) {
        patch.article = $("editArticle").value;
        patch.clause = $("editClause").value;
        patch.item = $("editItem").value;
      }
      if ($("editApplyReason").checked) patch.reason = $("editReason").value;
      if ($("editApplyAdjust").checked) {
        patch.adjustReason = $("editAdjustReason").value;
        patch.adjustAmount = patch.adjustReason ? $("editAdjustAmount").value : "";
      }
      const records = core.editRecords(output, editingRows, template, patch);
      const count = editingRows.length;
      replaceOutput(records);
      $("recordEditor").close();
      reportOutput("已更新 " + count + " 筆紀錄；目前共 " + output.length + " 筆，可直接匯出");
    } catch (error) {
      $("editStatus").textContent = error.message;
      $("editStatus").className = "error";
    }
  }

  function preview() {
    try {
      const result = core.buildRecords($("studentInput").value, roster, template, selection());
      replaceOutput(output.concat(result.records));
      previewSummary = {skipped: result.resolved.errors.length};
      reportOutput("已新增 " + result.records.length + " 筆紀錄；目前共 " + output.length + " 筆，可直接匯出");
    } catch (error) {
      status("resultStatus", "新增失敗：" + error.message +
        (output.length ? "\n原有 " + output.length + " 筆紀錄已保留，可繼續匯出。" : ""), "error");
    }
  }

  $("rosterFile").addEventListener("change", event => { importFile("roster", event.target); });
  $("templateFile").addEventListener("change", event => { importFile("template", event.target); });
  $("studentInput").addEventListener("input", checkStudents);
  [["article", updateClause], ["clause", updateItem], ["item", syncRuleDetails]].forEach(([id, handler]) => {
    $(id).addEventListener("change", handler);
  });
  $("adjustReason").addEventListener("change", () => {
    $("adjustAmount").disabled = !$("adjustReason").value;
    if (!$("adjustReason").value) $("adjustAmount").value = "";
  });
  $("previewBtn").addEventListener("click", preview);
  $("selectAll").addEventListener("change", () => {
    selectedRows.clear();
    if ($("selectAll").checked) output.forEach((record, index) => selectedRows.add(index));
    renderOutput();
  });
  $("editSelectedBtn").addEventListener("click", () => openEditor(Array.from(selectedRows).sort((a, b) => a - b)));
  $("clearRecordsBtn").addEventListener("click", () => {
    invalidatePreview("已清空預覽紀錄；可再次批次新增。");
  });
  $("editArticle").addEventListener("change", () => fillEditClauses());
  $("editClause").addEventListener("change", () => fillEditItems());
  $("editItem").addEventListener("change", () => { fillEditorReason(); syncEditor(); });
  ["editApplyRule", "editApplyReason", "editApplyAdjust"].forEach(id => $(id).addEventListener("change", syncEditor));
  $("editAdjustReason").addEventListener("change", () => {
    if (!$("editAdjustReason").value) $("editAdjustAmount").value = "";
    syncEditor();
  });
  $("recordEditForm").addEventListener("submit", saveEdits);
  $("cancelEditBtn").addEventListener("click", () => $("recordEditor").close());
  $("recordEditor").addEventListener("close", () => { if (!$("recordEditor").open) editingRows = []; });
  $("exportBtn").addEventListener("click", event => {
    if (!output.length || !downloadURL) event.preventDefault();
  });
  $("clearBtn").addEventListener("click", () => {
    $("studentInput").value = "";
    resetRuleControls();
    checkStudents();
    if (output.length) reportOutput("已清除上方輸入；原有 " + output.length + " 筆預覽紀錄保留，可直接匯出");
    else status("resultStatus", "已清除輸入；匯入的名冊與範本仍可使用。");
  });
  resetRuleControls();
  updateSelection();
  refreshImports();
  checkStudents();
  status("startupStatus", "程式已就緒，可選取 Excel 檔案。", "ok");
  document.documentElement.dataset.appReady = "true";
})();
