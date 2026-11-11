(function () {
  "use strict";

  var DEFAULT_GROUPS = [
    { region: "雲東區", members: ["斗六", "古坑", "林內", "西螺", "莿桐", "斗南"] },
    { region: "雲西區", members: ["虎尾", "土庫", "崙背", "褒忠", "二崙", "麥寮", "北港", "口湖"] },
    { region: "嘉義區", members: ["嘉義市（梅山）", "中埔", "竹崎", "番路"] },
    { region: "民雄區", members: ["民雄", "溪口", "大林", "新港", "(六腳)"] },
    { region: "朴子區", members: ["朴子", "布袋", "鹿草", "太保", "水上"] }
  ];
  var DEFAULT_REGION_NAME = "未分區";

  function normalizeGroups(raw) {
    if (!Array.isArray(raw)) return null;
    var seen = {};
    var groups = [];
    raw.forEach(function (g) {
      if (!g || typeof g.region !== "string" || !Array.isArray(g.members)) return;
      var region = g.region.trim();
      if (!region) return;
      var members = [];
      g.members.forEach(function (name) {
        var n = typeof name === "string" ? name.trim() : "";
        if (!n || seen[n]) return;
        seen[n] = true;
        members.push(n);
      });
      groups.push({ region: region, members: members });
    });
    return groups;
  }

  function flattenGroups(groups) {
    var flat = [];
    groups.forEach(function (g) {
      g.members.forEach(function (name) { flat.push(name); });
    });
    return flat;
  }

  var CONGREGATION_GROUPS = [];
  var CONGREGATIONS = [];
  var dataState = {};
  var rollcallState = {};
  var ROOM_ID = null;

  var TOTAL_LABEL = "合計";

  function getSheet(workbook, preferredName, fallbackIndex) {
    var sheet = workbook.Sheets[preferredName];
    if (!sheet) {
      var fallbackName = workbook.SheetNames[fallbackIndex];
      sheet = fallbackName ? workbook.Sheets[fallbackName] : null;
    }
    return sheet;
  }

  function formatNum(n) {
    return Math.round(n * 10) / 10;
  }

  // .xlsx files are ZIP archives (magic bytes "PK") and must be read as
  // binary; .csv exports are plain UTF-8 text, and handing their raw bytes
  // to XLSX.read as "array" silently mojibake's every Chinese character
  // (it guesses a single-byte codepage instead of UTF-8). Sniff the actual
  // bytes rather than trusting the file extension, and decode CSV text
  // ourselves first so SheetJS parses already-correct UTF-8.
  function readWorkbook(buf) {
    var bytes = new Uint8Array(buf);
    var isZip = bytes.length > 2 && bytes[0] === 0x50 && bytes[1] === 0x4b;
    if (isZip) {
      return XLSX.read(buf, { type: "array" });
    }
    var text = new TextDecoder("utf-8").decode(buf);
    return XLSX.read(text, { type: "string" });
  }

  // ---- Room API ----

  function apiCreateRoom() {
    return fetch("/api/rooms", { method: "POST" }).then(function (res) {
      if (!res.ok) throw new Error("建立房間失敗");
      return res.json();
    });
  }

  function apiFetchRoom(id) {
    return fetch("/api/rooms/" + encodeURIComponent(id)).then(function (res) {
      if (res.status === 404) {
        var err = new Error("找不到這個房間");
        err.notFound = true;
        throw err;
      }
      if (!res.ok) throw new Error("連線失敗，請稍後再試");
      return res.json();
    });
  }

  function apiSaveGroups(id, groups) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/groups", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ groups: groups })
    }).then(function (res) {
      if (!res.ok) throw new Error("儲存失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiUpsertCongregation(id, name, entry) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/congregations/" + encodeURIComponent(name), {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(entry)
    }).then(function (res) {
      if (!res.ok) throw new Error("儲存失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiDeleteCongregation(id, name) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/congregations/" + encodeURIComponent(name), {
      method: "DELETE"
    }).then(function (res) {
      if (!res.ok) throw new Error("刪除失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiResetStats(id) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/reset", { method: "POST" }).then(function (res) {
      if (!res.ok) throw new Error("清除失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiSetRoomName(id, name) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/name", {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: name })
    }).then(function (res) {
      if (!res.ok) throw new Error("命名失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiUpsertRollcall(id, name, entry) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/rollcall/" + encodeURIComponent(name), {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(entry)
    }).then(function (res) {
      if (!res.ok) throw new Error("儲存失敗，請檢查網路連線（若剛更新版本，請確認資料庫已套用 migrations/0003）");
      return res.json();
    });
  }

  function apiDeleteRollcall(id, name) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/rollcall/" + encodeURIComponent(name), {
      method: "DELETE"
    }).then(function (res) {
      if (!res.ok) throw new Error("刪除失敗，請檢查網路連線");
      return res.json();
    });
  }

  function apiResetRollcall(id) {
    return fetch("/api/rooms/" + encodeURIComponent(id) + "/rollcall", { method: "DELETE" }).then(function (res) {
      if (!res.ok) throw new Error("清除失敗，請檢查網路連線");
      return res.json();
    });
  }

  function getRoomIdFromUrl() {
    var v = new URLSearchParams(window.location.search).get("room");
    return v ? v.trim().toUpperCase() : "";
  }

  function goToRoom(id) {
    var url = new URL(window.location.href);
    url.search = "?room=" + encodeURIComponent(id);
    window.location.href = url.toString();
  }

  // ---- 週報網格：一次掃描，兒童／青職共用 ----

  var GRID_SHEET_NAME = "週報網格";
  var GRID_WEEK_PATTERN = /^\d+W\d+$/;

  function fillForward(row, startCol) {
    var out = [];
    var last = null;
    var len = row ? row.length : 0;
    for (var c = 0; c < len; c++) {
      var v = row[c];
      if (v != null && String(v).trim() !== "") last = String(v).trim();
      out[c] = c >= startCol ? last : null;
    }
    return out;
  }

  // Standalone single-column metrics (兒童, 合計 - no age breakdown) are
  // labelled with age "小計" in the per-week columns, but the report's own
  // "N 週平均"/"N 週合計" summary blocks instead repeat the category name
  // itself as the age label (e.g. category "兒童" / age "兒童"). Collapse
  // that back to "小計" so both blocks key the same metric identically.
  function normalizeAgeLabel(cat, ageCell) {
    var age = ageCell != null ? String(ageCell).trim() : "";
    return age === cat ? "小計" : age;
  }

  function scanWeeklyGrid(workbook) {
    var sheet = getSheet(workbook, GRID_SHEET_NAME, 0);
    if (!sheet) {
      throw new Error("找不到「" + GRID_SHEET_NAME + "」分頁，請確認上傳的檔案格式");
    }

    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
    if (!rows.length) {
      throw new Error("「" + GRID_SHEET_NAME + "」分頁是空的");
    }

    // The header row's exact label text has changed between report template
    // versions (區/排 vs 大區/區), so anchor on the first week-code cell
    // (e.g. "2026W28") instead - it's the stable structural marker. The
    // group/合計 label column is always the one immediately to its left.
    var headerRowIdx = -1;
    var groupColIdx = -1;
    for (var r = 0; r < rows.length && headerRowIdx === -1; r++) {
      var row = rows[r];
      if (!row) continue;
      for (var c = 1; c < row.length; c++) {
        if (row[c] != null && GRID_WEEK_PATTERN.test(String(row[c]).trim())) {
          headerRowIdx = r;
          groupColIdx = c - 1;
          break;
        }
      }
    }

    if (headerRowIdx === -1) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到週別欄位（例如 2026W28）");
    }

    var weekRow = rows[headerRowIdx];
    var categoryRow = rows[headerRowIdx + 1];
    var ageRow = rows[headerRowIdx + 2];

    if (!categoryRow || !ageRow) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到分類或身份列");
    }

    var dataStart = headerRowIdx + 3;
    var aggRow = null;
    var dataRows = [];
    for (var i = dataStart; i < rows.length; i++) {
      var dr = rows[i];
      if (!dr) continue;
      var groupLabel = dr[groupColIdx];
      if (groupLabel == null || String(groupLabel).trim() === "") continue;
      if (String(groupLabel).trim() === TOTAL_LABEL) {
        aggRow = dr;
      } else {
        dataRows.push(dr);
      }
    }

    if (!aggRow && !dataRows.length) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到任何資料列");
    }

    function valueAtCol(col) {
      if (aggRow) return Number(aggRow[col]) || 0;
      return dataRows.reduce(function (acc, dr) { return acc + (Number(dr[col]) || 0); }, 0);
    }

    var width = Math.max(weekRow.length, categoryRow.length, ageRow.length);

    // Prefer the report's own pre-computed "N 週平均" column block over
    // summing individual week columns ourselves - the sheet's own tool
    // decides which weeks count toward that average (e.g. it may exclude
    // weeks the raw grid still lists), so re-deriving it independently can
    // silently disagree with the number the congregation actually reports.
    var avgBlockStartCol = -1;
    var avgWeeksLabel = null;
    for (var col = groupColIdx + 1; col < weekRow.length; col++) {
      var cell = weekRow[col];
      if (cell == null) continue;
      var m = String(cell).trim().match(/^(\d+)\s*週平均$/);
      if (m) {
        avgBlockStartCol = col;
        avgWeeksLabel = m[1];
        break;
      }
    }

    if (avgBlockStartCol !== -1) {
      var avgBlockEndCol = width;
      for (var col2 = avgBlockStartCol + 1; col2 < width; col2++) {
        if (weekRow[col2] != null && String(weekRow[col2]).trim() !== "") {
          avgBlockEndCol = col2;
          break;
        }
      }

      var catPerColAvg = fillForward(categoryRow, avgBlockStartCol);
      var avgSums = {};
      for (var c = avgBlockStartCol; c < avgBlockEndCol; c++) {
        var cat = catPerColAvg[c];
        if (!cat) continue;
        var age = normalizeAgeLabel(cat, ageRow[c]);
        avgSums[cat + "|" + age] = valueAtCol(c);
      }

      return { weeks: Number(avgWeeksLabel) || 0, sums: avgSums };
    }

    // Fallback: no pre-computed average block found - sum each real week
    // column ourselves and divide by how many week columns were detected.
    var weekPerCol = fillForward(weekRow, groupColIdx + 1);
    var catPerCol = fillForward(categoryRow, groupColIdx + 1);

    var matches = [];
    for (var col3 = groupColIdx + 1; col3 < width; col3++) {
      var wk = weekPerCol[col3];
      if (!wk || !GRID_WEEK_PATTERN.test(wk)) continue;
      var cat3 = catPerCol[col3];
      if (!cat3) continue;
      var age3 = normalizeAgeLabel(cat3, ageRow[col3]);
      matches.push({ week: wk, key: cat3 + "|" + age3, col: col3 });
    }

    if (!matches.length) {
      throw new Error("在「" + GRID_SHEET_NAME + "」分頁找不到任何週別欄位資料");
    }

    var sums = {};
    var weeksSet = {};
    var weeksCount = 0;

    matches.forEach(function (m) {
      if (!weeksSet[m.week]) {
        weeksSet[m.week] = true;
        weeksCount++;
      }
      sums[m.key] = (sums[m.key] || 0) + valueAtCol(m.col);
    });

    Object.keys(sums).forEach(function (key) { sums[key] = sums[key] / weeksCount; });

    return { weeks: weeksCount, sums: sums };
  }

  function deriveChildrenResult(scan) {
    var get = function (cat, age) { return scan.sums[cat + "|" + age] || 0; };
    return {
      weeks: scan.weeks,
      avgSunday: get("兒童", "小計") + get("主日", "國小"),
      avgGroup: get("小排", "學齡前") + get("小排", "國小"),
      avgLife: get("召會生活", "學齡前") + get("召會生活", "國小")
    };
  }

  function deriveYouthResult(scan) {
    var get = function (cat, age) { return scan.sums[cat + "|" + age] || 0; };
    return {
      weeks: scan.weeks,
      avgSunday: get("主日", "青職"),
      avgFamily: get("家聚會出訪", "青職") + get("家聚會受訪", "青職"),
      avgGroup: get("小排", "青職"),
      avgLifeReading: get("生命讀經", "青職")
    };
  }

  // ---- 共用區塊（進度／表格／總計）算繪工廠 ----

  function createSection(cfg) {
    var prefix = cfg.prefix;
    var metrics = cfg.metrics;
    var extract = cfg.extract;

    function id(name) { return document.getElementById(prefix + "-" + name); }

    var els = {
      progressCount: id("progress-count"),
      progressTotal: id("progress-total"),
      progressBar: id("progress-bar"),
      chipGrid: id("chip-grid"),
      resultsTheadRow: id("results-thead-row"),
      resultsTbody: id("results-tbody"),
      resultsEmpty: id("results-empty"),
      downloadBtn: id("download-btn"),
      summaryGrid: id("summary-grid"),
      summaryWarn: id("summary-warn")
    };

    function initTableHeader() {
      var html = "<th>召會 / 區域</th><th>週數</th>";
      metrics.forEach(function (m) { html += "<th>" + m.label + "</th>"; });
      els.resultsTheadRow.innerHTML = html;
    }

    function initSummaryGrid() {
      var html = "";
      metrics.forEach(function (m) {
        html +=
          '<div class="stat-tile">' +
          '<div class="stat-label">' + m.totalLabel + "</div>" +
          '<div class="stat-value" id="' + prefix + "-total-" + m.key + '">0</div>' +
          "</div>";
      });
      els.summaryGrid.innerHTML = html;
    }

    function renderChecklist(state) {
      var uploadedCount = 0;
      els.chipGrid.innerHTML = "";
      CONGREGATIONS.forEach(function (name) {
        var chip = document.createElement("span");
        var done = Object.prototype.hasOwnProperty.call(state, name);
        if (done) uploadedCount++;
        chip.className = "chip" + (done ? " done" : "");
        chip.textContent = name;
        els.chipGrid.appendChild(chip);
      });
      els.progressCount.textContent = uploadedCount;
      els.progressTotal.textContent = CONGREGATIONS.length;
      els.progressBar.style.width = (CONGREGATIONS.length ? (uploadedCount / CONGREGATIONS.length * 100) : 0) + "%";
    }

    function regionTotalRow(entries) {
      // 週數 stays a mean (it's informational context, not one of the
      // stats being reported); every actual metric column is a straight
      // sum of the region's uploaded congregations, per how this region
      // total is meant to be read.
      var total = { weeks: 0 };
      metrics.forEach(function (m) { total[m.key] = 0; });
      if (!entries.length) return total;
      total.weeks = entries.reduce(function (acc, d) { return acc + d.weeks; }, 0) / entries.length;
      metrics.forEach(function (m) {
        total[m.key] = entries.reduce(function (acc, d) { return acc + (d[m.key] || 0); }, 0);
      });
      return total;
    }

    function buildRows(state) {
      var out = [];
      CONGREGATION_GROUPS.forEach(function (group) {
        var uploaded = group.members.filter(function (name) {
          return Object.prototype.hasOwnProperty.call(state, name);
        });
        if (!uploaded.length) return;
        var entries = uploaded.map(function (name) { return extract(state[name]); });
        uploaded.forEach(function (name) {
          out.push({ type: "congregation", name: name, data: extract(state[name]) });
        });
        out.push({ type: "region", region: group.region, data: regionTotalRow(entries) });
      });
      return out;
    }

    function renderTable(state, onRemove) {
      els.resultsTbody.innerHTML = "";
      var rowsData = buildRows(state);

      els.resultsEmpty.style.display = rowsData.length ? "none" : "block";
      els.downloadBtn.disabled = !rowsData.length;

      rowsData.forEach(function (r) {
        var tr = document.createElement("tr");
        var html;
        if (r.type === "region") {
          tr.className = "region-row";
          html = "<td>" + r.region + "</td><td>" + formatNum(r.data.weeks) + "</td>";
          metrics.forEach(function (m) { html += "<td>" + formatNum(r.data[m.key]) + "</td>"; });
          html += "<td></td>";
        } else {
          html = "<td class=\"congregation-name\">" + r.name + "</td><td>" + r.data.weeks + "</td>";
          metrics.forEach(function (m) { html += "<td>" + formatNum(r.data[m.key]) + "</td>"; });
          html += "<td><button class=\"btn-danger-ghost\" data-remove=\"" + r.name + "\">移除</button></td>";
        }
        tr.innerHTML = html;
        els.resultsTbody.appendChild(tr);
      });

      els.resultsTbody.querySelectorAll("[data-remove]").forEach(function (btn) {
        btn.addEventListener("click", function () {
          onRemove(btn.getAttribute("data-remove"));
        });
      });
    }

    function renderSummary(state) {
      var totals = {};
      metrics.forEach(function (m) { totals[m.key] = 0; });
      Object.keys(state).forEach(function (name) {
        var d = extract(state[name]);
        metrics.forEach(function (m) { totals[m.key] += d[m.key] || 0; });
      });

      metrics.forEach(function (m) {
        var el = document.getElementById(prefix + "-total-" + m.key);
        if (el) el.textContent = formatNum(totals[m.key]);
      });

      var missing = CONGREGATIONS.filter(function (name) {
        return !Object.prototype.hasOwnProperty.call(state, name);
      });

      if (missing.length) {
        els.summaryWarn.style.display = "block";
        els.summaryWarn.textContent =
          "尚有 " + missing.length + " 個召會未上傳（" + missing.join("、") + "），以上總計僅計入已上傳的召會。";
      } else {
        els.summaryWarn.style.display = "none";
      }
    }

    function onDownloadClick(state) {
      var rowsData = buildRows(state);
      if (!rowsData.length) return;

      var header = ["召會 / 區域", "週數"].concat(metrics.map(function (m) { return m.label; }));
      var aoa = [header];
      rowsData.forEach(function (r) {
        var label = r.type === "region" ? r.region : r.name;
        var row = [label, formatNum(r.data.weeks)];
        metrics.forEach(function (m) { row.push(formatNum(r.data[m.key])); });
        aoa.push(row);
      });

      var ws = XLSX.utils.aoa_to_sheet(aoa);
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, cfg.summaryTitle);
      XLSX.writeFile(wb, cfg.fileBaseName + ".xlsx");
    }

    initTableHeader();
    initSummaryGrid();

    return {
      renderAll: function (state, onRemove) {
        renderChecklist(state);
        renderTable(state, onRemove);
        renderSummary(state);
      },
      bindDownloadButton: function (getState) {
        els.downloadBtn.addEventListener("click", function () {
          onDownloadClick(getState());
        });
      }
    };
  }

  function initUploadAndSections() {
    var els = {
      select: document.getElementById("congregation-select"),
      dropzone: document.getElementById("dropzone"),
      fileInput: document.getElementById("file-input"),
      filePicked: document.getElementById("file-picked"),
      parseBtn: document.getElementById("parse-btn"),
      resetAllBtn: document.getElementById("reset-all-btn"),
      statusMsg: document.getElementById("status-msg")
    };

    var pickedFile = null;
    var pendingWrites = 0;

    function trackWrite(promise) {
      pendingWrites++;
      var done = function () { pendingWrites--; };
      promise.then(done, done);
      return promise;
    }

    function showStatus(message, type) {
      els.statusMsg.textContent = message;
      els.statusMsg.className = "status-msg show " + type;
    }

    function clearStatus() {
      els.statusMsg.className = "status-msg";
    }

    function initSelect() {
      els.select.innerHTML = "";
      if (!CONGREGATIONS.length) {
        var placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "尚未設定召會，請先到設定新增";
        els.select.appendChild(placeholder);
        els.parseBtn.disabled = true;
        return;
      }
      els.parseBtn.disabled = false;
      CONGREGATION_GROUPS.forEach(function (group) {
        var optgroup = document.createElement("optgroup");
        optgroup.label = group.region;
        group.members.forEach(function (name) {
          var opt = document.createElement("option");
          opt.value = name;
          opt.textContent = name;
          optgroup.appendChild(opt);
        });
        els.select.appendChild(optgroup);
      });
    }

    function handleFileSelected(file) {
      pickedFile = file;
      els.filePicked.textContent = file ? "已選擇檔案：" + file.name : "";
    }

    function setupDropzone() {
      var dz = els.dropzone;
      ["dragenter", "dragover"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) {
          e.preventDefault();
          dz.classList.add("drag-over");
        });
      });
      ["dragleave", "drop"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) {
          e.preventDefault();
          dz.classList.remove("drag-over");
        });
      });
      dz.addEventListener("drop", function (e) {
        var files = e.dataTransfer.files;
        if (files && files.length) handleFileSelected(files[0]);
      });
      els.fileInput.addEventListener("change", function (e) {
        if (e.target.files && e.target.files.length) handleFileSelected(e.target.files[0]);
      });
    }

    var childrenSection = createSection({
      prefix: "children",
      summaryTitle: "全台兒童統計",
      fileBaseName: "children-stats",
      extract: function (entry) { return entry.children; },
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "avgLife", label: "召會生活（週平均）", totalLabel: "召會生活總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" }
      ]
    });

    var youthSection = createSection({
      prefix: "youth",
      summaryTitle: "全台青職統計",
      fileBaseName: "youth-stats",
      extract: function (entry) { return entry.youth; },
      metrics: [
        { key: "avgSunday", label: "主日（週平均）", totalLabel: "主日總計" },
        { key: "avgFamily", label: "家聚會（出訪+受訪，週平均）", totalLabel: "家聚會總計" },
        { key: "avgGroup", label: "小排（週平均）", totalLabel: "小排總計" },
        { key: "avgLifeReading", label: "生命讀經（週平均）", totalLabel: "生命讀經總計" }
      ]
    });

    function removeCongregation(name) {
      var ok = window.confirm("確定要移除「" + name + "」的資料嗎？（兒童、青職資料會一併移除）");
      if (!ok) return;
      delete dataState[name];
      renderAll();
      trackWrite(apiDeleteCongregation(ROOM_ID, name)).catch(function (err) {
        showStatus("移除失敗：" + err.message, "error");
      });
    }

    function renderAll() {
      childrenSection.renderAll(dataState, removeCongregation);
      youthSection.renderAll(dataState, removeCongregation);
    }

    function onParseClick() {
      clearStatus();
      var congregation = els.select.value;
      if (!congregation) {
        showStatus("請先到設定新增召會", "error");
        return;
      }
      if (!pickedFile) {
        showStatus("請先選擇一個 .xlsx 檔案", "error");
        return;
      }
      els.parseBtn.disabled = true;

      pickedFile.arrayBuffer().then(function (buf) {
        var workbook = readWorkbook(buf);
        var scan = scanWeeklyGrid(workbook);
        var childrenResult = deriveChildrenResult(scan);
        var youthResult = deriveYouthResult(scan);

        var isUpdate = Object.prototype.hasOwnProperty.call(dataState, congregation);
        if (isUpdate) {
          var ok = window.confirm("「" + congregation + "」已經上傳過資料（兒童、青職），是否要用這個新檔案覆蓋？");
          if (!ok) {
            els.parseBtn.disabled = false;
            return;
          }
        }

        var entry = {
          weeks: scan.weeks,
          children: childrenResult,
          youth: youthResult,
          fileName: pickedFile.name,
          updatedAt: new Date().toISOString()
        };
        dataState[congregation] = entry;
        renderAll();

        showStatus(
          "已加入「" + congregation + "」：共 " + scan.weeks + " 週。" +
          "兒童 — 主日 " + formatNum(childrenResult.avgSunday) + "、召會生活 " + formatNum(childrenResult.avgLife) + "、小排 " + formatNum(childrenResult.avgGroup) + "；" +
          "青職 — 主日 " + formatNum(youthResult.avgSunday) + "、家聚會 " + formatNum(youthResult.avgFamily) + "、小排 " + formatNum(youthResult.avgGroup) + "、生命讀經 " + formatNum(youthResult.avgLifeReading) +
          "（儲存中…）",
          "ok"
        );

        pickedFile = null;
        els.fileInput.value = "";
        els.filePicked.textContent = "";
        els.parseBtn.disabled = false;

        trackWrite(apiUpsertCongregation(ROOM_ID, congregation, entry)).then(function () {
          showStatus("已加入「" + congregation + "」並儲存到房間。", "ok");
        }).catch(function (err) {
          showStatus("已加入「" + congregation + "」，但儲存失敗：" + err.message, "error");
        });
      }).catch(function (err) {
        showStatus("解析失敗：" + err.message, "error");
        els.parseBtn.disabled = false;
      });
    }

    function onResetAllClick() {
      if (!Object.keys(dataState).length) return;
      var ok = window.confirm("確定要清除所有已上傳的召會資料嗎？（兒童、青職資料會一併清除，此操作無法復原）");
      if (!ok) return;
      dataState = {};
      renderAll();
      clearStatus();
      trackWrite(apiResetStats(ROOM_ID)).catch(function (err) {
        showStatus("清除失敗：" + err.message, "error");
      });
    }

    initSelect();
    setupDropzone();
    els.parseBtn.addEventListener("click", onParseClick);
    els.resetAllBtn.addEventListener("click", onResetAllClick);
    childrenSection.bindDownloadButton(function () { return dataState; });
    youthSection.bindDownloadButton(function () { return dataState; });

    renderAll();

    return {
      initSelect: initSelect,
      renderAll: renderAll,
      hasPendingWrites: function () { return pendingWrites > 0; }
    };
  }

  // ---- 點名報表：每月出席一次以上的人數（可依群組篩選） ----

  var ROLLCALL_SHEET_NAME = "出席";
  var ROLLCALL_GROUP_ORDER = ["學齡前", "國小", "國中", "高中", "大學", "青職", "青壯", "中壯", "年長"];
  var ROLLCALL_NO_GROUP = "未分組";
  var DAY_MS = 86400000;

  // Week columns are headed by the week's Monday ("2026-08-31"). Excel may
  // also hand them over as real date cells (serial numbers) or "2026/8/31".
  function parseWeekDate(cell) {
    if (cell == null) return null;
    if (typeof cell === "number") {
      if (cell < 20000 || cell > 80000) return null;
      return new Date(Date.UTC(1899, 11, 30) + Math.round(cell) * DAY_MS);
    }
    var m = String(cell).trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
    if (!m) return null;
    return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  }

  // A week belongs to the month its Lord's day (Monday + 6) falls in - this
  // matches the report's own range, e.g. "2026年9月" starts at week 2026-08-31.
  function weekMonthKey(monday) {
    var sunday = new Date(monday.getTime() + 6 * DAY_MS);
    var mm = sunday.getUTCMonth() + 1;
    return sunday.getUTCFullYear() + "-" + (mm < 10 ? "0" : "") + mm;
  }

  function monthLabel(key) {
    var parts = key.split("-");
    return parts[0] + "年" + Number(parts[1]) + "月";
  }

  function isAttendedCell(v) {
    if (v == null) return false;
    if (typeof v === "number") return v > 0;
    var t = String(v).trim();
    if (!t) return false;
    if (!isNaN(Number(t))) return Number(t) > 0;
    return /^(v|✓|✔|○|o|y|yes|出席)$/i.test(t);
  }

  function parseRollcall(workbook) {
    var sheet = getSheet(workbook, ROLLCALL_SHEET_NAME, 0);
    if (!sheet) throw new Error("找不到「" + ROLLCALL_SHEET_NAME + "」分頁，請確認上傳的是點名報表");
    var rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

    var headerIdx = -1, nameCol = -1, groupCol = -1;
    for (var r = 0; r < rows.length && headerIdx === -1; r++) {
      var row = rows[r] || [];
      for (var c = 0; c < row.length; c++) {
        if (row[c] != null && String(row[c]).trim() === "姓名") {
          headerIdx = r;
          nameCol = c;
          break;
        }
      }
    }
    if (headerIdx === -1) throw new Error("找不到「姓名」欄位，請確認上傳的是點名報表");

    var header = rows[headerIdx];
    var weekCols = [];
    for (var c2 = 0; c2 < header.length; c2++) {
      var label = header[c2] != null ? String(header[c2]).trim() : "";
      if (label === "群組") groupCol = c2;
      var d = parseWeekDate(header[c2]);
      if (d) weekCols.push({ col: c2, month: weekMonthKey(d) });
    }
    if (!weekCols.length) throw new Error("找不到週別日期欄位（例如 2026-08-31）");

    var months = {};
    weekCols.forEach(function (w) {
      if (!months[w.month]) months[w.month] = { weeks: 0, attended: {} };
      months[w.month].weeks++;
    });

    var roster = {};
    var people = 0;
    for (var i = headerIdx + 1; i < rows.length; i++) {
      var dr = rows[i];
      if (!dr) continue;
      var name = dr[nameCol];
      if (name == null || String(name).trim() === "") continue;
      var grp = groupCol !== -1 && dr[groupCol] != null && String(dr[groupCol]).trim() !== ""
        ? String(dr[groupCol]).trim()
        : ROLLCALL_NO_GROUP;
      people++;
      roster[grp] = (roster[grp] || 0) + 1;

      var hitMonths = {};
      weekCols.forEach(function (w) {
        if (isAttendedCell(dr[w.col])) hitMonths[w.month] = true;
      });
      Object.keys(hitMonths).forEach(function (mk) {
        var att = months[mk].attended;
        att[grp] = (att[grp] || 0) + 1;
      });
    }
    if (!people) throw new Error("「" + ROLLCALL_SHEET_NAME + "」分頁沒有任何名單資料");

    var title = rows[0] && rows[0][0] != null ? String(rows[0][0]).trim() : "";
    return { months: months, roster: roster, people: people, title: title };
  }

  function sortGroups(list) {
    return list.slice().sort(function (a, b) {
      var ia = ROLLCALL_GROUP_ORDER.indexOf(a), ib = ROLLCALL_GROUP_ORDER.indexOf(b);
      if (ia === -1) ia = a === ROLLCALL_NO_GROUP ? 1000 : 999;
      if (ib === -1) ib = b === ROLLCALL_NO_GROUP ? 1000 : 999;
      if (ia !== ib) return ia - ib;
      return a < b ? -1 : a > b ? 1 : 0;
    });
  }

  function initRollcallSection() {
    function $(name) { return document.getElementById("rollcall-" + name); }
    var els = {
      select: $("congregation-select"),
      dropzone: $("dropzone"),
      fileInput: $("file-input"),
      filePicked: $("file-picked"),
      parseBtn: $("parse-btn"),
      resetBtn: $("reset-btn"),
      statusMsg: $("status-msg"),
      progressCount: $("progress-count"),
      progressTotal: $("progress-total"),
      progressBar: $("progress-bar"),
      chipGrid: $("chip-grid"),
      groupFilter: $("group-filter"),
      groupEmpty: $("group-empty"),
      groupAllBtn: $("group-all-btn"),
      groupNoneBtn: $("group-none-btn"),
      theadRow: $("results-thead-row"),
      tbody: $("results-tbody"),
      resultsEmpty: $("results-empty"),
      downloadBtn: $("download-btn"),
      summaryGrid: $("summary-grid"),
      summaryWarn: $("summary-warn")
    };

    var pickedFile = null;
    var pendingWrites = 0;
    // Remembered per viewer (not shared): which groups this viewer unticked.
    // Groups appearing for the first time are therefore ticked by default.
    var storageKey = "rollcall-deselected-" + ROOM_ID;
    var deselected = {};
    try {
      var saved = JSON.parse(localStorage.getItem(storageKey) || "[]");
      if (Array.isArray(saved)) saved.forEach(function (g) { deselected[g] = true; });
    } catch (e) { /* storage unavailable - default to all groups */ }

    function saveDeselected() {
      try { localStorage.setItem(storageKey, JSON.stringify(Object.keys(deselected))); } catch (e) { /* ignore */ }
    }

    function trackWrite(promise) {
      pendingWrites++;
      var done = function () { pendingWrites--; };
      promise.then(done, done);
      return promise;
    }

    function showStatus(message, type) {
      els.statusMsg.textContent = message;
      els.statusMsg.className = "status-msg show " + type;
    }

    function has(name) { return Object.prototype.hasOwnProperty.call(rollcallState, name); }

    function initSelect() {
      var prev = els.select.value;
      els.select.innerHTML = "";
      if (!CONGREGATIONS.length) {
        var placeholder = document.createElement("option");
        placeholder.value = "";
        placeholder.textContent = "尚未設定召會，請先到設定新增";
        els.select.appendChild(placeholder);
        els.parseBtn.disabled = true;
        return;
      }
      els.parseBtn.disabled = false;
      CONGREGATION_GROUPS.forEach(function (group) {
        var optgroup = document.createElement("optgroup");
        optgroup.label = group.region;
        group.members.forEach(function (name) {
          var opt = document.createElement("option");
          opt.value = name;
          opt.textContent = name;
          optgroup.appendChild(opt);
        });
        els.select.appendChild(optgroup);
      });
      if (prev && CONGREGATIONS.indexOf(prev) !== -1) els.select.value = prev;
    }

    function handleFileSelected(file) {
      pickedFile = file;
      els.filePicked.textContent = file ? "已選擇檔案：" + file.name : "";
    }

    function setupDropzone() {
      var dz = els.dropzone;
      ["dragenter", "dragover"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.add("drag-over"); });
      });
      ["dragleave", "drop"].forEach(function (evt) {
        dz.addEventListener(evt, function (e) { e.preventDefault(); dz.classList.remove("drag-over"); });
      });
      dz.addEventListener("drop", function (e) {
        var files = e.dataTransfer.files;
        if (files && files.length) handleFileSelected(files[0]);
      });
      els.fileInput.addEventListener("change", function (e) {
        if (e.target.files && e.target.files.length) handleFileSelected(e.target.files[0]);
      });
    }

    function visibleNames() {
      return CONGREGATIONS.filter(has);
    }

    function allGroups() {
      var seen = {};
      visibleNames().forEach(function (name) {
        Object.keys(rollcallState[name].roster || {}).forEach(function (g) { seen[g] = true; });
      });
      return sortGroups(Object.keys(seen));
    }

    function selectedGroups() {
      return allGroups().filter(function (g) { return !deselected[g]; });
    }

    function allMonths() {
      var seen = {};
      visibleNames().forEach(function (name) {
        Object.keys(rollcallState[name].months || {}).forEach(function (m) { seen[m] = true; });
      });
      return Object.keys(seen).sort();
    }

    // { roster, months: { key: count | null } } for one entry, summed over
    // the selected groups. null = this report doesn't cover that month.
    function extract(entry, groups, months) {
      var out = { roster: 0, months: {} };
      groups.forEach(function (g) { out.roster += (entry.roster && entry.roster[g]) || 0; });
      months.forEach(function (mk) {
        var m = entry.months && entry.months[mk];
        if (!m) { out.months[mk] = null; return; }
        var n = 0;
        groups.forEach(function (g) { n += (m.attended && m.attended[g]) || 0; });
        out.months[mk] = n;
      });
      return out;
    }

    function sumRows(list, months) {
      var total = { roster: 0, months: {} };
      months.forEach(function (mk) { total.months[mk] = null; });
      list.forEach(function (d) {
        total.roster += d.roster;
        months.forEach(function (mk) {
          if (d.months[mk] == null) return;
          total.months[mk] = (total.months[mk] || 0) + d.months[mk];
        });
      });
      return total;
    }

    function buildRows(groups, months) {
      var out = [];
      CONGREGATION_GROUPS.forEach(function (group) {
        var uploaded = group.members.filter(has);
        if (!uploaded.length) return;
        var entries = [];
        uploaded.forEach(function (name) {
          var d = extract(rollcallState[name], groups, months);
          entries.push(d);
          out.push({ type: "congregation", name: name, data: d });
        });
        out.push({ type: "region", region: group.region, data: sumRows(entries, months) });
      });
      return out;
    }

    function cell(v) { return v == null ? "—" : String(v); }

    function escapeHtml(t) {
      return String(t).replace(/[&<>"]/g, function (ch) {
        return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[ch];
      });
    }

    function renderChecklist() {
      var count = 0;
      els.chipGrid.innerHTML = "";
      CONGREGATIONS.forEach(function (name) {
        var chip = document.createElement("span");
        var done = has(name);
        if (done) count++;
        chip.className = "chip" + (done ? " done" : "");
        chip.textContent = name;
        els.chipGrid.appendChild(chip);
      });
      els.progressCount.textContent = count;
      els.progressTotal.textContent = CONGREGATIONS.length;
      els.progressBar.style.width = (CONGREGATIONS.length ? count / CONGREGATIONS.length * 100 : 0) + "%";
    }

    function renderGroupFilter() {
      var groups = allGroups();
      els.groupFilter.innerHTML = "";
      els.groupEmpty.style.display = groups.length ? "none" : "block";
      els.groupAllBtn.disabled = els.groupNoneBtn.disabled = !groups.length;
      groups.forEach(function (g) {
        var label = document.createElement("label");
        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = !deselected[g];
        label.className = cb.checked ? "checked" : "";
        cb.addEventListener("change", function () {
          if (cb.checked) delete deselected[g]; else deselected[g] = true;
          saveDeselected();
          renderAll();
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(g));
        els.groupFilter.appendChild(label);
      });
    }

    function renderTable() {
      var groups = selectedGroups();
      var months = allMonths();
      var rowsData = buildRows(groups, months);

      var head = "<th>召會 / 區域</th><th>名單人數</th>";
      months.forEach(function (mk) { head += "<th>" + monthLabel(mk) + "</th>"; });
      head += "<th></th>";
      els.theadRow.innerHTML = head;

      els.tbody.innerHTML = "";
      els.resultsEmpty.style.display = rowsData.length ? "none" : "block";
      els.downloadBtn.disabled = !rowsData.length;

      rowsData.forEach(function (r) {
        var tr = document.createElement("tr");
        var html;
        if (r.type === "region") {
          tr.className = "region-row";
          html = "<td>" + escapeHtml(r.region) + "</td><td>" + r.data.roster + "</td>";
          months.forEach(function (mk) { html += "<td>" + cell(r.data.months[mk]) + "</td>"; });
          html += "<td></td>";
        } else {
          html = "<td class=\"congregation-name\">" + escapeHtml(r.name) + "</td><td>" + r.data.roster + "</td>";
          months.forEach(function (mk) { html += "<td>" + cell(r.data.months[mk]) + "</td>"; });
          html += "<td><button class=\"btn-danger-ghost\" data-rollcall-remove=\"" + escapeHtml(r.name) + "\">移除</button></td>";
        }
        tr.innerHTML = html;
        els.tbody.appendChild(tr);
      });

      els.tbody.querySelectorAll("[data-rollcall-remove]").forEach(function (btn) {
        btn.addEventListener("click", function () { removeCongregation(btn.getAttribute("data-rollcall-remove")); });
      });
    }

    function renderSummary() {
      var groups = selectedGroups();
      var months = allMonths();
      var total = sumRows(visibleNames().map(function (n) { return extract(rollcallState[n], groups, months); }), months);

      var html = "";
      months.forEach(function (mk) {
        html += '<div class="stat-tile"><div class="stat-label">' + monthLabel(mk) + " 出席≥1次</div>" +
          '<div class="stat-value">' + cell(total.months[mk]) + "</div></div>";
      });
      html += '<div class="stat-tile"><div class="stat-label">名單人數</div><div class="stat-value">' + total.roster + "</div></div>";
      els.summaryGrid.innerHTML = html;

      var missing = CONGREGATIONS.filter(function (n) { return !has(n); });
      if (missing.length) {
        els.summaryWarn.style.display = "block";
        els.summaryWarn.textContent = "尚有 " + missing.length + " 個召會未上傳點名報表（" + missing.join("、") + "），以上總計僅計入已上傳的召會。";
      } else {
        els.summaryWarn.style.display = "none";
      }
    }

    function renderAll() {
      renderChecklist();
      renderGroupFilter();
      renderTable();
      renderSummary();
    }

    function removeCongregation(name) {
      if (!window.confirm("確定要移除「" + name + "」的點名資料嗎？")) return;
      delete rollcallState[name];
      renderAll();
      trackWrite(apiDeleteRollcall(ROOM_ID, name)).catch(function (err) {
        showStatus("移除失敗：" + err.message, "error");
      });
    }

    function onParseClick() {
      els.statusMsg.className = "status-msg";
      var congregation = els.select.value;
      if (!congregation) { showStatus("請先到設定新增召會", "error"); return; }
      if (!pickedFile) { showStatus("請先選擇點名報表檔案", "error"); return; }
      els.parseBtn.disabled = true;

      pickedFile.arrayBuffer().then(function (buf) {
        var parsed = parseRollcall(readWorkbook(buf));

        if (has(congregation) && !window.confirm("「" + congregation + "」已經上傳過點名報表，是否要用這個新檔案覆蓋？")) {
          els.parseBtn.disabled = false;
          return;
        }

        var entry = {
          months: parsed.months,
          roster: parsed.roster,
          title: parsed.title,
          fileName: pickedFile.name,
          updatedAt: new Date().toISOString()
        };
        rollcallState[congregation] = entry;
        renderAll();

        var summary = Object.keys(parsed.months).sort().map(function (mk) {
          var n = 0;
          Object.keys(parsed.months[mk].attended).forEach(function (g) { n += parsed.months[mk].attended[g]; });
          return monthLabel(mk) + " " + n + " 人";
        }).join("、");
        showStatus("已加入「" + congregation + "」：名單 " + parsed.people + " 人；出席≥1次（全部群組）— " + summary + "（儲存中…）", "ok");

        pickedFile = null;
        els.fileInput.value = "";
        els.filePicked.textContent = "";
        els.parseBtn.disabled = false;

        trackWrite(apiUpsertRollcall(ROOM_ID, congregation, entry)).then(function () {
          showStatus("已加入「" + congregation + "」：名單 " + parsed.people + " 人；出席≥1次（全部群組）— " + summary + "。已儲存到房間。", "ok");
        }).catch(function (err) {
          showStatus("已加入「" + congregation + "」，但儲存失敗：" + err.message, "error");
        });
      }).catch(function (err) {
        showStatus("解析失敗：" + err.message, "error");
        els.parseBtn.disabled = false;
      });
    }

    function onResetClick() {
      if (!Object.keys(rollcallState).length) return;
      if (!window.confirm("確定要清除所有召會的點名資料嗎？（兒童、青職資料不受影響，此操作無法復原）")) return;
      rollcallState = {};
      renderAll();
      trackWrite(apiResetRollcall(ROOM_ID)).catch(function (err) {
        showStatus("清除失敗：" + err.message, "error");
      });
    }

    function onDownloadClick() {
      var groups = selectedGroups();
      var months = allMonths();
      var rowsData = buildRows(groups, months);
      if (!rowsData.length) return;
      var aoa = [["群組：" + (groups.length ? groups.join("、") : "（未選擇）")]];
      aoa.push(["召會 / 區域", "名單人數"].concat(months.map(function (mk) { return monthLabel(mk) + " 出席≥1次"; })));
      rowsData.forEach(function (r) {
        var row = [r.type === "region" ? r.region : r.name, r.data.roster];
        months.forEach(function (mk) { row.push(r.data.months[mk] == null ? "" : r.data.months[mk]); });
        aoa.push(row);
      });
      var wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), "點名出席統計");
      XLSX.writeFile(wb, "rollcall-stats.xlsx");
    }

    function setAll(on) {
      allGroups().forEach(function (g) { if (on) delete deselected[g]; else deselected[g] = true; });
      saveDeselected();
      renderAll();
    }

    initSelect();
    setupDropzone();
    els.parseBtn.addEventListener("click", onParseClick);
    els.resetBtn.addEventListener("click", onResetClick);
    els.downloadBtn.addEventListener("click", onDownloadClick);
    els.groupAllBtn.addEventListener("click", function () { setAll(true); });
    els.groupNoneBtn.addEventListener("click", function () { setAll(false); });
    renderAll();

    return {
      initSelect: initSelect,
      renderAll: renderAll,
      hasPendingWrites: function () { return pendingWrites > 0; }
    };
  }

  function initTabs() {
    var buttons = document.querySelectorAll("[data-tab-target]");
    buttons.forEach(function (btn) {
      btn.addEventListener("click", function () {
        var target = btn.getAttribute("data-tab-target");
        buttons.forEach(function (b) { b.classList.toggle("active", b === btn); });
        document.querySelectorAll("[data-tab-panel]").forEach(function (panel) {
          panel.style.display = panel.getAttribute("data-tab-panel") === target ? "" : "none";
        });
        var weeklyCard = document.getElementById("weekly-upload-card");
        if (weeklyCard) weeklyCard.style.display = target === "rollcall" ? "none" : "";
      });
    });
  }

  function groupsToText(groups) {
    return groups.map(function (g) {
      return g.region + "\n" + g.members.map(function (m) { return "  " + m; }).join("\n");
    }).join("\n");
  }

  function textToGroups(text) {
    var seen = {};
    var groups = [];
    var current = null;

    text.split("\n").forEach(function (line) {
      if (/^\s/.test(line)) {
        var member = line.trim();
        if (!member || seen[member]) return;
        if (!current) {
          current = { region: DEFAULT_REGION_NAME, members: [] };
          groups.push(current);
        }
        seen[member] = true;
        current.members.push(member);
      } else {
        var region = line.trim();
        if (!region) return;
        current = { region: region, members: [] };
        groups.push(current);
      }
    });

    return groups.filter(function (g) { return g.members.length > 0; });
  }

  function initSettingsModal(app) {
    var openBtn = document.getElementById("open-settings-btn");
    var backdrop = document.getElementById("settings-backdrop");
    var textarea = document.getElementById("settings-textarea");
    var countEl = document.getElementById("settings-count");
    var clearBtn = document.getElementById("settings-clear-btn");
    var defaultBtn = document.getElementById("settings-default-btn");
    var cancelBtn = document.getElementById("settings-cancel-btn");
    var saveBtn = document.getElementById("settings-save-btn");

    function updateCount() {
      var groups = textToGroups(textarea.value);
      var total = groups.reduce(function (acc, g) { return acc + g.members.length; }, 0);
      countEl.textContent = "目前共 " + groups.length + " 個區域、" + total + " 個召會";
    }

    function open() {
      textarea.value = groupsToText(CONGREGATION_GROUPS);
      updateCount();
      backdrop.style.display = "flex";
    }

    function close() {
      backdrop.style.display = "none";
    }

    openBtn.style.display = "";
    openBtn.addEventListener("click", open);
    cancelBtn.addEventListener("click", close);
    backdrop.addEventListener("click", function (e) {
      if (e.target === backdrop) close();
    });
    textarea.addEventListener("input", updateCount);

    clearBtn.addEventListener("click", function () {
      textarea.value = "";
      updateCount();
    });

    defaultBtn.addEventListener("click", function () {
      textarea.value = groupsToText(DEFAULT_GROUPS);
      updateCount();
    });

    saveBtn.addEventListener("click", function () {
      saveBtn.disabled = true;
      var groups = textToGroups(textarea.value);
      var normalized = normalizeGroups(groups) || [];
      apiSaveGroups(ROOM_ID, normalized).then(function () {
        CONGREGATION_GROUPS = normalized;
        CONGREGATIONS = flattenGroups(CONGREGATION_GROUPS);
        saveBtn.disabled = false;
        close();
        app.initSelect();
        app.renderAll();
      }).catch(function (err) {
        saveBtn.disabled = false;
        window.alert("儲存失敗：" + err.message);
      });
    });
  }

  var POLL_INTERVAL_MS = 8000;

  function initLandingAndRoom() {
    var landing = document.getElementById("landing-screen");
    var roomBar = document.getElementById("room-bar");
    var appMain = document.getElementById("app-main");
    var createBtn = document.getElementById("create-room-btn");
    var joinInput = document.getElementById("join-room-input");
    var joinBtn = document.getElementById("join-room-btn");
    var landingStatus = document.getElementById("landing-status-msg");
    var roomCodeDisplay = document.getElementById("room-code-display");
    var roomNameDisplay = document.getElementById("room-name-display");
    var renameBtn = document.getElementById("rename-room-btn");
    var copyCodeBtn = document.getElementById("copy-room-code-btn");
    var refreshBtn = document.getElementById("refresh-room-btn");
    var leaveBtn = document.getElementById("leave-room-btn");

    var roomName = "";
    var app = null;
    var pollTimer = null;

    function showLandingStatus(message, type) {
      landingStatus.textContent = message;
      landingStatus.className = "status-msg show " + type;
    }

    function renderRoomName() {
      if (roomName) {
        roomNameDisplay.textContent = roomName;
        roomNameDisplay.classList.remove("unnamed");
      } else {
        roomNameDisplay.textContent = "（尚未命名）";
        roomNameDisplay.classList.add("unnamed");
      }
    }

    createBtn.addEventListener("click", function () {
      createBtn.disabled = true;
      apiCreateRoom().then(function (data) {
        goToRoom(data.id);
      }).catch(function (err) {
        createBtn.disabled = false;
        showLandingStatus("建立房間失敗：" + err.message, "error");
      });
    });

    function doJoin() {
      var code = joinInput.value.trim().toUpperCase();
      if (!code) {
        showLandingStatus("請輸入房間代碼", "error");
        return;
      }
      goToRoom(code);
    }

    joinBtn.addEventListener("click", doJoin);
    joinInput.addEventListener("keydown", function (e) {
      if (e.key === "Enter") doJoin();
    });

    leaveBtn.addEventListener("click", function () {
      var url = new URL(window.location.href);
      url.search = "";
      window.location.href = url.toString();
    });

    copyCodeBtn.addEventListener("click", function () {
      var restoreText = copyCodeBtn.textContent;
      var flash = function (text) {
        copyCodeBtn.textContent = text;
        setTimeout(function () { copyCodeBtn.textContent = restoreText; }, 1600);
      };
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(ROOM_ID).then(function () {
          flash("已複製！");
        }).catch(function () {
          window.prompt("複製這個房間代碼：", ROOM_ID);
        });
      } else {
        window.prompt("複製這個房間代碼：", ROOM_ID);
      }
    });

    renameBtn.addEventListener("click", function () {
      var next = window.prompt("幫這個房間取個名字（例如「雲嘉區兒童統計」）：", roomName);
      if (next == null) return;
      next = next.trim();
      renameBtn.disabled = true;
      apiSetRoomName(ROOM_ID, next).then(function () {
        roomName = next;
        renderRoomName();
        renameBtn.disabled = false;
      }).catch(function (err) {
        renameBtn.disabled = false;
        window.alert("命名失敗：" + err.message);
      });
    });

    // Applies a freshly-fetched room snapshot to the in-memory model and
    // re-renders. Used for the initial load, the manual refresh button, and
    // background polling alike, so teammates' uploads/settings changes show
    // up here without everyone having to remember to click refresh.
    function applyRoomData(data) {
      CONGREGATION_GROUPS = normalizeGroups(data.groups) || CONGREGATION_GROUPS;
      CONGREGATIONS = flattenGroups(CONGREGATION_GROUPS);
      dataState = data.stats && typeof data.stats === "object" ? data.stats : {};
      rollcallState = data.rollcall && typeof data.rollcall === "object" ? data.rollcall : {};
      roomName = data.name || "";
      renderRoomName();

      if (!app) {
        var weekly = initUploadAndSections();
        var rollcall = initRollcallSection();
        app = {
          initSelect: function () { weekly.initSelect(); rollcall.initSelect(); },
          renderAll: function () { weekly.renderAll(); rollcall.renderAll(); },
          hasPendingWrites: function () { return weekly.hasPendingWrites() || rollcall.hasPendingWrites(); }
        };
        initSettingsModal(app);
      } else {
        app.initSelect();
        app.renderAll();
      }
    }

    function startPolling() {
      if (pollTimer) return;
      pollTimer = setInterval(function () {
        if (app && app.hasPendingWrites && app.hasPendingWrites()) return;
        apiFetchRoom(ROOM_ID).then(function (fresh) {
          CONGREGATION_GROUPS = normalizeGroups(fresh.groups) || CONGREGATION_GROUPS;
          CONGREGATIONS = flattenGroups(CONGREGATION_GROUPS);
          dataState = fresh.stats && typeof fresh.stats === "object" ? fresh.stats : {};
          rollcallState = fresh.rollcall && typeof fresh.rollcall === "object" ? fresh.rollcall : {};
          roomName = fresh.name || "";
          renderRoomName();
          if (app) app.renderAll();
        }).catch(function () {
          /* transient network hiccup - just try again next tick */
        });
      }, POLL_INTERVAL_MS);
    }

    refreshBtn.addEventListener("click", function () {
      refreshBtn.disabled = true;
      apiFetchRoom(ROOM_ID).then(function (fresh) {
        applyRoomData(fresh);
        refreshBtn.disabled = false;
      }).catch(function (err) {
        refreshBtn.disabled = false;
        showLandingStatus("重新整理失敗：" + err.message, "error");
      });
    });

    function boot() {
      ROOM_ID = getRoomIdFromUrl();
      if (!ROOM_ID) {
        landing.style.display = "";
        return;
      }

      landing.style.display = "";
      showLandingStatus("正在連線到房間 " + ROOM_ID + " …", "ok");
      createBtn.disabled = true;
      joinBtn.disabled = true;

      apiFetchRoom(ROOM_ID).then(function (data) {
        landing.style.display = "none";
        roomBar.style.display = "flex";
        appMain.style.display = "";
        roomCodeDisplay.textContent = ROOM_ID;

        applyRoomData(data);
        startPolling();
      }).catch(function (err) {
        landing.style.display = "";
        roomBar.style.display = "none";
        appMain.style.display = "none";
        createBtn.disabled = false;
        joinBtn.disabled = false;
        if (err.notFound) {
          showLandingStatus("找不到房間代碼「" + ROOM_ID + "」，請確認代碼或建立新房間。", "error");
        } else {
          showLandingStatus("連線失敗：" + err.message, "error");
        }
      });
    }

    boot();
  }

  function init() {
    initTabs();
    initLandingAndRoom();

    if ("serviceWorker" in navigator) {
      window.addEventListener("load", function () {
        navigator.serviceWorker.register("sw.js").catch(function () {
          /* offline support unavailable, app still works online */
        });
      });
    }
  }

  init();
})();
