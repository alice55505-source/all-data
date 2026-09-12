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

  function getRoomIdFromUrl() {
    var v = new URLSearchParams(window.location.search).get("room");
    return v ? v.trim().toUpperCase() : "";
  }

  function goToRoom(id) {
    var url = new URL(window.location.href);
    url.search = "?room=" + encodeURIComponent(id);
    window.location.href = url.toString();
  }

  // Remembers rooms this browser has actually connected to, so someone who
  // forgot to copy down the room code (closed the tab, lost the link) can
  // still find their way back instead of losing access to their data.
  var ROOM_HISTORY_KEY = "roomHistory";
  var ROOM_HISTORY_MAX = 10;

  function loadRoomHistory() {
    try {
      var list = JSON.parse(window.localStorage.getItem(ROOM_HISTORY_KEY) || "[]");
      return Array.isArray(list) ? list : [];
    } catch (e) {
      return [];
    }
  }

  function saveRoomToHistory(id, name) {
    try {
      var list = loadRoomHistory().filter(function (r) { return r.id !== id; });
      list.unshift({ id: id, name: name || "" });
      window.localStorage.setItem(ROOM_HISTORY_KEY, JSON.stringify(list.slice(0, ROOM_HISTORY_MAX)));
    } catch (e) {
      /* localStorage unavailable (private mode, quota, etc.) - not critical */
    }
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

  function initTabs() {
    var buttons = document.querySelectorAll("[data-tab-target]");
    buttons.forEach(function (btn) {
      btn.addEventListener("click", function () {
        var target = btn.getAttribute("data-tab-target");
        buttons.forEach(function (b) { b.classList.toggle("active", b === btn); });
        document.querySelectorAll("[data-tab-panel]").forEach(function (panel) {
          panel.style.display = panel.getAttribute("data-tab-panel") === target ? "" : "none";
        });
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
    var roomHistoryEl = document.getElementById("room-history");

    var roomName = "";
    var app = null;
    var pollTimer = null;

    function showLandingStatus(message, type) {
      landingStatus.textContent = message;
      landingStatus.className = "status-msg show " + type;
    }

    function renderRoomHistory() {
      if (!roomHistoryEl) return;
      var list = loadRoomHistory().filter(function (r) { return r.id !== ROOM_ID; });
      roomHistoryEl.innerHTML = "";
      if (!list.length) {
        roomHistoryEl.style.display = "none";
        return;
      }
      roomHistoryEl.style.display = "";
      var label = document.createElement("div");
      label.className = "room-history-label";
      label.textContent = "最近使用過的房間：";
      roomHistoryEl.appendChild(label);
      list.forEach(function (r) {
        var btn = document.createElement("button");
        btn.type = "button";
        btn.className = "btn-secondary room-history-item";
        btn.textContent = r.name ? (r.name + "（" + r.id + "）") : r.id;
        btn.addEventListener("click", function () { goToRoom(r.id); });
        roomHistoryEl.appendChild(btn);
      });
    }

    function renderRoomName() {
      if (roomName) {
        roomNameDisplay.textContent = roomName;
        roomNameDisplay.classList.remove("unnamed");
      } else {
        roomNameDisplay.textContent = "（尚未命名）";
        roomNameDisplay.classList.add("unnamed");
      }
      if (ROOM_ID) saveRoomToHistory(ROOM_ID, roomName);
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
      roomName = data.name || "";
      renderRoomName();

      if (!app) {
        app = initUploadAndSections();
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
        renderRoomHistory();
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
        renderRoomHistory();
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
