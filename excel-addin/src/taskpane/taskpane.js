/* Excel AI Copilot -- Excel Add-in task pane logic.
 *
 * Reads a snapshot of the active worksheet through Office.js, sends the
 * natural-language command plus the snapshot to the Flask backend and
 * applies the returned writes back to the open workbook.
 */

import { buildSnapshot, isEmptyValues } from "./snapshot.mjs";

/* ---------------------------------------------------------------------- */
/* Configuration                                                          */
/* ---------------------------------------------------------------------- */

const BACKEND_URL = "http://127.0.0.1:5000";
const API_BASE = `${BACKEND_URL}/api/excel`;

// Largest snapshot we will send. Keeps payloads small; larger sheets are
// trimmed and the user is told the context is partial.
const MAX_ROWS = 3000;
const MAX_COLS = 100;

/* ---------------------------------------------------------------------- */
/* State                                                                  */
/* ---------------------------------------------------------------------- */

const state = {
  ready: false,
  busy: false,
  sheetName: null,
  used: null, // { address, values, rows, cols, originRow, originCol }
  usedSource: null, // worksheet | region | selection
  selection: null, // { address, values, row, column, rowCount, columnCount }
  truncated: false,
  health: "unknown", // unknown | online | offline
};

const $ = (id) => document.getElementById(id);
const messagesEl = $("messages");
const inputEl = $("commandInput");
const sendBtn = $("sendBtn");
const typingEl = $("typing");

/* ---------------------------------------------------------------------- */
/* Office.js lifecycle                                                    */
/* ---------------------------------------------------------------------- */

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    showError("This add-in only works inside Microsoft Excel.");
    return;
  }
  state.ready = true;
  sendBtn.disabled = false;
  addMessage(
    "bot",
    "Connected to Excel. Select some data and type a command like " +
      '"Revenue ka total karo".'
  );
  startup();
});

async function startup() {
  try {
    await refreshContext();
  } catch (err) {
    showError("Could not read the current workbook: " + friendlyError(err));
  }
  checkBackend();
}

/* ---------------------------------------------------------------------- */
/* Backend health                                                         */
/* ---------------------------------------------------------------------- */

async function checkBackend() {
  const el = $("backendStatus");
  try {
    const res = await fetch(`${API_BASE}/health`);
    if (!res.ok) throw new Error("health check failed");
    const data = await res.json();
    state.health = data.status === "ok" ? "online" : "offline";
  } catch {
    state.health = "offline";
  }
  el.textContent = state.health === "online" ? "backend online" : "backend offline";
  el.className = "backend-status " + state.health;
  if (state.health !== "online") {
    addMessage(
      "error",
      "Cannot reach the Excel AI Copilot backend at " +
        BACKEND_URL +
        ".\nStart it with:\n\n  python main.py --serve --port 5000\n\n" +
        "then run a command again."
    );
  }
}

/* ---------------------------------------------------------------------- */
/* Reading the workbook context (Office.js)                               */
/* ---------------------------------------------------------------------- */

async function refreshContext() {
  let info;
  await Excel.run(async (context) => {
    context.workbook.worksheets.load("items/name");
    const ws = context.workbook.worksheets.getActiveWorksheet();
    ws.load("name");

    // ---- used range snapshot -----------------------------------------
    // Strategy (each step is a cross-check against real cell contents):
    //   1. getUsedRangeOrNullObject() when it genuinely contains data.
    //   2. a fixed A1-window scan, when the used range reports "empty" so we
    //      never trust a null-object result blindly.
    //   3. finally, the user's live selection when both reads came up empty
    //      but the selection really has content.
    const usedRange = ws.getUsedRangeOrNullObject();
    usedRange.load([
      "address",
      "values",
      "rowCount",
      "columnCount",
      "isNullObject",
      "rowIndex",
      "columnIndex",
    ]);

    const selection = context.workbook.getSelectedRange();
    selection.load([
      "address",
      "values",
      "rowCount",
      "columnCount",
      "rowIndex",
      "columnIndex",
    ]);

    await context.sync();

    let used = usedRangeToSnapshot(usedRange);
    let usedSource = used ? "worksheet" : null;
    let truncated = Boolean(used && used.truncated);

    // Fallback #1: the used range appeared to be null/empty. Scan an
    // absolute window of the sheet so we read the real cells rather than a
    // possibly-wrong null object.
    if (!used) {
      const region = ws.getRangeByIndexes(0, 0, MAX_ROWS, MAX_COLS);
      region.load("values");
      await context.sync();
      used = buildSnapshot(region.values, {
        baseRow: 1,
        baseCol: 1,
        maxRows: MAX_ROWS,
        maxCols: MAX_COLS,
      });
      usedSource = used ? "region" : "empty";
      truncated = Boolean(used && used.truncated);
    }

    const sel = selectionToState(selection);

    // Fallback #2: still no worksheet data, but the live selection has
    // content. Operate on the selected cells instead.
    if (!used && sel && !isEmptyValues(sel.values)) {
      used = {
        values: sel.values,
        rows: sel.rowCount || sel.values.length,
        cols: sel.columnCount || (sel.values[0] ? sel.values[0].length : 0),
        originRow: sel.row,
        originCol: sel.column,
        truncated: truncated || sel.rowCount > MAX_ROWS || sel.columnCount > MAX_COLS,
      };
      usedSource = "selection";
    }

    info = {
      sheetName: ws.name,
      used,
      usedSource,
      truncated,
      selection: sel,
    };
  });

  state.sheetName = info.sheetName;
  state.used = info.used;
  state.usedSource = info.usedSource;
  state.truncated = info.truncated;
  state.selection = info.selection;

  renderContext();
  return info;
}

/**
 * Extract a snapshot from a loaded used-range proxy. Returns null when the
 * range is a null object or contains no actual data.
 */
function usedRangeToSnapshot(usedRange) {
  if (!usedRange || usedRange.isNullObject) return null;
  if (!Array.isArray(usedRange.values)) return null;
  return buildSnapshot(usedRange.values, {
    baseRow: (usedRange.rowIndex === undefined ? 0 : usedRange.rowIndex) + 1,
    baseCol: (usedRange.columnIndex === undefined ? 0 : usedRange.columnIndex) + 1,
    maxRows: MAX_ROWS,
    maxCols: MAX_COLS,
  });
}

function selectionToState(selection) {
  if (!selection || !Array.isArray(selection.values) || !selection.values.length) {
    return null;
  }
  const parsed = parseAddress(
    String(selection.address || "").split(",")[0].replace(/^.*!/, "")
  );
  return {
    address: selection.address,
    values: selection.values.slice(0, MAX_ROWS).map((r) => r.slice(0, MAX_COLS)),
    row: parsed ? parsed.row : (selection.rowIndex === undefined ? 0 : selection.rowIndex) + 1,
    column: parsed
      ? parsed.col
      : (selection.columnIndex === undefined ? 0 : selection.columnIndex) + 1,
    rowCount: selection.rowCount,
    columnCount: selection.columnCount,
  };
}

function renderContext() {
  $("ctxSheet").textContent = state.sheetName || "—";
  $("ctxSelection").textContent = state.selection
    ? `${state.selection.address}${
        state.selection.values && state.selection.values[0]
          ? " (" + describeSelection(state.selection) + ")"
          : ""
      }`
    : "—";
}

function describeSelection(sel) {
  if (!sel || !sel.values) return "empty";
  const flat = sel.values.flat();
  const nums = flat.filter(
    (v) => typeof v === "number" || (typeof v === "string" && v.trim() !== "" && !isNaN(v))
  );
  return nums.length ? `${nums.length} number(s)` : "text only";
}

/* ---------------------------------------------------------------------- */
/* Command pipeline                                                       */
/* ---------------------------------------------------------------------- */

function buildPayload(cmd, allowOverwrite) {
  const used = state.used;
  let headers = null;
  let dataRows = [];

  if (used && used.values && used.values.length) {
    const first = used.values[0];
    const hasHeader = first.some(isTextLike);
    if (hasHeader) {
      headers = first;
      dataRows = used.values.slice(1);
    } else {
      dataRows = used.values;
    }
  }

  const sel = state.selection;
  return {
    command: cmd,
    sheet_name: state.sheetName || "Sheet1",
    origin: used
      ? { row: used.originRow, column: used.originCol }
      : { row: 1, column: 1 },
    headers: headers || [],
    rows: dataRows,
    selection:
      sel && sel.values && sel.values.length
        ? {
            address: sel.address,
            values: sel.values,
            row: sel.row,
            column: sel.column,
            row_count: sel.rowCount,
            column_count: sel.columnCount,
          }
        : null,
    allow_overwrite: !!allowOverwrite,
  };
}

async function postCommand(payload) {
  let res;
  try {
    res = await fetch(`${API_BASE}/command`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  } catch (err) {
    throw new Error(
      "Cannot reach the Excel AI Copilot backend at " + BACKEND_URL + "."
    );
  }
  let data;
  try {
    data = await res.json();
  } catch {
    throw new Error("The backend returned an unreadable response (HTTP " + res.status + ").");
  }
  return data;
}

async function handleSubmit(cmd) {
  if (state.busy || !state.ready) return;

  state.busy = true;
  setBusy(true);
  addMessage("user", cmd);
  inputEl.value = "";
  inputEl.placeholder = 'Try: "Revenue ka total karo"';

  try {
    // LIVE RE-READ: the snapshot captured at pane startup (or after the last
    // command) can be stale — the worksheet may have been empty when the pane
    // loaded, or the user may have edited cells since. Every command must run
    // against the CURRENT worksheet, so refresh the context before the
    // emptiness check. This prevents a pane that loaded on an empty sheet
    // from permanently rejecting commands with "worksheet is empty" even
    // after the user adds data.
    await refreshContext();

    if (!state.used || !state.used.values || !state.used.values.length) {
      showError(
        "The active worksheet is empty — Excel reported no data in the " +
          "used range, the sheet scan, or your selection. Add some data first."
      );
      return;
    }
    if (state.truncated) {
      addMessage(
        "bot",
        "Note: the sheet is larger than the snapshot limit (" +
          MAX_ROWS +
          " rows x " +
          MAX_COLS +
          " cols), so results may be based on the visible portion."
      );
    }

    const payload = buildPayload(cmd, false);
    const data = await postCommand(payload);
    await handleResponse(cmd, data, false);
  } catch (err) {
    showError(friendlyError(err));
  } finally {
    state.busy = false;
    setBusy(false);
  }
}

async function handleResponse(cmd, data, overwrote) {
  if (data.success) {
    addMessage("bot", data.message);

    const writes = data.writes || { cells: [], sheets: {}, charts: [] };

    if (data.overwrite_needed && !overwrote) {
      showConfirm(data.message, async () => {
        const retry = await postCommand(buildPayload(cmd, true));
        await handleResponse(cmd, retry, true);
      });
      return;
    }

    const blocked = await findBlockedWrites(writes.cells);
    if (blocked && blocked.length) {
      showConfirm(
        `${blocked
          .map((b) => `${b.cell} already contains "${b.current}"`)
          .join(", ")}.\nReplace it?`,
        async () => {
          await applyWrites(writes);
          addMessage("bot", "Applied the changes to the worksheet.");
          await finish();
        }
      );
      return;
    }

    await applyWrites(writes);
    if (hasWrites(writes)) {
      addMessage("bot", "Applied the changes to the worksheet.");
    }

    if (data.selection_auto_write) {
      await handleAutoWrite(data.selection_auto_write);
    }

    await finish();
    return;
  }

  if (data.overwrite_needed && !overwrote) {
    showConfirm(data.message, async () => {
      const retry = await postCommand(buildPayload(cmd, true));
      await handleResponse(cmd, retry, true);
    });
    return;
  }

  showError(data.message || "Something went wrong. Try again.");
}

async function finish() {
  try {
    await refreshContext();
  } catch (err) {
    // Non-fatal: the next command will retry the context read.
    console.warn(err);
  }
}

/* ---------------------------------------------------------------------- */
/* Applying writes back to Excel                                          */
/* ---------------------------------------------------------------------- */

function hasWrites(writes) {
  return Boolean(
    (writes.cells && writes.cells.length) ||
      (writes.sheets && Object.keys(writes.sheets).length) ||
      (writes.charts && writes.charts.length)
  );
}

function isInKnownRegion(row, col) {
  if (!state.used) return false;
  const { originRow, originCol, rows, cols } = state.used;
  return row >= originRow && row < originRow + rows && col >= originCol && col < originCol + cols;
}

async function findBlockedWrites(cells) {
  const external = (cells || []).filter(
    (c) =>
      !isInKnownRegion(c.row, c.column) && (c.value !== undefined || c.formula !== undefined)
  );
  if (!external.length) return [];

  return Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getActiveWorksheet();
    const ranges = external.map((c) => ws.getRange(cellRef(c.row, c.column)));
    ranges.forEach((r) => r.load("values"));
    await context.sync();
    const blocked = [];
    external.forEach((c, i) => {
      const current = ranges[i].values[0][0];
      if (current !== null && current !== undefined && current !== "") {
        blocked.push({ cell: cellRef(c.row, c.column), current });
      }
    });
    return blocked;
  });
}

async function applyWrites(writes) {
  if (!hasWrites(writes)) return;
  await Excel.run(async (context) => {
    const ws = context.workbook.worksheets.getActiveWorksheet();

    for (const c of writes.cells || []) {
      const rng = ws.getRange(cellRef(c.row, c.column));
      if (c.clear) {
        rng.clear();
      } else if (c.formula !== undefined) {
        rng.formulas = [[c.formula]];
      } else if (c.value !== undefined) {
        rng.values = [[c.value]];
      }
    }

    for (const [name, def] of Object.entries(writes.sheets || {})) {
      const values = def.values || [];
      const maxCols = Math.max(0, ...values.map((r) => (r ? r.length : 0)));
      if (!values.length || !maxCols) continue;

      const existing = context.workbook.worksheets.getItemOrNullObject(name);
      existing.load("name");
      await context.sync();

      let target;
      if (existing.isNullObject) {
        target = context.workbook.worksheets.add(name);
      } else {
        target = existing;
        target.getUsedRange().clear("All");
      }
      target.getRangeByIndexes(0, 0, values.length, maxCols).values = values;
    }

    for (const chartInfo of writes.charts || []) {
      applyChart(context, ws, chartInfo);
    }

    await context.sync();
  });
}

function applyChart(context, ws, chartInfo) {
  const type = chartInfo.type === "LineChart" ? "Line" : "ColumnClustered";
  const source = ws.getRange(chartInfo.source_range);

  const chart = ws.charts.add(type, source, "Auto");
  chart.title.text = chartInfo.title;
  chart.title.visible = true;

  const parsed = parseAddress(chartInfo.source_range);
  if (parsed) {
    chart.setPosition(cellRef(parsed.row + 2, parsed.col + 2));
  }
}

async function handleAutoWrite(autoWrite) {
  const { row, column, value } = autoWrite;
  const addr = cellRef(row, column);
  let current;
  try {
    current = await readCell(row, column);
  } catch (err) {
    showError("Could not check the destination cell: " + friendlyError(err));
    return;
  }
  const write = async () => {
    try {
      await writeValue(row, column, value);
      addMessage("bot", `Wrote ${formatValue(value)} into ${addr}.`);
    } catch (err) {
      showError("Could not write to " + addr + ": " + friendlyError(err));
    }
  };
  if (isEmptyValue(current)) {
    await write();
  } else {
    showConfirm(
      `${addr} already contains "${current}". Replace it with ${formatValue(value)}?`,
      write
    );
  }
}

async function readCell(row, column) {
  return Excel.run(async (context) => {
    const rng = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(cellRef(row, column));
    rng.load("values");
    await context.sync();
    return rng.values[0][0];
  });
}

async function writeValue(row, column, value) {
  return Excel.run(async (context) => {
    const rng = context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(cellRef(row, column));
    rng.values = [[value]];
    await context.sync();
  });
}

/* ---------------------------------------------------------------------- */
/* UI helpers                                                             */
/* ---------------------------------------------------------------------- */

function addMessage(kind, text, html = false) {
  const div = document.createElement("div");
  div.className = "msg " + (kind === "error" ? "error" : kind);
  const content = document.createElement("div");
  if (html) {
    content.innerHTML = text;
  } else {
    content.textContent = text;
  }
  div.appendChild(content);
  messagesEl.appendChild(div);
  messagesEl.scrollTop = messagesEl.scrollHeight;
  return div;
}

function showError(text) {
  addMessage("error", text);
}

function setBusy(isBusy) {
  sendBtn.disabled = !state.ready || isBusy || state.busy;
  sendBtn.classList.toggle("loading", isBusy);
  typingEl.hidden = !isBusy;
  if (isBusy) {
    $("dotA").style.animationDelay = "0s";
    $("dotB").style.animationDelay = "-0.16s";
    $("dotC").style.animationDelay = "-0.32s";
  }
  inputEl.disabled = isBusy;
}

function showConfirm(message, onYes) {
  const div = addMessage("bot", message);
  const row = document.createElement("div");
  row.className = "confirm-row";

  const yesBtn = document.createElement("button");
  yesBtn.type = "button";
  yesBtn.className = "yes";
  yesBtn.textContent = "Yes, replace";
  yesBtn.addEventListener("click", async () => {
    row.remove();
    await onYes();
  });

  const noBtn = document.createElement("button");
  noBtn.type = "button";
  noBtn.className = "no";
  noBtn.textContent = "No, cancel";
  noBtn.addEventListener("click", () => {
    row.remove();
    addMessage("bot", "OK, I didn't change anything.");
  });

  row.appendChild(yesBtn);
  row.appendChild(noBtn);
  div.appendChild(row);
}

function friendlyError(err) {
  if (!err) return "An unexpected error occurred.";
  const msg = String(err.message || err);
  if (/Failed to fetch/i.test(msg)) {
    return (
      "Cannot reach the Excel AI Copilot backend at " +
      BACKEND_URL +
      ". Make sure it is running with 'python main.py --serve --port 5000'."
    );
  }
  return msg;
}

function formatValue(value) {
  if (typeof value === "number") {
    return Number.isInteger(value) ? String(value) : String(value);
  }
  return String(value);
}

/* ---------------------------------------------------------------------- */
/* Coordinate helpers                                                     */
/* ---------------------------------------------------------------------- */

function colLetter(num) {
  let s = "";
  let n = num;
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function cellRef(row, col) {
  return colLetter(col) + row;
}

function parseAddress(address) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(String(address || "").trim());
  if (!m) return null;
  return {
    col: colToNum(m[1]),
    row: parseInt(m[2], 10),
    endCol: m[3] ? colToNum(m[3]) : colToNum(m[1]),
    endRow: m[4] ? parseInt(m[4], 10) : parseInt(m[2], 10),
  };
}

function colToNum(letters) {
  let n = 0;
  for (const ch of letters) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

function isTextLike(v) {
  return typeof v === "string" && v.trim() !== "" && isNaN(Number(v));
}

function isEmptyValue(v) {
  return v === null || v === undefined || v === "";
}

/* ---------------------------------------------------------------------- */
/* Wiring                                                                 */
/* ---------------------------------------------------------------------- */

$("commandForm").addEventListener("submit", (e) => {
  e.preventDefault();
  const cmd = inputEl.value.trim();
  if (!cmd || state.busy) return;
  handleSubmit(cmd);
});

document.querySelectorAll("#examplesList li").forEach((li) => {
  li.addEventListener("click", () => {
    if (state.busy) return;
    inputEl.value = li.textContent.trim();
    inputEl.focus();
  });
});