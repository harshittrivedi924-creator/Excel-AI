const fileInput = document.getElementById("fileInput");
const fileStatus = document.getElementById("fileStatus");
const sheets = document.getElementById("sheets");
const tableWrap = document.getElementById("tableWrap");
const messages = document.getElementById("messages");
const form = document.getElementById("commandForm");
const input = document.getElementById("commandInput");
const history = document.getElementById("history");

let pendingOverwrite = null;

function addMessage(text, who) {
  const div = document.createElement("div");
  div.className = `msg ${who}`;
  div.textContent = text;
  messages.appendChild(div);
  messages.scrollTop = messages.scrollHeight;
}

async function loadWorkbookInfo() {
  const res = await fetch("/api/workbook");
  const data = await res.json();
  if (data.loaded) {
    fileStatus.textContent = `Loaded: ${data.file || "workbook"}`;
    sheets.innerHTML = "";
    for (const name of data.sheets) {
      const chip = document.createElement("span");
      chip.className = "sheet-chip" + (name === data.active_sheet ? " active" : "");
      chip.textContent = name;
      sheets.appendChild(chip);
    }
    renderTable(data.rows);
  }
}

function renderTable(rows) {
  if (!rows || !rows.length) {
    tableWrap.innerHTML = '<p class="hint">No data in this sheet.</p>';
    return;
  }
  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  rows[0].forEach((cell) => {
    const th = document.createElement("th");
    th.textContent = cell;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.slice(1).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  tableWrap.innerHTML = "";
  tableWrap.appendChild(table);
}

async function refreshHistory() {
  const res = await fetch("/api/history");
  const items = await res.json();
  history.innerHTML = "";
  for (const item of items) {
    const li = document.createElement("li");
    li.textContent = `${item.operation}(...) = ${item.result}${item.destination ? " -> " + item.destination : ""}`;
    history.appendChild(li);
  }
}

async function sendCommand(cmd, allowOverwrite = false) {
  const res = await fetch("/api/command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ command: cmd, allow_overwrite }),
  });
  return res.json();
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const cmd = input.value.trim();
  if (!cmd || pendingOverwrite) return;

  addMessage(cmd, "user");
  input.value = "";

  if (cmd.toLowerCase() === "history") {
    refreshHistory();
    return;
  }

  const data = await sendCommand(cmd);
  addMessage(data.message, "bot");

  if (data.overwrite_needed) {
    pendingOverwrite = { cmd };
    addMessage("Replace the existing data? (reply 'yes' or 'no')", "bot");
    input.placeholder = "Type 'yes' or 'no'";
  }

  refreshHistory();
  loadWorkbookInfo();
});

input.addEventListener("keydown", async (e) => {
  if (e.key !== "Enter") return;
  const text = input.value.trim().toLowerCase();
  if (!pendingOverwrite) return;

  e.preventDefault();
  if (text === "yes" || text === "y" || text === "haan") {
    const data = await sendCommand(pendingOverwrite.cmd, true);
    addMessage(data.message, "bot");
  } else {
    addMessage("OK, I didn't change anything.", "bot");
  }
  pendingOverwrite = null;
  input.placeholder = 'Try: "Column D ka total karo aur D21 mein daal do"';
  input.value = "";
  refreshHistory();
  loadWorkbookInfo();
});

fileInput.addEventListener("change", async () => {
  if (!fileInput.files.length) return;
  const fd = new FormData();
  fd.append("file", fileInput.files[0]);
  const res = await fetch("/api/upload", { method: "POST", body: fd });
  const data = await res.json();
  if (data.loaded) {
    addMessage(`Workbook "${data.file}" loaded. Ask away!`, "bot");
    loadWorkbookInfo();
  } else {
    addMessage(data.error || "Couldn't load the workbook.", "bot");
  }
});

loadWorkbookInfo();
refreshHistory();