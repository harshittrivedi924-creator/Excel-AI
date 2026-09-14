const fileInput = document.getElementById("fileInput");
const fileStatus = document.getElementById("fileStatus");
const sheets = document.getElementById("sheets");
const tableWrap = document.getElementById("tableWrap");
const previewHint = document.getElementById("previewHint");
const messages = document.getElementById("messages");
const form = document.getElementById("commandForm");
const input = document.getElementById("commandInput");
const sendBtn = document.getElementById("sendBtn");
const history = document.getElementById("history");
const typingIndicator = document.getElementById("typingIndicator");

let pendingOverwrite = null;
let isLoading = false;

// Add a message to the chat
function addMessage(text, who, isHtml = false) {
  const div = document.createElement("div");
  div.className = `msg ${who}`;

  if (who === "bot") {
    const icon = document.createElement("span");
    icon.className = "msg-icon";
    icon.textContent = "🤖";
    div.appendChild(icon);
  }

  const content = document.createElement("div");
  content.className = "msg-content";
  if (isHtml) {
    content.innerHTML = text;
  } else {
    content.textContent = text;
  }
  div.appendChild(content);

  messages.appendChild(div);
  messages.scrollTop = messages.scrollHeight;
  return div;
}

// Show/hide typing indicator
function setLoading(loading) {
  isLoading = loading;
  sendBtn.disabled = loading;
  if (loading) {
    sendBtn.classList.add("loading");
    typingIndicator.classList.add("active");
  } else {
    sendBtn.classList.remove("loading");
    typingIndicator.classList.remove("active");
  }
}

// Load workbook info from API
async function loadWorkbookInfo() {
  try {
    const res = await fetch("/api/workbook");
    const data = await res.json();

    if (data.loaded) {
      fileStatus.textContent = `Loaded: ${data.file || "workbook"}`;
      fileStatus.style.color = "#22c55e";

      sheets.innerHTML = "";
      for (const name of data.sheets) {
        const chip = document.createElement("span");
        chip.className = "sheet-chip" + (name === data.active_sheet ? " active" : "");
        chip.textContent = name;
        sheets.appendChild(chip);
      }

      renderTable(data.rows);
      previewHint.textContent = `${data.rows.length - 1} rows shown`;
    } else {
      fileStatus.textContent = "No workbook loaded";
      fileStatus.style.color = "";
      sheets.innerHTML = "";
      tableWrap.innerHTML = '<p class="hint">Upload an Excel file to see a preview.</p>';
      previewHint.textContent = "";
    }
  } catch (err) {
    console.error("Failed to load workbook info:", err);
  }
}

// Render data table
function renderTable(rows) {
  if (!rows || !rows.length) {
    tableWrap.innerHTML = '<p class="hint">No data in this sheet.</p>';
    return;
  }

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");

  rows[0].forEach((cell, i) => {
    const th = document.createElement("th");
    th.textContent = cell || `Col ${i + 1}`;
    headRow.appendChild(th);
  });

  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  rows.slice(1).forEach((row) => {
    const tr = document.createElement("tr");
    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell === "" ? "—" : cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(tbody);
  tableWrap.innerHTML = "";
  tableWrap.appendChild(table);
}

// Refresh command history
async function refreshHistory() {
  try {
    const res = await fetch("/api/history");
    const items = await res.json();

    history.innerHTML = "";
    if (items.length === 0) {
      const li = document.createElement("li");
      li.textContent = "No commands yet";
      li.style.fontStyle = "italic";
      history.appendChild(li);
      return;
    }

    for (const item of items.reverse()) {
      const li = document.createElement("li");
      const dest = item.destination ? ` → ${item.destination}` : "";
      li.textContent = `${item.operation}(...) = ${item.result}${dest}`;
      history.appendChild(li);
    }
  } catch (err) {
    console.error("Failed to load history:", err);
  }
}

// Send command to API
async function sendCommand(cmd, allowOverwrite = false) {
  const res = await fetch("/api/command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ command: cmd, allow_overwrite: allowOverwrite }),
  });
  return res.json();
}

// Handle form submission
form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const cmd = input.value.trim();
  if (!cmd || pendingOverwrite || isLoading) return;

  addMessage(cmd, "user");
  input.value = "";
  setLoading(true);

  if (cmd.toLowerCase() === "history") {
    refreshHistory();
    setLoading(false);
    return;
  }

  try {
    const data = await sendCommand(cmd);

    if (data.overwrite_needed) {
      pendingOverwrite = { cmd };
      addMessage(
        "⚠️ " + data.message + "<br><br>Reply <strong>'yes'</strong> to replace or <strong>'no'</strong> to cancel.",
        "bot",
        true
      );
      input.placeholder = "Type 'yes' or 'no'";
    } else if (data.success) {
      addMessage(data.message, "bot");
    } else {
      const errorDiv = addMessage(data.message, "bot error");
    }

    refreshHistory();
    loadWorkbookInfo();
  } catch (err) {
    addMessage("Failed to send command. Please try again.", "bot error");
  } finally {
    setLoading(false);
  }
});

// Handle overwrite confirmation
input.addEventListener("keydown", async (e) => {
  if (e.key !== "Enter" || !pendingOverwrite) return;

  e.preventDefault();
  const text = input.value.trim().toLowerCase();

  if (text === "yes" || text === "y" || text === "haan") {
    setLoading(true);
    try {
      const data = await sendCommand(pendingOverwrite.cmd, true);
      addMessage(data.message, "bot");
      refreshHistory();
      loadWorkbookInfo();
    } catch (err) {
      addMessage("Failed to execute command.", "bot error");
    } finally {
      setLoading(false);
    }
  } else {
    addMessage("OK, I didn't change anything.", "bot");
  }

  pendingOverwrite = null;
  input.placeholder = 'Try: "Column D ka total karo aur D21 mein daal do"';
  input.value = "";
});

// Handle file upload
fileInput.addEventListener("change", async () => {
  if (!fileInput.files.length) return;

  const fd = new FormData();
  fd.append("file", fileInput.files[0]);

  setLoading(true);
  try {
    const res = await fetch("/api/upload", { method: "POST", body: fd });
    const data = await res.json();

    if (data.loaded) {
      addMessage(`Workbook "${data.file}" loaded successfully. Ask me anything!`, "bot");
      loadWorkbookInfo();
    } else {
      addMessage(data.error || "Couldn't load the workbook.", "bot error");
    }
  } catch (err) {
    addMessage("Failed to upload file. Please try again.", "bot error");
  } finally {
    setLoading(false);
  }
});

// Example command click handlers
document.querySelectorAll(".examples li").forEach((li) => {
  li.addEventListener("click", () => {
    if (!isLoading && !pendingOverwrite) {
      input.value = li.textContent;
      input.focus();
    }
  });
});

// Initialize
loadWorkbookInfo();
refreshHistory();
