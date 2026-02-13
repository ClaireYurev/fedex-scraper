"use strict";

const amountsInput = document.getElementById("amounts-input");
const parsedCount = document.getElementById("parsed-count");
const btnStart = document.getElementById("btn-start");
const btnCancel = document.getElementById("btn-cancel");
const progressSection = document.getElementById("progress-section");
const progressBar = document.getElementById("progress-bar");
const progressPct = document.getElementById("progress-pct");
const progressText = document.getElementById("progress-text");
const logList = document.getElementById("log-list");
const resultSection = document.getElementById("result-section");
const resultIcon = document.getElementById("result-icon");
const resultText = document.getElementById("result-text");
const statusDot = document.getElementById("status-dot");

let isRunning = false;

// ---------------------------------------------------------------------------
// Parse pasted amounts into an array of normalized strings like "452.67"
// ---------------------------------------------------------------------------
function parseAmounts(raw) {
  const tokens = raw.split(/[,\n\r;]+/).map((t) => t.trim()).filter(Boolean);
  const amounts = [];
  for (const tok of tokens) {
    const cleaned = tok.replace(/[^0-9.]/g, "");
    if (cleaned && /^\d+(\.\d{1,2})?$/.test(cleaned)) {
      amounts.push(cleaned);
    }
  }
  return [...new Set(amounts)];
}

// Update parsed count badge as user types
amountsInput.addEventListener("input", () => {
  const amounts = parseAmounts(amountsInput.value);
  if (amounts.length > 0) {
    parsedCount.textContent = `${amounts.length} amount${amounts.length > 1 ? "s" : ""}`;
    parsedCount.classList.add("visible");
  } else {
    parsedCount.classList.remove("visible");
    parsedCount.textContent = "";
  }
});

// ---------------------------------------------------------------------------
// Status dot management
// ---------------------------------------------------------------------------
function setStatus(state) {
  statusDot.className = "status-dot";
  statusDot.classList.add("status-" + state);
  const titles = { idle: "Idle", running: "Processing...", done: "Complete", error: "Error" };
  statusDot.title = titles[state] || "";
}

// ---------------------------------------------------------------------------
// Logging helper — dark terminal style with animated entries
// ---------------------------------------------------------------------------
function addLog(message, type = "info") {
  const div = document.createElement("div");
  div.className = "log-entry " + type;

  // Detect separator lines (--- Processing ... ---)
  if (message.startsWith("---") && message.endsWith("---")) {
    div.className = "log-entry separator";
    div.textContent = message.replace(/^-+\s*/, "").replace(/\s*-+$/, "");
  } else {
    div.textContent = message;
  }

  logList.appendChild(div);

  // Auto-scroll
  const container = document.getElementById("log-container");
  requestAnimationFrame(() => {
    container.scrollTop = container.scrollHeight;
  });
}

// ---------------------------------------------------------------------------
// Listen for progress updates from the background service worker
// ---------------------------------------------------------------------------
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "PROGRESS") {
    const pct = Math.round(msg.percent);
    progressBar.style.width = pct + "%";
    progressPct.textContent = pct + "%";
    progressText.textContent = msg.text || `${pct}%`;
  } else if (msg.type === "LOG") {
    addLog(msg.text, msg.level || "info");
  } else if (msg.type === "DONE") {
    isRunning = false;
    btnStart.disabled = false;
    btnCancel.disabled = true;

    resultSection.classList.remove("hidden");

    if (msg.error) {
      setStatus("error");
      resultIcon.className = "result-icon error";
      resultIcon.innerHTML = "&#x2717;";
      resultText.textContent = "Extraction failed: " + msg.error;
    } else {
      setStatus("done");
      resultIcon.className = "result-icon success";
      resultIcon.innerHTML = "&#x2713;";
      const sc = msg.shipmentCount || 0;
      const ic = msg.invoiceCount || 0;
      resultText.textContent =
        `Complete! ${sc} shipment${sc !== 1 ? "s" : ""} across ${ic} invoice${ic !== 1 ? "s" : ""}. File downloaded.`;
    }
  }
});

// ---------------------------------------------------------------------------
// Start button
// ---------------------------------------------------------------------------
btnStart.addEventListener("click", async () => {
  const amounts = parseAmounts(amountsInput.value);
  if (amounts.length === 0) {
    addLog("No valid amounts found. Please paste invoice amounts.", "error");
    return;
  }

  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !tab.url || !tab.url.includes("fedex.com/online/billing")) {
    addLog("Please navigate to fedex.com/online/billing first.", "error");
    return;
  }

  isRunning = true;
  setStatus("running");
  btnStart.disabled = true;
  btnCancel.disabled = false;
  progressSection.classList.remove("hidden");
  resultSection.classList.add("hidden");
  logList.innerHTML = "";
  progressBar.style.width = "0%";
  progressPct.textContent = "0%";
  progressText.textContent = "Starting...";

  addLog(`Starting analysis for ${amounts.length} amount${amounts.length > 1 ? "s" : ""}...`);

  chrome.runtime.sendMessage({
    type: "START_EXTRACTION",
    amounts,
    tabId: tab.id,
  });
});

// ---------------------------------------------------------------------------
// Cancel button
// ---------------------------------------------------------------------------
btnCancel.addEventListener("click", () => {
  chrome.runtime.sendMessage({ type: "CANCEL_EXTRACTION" });
  isRunning = false;
  setStatus("idle");
  btnStart.disabled = false;
  btnCancel.disabled = true;
  progressText.textContent = "Cancelled.";
  addLog("Analysis cancelled by user.", "error");
});
