"use strict";

const amountsInput = document.getElementById("amounts-input");
const parsedCount = document.getElementById("parsed-count");
const btnStart = document.getElementById("btn-start");
const btnCancel = document.getElementById("btn-cancel");
const progressSection = document.getElementById("progress-section");
const progressBar = document.getElementById("progress-bar");
const progressText = document.getElementById("progress-text");
const logList = document.getElementById("log-list");
const resultSection = document.getElementById("result-section");
const resultText = document.getElementById("result-text");

let isRunning = false;

// ---------------------------------------------------------------------------
// Parse pasted amounts into an array of normalized strings like "452.67"
// ---------------------------------------------------------------------------
function parseAmounts(raw) {
  // Split by commas, newlines, semicolons, or whitespace
  const tokens = raw.split(/[,\n\r;]+/).map((t) => t.trim()).filter(Boolean);
  const amounts = [];
  for (const tok of tokens) {
    // Strip dollar sign and whitespace, keep digits and decimal
    const cleaned = tok.replace(/[^0-9.]/g, "");
    if (cleaned && /^\d+(\.\d{1,2})?$/.test(cleaned)) {
      amounts.push(cleaned);
    }
  }
  return [...new Set(amounts)]; // deduplicate
}

// Update parsed count hint as user types
amountsInput.addEventListener("input", () => {
  const amounts = parseAmounts(amountsInput.value);
  if (amounts.length > 0) {
    parsedCount.textContent = `${amounts.length} unique amount(s) detected`;
  } else {
    parsedCount.textContent = "";
  }
});

// ---------------------------------------------------------------------------
// Logging helper
// ---------------------------------------------------------------------------
function addLog(message, type = "info") {
  const li = document.createElement("li");
  li.textContent = message;
  if (type === "error") li.classList.add("error");
  if (type === "success") li.classList.add("success");
  logList.appendChild(li);
  li.scrollIntoView({ behavior: "smooth" });
}

// ---------------------------------------------------------------------------
// Listen for progress updates from the background service worker
// ---------------------------------------------------------------------------
chrome.runtime.onMessage.addListener((msg) => {
  if (msg.type === "PROGRESS") {
    const pct = Math.round(msg.percent);
    progressBar.style.width = pct + "%";
    progressText.textContent = msg.text || `${pct}%`;
  } else if (msg.type === "LOG") {
    addLog(msg.text, msg.level || "info");
  } else if (msg.type === "DONE") {
    isRunning = false;
    btnStart.disabled = false;
    btnCancel.disabled = true;
    resultSection.classList.remove("hidden");
    if (msg.error) {
      resultText.textContent = "Extraction failed: " + msg.error;
      resultText.style.color = "#de002e";
    } else {
      resultText.textContent =
        `Extraction complete. ${msg.shipmentCount || 0} shipment(s) across ${msg.invoiceCount || 0} invoice(s). File downloaded.`;
      resultText.style.color = "#00805a";
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

  // Check that the active tab is on FedEx Billing
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab || !tab.url || !tab.url.includes("fedex.com/online/billing")) {
    addLog(
      "Please navigate to fedex.com/online/billing/cbs/invoices first.",
      "error"
    );
    return;
  }

  isRunning = true;
  btnStart.disabled = true;
  btnCancel.disabled = false;
  progressSection.classList.remove("hidden");
  resultSection.classList.add("hidden");
  logList.innerHTML = "";
  progressBar.style.width = "0%";
  progressText.textContent = "Starting...";

  addLog(`Starting extraction for ${amounts.length} amount(s)...`);

  // Send to background service worker
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
  btnStart.disabled = false;
  btnCancel.disabled = true;
  progressText.textContent = "Cancelled.";
  addLog("Extraction cancelled by user.", "error");
});
