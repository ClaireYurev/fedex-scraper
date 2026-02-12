"use strict";

// ==========================================================================
// FedEx Invoice Scraper — Background Service Worker
//
// Orchestrates the multi-page scraping workflow:
//   1. For each input amount → find invoice, click it, scrape tracking IDs
//      (all in one content script call to avoid SPA navigation issues)
//   2. For each tracking ID → click it, wait, scrape shipment details
//   3. Generate XLSX with one sheet per invoice amount
// ==========================================================================

let cancelled = false;

// ---------------------------------------------------------------------------
// Broadcast helpers
// ---------------------------------------------------------------------------
function sendProgress(percent, text) {
  chrome.runtime.sendMessage({ type: "PROGRESS", percent, text }).catch(() => {});
}

function sendLog(text, level = "info") {
  chrome.runtime.sendMessage({ type: "LOG", text, level }).catch(() => {});
}

function sendDone(payload) {
  chrome.runtime.sendMessage({ type: "DONE", ...payload }).catch(() => {});
}

// ---------------------------------------------------------------------------
// Helper: send a message to a tab's content script and await response
// ---------------------------------------------------------------------------
function sendToTab(tabId, message, timeout = 120000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error(`sendToTab timeout (${timeout}ms) for action: ${message.action}`));
    }, timeout);

    chrome.tabs.sendMessage(tabId, message, (response) => {
      clearTimeout(timer);
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
      } else {
        resolve(response);
      }
    });
  });
}

// ---------------------------------------------------------------------------
// Helper: inject content script if not already present
// ---------------------------------------------------------------------------
async function ensureContentScript(tabId) {
  try {
    const resp = await sendToTab(tabId, { action: "PING" }, 3000);
    if (resp && resp.success) return;
  } catch {
    // Content script not loaded — inject it
  }

  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content.js"],
  });
  await sleep(1000);
}

// ---------------------------------------------------------------------------
// Helper: wait for the tab to finish loading after a navigation
// ---------------------------------------------------------------------------
function waitForTabLoad(tabId, timeout = 30000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      chrome.tabs.onUpdated.removeListener(listener);
      reject(new Error("Tab load timeout"));
    }, timeout);

    function listener(updatedTabId, changeInfo) {
      if (updatedTabId === tabId && changeInfo.status === "complete") {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    }

    // Check if already complete
    chrome.tabs.get(tabId, (tab) => {
      if (tab && tab.status === "complete") {
        clearTimeout(timer);
        chrome.tabs.onUpdated.removeListener(listener);
        resolve();
      }
    });

    chrome.tabs.onUpdated.addListener(listener);
  });
}

// ---------------------------------------------------------------------------
// Helper: random delay (anti-bot throttle)
// ---------------------------------------------------------------------------
function throttle() {
  const ms = 2000 + Math.random() * 3000;
  return sleep(ms);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ---------------------------------------------------------------------------
// Helper: navigate a tab to a URL and wait for it to load
// ---------------------------------------------------------------------------
async function navigateTab(tabId, url) {
  sendLog(`Navigating to ${url.slice(0, 60)}...`);
  await chrome.tabs.update(tabId, { url });
  await waitForTabLoad(tabId);
  await sleep(2000); // Extra time for Angular to bootstrap
  await ensureContentScript(tabId);
}

// ---------------------------------------------------------------------------
// Helper: wait for navigation after a click, then ensure content script
// (Used for tracking ID → shipment details navigation)
// ---------------------------------------------------------------------------
async function waitForNavAfterClick(tabId, expectedUrlPart, timeout = 30000) {
  const t0 = Date.now();

  while (Date.now() - t0 < timeout) {
    await sleep(500);
    try {
      const tab = await chrome.tabs.get(tabId);
      if (tab.url && tab.url.includes(expectedUrlPart)) {
        if (tab.status === "loading") {
          await waitForTabLoad(tabId);
        }
        await sleep(2000);
        await ensureContentScript(tabId);
        return true;
      }
    } catch { /* tab mid-navigation */ }
  }

  // Fallback: just ensure content script is ready
  await sleep(2000);
  await ensureContentScript(tabId);
  return false;
}

// ---------------------------------------------------------------------------
// XLSX generation
// ---------------------------------------------------------------------------
try {
  importScripts("lib/xlsx.full.min.js");
} catch (e) {
  console.error("Failed to load SheetJS:", e);
}

function generateXlsx(allData) {
  const wb = XLSX.utils.book_new();

  for (const [amount, invoiceData] of Object.entries(allData)) {
    const sheetName = sanitizeSheetName(`$${amount}`);
    const shipments = invoiceData.shipments || [];

    if (shipments.length === 0) {
      const ws = XLSX.utils.aoa_to_sheet([
        ["Invoice Number", invoiceData.invoiceNumber || "N/A"],
        ["Amount", `$${amount}`],
        ["Note", "No shipment details found"],
      ]);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      continue;
    }

    const allKeys = new Set();
    shipments.forEach((s) => Object.keys(s).forEach((k) => allKeys.add(k)));

    const preferredOrder = [
      "Tracking ID number", "Invoice number", "Account number",
      "Invoice date", "Due date", "Status",
      "Total billed", "Tracking ID balance due",
      "Sender Name", "Sender Company", "Sender Address",
      "Sender City/State/Zip", "Sender Country",
      "Recipient Name", "Recipient Company", "Recipient Address",
      "Recipient City/State/Zip", "Recipient Country",
    ];

    const orderedKeys = [];
    for (const pk of preferredOrder) {
      if (allKeys.has(pk)) {
        orderedKeys.push(pk);
        allKeys.delete(pk);
      }
    }
    orderedKeys.push(...[...allKeys].sort());

    const header = orderedKeys;
    const rows = shipments.map((s) => orderedKeys.map((k) => s[k] || ""));
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
    ws["!cols"] = orderedKeys.map((key, idx) => {
      const maxLen = Math.max(key.length, ...rows.map((r) => (r[idx] || "").toString().length));
      return { wch: Math.min(maxLen + 2, 50) };
    });
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  return XLSX.write(wb, { bookType: "xlsx", type: "array" });
}

function sanitizeSheetName(name) {
  return name.replace(/[[\]:*?/\\]/g, "_").slice(0, 31);
}

// ---------------------------------------------------------------------------
// Main extraction orchestrator
// ---------------------------------------------------------------------------
async function runExtraction(amounts, tabId) {
  cancelled = false;
  const allData = {};
  const totalSteps = amounts.length;

  try {
    for (let i = 0; i < amounts.length; i++) {
      if (cancelled) {
        sendLog("Extraction cancelled.", "error");
        sendDone({ error: "Cancelled by user" });
        return;
      }

      const amount = amounts[i];
      const pctBase = (i / totalSteps) * 100;
      sendProgress(pctBase, `Processing $${amount} (${i + 1}/${totalSteps})`);
      sendLog(`--- Processing $${amount} (${i + 1}/${totalSteps}) ---`);

      // Step 1: Navigate to the invoice list page (fresh load each time)
      await navigateTab(tabId, "https://www.fedex.com/online/billing/cbs/invoices");
      await throttle();

      // Step 2: Combined find + click + wait for navigation + scrape tracking IDs
      // All done inside one content script call to survive Angular SPA routing
      sendLog(`Searching for invoice with amount $${amount}...`);
      let result;
      try {
        result = await sendToTab(tabId, {
          action: "FIND_CLICK_AND_SCRAPE",
          amount,
        }, 90000); // 90 second timeout for the combined operation
      } catch (err) {
        sendLog(`Error in combined find/scrape: ${err.message}`, "error");
        allData[amount] = { invoiceNumber: "ERROR", shipments: [] };
        continue;
      }

      if (!result || !result.success) {
        sendLog(`Invoice not found for $${amount}: ${result?.error || "unknown"}`, "error");
        if (result?.diagnostics) {
          const d = result.diagnostics;
          sendLog(`  Page: ${d.url}`, "error");
          sendLog(`  Data-labels: [${d.dataLabels?.join(", ") || "none"}]`, "error");
          sendLog(`  Components: [${d.appComponents?.join(", ") || "none"}]`, "error");
        }
        allData[amount] = { invoiceNumber: "NOT FOUND", shipments: [] };
        continue;
      }

      sendLog(`Found invoice #${result.invoiceNumber}`, "success");

      const trackingIds = result.trackingIds || [];
      sendLog(`Found ${trackingIds.length} tracking ID(s): [${trackingIds.join(", ")}]`,
        trackingIds.length > 0 ? "success" : "error");

      if (result.diagnostics) {
        const d = result.diagnostics;
        sendLog(`  Post-nav URL: ${d.url}`);
        sendLog(`  Post-nav data-labels: [${d.dataLabels?.join(", ") || "none"}]`);
        sendLog(`  Post-nav tracking links: [${d.trackingLikeLinks?.join(", ") || "none"}]`);
      }

      if (trackingIds.length === 0) {
        allData[amount] = {
          invoiceNumber: result.invoiceNumber,
          shipments: [],
        };
        continue;
      }

      // Step 3: For each tracking ID, visit shipment details and scrape
      const shipments = [];
      for (let j = 0; j < trackingIds.length; j++) {
        if (cancelled) break;

        const tid = trackingIds[j];
        const subPct = pctBase + ((j + 1) / trackingIds.length / totalSteps) * 100;
        sendProgress(subPct, `$${amount}: Shipment ${j + 1}/${trackingIds.length}`);
        sendLog(`  Opening shipment ${tid} (${j + 1}/${trackingIds.length})...`);

        // If not the first tracking ID, we need to go back to invoice details
        if (j > 0) {
          sendLog("  Navigating back to invoice details...");
          try {
            await sendToTab(tabId, { action: "NAVIGATE_BACK" }, 5000);
          } catch { /* may fail if page reloaded */ }

          // Wait for invoice details page via content script
          try {
            await sendToTab(tabId, { action: "WAIT_FOR_INVOICE_DETAILS" }, 30000);
          } catch (err) {
            sendLog(`  Back-nav failed: ${err.message}`, "error");
            await ensureContentScript(tabId);
          }
          await throttle();
        }

        // Click the tracking ID link
        try {
          const clickResult = await sendToTab(tabId, {
            action: "CLICK_TRACKING_ID",
            trackingId: tid,
          }, 10000);

          if (!clickResult || !clickResult.success) {
            sendLog(`  Could not click tracking ID ${tid}`, "error");
            continue;
          }
        } catch (err) {
          sendLog(`  Error clicking tracking ID ${tid}: ${err.message}`, "error");
          continue;
        }

        // Wait for shipment details page — use content script's URL watcher
        try {
          await sendToTab(tabId, { action: "WAIT_FOR_SHIPMENT_PAGE" }, 30000);
        } catch (err) {
          sendLog(`  Shipment page wait failed: ${err.message}`, "error");
          // Try background-level URL detection as fallback
          await waitForNavAfterClick(tabId, "shipment-detail", 15000);
        }
        await throttle();

        // Scrape shipment details
        let shipmentResult;
        try {
          shipmentResult = await sendToTab(tabId, {
            action: "SCRAPE_SHIPMENT_DETAILS",
          }, 30000);
        } catch (err) {
          sendLog(`  Error scraping shipment ${tid}: ${err.message}`, "error");
          continue;
        }

        if (shipmentResult && shipmentResult.success && shipmentResult.data) {
          const fieldCount = Object.keys(shipmentResult.data).length;
          shipments.push(shipmentResult.data);
          sendLog(`  Scraped shipment ${tid} (${fieldCount} fields)`, "success");
        } else {
          sendLog(`  Empty data for ${tid}: ${shipmentResult?.error || "no fields"}`, "error");
        }

        await throttle();
      }

      allData[amount] = {
        invoiceNumber: result.invoiceNumber,
        shipments,
      };

      sendLog(
        `Completed $${amount}: ${shipments.length} shipment(s)`,
        shipments.length > 0 ? "success" : "error"
      );
    }

    // Step 4: Generate XLSX
    sendProgress(95, "Generating Excel file...");
    sendLog("Generating XLSX...");

    let xlsxBuffer;
    try {
      xlsxBuffer = generateXlsx(allData);
    } catch (err) {
      sendLog(`XLSX generation failed: ${err.message}`, "error");
      sendDone({ error: "XLSX generation failed: " + err.message });
      return;
    }

    // Step 5: Download
    const uint8 = new Uint8Array(xlsxBuffer);
    let binary = "";
    for (let k = 0; k < uint8.length; k++) {
      binary += String.fromCharCode(uint8[k]);
    }
    const base64 = btoa(binary);
    const dataUrl =
      "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + base64;
    const timestamp = new Date().toISOString().slice(0, 10);

    chrome.downloads.download(
      {
        url: dataUrl,
        filename: `FedEx_Invoices_${timestamp}.xlsx`,
        saveAs: true,
      },
      (downloadId) => {
        if (chrome.runtime.lastError) {
          sendLog(`Download error: ${chrome.runtime.lastError.message}`, "error");
          sendDone({ error: chrome.runtime.lastError.message });
        } else {
          const totalShipments = Object.values(allData).reduce(
            (sum, d) => sum + d.shipments.length, 0
          );
          sendProgress(100, "Done!");
          sendLog("XLSX downloaded!", "success");
          sendDone({ shipmentCount: totalShipments, invoiceCount: amounts.length });
        }
      }
    );
  } catch (err) {
    sendLog(`Fatal error: ${err.message}`, "error");
    sendDone({ error: err.message });
  }
}

// ---------------------------------------------------------------------------
// Listen for messages from popup
// ---------------------------------------------------------------------------
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === "START_EXTRACTION") {
    sendResponse({ ack: true });
    runExtraction(msg.amounts, msg.tabId);
    return true;
  } else if (msg.type === "CANCEL_EXTRACTION") {
    cancelled = true;
    sendResponse({ ack: true });
    return true;
  }
  return false;
});
