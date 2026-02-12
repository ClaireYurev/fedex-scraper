"use strict";

// ==========================================================================
// FedEx Invoice Scraper — Background Service Worker
//
// Orchestrates the multi-page scraping workflow:
//   1. For each input amount → find invoice on list page → click it
//   2. On invoice details → collect all tracking IDs
//   3. For each tracking ID → open shipment details → scrape data
//   4. Generate XLSX with one sheet per invoice amount
// ==========================================================================

let cancelled = false;

// ---------------------------------------------------------------------------
// Broadcast helpers — send messages to popup for progress/logging
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
function sendToTab(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
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
    const resp = await sendToTab(tabId, { action: "PING" });
    if (resp && resp.success) return;
  } catch {
    // Content script not loaded — inject it
  }

  await chrome.scripting.executeScript({
    target: { tabId },
    files: ["content.js"],
  });
  // Wait for it to initialize
  await sleep(500);
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

    chrome.tabs.onUpdated.addListener(listener);
  });
}

// ---------------------------------------------------------------------------
// Helper: random delay between 2-5 seconds (anti-bot throttle)
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
  await chrome.tabs.update(tabId, { url });
  await waitForTabLoad(tabId);
  await sleep(1000); // Extra time for Angular to bootstrap
  await ensureContentScript(tabId);
}

// ---------------------------------------------------------------------------
// Helper: wait for page navigation after a click (content script clicks
// something, which triggers Angular routing) and re-inject content script
// ---------------------------------------------------------------------------
async function waitForNavigation(tabId, expectedUrlPart, timeout = 25000) {
  const t0 = Date.now();

  while (Date.now() - t0 < timeout) {
    await sleep(500);
    try {
      const tab = await chrome.tabs.get(tabId);
      if (tab.url && tab.url.includes(expectedUrlPart)) {
        // Wait for page to settle
        if (tab.status === "loading") {
          await waitForTabLoad(tabId);
        }
        await sleep(1500);
        await ensureContentScript(tabId);
        return;
      }
    } catch {
      // Tab might be mid-navigation
    }
  }

  // If Angular uses client-side routing, the URL may have changed without
  // a full page load. Just ensure content script is ready.
  await sleep(1500);
  await ensureContentScript(tabId);
}

// ---------------------------------------------------------------------------
// XLSX generation using SheetJS loaded into the service worker
// ---------------------------------------------------------------------------

// We import SheetJS as an IIFE that's been bundled into lib/xlsx.full.min.js
// For Manifest V3 service workers, we use importScripts.
try {
  importScripts("lib/xlsx.full.min.js");
} catch (e) {
  console.error("Failed to load SheetJS:", e);
}

function generateXlsx(allData) {
  // allData: { [amount]: { invoiceNumber, shipments: [ { field: value, ... }, ... ] } }
  const wb = XLSX.utils.book_new();

  for (const [amount, invoiceData] of Object.entries(allData)) {
    const sheetName = sanitizeSheetName(`$${amount}`);
    const shipments = invoiceData.shipments || [];

    if (shipments.length === 0) {
      // Create a sheet with just headers noting no shipments found
      const ws = XLSX.utils.aoa_to_sheet([
        ["Invoice Number", invoiceData.invoiceNumber || "N/A"],
        ["Amount", `$${amount}`],
        ["Note", "No shipment details found"],
      ]);
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      continue;
    }

    // Collect all unique keys across shipments for this invoice
    const allKeys = new Set();
    shipments.forEach((s) => Object.keys(s).forEach((k) => allKeys.add(k)));

    // Define a preferred column order
    const preferredOrder = [
      "Tracking ID number",
      "Invoice number",
      "Account number",
      "Invoice date",
      "Due date",
      "Status",
      "Total billed",
      "Tracking ID balance due",
      "Sender Name",
      "Sender Company",
      "Sender Address",
      "Sender City/State/Zip",
      "Sender Country",
      "Recipient Name",
      "Recipient Company",
      "Recipient Address",
      "Recipient City/State/Zip",
      "Recipient Country",
    ];

    const orderedKeys = [];
    for (const pk of preferredOrder) {
      if (allKeys.has(pk)) {
        orderedKeys.push(pk);
        allKeys.delete(pk);
      }
    }
    // Append remaining keys alphabetically
    const remaining = [...allKeys].sort();
    orderedKeys.push(...remaining);

    // Build rows
    const header = orderedKeys;
    const rows = shipments.map((s) => orderedKeys.map((k) => s[k] || ""));
    const wsData = [header, ...rows];

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Auto-size columns (approximate)
    ws["!cols"] = orderedKeys.map((key) => {
      const maxLen = Math.max(
        key.length,
        ...rows.map((r) => (r[orderedKeys.indexOf(key)] || "").toString().length)
      );
      return { wch: Math.min(maxLen + 2, 50) };
    });

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  // Write to binary
  const wbOut = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  return wbOut;
}

function sanitizeSheetName(name) {
  // Excel sheet names: max 31 chars, no []:*?/\
  return name.replace(/[[\]:*?/\\]/g, "_").slice(0, 31);
}

// ---------------------------------------------------------------------------
// Main extraction orchestrator
// ---------------------------------------------------------------------------
async function runExtraction(amounts, tabId) {
  cancelled = false;
  const allData = {};
  const totalSteps = amounts.length;
  const invoiceListUrl = "fedex.com/online/billing/cbs/invoices";

  try {
    for (let i = 0; i < amounts.length; i++) {
      if (cancelled) {
        sendLog("Extraction cancelled.", "error");
        sendDone({ error: "Cancelled by user" });
        return;
      }

      const amount = amounts[i];
      const pctBase = (i / totalSteps) * 100;
      sendProgress(pctBase, `Processing amount $${amount} (${i + 1}/${totalSteps})`);
      sendLog(`--- Processing $${amount} (${i + 1}/${totalSteps}) ---`);

      // Step 1: Navigate back to the invoice list page
      sendLog("Navigating to invoice list...");
      await navigateTab(tabId, "https://www.fedex.com/online/billing/cbs/invoices");
      await throttle();

      // Step 2: Find and click the matching invoice
      sendLog(`Searching for invoice with amount $${amount}...`);
      let findResult;
      try {
        findResult = await sendToTab(tabId, {
          action: "FIND_AND_CLICK_INVOICE",
          amount,
        });
      } catch (err) {
        sendLog(`Error finding invoice: ${err.message}`, "error");
        allData[amount] = { invoiceNumber: "ERROR", shipments: [] };
        continue;
      }

      if (!findResult || !findResult.success) {
        sendLog(
          `Could not find invoice for $${amount}: ${findResult?.error || "unknown"}`,
          "error"
        );
        allData[amount] = { invoiceNumber: "NOT FOUND", shipments: [] };
        continue;
      }

      sendLog(`Found invoice #${findResult.invoiceNumber}`, "success");

      // Step 3: Wait for invoice details page to load
      await waitForNavigation(tabId, "invoice-details");
      await throttle();

      // Step 4: Scrape tracking IDs from the invoice details page
      sendLog("Scraping tracking IDs...");
      let trackingResult;
      try {
        trackingResult = await sendToTab(tabId, {
          action: "SCRAPE_TRACKING_IDS",
        });
      } catch (err) {
        sendLog(`Error scraping tracking IDs: ${err.message}`, "error");
        allData[amount] = {
          invoiceNumber: findResult.invoiceNumber,
          shipments: [],
        };
        continue;
      }

      const trackingIds = trackingResult?.trackingIds || [];
      sendLog(`Found ${trackingIds.length} tracking ID(s)`, "success");

      if (trackingIds.length === 0) {
        allData[amount] = {
          invoiceNumber: findResult.invoiceNumber,
          shipments: [],
        };
        continue;
      }

      // Step 5: For each tracking ID, navigate to shipment details and scrape
      const shipments = [];
      for (let j = 0; j < trackingIds.length; j++) {
        if (cancelled) break;

        const tid = trackingIds[j];
        const subPct = pctBase + ((j + 1) / trackingIds.length / totalSteps) * 100;
        sendProgress(
          subPct,
          `$${amount}: Shipment ${j + 1}/${trackingIds.length} (${tid})`
        );
        sendLog(`  Opening shipment ${tid} (${j + 1}/${trackingIds.length})...`);

        // We need to be on the invoice details page to click the tracking ID.
        // If we navigated away for a previous tracking ID, go back.
        if (j > 0) {
          // Navigate back to invoice details
          try {
            await sendToTab(tabId, { action: "NAVIGATE_BACK" });
          } catch {
            // If content script is gone, navigate manually
          }
          await waitForNavigation(tabId, "invoice-details");
          await throttle();
          await ensureContentScript(tabId);
        }

        // Click the tracking ID link
        try {
          await sendToTab(tabId, {
            action: "CLICK_TRACKING_ID",
            trackingId: tid,
          });
        } catch (err) {
          sendLog(`  Error clicking tracking ID ${tid}: ${err.message}`, "error");
          continue;
        }

        // Wait for shipment details page
        await waitForNavigation(tabId, "shipment-details");
        await throttle();

        // Scrape shipment details
        let shipmentResult;
        try {
          shipmentResult = await sendToTab(tabId, {
            action: "SCRAPE_SHIPMENT_DETAILS",
          });
        } catch (err) {
          sendLog(`  Error scraping shipment ${tid}: ${err.message}`, "error");
          continue;
        }

        if (shipmentResult && shipmentResult.success && shipmentResult.data) {
          shipments.push(shipmentResult.data);
          sendLog(`  Scraped shipment ${tid}`, "success");
        } else {
          sendLog(
            `  No data for shipment ${tid}: ${shipmentResult?.error || "empty"}`,
            "error"
          );
        }

        await throttle();
      }

      allData[amount] = {
        invoiceNumber: findResult.invoiceNumber,
        shipments,
      };

      sendLog(
        `Completed $${amount}: ${shipments.length} shipment(s) scraped`,
        "success"
      );
    }

    // Step 6: Generate XLSX
    sendProgress(95, "Generating Excel file...");
    sendLog("Generating XLSX file...");

    let xlsxBuffer;
    try {
      xlsxBuffer = generateXlsx(allData);
    } catch (err) {
      sendLog(`XLSX generation failed: ${err.message}`, "error");
      sendDone({ error: "XLSX generation failed: " + err.message });
      return;
    }

    // Step 7: Download the file
    // Service workers don't have Blob/URL.createObjectURL, so use a data URL
    const uint8 = new Uint8Array(xlsxBuffer);
    let binary = "";
    for (let k = 0; k < uint8.length; k++) {
      binary += String.fromCharCode(uint8[k]);
    }
    const base64 = btoa(binary);
    const dataUrl =
      "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," +
      base64;
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
            (sum, d) => sum + d.shipments.length,
            0
          );
          sendProgress(100, "Done!");
          sendLog("XLSX file downloaded successfully!", "success");
          sendDone({
            shipmentCount: totalShipments,
            invoiceCount: amounts.length,
          });
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
  } else if (msg.type === "CANCEL_EXTRACTION") {
    cancelled = true;
    sendResponse({ ack: true });
  }
  return true;
});
