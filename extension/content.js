"use strict";

// ==========================================================================
// FedEx Invoice Scraper — Content Script
//
// Injected into https://www.fedex.com/online/billing/*
// Listens for messages from the background service worker and performs
// DOM queries / clicks on the live FedEx Billing pages.
// ==========================================================================

// Guard against double injection
if (window.__fedexScraperLoaded) {
  // Already loaded — skip re-registration
} else {
  window.__fedexScraperLoaded = true;

// ---------------------------------------------------------------------------
// Utility: send a debug log back to the background/popup
// ---------------------------------------------------------------------------
function debugLog(text) {
  try {
    chrome.runtime.sendMessage({ type: "LOG", text: "[CS] " + text, level: "info" });
  } catch { /* popup may be closed */ }
}

// ---------------------------------------------------------------------------
// Utility: wait for an element to appear in the DOM
// ---------------------------------------------------------------------------
function waitForElement(selector, timeout = 15000) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(selector);
    if (existing) return resolve(existing);

    const timer = setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timeout waiting for "${selector}" (${timeout}ms)`));
    }, timeout);

    const observer = new MutationObserver(() => {
      const el = document.querySelector(selector);
      if (el) {
        clearTimeout(timer);
        observer.disconnect();
        resolve(el);
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  });
}

// ---------------------------------------------------------------------------
// Utility: wait for ANY of multiple selectors to appear
// ---------------------------------------------------------------------------
function waitForAny(selectors, timeout = 20000) {
  return new Promise((resolve, reject) => {
    for (const sel of selectors) {
      const existing = document.querySelector(sel);
      if (existing) return resolve(existing);
    }

    const timer = setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timeout waiting for any of: ${selectors.join(", ")}`));
    }, timeout);

    const observer = new MutationObserver(() => {
      for (const sel of selectors) {
        const el = document.querySelector(sel);
        if (el) {
          clearTimeout(timer);
          observer.disconnect();
          resolve(el);
          return;
        }
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  });
}

// ---------------------------------------------------------------------------
// Utility: wait until the URL changes to contain a substring
// (works for Angular SPA pushState routing)
// ---------------------------------------------------------------------------
function waitForUrlChange(substring, timeout = 30000) {
  return new Promise((resolve, reject) => {
    if (window.location.href.includes(substring)) return resolve();

    const t0 = Date.now();
    const interval = setInterval(() => {
      if (window.location.href.includes(substring)) {
        clearInterval(interval);
        resolve();
      } else if (Date.now() - t0 > timeout) {
        clearInterval(interval);
        reject(new Error(`URL never contained "${substring}" (timeout ${timeout}ms). Current: ${window.location.href}`));
      }
    }, 200);
  });
}

// ---------------------------------------------------------------------------
// Utility: wait until a specific data-label appears in the DOM
// ---------------------------------------------------------------------------
function waitForDataLabel(labelName, timeout = 25000) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`[data-label="${labelName}"]`);
    if (existing) return resolve(existing);

    const timer = setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timeout waiting for data-label="${labelName}"`));
    }, timeout);

    const observer = new MutationObserver(() => {
      const el = document.querySelector(`[data-label="${labelName}"]`);
      if (el) {
        clearTimeout(timer);
        observer.disconnect();
        resolve(el);
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  });
}

// ---------------------------------------------------------------------------
// Utility: small random delay (anti-bot throttle)
// ---------------------------------------------------------------------------
function throttle() {
  const ms = 2000 + Math.random() * 3000;
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ---------------------------------------------------------------------------
// Utility: normalize an amount string to just digits and decimal
// ---------------------------------------------------------------------------
function normalizeAmount(str) {
  return str.replace(/[^0-9.]/g, "");
}

// ---------------------------------------------------------------------------
// DIAGNOSE: Dump the current page structure for debugging
// ---------------------------------------------------------------------------
function diagnosePage() {
  const info = {
    url: window.location.href,
    title: document.title,
  };

  const tableSelectors = [
    ".fdx-c-table__tbody.invoice-table-data",
    ".fdx-c-table__tbody",
    ".invoice-table-data",
    "table",
    "app-invoices",
    "app-invoice-detail",
    "app-shipment-table",
    "app-shipment-detail",
    "cdk-virtual-scroll-viewport",
  ];

  info.selectors = {};
  for (const sel of tableSelectors) {
    try {
      info.selectors[sel] = document.querySelectorAll(sel).length;
    } catch {
      info.selectors[sel] = "err";
    }
  }

  const dataLabels = new Set();
  document.querySelectorAll("[data-label]").forEach((el) => {
    dataLabels.add(el.getAttribute("data-label"));
  });
  info.dataLabels = [...dataLabels];

  info.totalTrElements = document.querySelectorAll("tr").length;
  info.totalLinks = document.querySelectorAll("a").length;

  const bodyText = document.body.innerText;
  const amountMatches = bodyText.match(/\$[\d,]+\.\d{2}/g);
  info.dollarAmountsOnPage = amountMatches ? [...new Set(amountMatches)].slice(0, 30) : [];

  // Count links that look like tracking numbers
  const trackingLikeLinks = [];
  document.querySelectorAll("a").forEach((a) => {
    const text = a.textContent.trim();
    if (/^\d{10,22}$/.test(text)) {
      trackingLikeLinks.push(text);
    }
  });
  info.trackingLikeLinks = trackingLikeLinks;

  const appComponents = new Set();
  document.querySelectorAll("*").forEach((el) => {
    if (el.tagName.startsWith("APP-")) {
      appComponents.add(el.tagName.toLowerCase());
    }
  });
  info.appComponents = [...appComponents];

  // Get first 300 chars of visible page text (for debugging)
  info.pageTextSnippet = bodyText.replace(/\s+/g, " ").slice(0, 300);

  return info;
}

// ---------------------------------------------------------------------------
// Invoice table finding and scraping (multiple strategies)
// ---------------------------------------------------------------------------
function findInvoiceTable() {
  const strategies = [
    ".fdx-c-table__tbody.invoice-table-data",
    "tbody.invoice-table-data",
    ".invoice-table-data",
    "app-invoices .fdx-c-table__tbody",
    "app-invoices tbody",
    "#content .fdx-c-table .fdx-c-table__tbody",
    "#content table tbody",
    "cdk-virtual-scroll-viewport .fdx-c-table__tbody",
    "cdk-virtual-scroll-viewport tbody",
    ".fdx-c-table__tbody",
    "table tbody",
  ];

  for (const sel of strategies) {
    try {
      const el = document.querySelector(sel);
      if (el) {
        const rows = el.querySelectorAll("tr");
        if (rows.length > 0) {
          debugLog(`Found invoice table via "${sel}" with ${rows.length} rows`);
          return { tableBody: el, selector: sel };
        }
      }
    } catch { /* skip */ }
  }
  return null;
}

function scrapeInvoiceRows(tableBody) {
  const rows = tableBody.querySelectorAll("tr");
  const results = [];

  rows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length === 0) return;

    // Strategy A: data-label attributes
    const cellMap = {};
    cells.forEach((td) => {
      const label = td.getAttribute("data-label");
      if (label) cellMap[label] = { text: td.textContent.trim(), el: td };
    });

    const amountCell =
      cellMap["currentBalanceStr"] || cellMap["originalAmountStr"] ||
      cellMap["currentBalance"] || cellMap["originalAmount"] ||
      cellMap["amount"] || cellMap["ORIGINAL_AMOUNT_DUE"] ||
      cellMap["CURRENT_BALANCE"] || cellMap["balance"];

    const invoiceCell =
      cellMap["invoiceNumber"] || cellMap["INVOICE_NUMBER"] || cellMap["invoice"];

    if (invoiceCell && amountCell) {
      const link = invoiceCell.el.querySelector("a") ||
                   invoiceCell.el.querySelector("[role='link']") || invoiceCell.el;
      results.push({
        invoiceNumber: invoiceCell.text,
        amount: normalizeAmount(amountCell.text),
        linkEl: link,
        strategy: "data-label",
      });
      return;
    }

    // Strategy B: scan cells for dollar amounts and links
    let foundAmount = "";
    let foundLink = null;
    let foundInvoiceNum = "";

    cells.forEach((td) => {
      const text = td.textContent.trim();
      const amtMatch = text.match(/^\$?([\d,]+\.\d{2})$/);
      if (amtMatch && !foundAmount) foundAmount = normalizeAmount(amtMatch[1]);
      const linkEl = td.querySelector("a");
      if (linkEl && !foundLink) {
        foundLink = linkEl;
        foundInvoiceNum = linkEl.textContent.trim();
      }
    });

    if (foundAmount && foundLink) {
      results.push({
        invoiceNumber: foundInvoiceNum,
        amount: foundAmount,
        linkEl: foundLink,
        strategy: "cell-scan",
      });
      return;
    }

    // Strategy C: full text scan of the row
    const rowText = row.textContent;
    const allAmounts = rowText.match(/\$[\d,]+\.\d{2}/g);
    const link = row.querySelector("a");
    if (allAmounts && link) {
      for (const amt of allAmounts) {
        results.push({
          invoiceNumber: link.textContent.trim(),
          amount: normalizeAmount(amt),
          linkEl: link,
          strategy: "text-scan",
        });
      }
    }
  });

  return results;
}

// ---------------------------------------------------------------------------
// FIND AND CLICK INVOICE + WAIT FOR DETAILS PAGE
// This now handles the entire flow: find → click → wait for navigation
// → scrape tracking IDs. Doing it all in one content script call avoids
// the SPA navigation / re-injection timing problems.
// ---------------------------------------------------------------------------
async function findClickAndScrapeInvoice(targetAmount) {
  const target = normalizeAmount(targetAmount);
  debugLog(`Looking for amount: ${target}`);

  // Wait for table content
  try {
    await waitForAny([
      ".fdx-c-table__tbody.invoice-table-data",
      ".invoice-table-data",
      "app-invoices table",
      "#content table tbody",
      ".fdx-c-table__tbody",
    ], 25000);
  } catch {
    debugLog("No table element found, trying text wait...");
    try {
      // Wait for dollar signs to appear
      await new Promise((resolve, reject) => {
        const t0 = Date.now();
        const iv = setInterval(() => {
          if (document.body.innerText.includes("$")) {
            clearInterval(iv);
            resolve();
          } else if (Date.now() - t0 > 15000) {
            clearInterval(iv);
            reject();
          }
        }, 300);
      });
    } catch { /* continue anyway */ }
  }

  await new Promise((r) => setTimeout(r, 2500));

  const diag = diagnosePage();
  debugLog(`URL: ${diag.url}`);
  debugLog(`Data-labels: [${diag.dataLabels.join(", ")}]`);
  debugLog(`TR elements: ${diag.totalTrElements}, Links: ${diag.totalLinks}`);

  // Find invoice table
  const tableResult = findInvoiceTable();
  let matchedInvoice = null;

  if (tableResult) {
    const invoices = scrapeInvoiceRows(tableResult.tableBody);
    debugLog(`Scraped ${invoices.length} invoice row(s)`);
    for (let i = 0; i < Math.min(5, invoices.length); i++) {
      debugLog(`  Row: inv="${invoices[i].invoiceNumber}" amt="${invoices[i].amount}" [${invoices[i].strategy}]`);
    }

    for (const inv of invoices) {
      if (inv.amount === target) {
        matchedInvoice = inv;
        break;
      }
    }

    // Try virtual scroll
    if (!matchedInvoice) {
      const viewport = document.querySelector("cdk-virtual-scroll-viewport");
      if (viewport) {
        debugLog("Scrolling virtual viewport...");
        for (let i = 0; i < 30; i++) {
          viewport.scrollTop += viewport.clientHeight;
          await new Promise((r) => setTimeout(r, 800));
          const more = scrapeInvoiceRows(tableResult.tableBody);
          matchedInvoice = more.find((inv) => inv.amount === target);
          if (matchedInvoice) break;
        }
      }
    }
  }

  // Fallback: full-page text search
  if (!matchedInvoice) {
    debugLog("Table search failed, trying full-page text search...");
    const targetFormatted = "$" + Number(target).toLocaleString("en-US", {
      minimumFractionDigits: 2, maximumFractionDigits: 2,
    });

    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT, {
      acceptNode: (node) =>
        (node.textContent.includes(target) || node.textContent.includes(targetFormatted))
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT,
    });

    while (walker.nextNode()) {
      let el = walker.currentNode.parentElement;
      for (let i = 0; i < 10 && el; i++) {
        if (el.tagName === "TR" || el.classList.contains("fdx-c-table__tbody__tr") ||
            el.classList.contains("invoice-grid-item")) {
          const link = el.querySelector("a");
          if (link) {
            matchedInvoice = {
              invoiceNumber: link.textContent.trim(),
              linkEl: link,
              strategy: "text-walker",
            };
            break;
          }
        }
        el = el.parentElement;
      }
      if (matchedInvoice) break;
    }
  }

  if (!matchedInvoice) {
    return {
      success: false,
      error: `Amount $${target} not found`,
      diagnostics: diag,
      invoiceNumber: null,
      trackingIds: [],
    };
  }

  debugLog(`MATCH: invoice #${matchedInvoice.invoiceNumber} via ${matchedInvoice.strategy}`);

  // --- CLICK THE INVOICE LINK ---
  matchedInvoice.linkEl.click();
  debugLog("Clicked invoice link. Waiting for navigation...");

  // --- WAIT FOR THE INVOICE DETAILS PAGE TO LOAD ---
  // Strategy 1: Wait for URL to change to invoice-details
  let navigatedToDetails = false;
  try {
    await waitForUrlChange("invoice-detail", 15000);
    navigatedToDetails = true;
    debugLog(`URL changed to: ${window.location.href}`);
  } catch (e) {
    debugLog(`URL did not change to invoice-details: ${e.message}`);
  }

  // Wait for Angular to finish rendering the new view
  await new Promise((r) => setTimeout(r, 3000));

  // Strategy 2: Wait for tracking-specific elements regardless of URL
  debugLog("Waiting for tracking ID elements to appear...");
  let trackingIds = [];

  try {
    // Wait for the data-label="trackingNumber" elements specifically
    await waitForDataLabel("trackingNumber", 20000);
    debugLog("Found data-label='trackingNumber' elements!");
  } catch {
    debugLog("data-label='trackingNumber' not found, trying alternatives...");
    // Try waiting for app-shipment-table or any tracking-like links
    try {
      await waitForAny([
        "app-shipment-table",
        "app-invoice-detail",
      ], 10000);
    } catch {
      debugLog("No shipment table component found either");
    }
    await new Promise((r) => setTimeout(r, 2000));
  }

  // --- NOW SCRAPE TRACKING IDs ---
  const postNavDiag = diagnosePage();
  debugLog(`Post-nav URL: ${postNavDiag.url}`);
  debugLog(`Post-nav data-labels: [${postNavDiag.dataLabels.join(", ")}]`);
  debugLog(`Post-nav tracking-like links: [${postNavDiag.trackingLikeLinks.join(", ")}]`);
  debugLog(`Post-nav app-components: [${postNavDiag.appComponents.join(", ")}]`);

  // Strategy 1: data-label approach
  document.querySelectorAll("[data-label]").forEach((td) => {
    const label = td.getAttribute("data-label");
    if (label === "trackingNumber" || label === "TRACKING_ID" || label === "trackingId") {
      const text = td.textContent.trim();
      if (text && !trackingIds.includes(text)) {
        trackingIds.push(text);
      }
    }
  });
  debugLog(`Strategy 1 (data-label): ${trackingIds.length} tracking IDs`);

  // Strategy 2: links with digit-only text in tables
  if (trackingIds.length === 0) {
    document.querySelectorAll("a").forEach((a) => {
      const text = a.textContent.trim();
      if (/^\d{10,22}$/.test(text)) {
        const href = a.getAttribute("href") || "";
        if (href.includes("shipment") || href.includes("tracking") ||
            a.closest("table") || a.closest("[class*='table']")) {
          if (!trackingIds.includes(text)) trackingIds.push(text);
        }
      }
    });
    debugLog(`Strategy 2 (link-digit): ${trackingIds.length} tracking IDs`);
  }

  // Strategy 3: ANY links with pure digit text
  if (trackingIds.length === 0) {
    document.querySelectorAll("a").forEach((a) => {
      const text = a.textContent.trim();
      if (/^\d{10,22}$/.test(text) && !trackingIds.includes(text)) {
        trackingIds.push(text);
      }
    });
    debugLog(`Strategy 3 (all-digit-links): ${trackingIds.length} tracking IDs`);
  }

  // Strategy 4: look for tracking numbers in the page text
  if (trackingIds.length === 0) {
    const pageText = document.body.innerText;
    // FedEx tracking numbers: 12-15 digits, often starting with 7 or 8
    const matches = pageText.match(/\b\d{12,15}\b/g);
    if (matches) {
      const unique = [...new Set(matches)];
      debugLog(`Strategy 4 (regex): found ${unique.length} digit sequences: ${unique.join(", ")}`);
      // Only include if they're actually clickable
      for (const num of unique) {
        const links = document.querySelectorAll("a");
        for (const a of links) {
          if (a.textContent.trim().includes(num)) {
            if (!trackingIds.includes(num)) trackingIds.push(num);
            break;
          }
        }
      }
    }
    debugLog(`Strategy 4 (clickable regex): ${trackingIds.length} tracking IDs`);
  }

  debugLog(`Total tracking IDs: ${trackingIds.length}: [${trackingIds.join(", ")}]`);

  return {
    success: true,
    invoiceNumber: matchedInvoice.invoiceNumber,
    trackingIds,
    navigatedToDetails,
    diagnostics: postNavDiag,
  };
}

// ---------------------------------------------------------------------------
// Click a tracking ID link
// ---------------------------------------------------------------------------
async function clickTrackingId(trackingId) {
  debugLog(`Clicking tracking ID: ${trackingId}`);

  // Strategy 1: data-label cells
  for (const td of document.querySelectorAll("[data-label]")) {
    const label = td.getAttribute("data-label");
    if ((label === "trackingNumber" || label === "TRACKING_ID" || label === "trackingId") &&
        td.textContent.trim() === trackingId) {
      const link = td.querySelector("a") || td;
      link.click();
      debugLog("Clicked via data-label");
      return true;
    }
  }

  // Strategy 2: any link with this text
  for (const a of document.querySelectorAll("a")) {
    if (a.textContent.trim() === trackingId) {
      a.click();
      debugLog("Clicked via link text match");
      return true;
    }
  }

  debugLog(`Could not find clickable element for ${trackingId}`);
  return false;
}

// ---------------------------------------------------------------------------
// Scrape shipment details (on the Shipment Details page)
// ---------------------------------------------------------------------------
async function scrapeShipmentDetails() {
  debugLog(`Scraping shipment details at: ${window.location.href}`);

  // Wait for shipment detail content
  try {
    await waitForAny([
      "app-shipment-summary",
      "app-shipment-detail",
      ".invoice-summary",
      "#SHIPMENT_DETAILS",
    ], 20000);
  } catch {
    debugLog("No shipment detail container found via waitForAny");
  }

  await new Promise((r) => setTimeout(r, 2500));

  const data = {};
  const summaryEl =
    document.querySelector("app-shipment-summary") ||
    document.querySelector(".invoice-summary") ||
    document.querySelector("app-shipment-detail");

  if (summaryEl) {
    // Eyebrow labels
    summaryEl.querySelectorAll(
      ".fdx-c-eyebrow, .fdx-c-eyebrow--small, [class*='eyebrow']"
    ).forEach((label) => {
      const key = label.textContent.trim();
      const parent = label.closest(".fdx-o-grid__item") ||
                     label.closest("[class*='grid__item']") ||
                     label.parentElement;
      if (parent) {
        const value = parent.textContent.trim().replace(key, "").trim();
        if (value) data[key] = value;
      }
    });

    // Sender / Recipient info
    const senderSection = extractAddressSection(summaryEl, "Sender information");
    if (senderSection) {
      data["Sender Name"] = senderSection.name;
      data["Sender Company"] = senderSection.company;
      data["Sender Address"] = senderSection.address;
      data["Sender City/State/Zip"] = senderSection.cityStateZip;
      data["Sender Country"] = senderSection.country;
    }
    const recipientSection = extractAddressSection(summaryEl, "Recipient information");
    if (recipientSection) {
      data["Recipient Name"] = recipientSection.name;
      data["Recipient Company"] = recipientSection.company;
      data["Recipient Address"] = recipientSection.address;
      data["Recipient City/State/Zip"] = recipientSection.cityStateZip;
      data["Recipient Country"] = recipientSection.country;
    }
  }

  // Fallback: generic label-value scan
  if (Object.keys(data).length === 0 && summaryEl) {
    debugLog("Eyebrow approach empty, trying computed style scan...");
    const allEls = summaryEl.querySelectorAll("*");
    let lastLabel = "";
    allEls.forEach((el) => {
      if (el.children.length === 0) {
        const text = el.textContent.trim();
        if (!text) return;
        const style = window.getComputedStyle(el);
        const fontSize = parseFloat(style.fontSize);
        const fontWeight = parseInt(style.fontWeight);
        if (fontSize <= 12 || fontWeight >= 600 || el.tagName === "H4" || el.tagName === "H5") {
          lastLabel = text;
        } else if (lastLabel && text !== lastLabel) {
          data[lastLabel] = text;
          lastLabel = "";
        }
      }
    });
  }

  // Expand accordions
  for (const btn of document.querySelectorAll(".fdx-c-accordion__button, [class*='accordion__button']")) {
    if (btn.getAttribute("aria-expanded") === "false") {
      btn.click();
      await new Promise((r) => setTimeout(r, 800));
    }
  }
  await new Promise((r) => setTimeout(r, 1000));

  // Accordion sections
  for (const [sectionId, prefix] of [
    ["SHIPMENT_DETAILS", ""],
    ["CHARGES", "Charge: "],
    ["REFERENCE", "Ref: "],
    ["CUSTOMS", "Customs: "],
  ]) {
    const section = document.getElementById(sectionId);
    if (section) {
      const pairs = extractLabelValuePairs(section);
      for (const [key, value] of Object.entries(pairs)) {
        const fullKey = prefix ? prefix + key : key;
        if (!data[fullKey]) data[fullKey] = value;
      }
    }
  }

  debugLog(`Scraped ${Object.keys(data).length} fields`);
  return data;
}

// ---------------------------------------------------------------------------
// Address extraction helper
// ---------------------------------------------------------------------------
function extractAddressSection(container, headerText) {
  let headerEl = null;
  for (const el of container.querySelectorAll("*")) {
    if (el.children.length === 0 && el.textContent.trim() === headerText) {
      headerEl = el;
      break;
    }
  }
  if (!headerEl) {
    for (const el of container.querySelectorAll("*")) {
      if (el.children.length === 0 &&
          el.textContent.trim().toLowerCase().includes(headerText.toLowerCase())) {
        headerEl = el;
        break;
      }
    }
  }
  if (!headerEl) return null;

  const section =
    headerEl.closest(".summary-col") ||
    headerEl.closest("[class*='summary']") ||
    headerEl.closest(".fdx-o-grid") ||
    headerEl.parentElement;
  if (!section) return null;

  const lines = [];
  let foundHeader = false;
  for (const p of section.querySelectorAll("p, span[class*='grid__item'], span")) {
    const text = p.textContent.trim();
    if (text.includes(headerText)) { foundHeader = true; continue; }
    if (foundHeader && text &&
        !text.includes("VIEW SIGNATURE") && !text.includes("Dispute") &&
        text !== headerText && !lines.includes(text)) {
      lines.push(text);
    }
  }
  if (lines.length === 0) return null;

  return {
    name: lines[0] || "",
    company: lines[1] || lines[0] || "",
    address: lines.length > 3 ? lines.slice(2, -2).join(", ") : (lines[2] || ""),
    cityStateZip: lines.length > 2 ? lines[lines.length - 2] : "",
    country: lines[lines.length - 1] || "",
  };
}

// ---------------------------------------------------------------------------
// Label-value pair extraction helper
// ---------------------------------------------------------------------------
function extractLabelValuePairs(container) {
  const pairs = {};

  container.querySelectorAll(".fdx-c-eyebrow, [class*='eyebrow']").forEach((label) => {
    const key = label.textContent.trim();
    const parent = label.closest("[class*='grid__item']") || label.parentElement;
    if (parent) {
      const value = parent.textContent.trim().replace(key, "").trim();
      if (key && value) pairs[key] = value;
    }
  });

  container.querySelectorAll("[class*='grid__row']").forEach((row) => {
    const label = row.querySelector("[class*='font-size--small'], [class*='color--text']");
    const value = row.querySelector("[class*='fontweight--medium']");
    if (label && value) {
      const k = label.textContent.trim();
      const v = value.textContent.trim();
      if (k && v) pairs[k] = v;
    }
  });

  container.querySelectorAll("tr").forEach((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length >= 2) {
      const k = cells[0].textContent.trim();
      const v = cells[cells.length - 1].textContent.trim();
      if (k && v) pairs[k] = v;
    }
  });

  return pairs;
}

// ---------------------------------------------------------------------------
// Navigate back
// ---------------------------------------------------------------------------
function navigateBack() {
  window.history.back();
}

// ==========================================================================
// Message listener
// ==========================================================================
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg.action) return false;

  (async () => {
    try {
      switch (msg.action) {
        case "FIND_CLICK_AND_SCRAPE": {
          const result = await findClickAndScrapeInvoice(msg.amount);
          sendResponse(result);
          break;
        }

        case "CLICK_TRACKING_ID": {
          const clicked = await clickTrackingId(msg.trackingId);
          sendResponse({ success: clicked });
          break;
        }

        case "SCRAPE_SHIPMENT_DETAILS": {
          const details = await scrapeShipmentDetails();
          sendResponse({ success: true, data: details });
          break;
        }

        case "WAIT_FOR_SHIPMENT_PAGE": {
          // Wait for URL to contain shipment-details
          debugLog("Waiting for shipment details page...");
          try {
            await waitForUrlChange("shipment-detail", 20000);
            debugLog(`Arrived at: ${window.location.href}`);
          } catch (e) {
            debugLog(`Shipment page wait failed: ${e.message}`);
          }
          await new Promise((r) => setTimeout(r, 2500));
          sendResponse({ success: true, url: window.location.href });
          break;
        }

        case "WAIT_FOR_INVOICE_DETAILS": {
          debugLog("Waiting for invoice details page...");
          try {
            await waitForUrlChange("invoice-detail", 20000);
            debugLog(`Arrived at: ${window.location.href}`);
          } catch (e) {
            debugLog(`Invoice details wait failed: ${e.message}`);
          }
          await new Promise((r) => setTimeout(r, 2500));
          sendResponse({ success: true, url: window.location.href });
          break;
        }

        case "NAVIGATE_BACK": {
          navigateBack();
          sendResponse({ success: true });
          break;
        }

        case "DIAGNOSE": {
          sendResponse({ success: true, diagnostics: diagnosePage() });
          break;
        }

        case "PING": {
          sendResponse({ success: true, url: window.location.href });
          break;
        }

        default:
          sendResponse({ success: false, error: "Unknown action: " + msg.action });
      }
    } catch (err) {
      debugLog(`ERROR in ${msg.action}: ${err.message}`);
      sendResponse({ success: false, error: err.message });
    }
  })();

  return true;
});

debugLog("Content script loaded on: " + window.location.href);

} // end of double-injection guard
