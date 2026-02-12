"use strict";

// ==========================================================================
// FedEx Invoice Scraper — Content Script
//
// Injected into https://www.fedex.com/online/billing/*
// Listens for messages from the background service worker and performs
// DOM queries / clicks on the live FedEx Billing pages.
// ==========================================================================

// ---------------------------------------------------------------------------
// Utility: send a debug log back to the background/popup
// ---------------------------------------------------------------------------
function debugLog(text) {
  try {
    chrome.runtime.sendMessage({ type: "LOG", text: "[CS] " + text, level: "info" });
  } catch { /* popup may be closed */ }
}

// ---------------------------------------------------------------------------
// Utility: wait for an element to appear in the DOM (MutationObserver-based)
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
// Utility: wait until text content appears anywhere in the page
// ---------------------------------------------------------------------------
function waitForText(text, timeout = 20000) {
  return new Promise((resolve, reject) => {
    if (document.body.innerText.includes(text)) return resolve();

    const timer = setTimeout(() => {
      observer.disconnect();
      reject(new Error(`Timeout waiting for text "${text}"`));
    }, timeout);

    const observer = new MutationObserver(() => {
      if (document.body.innerText.includes(text)) {
        clearTimeout(timer);
        observer.disconnect();
        resolve();
      }
    });

    observer.observe(document.body, { childList: true, subtree: true, characterData: true });
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
// "$1,431.43" → "1431.43", "452.67" → "452.67"
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

  // Check for common table containers
  const tableSelectors = [
    ".fdx-c-table__tbody.invoice-table-data",
    ".fdx-c-table__tbody",
    ".invoice-table-data",
    "table.fdx-c-table",
    ".fdx-c-table",
    "table",
    "[class*='invoice']",
    "[class*='table']",
    "app-invoices",
    "app-invoice-table",
    "app-shipment-table",
    "cdk-virtual-scroll-viewport",
  ];

  info.selectors = {};
  for (const sel of tableSelectors) {
    try {
      const els = document.querySelectorAll(sel);
      info.selectors[sel] = els.length;
    } catch {
      info.selectors[sel] = "invalid selector";
    }
  }

  // Find all elements with data-label attributes
  const dataLabels = new Set();
  document.querySelectorAll("[data-label]").forEach((el) => {
    dataLabels.add(el.getAttribute("data-label"));
  });
  info.dataLabels = [...dataLabels];

  // Find all table rows and their structure
  const allRows = document.querySelectorAll("tr");
  info.totalTrElements = allRows.length;

  // Look for dollar amounts on the page
  const bodyText = document.body.innerText;
  const amountMatches = bodyText.match(/\$[\d,]+\.\d{2}/g);
  info.dollarAmountsOnPage = amountMatches ? [...new Set(amountMatches)].slice(0, 30) : [];

  // Sample first table row structure
  if (allRows.length > 0) {
    const sampleRow = allRows[0];
    const cells = sampleRow.querySelectorAll("td, th");
    info.sampleRowCells = Array.from(cells).map((c) => ({
      tag: c.tagName,
      class: c.className.slice(0, 80),
      dataLabel: c.getAttribute("data-label"),
      text: c.textContent.trim().slice(0, 50),
    }));
  }

  // Look for Angular app components
  const appComponents = new Set();
  document.querySelectorAll("[_ngcontent]").length; // Just check presence
  document.querySelectorAll("*").forEach((el) => {
    if (el.tagName.startsWith("APP-")) {
      appComponents.add(el.tagName.toLowerCase());
    }
  });
  info.appComponents = [...appComponents];

  return info;
}

// ---------------------------------------------------------------------------
// ACTION: Find the invoice table on the page — tries multiple strategies
// Returns the table body element or null
// ---------------------------------------------------------------------------
async function findInvoiceTable() {
  // Strategy 1: exact class from static HTML
  const strategies = [
    ".fdx-c-table__tbody.invoice-table-data",
    "tbody.invoice-table-data",
    ".invoice-table-data",
    // Strategy 2: Angular component scoped
    "app-invoices .fdx-c-table__tbody",
    "app-invoices tbody",
    // Strategy 3: any table inside the content area
    "#content .fdx-c-table .fdx-c-table__tbody",
    "#content table tbody",
    // Strategy 4: virtual scroll table
    "cdk-virtual-scroll-viewport .fdx-c-table__tbody",
    "cdk-virtual-scroll-viewport tbody",
    // Strategy 5: broad fallbacks
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
    } catch { /* skip invalid */ }
  }

  return null;
}

// ---------------------------------------------------------------------------
// ACTION: Scrape invoice rows from whatever table we find.
// Uses multiple strategies to extract amount and invoice number.
// ---------------------------------------------------------------------------
function scrapeInvoiceRows(tableBody) {
  const rows = tableBody.querySelectorAll("tr");
  const results = [];

  debugLog(`Scraping ${rows.length} table rows...`);

  rows.forEach((row, idx) => {
    const cells = row.querySelectorAll("td");
    if (cells.length === 0) return; // skip header rows

    // --- Strategy A: data-label attributes ---
    const cellMap = {};
    cells.forEach((td) => {
      const label = td.getAttribute("data-label");
      if (label) {
        cellMap[label] = { text: td.textContent.trim(), el: td };
      }
    });

    // Try known data-label names for amount
    const amountCell =
      cellMap["currentBalanceStr"] ||
      cellMap["originalAmountStr"] ||
      cellMap["currentBalance"] ||
      cellMap["originalAmount"] ||
      cellMap["amount"] ||
      cellMap["ORIGINAL_AMOUNT_DUE"] ||
      cellMap["CURRENT_BALANCE"] ||
      cellMap["balance"];

    // Try known data-label names for invoice number
    const invoiceCell =
      cellMap["invoiceNumber"] ||
      cellMap["INVOICE_NUMBER"] ||
      cellMap["invoice"];

    if (invoiceCell && amountCell) {
      const amount = normalizeAmount(amountCell.text);
      const invoiceNumber = invoiceCell.text;
      const link =
        invoiceCell.el.querySelector("a") ||
        invoiceCell.el.querySelector("[role='link']") ||
        invoiceCell.el;
      results.push({ invoiceNumber, amount, linkEl: link, strategy: "data-label" });
      return;
    }

    // --- Strategy B: scan all cells for dollar amounts and links ---
    let foundAmount = "";
    let foundLink = null;
    let foundInvoiceNum = "";

    cells.forEach((td) => {
      const text = td.textContent.trim();

      // Check for dollar amount
      const amtMatch = text.match(/^\$?([\d,]+\.\d{2})$/);
      if (amtMatch && !foundAmount) {
        foundAmount = normalizeAmount(amtMatch[1]);
      }

      // Check for a link — the invoice number is usually a link
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

    // --- Strategy C: full text scan of the row ---
    const rowText = row.textContent;
    const allAmounts = rowText.match(/\$[\d,]+\.\d{2}/g);
    const link = row.querySelector("a");
    if (allAmounts && allAmounts.length > 0 && link) {
      // Use the last dollar amount in the row (often the balance/total column)
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
// ACTION: On the Invoice List page, find and click the row matching amount
// ---------------------------------------------------------------------------
async function findAndClickInvoice(targetAmount) {
  const target = normalizeAmount(targetAmount);
  debugLog(`Looking for amount: ${target} (raw: ${targetAmount})`);

  // Wait for the page to have some table content
  // Try multiple selectors since we don't know which one the live site uses
  try {
    await waitForAny([
      ".fdx-c-table__tbody.invoice-table-data",
      ".invoice-table-data",
      "app-invoices table",
      "app-invoices .fdx-c-table",
      "#content table tbody",
      ".fdx-c-table__tbody",
    ], 25000);
  } catch {
    debugLog("No table found via waitForAny, trying fallback text wait...");
    // Fallback: wait for any dollar amount to appear on the page
    try {
      await waitForText("$", 15000);
    } catch {
      debugLog("No dollar amounts appeared on page");
    }
  }

  // Give Angular extra time to finish rendering
  await new Promise((r) => setTimeout(r, 2500));

  // Diagnose what we actually see
  const diag = diagnosePage();
  debugLog(`Page URL: ${diag.url}`);
  debugLog(`Data-labels found: [${diag.dataLabels.join(", ")}]`);
  debugLog(`App components: [${diag.appComponents.join(", ")}]`);
  debugLog(`Dollar amounts on page: ${diag.dollarAmountsOnPage.length}`);
  debugLog(`TR elements: ${diag.totalTrElements}`);

  // Check if our target amount even exists on the page
  const targetFormatted = "$" + Number(target).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
  const bodyText = document.body.innerText;
  if (!bodyText.includes(target) && !bodyText.includes(targetFormatted)) {
    debugLog(`Amount ${targetFormatted} not visible on page text at all`);
    // It might be off-screen in virtual scroll, continue anyway
  } else {
    debugLog(`Amount ${targetFormatted} IS visible on page`);
  }

  // Try to find the invoice table
  const tableResult = await findInvoiceTable();

  if (tableResult) {
    const invoices = scrapeInvoiceRows(tableResult.tableBody);
    debugLog(`Scraped ${invoices.length} invoice row(s) via ${tableResult.selector}`);

    // Log first few for debugging
    for (let i = 0; i < Math.min(3, invoices.length); i++) {
      const inv = invoices[i];
      debugLog(`  Row ${i}: invoice="${inv.invoiceNumber}" amount="${inv.amount}" via ${inv.strategy}`);
    }

    // Try to match
    for (const inv of invoices) {
      if (inv.amount === target) {
        debugLog(`MATCH FOUND: invoice #${inv.invoiceNumber}`);
        inv.linkEl.click();
        return { success: true, invoiceNumber: inv.invoiceNumber };
      }
    }

    debugLog(`No exact match among ${invoices.length} rows. Target="${target}"`);

    // Virtual scroll: try scrolling to load more rows
    const viewport = document.querySelector("cdk-virtual-scroll-viewport");
    if (viewport) {
      debugLog("Virtual scroll detected, scrolling to find more rows...");
      const maxScrollAttempts = 30;
      for (let i = 0; i < maxScrollAttempts; i++) {
        viewport.scrollTop += viewport.clientHeight;
        await new Promise((r) => setTimeout(r, 800));

        const moreInvoices = scrapeInvoiceRows(tableResult.tableBody);
        for (const inv of moreInvoices) {
          if (inv.amount === target) {
            debugLog(`MATCH FOUND after scroll: invoice #${inv.invoiceNumber}`);
            inv.linkEl.click();
            return { success: true, invoiceNumber: inv.invoiceNumber };
          }
        }
      }
    }
  } else {
    debugLog("Could not find any invoice table element!");
  }

  // --- ULTIMATE FALLBACK: text-based search across the entire page ---
  debugLog("Trying ultimate fallback: full-page text search...");
  const result = await fullPageAmountSearch(target);
  if (result) {
    return result;
  }

  return {
    success: false,
    error: `Amount $${target} not found in invoice list`,
    diagnostics: diag,
  };
}

// ---------------------------------------------------------------------------
// FALLBACK: Search the entire page for the amount and find a nearby link
// ---------------------------------------------------------------------------
async function fullPageAmountSearch(targetNormalized) {
  const targetFormatted = "$" + Number(targetNormalized).toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

  // Find all text nodes containing this amount
  const walker = document.createTreeWalker(
    document.body,
    NodeFilter.SHOW_TEXT,
    {
      acceptNode: (node) =>
        node.textContent.includes(targetNormalized) ||
        node.textContent.includes(targetFormatted)
          ? NodeFilter.FILTER_ACCEPT
          : NodeFilter.FILTER_REJECT,
    }
  );

  const matchNodes = [];
  while (walker.nextNode()) {
    matchNodes.push(walker.currentNode);
  }

  debugLog(`Full-page search found ${matchNodes.length} text node(s) with amount`);

  for (const node of matchNodes) {
    // Walk up the DOM to find a table row or grid item
    let el = node.parentElement;
    for (let i = 0; i < 10 && el; i++) {
      const tagName = el.tagName.toLowerCase();
      if (tagName === "tr" || el.classList.contains("invoice-grid-item") ||
          el.classList.contains("fdx-c-table__tbody__tr")) {
        // Found a row-like container. Look for a clickable link.
        const link = el.querySelector("a");
        if (link) {
          const invoiceNumber = link.textContent.trim();
          debugLog(`Fallback MATCH: clicking link "${invoiceNumber}" in row`);
          link.click();
          return { success: true, invoiceNumber };
        }
      }
      el = el.parentElement;
    }
  }

  return null;
}

// ---------------------------------------------------------------------------
// ACTION: On the Invoice Details page, scrape all tracking IDs
// ---------------------------------------------------------------------------
async function scrapeTrackingIds() {
  debugLog("Scraping tracking IDs...");

  // Wait for the shipment table to appear — try multiple selectors
  try {
    await waitForAny([
      "app-shipment-table",
      "app-shipment-table-header",
      "#content .fdx-c-table__tbody",
      "table.fdx-c-table",
    ], 20000);
  } catch {
    debugLog("No shipment table found via waitForAny");
  }

  await new Promise((r) => setTimeout(r, 2500));

  const diag = diagnosePage();
  debugLog(`Invoice details - TR elements: ${diag.totalTrElements}, data-labels: [${diag.dataLabels.join(", ")}]`);

  const trackingIds = [];

  // Strategy 1: data-label approach
  const allCells = document.querySelectorAll("[data-label]");
  allCells.forEach((td) => {
    const label = td.getAttribute("data-label");
    if (label === "trackingNumber" || label === "TRACKING_ID" || label === "trackingId") {
      const text = td.textContent.trim();
      if (text && !trackingIds.some((t) => t.id === text)) {
        trackingIds.push({ id: text, el: td.querySelector("a") || td });
      }
    }
  });
  debugLog(`Strategy 1 (data-label): found ${trackingIds.length} tracking IDs`);

  // Strategy 2: find links that look like tracking numbers (digits, ~12-15 chars)
  if (trackingIds.length === 0) {
    const allLinks = document.querySelectorAll("a");
    allLinks.forEach((a) => {
      const text = a.textContent.trim();
      // FedEx tracking numbers are typically 12-15 digits
      if (/^\d{12,15}$/.test(text)) {
        const href = a.getAttribute("href") || "";
        // Make sure it's a shipment link, not some other number
        if (href.includes("shipment") || href.includes("tracking") ||
            a.closest("table") || a.closest(".fdx-c-table")) {
          if (!trackingIds.some((t) => t.id === text)) {
            trackingIds.push({ id: text, el: a });
          }
        }
      }
    });
    debugLog(`Strategy 2 (link scan): found ${trackingIds.length} tracking IDs`);
  }

  // Strategy 3: broader link scan (tracking IDs in any table row link)
  if (trackingIds.length === 0) {
    const rows = document.querySelectorAll("tr");
    rows.forEach((row) => {
      const links = row.querySelectorAll("a");
      links.forEach((a) => {
        const text = a.textContent.trim();
        if (/^\d{10,22}$/.test(text) && !trackingIds.some((t) => t.id === text)) {
          trackingIds.push({ id: text, el: a });
        }
      });
    });
    debugLog(`Strategy 3 (row link scan): found ${trackingIds.length} tracking IDs`);
  }

  debugLog(`Total tracking IDs found: ${trackingIds.length}`);
  trackingIds.forEach((t, i) => debugLog(`  Tracking ${i}: ${t.id}`));

  return trackingIds.map((t) => t.id);
}

// ---------------------------------------------------------------------------
// ACTION: Click a specific tracking ID link on the Invoice Details page
// ---------------------------------------------------------------------------
async function clickTrackingId(trackingId) {
  debugLog(`Clicking tracking ID: ${trackingId}`);

  // Strategy 1: data-label cells
  const allCells = document.querySelectorAll("[data-label]");
  for (const td of allCells) {
    const label = td.getAttribute("data-label");
    if ((label === "trackingNumber" || label === "TRACKING_ID" || label === "trackingId") &&
        td.textContent.trim() === trackingId) {
      const link = td.querySelector("a") || td;
      link.click();
      return true;
    }
  }

  // Strategy 2: find any link with this exact text
  const allLinks = document.querySelectorAll("a");
  for (const a of allLinks) {
    if (a.textContent.trim() === trackingId) {
      a.click();
      return true;
    }
  }

  debugLog(`Could not find clickable element for tracking ID ${trackingId}`);
  return false;
}

// ---------------------------------------------------------------------------
// ACTION: On the Shipment Details page, scrape all shipment data
// ---------------------------------------------------------------------------
async function scrapeShipmentDetails() {
  debugLog("Scraping shipment details...");

  // Wait for the shipment summary area to load
  try {
    await waitForAny([
      "app-shipment-summary",
      "app-shipment-detail",
      ".invoice-summary",
      "#SHIPMENT_DETAILS",
    ], 20000);
  } catch {
    debugLog("No shipment detail container found");
  }

  await new Promise((r) => setTimeout(r, 2500));

  const data = {};

  // --- Billing Information section ---
  const summaryEl =
    document.querySelector("app-shipment-summary") ||
    document.querySelector(".invoice-summary") ||
    document.querySelector("app-shipment-detail");

  if (summaryEl) {
    // Extract label/value pairs from eyebrow labels
    const eyebrows = summaryEl.querySelectorAll(
      ".fdx-c-eyebrow, .fdx-c-eyebrow--small, [class*='eyebrow']"
    );
    eyebrows.forEach((label) => {
      const key = label.textContent.trim();
      const parent = label.closest(".fdx-o-grid__item") ||
                     label.closest("[class*='grid__item']") ||
                     label.parentElement;
      if (parent) {
        const allText = parent.textContent.trim();
        const value = allText.replace(key, "").trim();
        if (value) data[key] = value;
      }
    });

    // --- Sender and Recipient information ---
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

  // If eyebrow approach found nothing, try a generic label-value scan
  if (Object.keys(data).length === 0 && summaryEl) {
    debugLog("Eyebrow approach empty, trying generic label-value scan...");
    // Look for any small/bold text pairings
    const allEls = summaryEl.querySelectorAll("*");
    let lastLabel = "";
    allEls.forEach((el) => {
      if (el.children.length === 0) {
        const text = el.textContent.trim();
        if (!text) return;
        const style = window.getComputedStyle(el);
        const fontSize = parseFloat(style.fontSize);
        const fontWeight = parseInt(style.fontWeight);
        // Heuristic: small text or bold text = label, normal text after = value
        if (fontSize <= 12 || fontWeight >= 600 || el.tagName === "H4" || el.tagName === "H5") {
          lastLabel = text;
        } else if (lastLabel && text !== lastLabel) {
          data[lastLabel] = text;
          lastLabel = "";
        }
      }
    });
  }

  // --- Accordion sections ---
  const accordionButtons = document.querySelectorAll(
    ".fdx-c-accordion__button, [class*='accordion__button']"
  );
  for (const btn of accordionButtons) {
    if (btn.getAttribute("aria-expanded") === "false") {
      btn.click();
      await new Promise((r) => setTimeout(r, 800));
    }
  }
  await new Promise((r) => setTimeout(r, 1000));

  // --- Shipment Details accordion ---
  const shipmentDetailsSection = document.getElementById("SHIPMENT_DETAILS");
  if (shipmentDetailsSection) {
    const pairs = extractLabelValuePairs(shipmentDetailsSection);
    for (const [key, value] of Object.entries(pairs)) {
      if (!data[key]) data[key] = value;
    }
  }

  // --- Charges accordion ---
  const chargesSection = document.getElementById("CHARGES");
  if (chargesSection) {
    const pairs = extractLabelValuePairs(chargesSection);
    for (const [key, value] of Object.entries(pairs)) {
      data["Charge: " + key] = value;
    }
    const chargeTables = chargesSection.querySelectorAll("table, .fdx-c-table");
    chargeTables.forEach((table) => {
      const rows = table.querySelectorAll("tr");
      rows.forEach((row) => {
        const cells = row.querySelectorAll("td, th");
        if (cells.length >= 2) {
          const key = cells[0].textContent.trim();
          const value = cells[cells.length - 1].textContent.trim();
          if (key && value) data["Charge: " + key] = value;
        }
      });
    });
  }

  // --- Reference accordion ---
  const refSection = document.getElementById("REFERENCE");
  if (refSection) {
    const pairs = extractLabelValuePairs(refSection);
    for (const [key, value] of Object.entries(pairs)) {
      data["Ref: " + key] = value;
    }
  }

  // --- Customs accordion ---
  const customsSection = document.getElementById("CUSTOMS");
  if (customsSection) {
    const pairs = extractLabelValuePairs(customsSection);
    for (const [key, value] of Object.entries(pairs)) {
      data["Customs: " + key] = value;
    }
  }

  debugLog(`Scraped ${Object.keys(data).length} fields from shipment details`);
  for (const [k, v] of Object.entries(data)) {
    debugLog(`  ${k}: ${v.toString().slice(0, 50)}`);
  }

  return data;
}

// ---------------------------------------------------------------------------
// Helper: Extract address block from the summary given a section header
// ---------------------------------------------------------------------------
function extractAddressSection(container, headerText) {
  const allElements = container.querySelectorAll("*");
  let headerEl = null;

  for (const el of allElements) {
    if (el.children.length === 0 && el.textContent.trim() === headerText) {
      headerEl = el;
      break;
    }
  }

  // Also try partial match
  if (!headerEl) {
    for (const el of allElements) {
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
  const pElements = section.querySelectorAll(
    "p, span[class*='grid__item'], span"
  );
  let foundHeader = false;
  for (const p of pElements) {
    const text = p.textContent.trim();
    if (text.includes(headerText)) {
      foundHeader = true;
      continue;
    }
    if (foundHeader && text &&
        !text.includes("VIEW SIGNATURE") &&
        !text.includes("Dispute") &&
        text !== headerText) {
      // Deduplicate: only add if not already present
      if (!lines.includes(text)) {
        lines.push(text);
      }
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
// Helper: Extract label-value pairs from an accordion section
// ---------------------------------------------------------------------------
function extractLabelValuePairs(container) {
  const pairs = {};

  // Pattern 1: eyebrow labels
  const eyebrows = container.querySelectorAll(
    ".fdx-c-eyebrow, [class*='eyebrow']"
  );
  eyebrows.forEach((label) => {
    const key = label.textContent.trim();
    const parent = label.closest("[class*='grid__item']") || label.parentElement;
    if (parent) {
      const allText = parent.textContent.trim();
      const value = allText.replace(key, "").trim();
      if (key && value) pairs[key] = value;
    }
  });

  // Pattern 2: grid rows with label + value
  const gridRows = container.querySelectorAll("[class*='grid__row']");
  gridRows.forEach((row) => {
    const label = row.querySelector(
      "[class*='font-size--small'], [class*='color--text']"
    );
    const value = row.querySelector("[class*='fontweight--medium']");
    if (label && value) {
      const key = label.textContent.trim();
      const val = value.textContent.trim();
      if (key && val) pairs[key] = val;
    }
  });

  // Pattern 3: table rows
  const tableRows = container.querySelectorAll("tr");
  tableRows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length >= 2) {
      const key = cells[0].textContent.trim();
      const value = cells[cells.length - 1].textContent.trim();
      if (key && value) pairs[key] = value;
    }
  });

  return pairs;
}

// ---------------------------------------------------------------------------
// ACTION: Navigate back (browser back button)
// ---------------------------------------------------------------------------
function navigateBack() {
  window.history.back();
}

// ==========================================================================
// Message listener — the background script sends commands here
// ==========================================================================
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  // Ignore messages without an action (these are progress/log broadcasts)
  if (!msg.action) return false;

  (async () => {
    try {
      switch (msg.action) {
        case "FIND_AND_CLICK_INVOICE": {
          const result = await findAndClickInvoice(msg.amount);
          sendResponse(result);
          break;
        }

        case "SCRAPE_TRACKING_IDS": {
          const trackingIds = await scrapeTrackingIds();
          sendResponse({ success: true, trackingIds });
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

        case "NAVIGATE_BACK": {
          navigateBack();
          sendResponse({ success: true });
          break;
        }

        case "DIAGNOSE": {
          const diag = diagnosePage();
          sendResponse({ success: true, diagnostics: diag });
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
      sendResponse({ success: false, error: err.message });
    }
  })();

  return true;
});

debugLog("Content script loaded on: " + window.location.href);
