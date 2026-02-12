"use strict";

// ==========================================================================
// FedEx Invoice Scraper — Content Script
//
// Injected into https://www.fedex.com/online/billing/*
// Listens for messages from the background service worker and performs
// DOM queries / clicks on the live FedEx Billing pages.
// ==========================================================================

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
// Utility: wait for multiple elements to appear
// ---------------------------------------------------------------------------
function waitForElements(selector, timeout = 15000) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelectorAll(selector);
    if (existing.length > 0) return resolve(existing);

    const timer = setTimeout(() => {
      observer.disconnect();
      // Resolve with whatever we have (might be 0)
      resolve(document.querySelectorAll(selector));
    }, timeout);

    const observer = new MutationObserver(() => {
      const els = document.querySelectorAll(selector);
      if (els.length > 0) {
        clearTimeout(timer);
        observer.disconnect();
        resolve(els);
      }
    });

    observer.observe(document.body, { childList: true, subtree: true });
  });
}

// ---------------------------------------------------------------------------
// Utility: wait until the page URL contains a given substring
// ---------------------------------------------------------------------------
function waitForUrl(substring, timeout = 20000) {
  return new Promise((resolve, reject) => {
    if (window.location.href.includes(substring)) return resolve();
    const t0 = Date.now();
    const interval = setInterval(() => {
      if (window.location.href.includes(substring)) {
        clearInterval(interval);
        resolve();
      } else if (Date.now() - t0 > timeout) {
        clearInterval(interval);
        reject(new Error(`Timeout waiting for URL containing "${substring}"`));
      }
    }, 300);
  });
}

// ---------------------------------------------------------------------------
// Utility: small random delay (anti-bot throttle)
// ---------------------------------------------------------------------------
function throttle() {
  const ms = 2000 + Math.random() * 3000; // 2–5 seconds
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ---------------------------------------------------------------------------
// ACTION: Scrape the invoice list table on the Account Summary page
// Returns array of { invoiceNumber, amount, linkEl } for ALL visible rows
// ---------------------------------------------------------------------------
function scrapeInvoiceList() {
  const rows = document.querySelectorAll(
    ".fdx-c-table__tbody.invoice-table-data .fdx-c-table__tbody__tr"
  );
  const results = [];

  rows.forEach((row) => {
    const cells = row.querySelectorAll(".fdx-c-table__tbody__td");
    // Build a map from data-label to cell text
    const cellMap = {};
    cells.forEach((td) => {
      const label = td.getAttribute("data-label");
      if (label) {
        cellMap[label] = {
          text: td.textContent.trim(),
          el: td,
        };
      }
    });

    // The amount we compare against is in the ORIGINAL_AMOUNT_DUE or
    // currentBalanceStr column. We try several column names.
    const amountCell =
      cellMap["currentBalanceStr"] ||
      cellMap["originalAmountStr"] ||
      cellMap["ORIGINAL_AMOUNT_DUE"] ||
      cellMap["CURRENT_BALANCE"];

    const invoiceCell =
      cellMap["invoiceNumber"] ||
      cellMap["INVOICE_NUMBER"];

    if (invoiceCell) {
      const amount = amountCell ? amountCell.text.replace(/[^0-9.]/g, "") : "";
      const invoiceNumber = invoiceCell.text;
      // The invoice number cell typically contains a clickable link
      const link = invoiceCell.el.querySelector("a") || invoiceCell.el;
      results.push({ invoiceNumber, amount, linkEl: link });
    }
  });

  return results;
}

// ---------------------------------------------------------------------------
// ACTION: On the Invoice List page, find and click the row matching amount
// ---------------------------------------------------------------------------
async function findAndClickInvoice(targetAmount) {
  await waitForElement(".fdx-c-table__tbody.invoice-table-data", 20000);
  // Give Angular a moment to finish rendering rows
  await new Promise((r) => setTimeout(r, 1500));

  const invoices = scrapeInvoiceList();

  // Normalize target: strip $ and commas
  const target = targetAmount.replace(/[^0-9.]/g, "");

  for (const inv of invoices) {
    if (inv.amount === target) {
      // Found the matching invoice — click its link
      inv.linkEl.click();
      return { success: true, invoiceNumber: inv.invoiceNumber };
    }
  }

  // If not found in visible rows, try scrolling through virtual scroll
  const viewport = document.querySelector("cdk-virtual-scroll-viewport");
  if (viewport) {
    // Attempt to scroll through the list to find more rows
    const maxScrollAttempts = 20;
    for (let i = 0; i < maxScrollAttempts; i++) {
      viewport.scrollTop += viewport.clientHeight;
      await new Promise((r) => setTimeout(r, 800));

      const invoicesAfterScroll = scrapeInvoiceList();
      for (const inv of invoicesAfterScroll) {
        if (inv.amount === target) {
          inv.linkEl.click();
          return { success: true, invoiceNumber: inv.invoiceNumber };
        }
      }
    }
  }

  return { success: false, error: `Amount $${target} not found in invoice list` };
}

// ---------------------------------------------------------------------------
// ACTION: On the Invoice Details page, scrape all tracking IDs
// ---------------------------------------------------------------------------
async function scrapeTrackingIds() {
  await waitForElement("app-shipment-table", 20000);
  await new Promise((r) => setTimeout(r, 1500));

  const rows = document.querySelectorAll(
    "app-shipment-table .fdx-c-table__tbody .fdx-c-table__tbody__tr"
  );
  const trackingIds = [];

  rows.forEach((row) => {
    const cells = row.querySelectorAll(".fdx-c-table__tbody__td");
    cells.forEach((td) => {
      const label = td.getAttribute("data-label");
      if (label === "trackingNumber" || label === "TRACKING_ID") {
        const text = td.textContent.trim();
        if (text) {
          trackingIds.push({
            id: text,
            el: td.querySelector("a") || td,
          });
        }
      }
    });
  });

  // Also check if there's pagination / "Show more" in the shipment table
  // and load all pages
  let showMore = document.querySelector(
    "app-shipment-table .fdx-c-button--text"
  );
  while (showMore && showMore.textContent.toLowerCase().includes("show")) {
    showMore.click();
    await new Promise((r) => setTimeout(r, 1500));

    const additionalRows = document.querySelectorAll(
      "app-shipment-table .fdx-c-table__tbody .fdx-c-table__tbody__tr"
    );
    additionalRows.forEach((row) => {
      const cells = row.querySelectorAll(".fdx-c-table__tbody__td");
      cells.forEach((td) => {
        const label = td.getAttribute("data-label");
        if (label === "trackingNumber" || label === "TRACKING_ID") {
          const text = td.textContent.trim();
          if (text && !trackingIds.some((t) => t.id === text)) {
            trackingIds.push({ id: text, el: td.querySelector("a") || td });
          }
        }
      });
    });

    showMore = document.querySelector(
      "app-shipment-table .fdx-c-button--text"
    );
    if (showMore && !showMore.textContent.toLowerCase().includes("show")) {
      break;
    }
  }

  return trackingIds.map((t) => t.id);
}

// ---------------------------------------------------------------------------
// ACTION: Click a specific tracking ID link on the Invoice Details page
// ---------------------------------------------------------------------------
async function clickTrackingId(trackingId) {
  const rows = document.querySelectorAll(
    "app-shipment-table .fdx-c-table__tbody .fdx-c-table__tbody__tr"
  );

  for (const row of rows) {
    const cells = row.querySelectorAll(".fdx-c-table__tbody__td");
    for (const td of cells) {
      const label = td.getAttribute("data-label");
      if (
        (label === "trackingNumber" || label === "TRACKING_ID") &&
        td.textContent.trim() === trackingId
      ) {
        const link = td.querySelector("a") || td;
        link.click();
        return true;
      }
    }
  }
  return false;
}

// ---------------------------------------------------------------------------
// ACTION: On the Shipment Details page, scrape all shipment data
// ---------------------------------------------------------------------------
async function scrapeShipmentDetails() {
  // Wait for the shipment summary area to load
  await waitForElement("app-shipment-summary", 20000);
  await new Promise((r) => setTimeout(r, 1500));

  const data = {};

  // --- Billing Information section ---
  const summaryEl = document.querySelector("app-shipment-summary");
  if (summaryEl) {
    // Extract label/value pairs from the summary grid
    // Labels use: span.fdx-c-eyebrow.fdx-c-eyebrow--small
    // Values follow as sibling elements

    const eyebrows = summaryEl.querySelectorAll(
      "span.fdx-c-eyebrow.fdx-c-eyebrow--small"
    );
    eyebrows.forEach((label) => {
      const key = label.textContent.trim();
      // The value is typically the next sibling or inside a sibling element
      const parent = label.closest(".fdx-o-grid__item") || label.parentElement;
      if (parent) {
        // Get all text after the label
        const allText = parent.textContent.trim();
        const labelText = label.textContent.trim();
        let value = allText.replace(labelText, "").trim();
        if (value) {
          data[key] = value;
        }
      }
    });

    // --- Sender and Recipient information ---
    const summaryText = summaryEl.innerHTML;

    // Find "Sender information" section
    const senderSection = extractAddressSection(summaryEl, "Sender information");
    if (senderSection) {
      data["Sender Name"] = senderSection.name;
      data["Sender Company"] = senderSection.company;
      data["Sender Address"] = senderSection.address;
      data["Sender City/State/Zip"] = senderSection.cityStateZip;
      data["Sender Country"] = senderSection.country;
    }

    // Find "Recipient information" section
    const recipientSection = extractAddressSection(summaryEl, "Recipient information");
    if (recipientSection) {
      data["Recipient Name"] = recipientSection.name;
      data["Recipient Company"] = recipientSection.company;
      data["Recipient Address"] = recipientSection.address;
      data["Recipient City/State/Zip"] = recipientSection.cityStateZip;
      data["Recipient Country"] = recipientSection.country;
    }
  }

  // --- Accordion sections: Shipment details, Charges, etc. ---
  // These are dynamically rendered. Try to expand and read them.
  const accordionButtons = document.querySelectorAll(
    ".fdx-c-accordion__button"
  );

  for (const btn of accordionButtons) {
    // Expand collapsed sections
    if (btn.getAttribute("aria-expanded") === "false") {
      btn.click();
      await new Promise((r) => setTimeout(r, 800));
    }
  }

  // Wait a moment for accordion content to render
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

    // Also look for charge tables
    const chargeTables = chargesSection.querySelectorAll("table, .fdx-c-table");
    chargeTables.forEach((table) => {
      const rows = table.querySelectorAll("tr");
      rows.forEach((row) => {
        const cells = row.querySelectorAll("td, th");
        if (cells.length >= 2) {
          const key = cells[0].textContent.trim();
          const value = cells[cells.length - 1].textContent.trim();
          if (key && value) {
            data["Charge: " + key] = value;
          }
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

  return data;
}

// ---------------------------------------------------------------------------
// Helper: Extract address block from the summary given a section header
// ---------------------------------------------------------------------------
function extractAddressSection(container, headerText) {
  // Find the heading element containing the header text
  const allElements = container.querySelectorAll("*");
  let headerEl = null;

  for (const el of allElements) {
    if (
      el.children.length === 0 &&
      el.textContent.trim() === headerText
    ) {
      headerEl = el;
      break;
    }
  }

  if (!headerEl) return null;

  // Walk up to find the summary-col or grid section containing address lines
  const section =
    headerEl.closest(".summary-col") ||
    headerEl.closest(".fdx-o-grid") ||
    headerEl.parentElement;

  if (!section) return null;

  // Get all paragraph/span elements after the header
  const lines = [];
  const pElements = section.querySelectorAll(
    "p, span.fdx-o-grid__item--12, span.fdx-o-grid__item"
  );
  let foundHeader = false;
  for (const p of pElements) {
    const text = p.textContent.trim();
    if (text === headerText) {
      foundHeader = true;
      continue;
    }
    if (foundHeader && text && text !== "VIEW SIGNATURE PROOF OF DELIVERY" && text !== "Dispute shipment") {
      lines.push(text);
    }
  }

  // Typical structure: Name, Company, Street, Suite, City State Zip, Country
  return {
    name: lines[0] || "",
    company: lines[1] || "",
    address: lines.slice(2, -2).join(", ") || "",
    cityStateZip: lines[lines.length - 2] || "",
    country: lines[lines.length - 1] || "",
  };
}

// ---------------------------------------------------------------------------
// Helper: Extract label-value pairs from an accordion section using
// common FedEx grid/eyebrow patterns
// ---------------------------------------------------------------------------
function extractLabelValuePairs(container) {
  const pairs = {};

  // Pattern 1: eyebrow labels
  const eyebrows = container.querySelectorAll(
    ".fdx-c-eyebrow, .fdx-c-eyebrow--small"
  );
  eyebrows.forEach((label) => {
    const key = label.textContent.trim();
    const parent = label.closest(".fdx-o-grid__item") || label.parentElement;
    if (parent) {
      const allText = parent.textContent.trim();
      const value = allText.replace(key, "").trim();
      if (key && value) pairs[key] = value;
    }
  });

  // Pattern 2: grid rows with small-font label + medium-font value
  const gridRows = container.querySelectorAll(".fdx-o-grid__row");
  gridRows.forEach((row) => {
    const label = row.querySelector(
      ".fdx-u-font-size--small, .fdx-u-color--text"
    );
    const value = row.querySelector(".fdx-u-fontweight--medium");
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

  // Return true to indicate async sendResponse
  return true;
});
