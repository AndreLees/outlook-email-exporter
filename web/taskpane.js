/* global Office, XlsxLite */
"use strict";

const FIELD_DEFINITIONS = [
  { key: "fromName", label: "From Name", default: true },
  { key: "fromEmail", label: "From Email", default: true },
  { key: "toName", label: "To Name", default: true },
  { key: "toEmail", label: "To Email", default: true },
  { key: "subject", label: "Subject", default: true },
  { key: "senderName", label: "Sender / Sent-by Name", default: false },
  { key: "senderEmail", label: "Sender / Sent-by Email", default: false },
  { key: "recipientSource", label: "To Source", default: false },
  { key: "replyTo", label: "Reply-To", default: false },
  { key: "returnPath", label: "Return-Path", default: false },
  { key: "internetMessageId", label: "Internet Message ID", default: false },
  { key: "dateHeader", label: "Date Header", default: false },
];

const state = {
  selectedItems: [],
  failures: [],
  selectionVersion: 0,
  busy: false,
};

const els = {};

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  cacheElements();
  renderFieldList();
  wireUi();

  const apiSupported = Office.context.requirements.isSetSupported("Mailbox", "1.15");
  if (!apiSupported) {
    els.compatibilityCard.classList.remove("hidden");
    setSelectionUi(0, false, "Mailbox API 1.15 is not available in this Outlook client.");
    return;
  }

  Office.context.mailbox.addHandlerAsync(
    Office.EventType.SelectedItemsChanged,
    handleSelectionChanged,
    (result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        console.warn("Could not register SelectedItemsChanged handler:", result.error.message);
      }
    }
  );

  refreshSelection();
});

function cacheElements() {
  [
    "selectionCount",
    "selectionHint",
    "fieldCount",
    "fieldList",
    "defaultFieldsButton",
    "selectAllFieldsButton",
    "clearFieldsButton",
    "exportButton",
    "progressWrap",
    "progressBar",
    "progressText",
    "completionText",
    "errorCard",
    "errorSummary",
    "errorList",
    "compatibilityCard",
  ].forEach((id) => {
    els[id] = document.getElementById(id);
  });
}

function renderFieldList() {
  const fragment = document.createDocumentFragment();
  for (const field of FIELD_DEFINITIONS) {
    const label = document.createElement("label");
    label.className = "field-option";

    const input = document.createElement("input");
    input.type = "checkbox";
    input.name = "exportField";
    input.value = field.key;
    input.checked = field.default;
    input.addEventListener("change", updateFieldCount);

    const text = document.createElement("span");
    text.textContent = field.label;

    label.append(input, text);
    fragment.appendChild(label);
  }
  els.fieldList.replaceChildren(fragment);
  updateFieldCount();
}

function wireUi() {
  els.exportButton.addEventListener("click", exportSelectedMessages);
  els.defaultFieldsButton.addEventListener("click", () => setFieldSelection("default"));
  els.selectAllFieldsButton.addEventListener("click", () => setFieldSelection("all"));
  els.clearFieldsButton.addEventListener("click", () => setFieldSelection("none"));
}

function setFieldSelection(mode) {
  const boxes = [...document.querySelectorAll('input[name="exportField"]')];
  for (const box of boxes) {
    const definition = FIELD_DEFINITIONS.find((field) => field.key === box.value);
    box.checked = mode === "all" ? true : mode === "none" ? false : Boolean(definition?.default);
  }
  updateFieldCount();
}

function updateFieldCount() {
  const count = getSelectedFields().length;
  els.fieldCount.textContent = `${count} selected`;
  updateExportButton();
}

function getSelectedFields() {
  const selectedKeys = new Set(
    [...document.querySelectorAll('input[name="exportField"]:checked')].map((input) => input.value)
  );
  return FIELD_DEFINITIONS.filter((field) => selectedKeys.has(field.key));
}

function handleSelectionChanged() {
  state.selectionVersion += 1;
  state.failures = [];
  clearErrors();
  clearCompletion();
  refreshSelection();
}

async function refreshSelection() {
  if (state.busy) return;

  try {
    const items = await getSelectedItems();
    state.selectedItems = items;
    const count = items.length;

    if (count === 0) {
      setSelectionUi(0, false, "Select one or more messages in the same Outlook folder.");
    } else {
      setSelectionUi(
        count,
        true,
        count === 100
          ? "100 messages selected — Outlook's multi-select limit has been reached."
          : "Ready to export the selected messages."
      );
    }
  } catch (error) {
    console.error(error);
    setSelectionUi(0, false, `Could not read the current selection: ${friendlyError(error)}`);
  }
}

function getSelectedItems() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getSelectedItemsAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }
      resolve(Array.isArray(result.value) ? result.value : []);
    });
  });
}

async function exportSelectedMessages() {
  if (state.busy) return;

  const selectedFields = getSelectedFields();
  if (!selectedFields.length) {
    alert("Select at least one export field.");
    return;
  }

  state.busy = true;
  state.failures = [];
  clearErrors();
  clearCompletion();
  setBusyUi(true);

  const startVersion = state.selectionVersion;

  try {
    const selected = await getSelectedItems();
    state.selectedItems = selected;

    if (!selected.length) {
      throw new Error("No messages are selected.");
    }

    const messages = [];
    updateProgress(0, selected.length, "Starting…");

    for (let i = 0; i < selected.length; i += 1) {
      if (startVersion !== state.selectionVersion) {
        throw new Error("The Outlook selection changed while messages were being read. Please run the export again.");
      }

      const selectedItem = selected[i];
      updateProgress(i, selected.length, `Reading ${i + 1} of ${selected.length}…`);

      try {
        messages.push(await loadMessageDetails(selectedItem, new Set(selectedFields.map((field) => field.key))));
      } catch (error) {
        state.failures.push({
          subject: selectedItem.subject || "(no subject)",
          message: friendlyError(error),
        });
      }

      updateProgress(i + 1, selected.length, `Read ${i + 1} of ${selected.length}`);
    }

    if (!messages.length) {
      throw new Error("None of the selected messages could be read.");
    }

    if (startVersion !== state.selectionVersion) {
      throw new Error("The Outlook selection changed before the workbook was created. Please run the export again.");
    }

    updateProgress(selected.length, selected.length, "Creating Excel workbook…");
    const layout = document.querySelector('input[name="recipientLayout"]:checked')?.value || "message";
    const rows = buildExportRows(sortMessages(messages), layout, selectedFields);
    const headers = selectedFields.map((field) => field.label);

    const blob = XlsxLite.writeWorkbookBlob({
      sheetName: "Emails",
      headers,
      rows,
    });

    downloadBlob(blob, `Outlook_Email_Export_${timestampForFilename()}.xlsx`);
    els.completionText.textContent = `${messages.length} message(s) exported to XLSX${
      state.failures.length ? `; ${state.failures.length} message(s) could not be read` : ""
    }.`;
    els.completionText.classList.remove("hidden");
    renderErrors();
  } catch (error) {
    state.failures.push({ subject: "Export stopped", message: friendlyError(error) });
    renderErrors();
  } finally {
    state.busy = false;
    setBusyUi(false);
    refreshSelection();
  }
}

function loadMessageDetails(selectedItem, _selectedFieldKeys) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.loadItemByIdAsync(selectedItem.itemId, (loadResult) => {
      if (loadResult.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(loadResult.error.message));
        return;
      }

      const loadedItem = loadResult.value;

      (async () => {
        let data;
        let extractionError = null;

        try {
          const structuredFrom = normaliseAddress(loadedItem.from);
          const structuredSender = normaliseAddress(loadedItem.sender);
          const structuredTo = normaliseAddressArray(loadedItem.to);

          // v1.3: MIME headers are now the authoritative source for From and To.
          // Testing with spam showed Outlook's structured address properties can
          // contain a delegate/sender address or omit the actual To alias, while
          // the raw EML From:/To: headers contain the values displayed in Outlook.
          // getAsFileAsync is therefore performed once per message and the parsed
          // header addresses are preferred. Structured Outlook values remain only
          // as a fallback if EML/header extraction is unavailable for a message.
          let rawHeaders = "";
          let headers = new Map();
          try {
            rawHeaders = await getEmlHeaderBlockWithTimeout(loadedItem, 5000);
            if (rawHeaders) headers = parseInternetHeaders(rawHeaders);
          } catch (headerError) {
            console.warn("EML header extraction unavailable; using Outlook address fallback:", friendlyError(headerError));
          }

          const rawFromAddresses = parseAddressList(getFirstHeader(headers, "from"));
          const rawFrom = rawFromAddresses[0] || emptyAddress();
          const from = mergeHeaderAddress(rawFrom, structuredFrom);

          const rawSender = parseAddressList(getFirstHeader(headers, "sender"))[0] || emptyAddress();
          const sender = hasAddressData(structuredSender)
            ? structuredSender
            : hasAddressData(rawSender)
              ? rawSender
              : from;

          const rawTo = getAllHeaders(headers, "to").flatMap(parseAddressList);
          const toResult = resolveToRecipients(rawTo, structuredTo, headers);

          if (!toResult.recipients.length) {
            const mailboxFallback = getMailboxProfileAddress();
            if (hasAddressData(mailboxFallback)) {
              toResult.recipients = [mailboxFallback];
              toResult.source = "Mailbox profile fallback";
            }
          }

          data = {
            itemId: selectedItem.itemId,
            from,
            sender,
            to: toResult.recipients,
            toSource: toResult.source,
            subject: loadedItem.subject || selectedItem.subject || "",
            internetMessageId: String(loadedItem.internetMessageId || "").trim(),
            replyTo: formatAddressList(getAllHeaders(headers, "reply-to").flatMap(parseAddressList)),
            returnPath: cleanHeaderValue(getFirstHeader(headers, "return-path")),
            dateHeader: cleanHeaderValue(getFirstHeader(headers, "date")),
          };
        } catch (error) {
          extractionError = error;
        }

        loadedItem.unloadAsync((unloadResult) => {
          if (unloadResult.status === Office.AsyncResultStatus.Failed) {
            reject(new Error(`Could not unload the message: ${unloadResult.error.message}`));
            return;
          }

          if (extractionError) {
            reject(extractionError);
            return;
          }

          resolve(data);
        });
      })().catch((error) => {
        loadedItem.unloadAsync(() => reject(error));
      });
    });
  });
}

function getEmlHeaderBlockWithTimeout(loadedItem, timeoutMs) {
  return new Promise((resolve, reject) => {
    if (typeof loadedItem.getAsFileAsync !== "function") {
      reject(new Error("EML API is not available for this message."));
      return;
    }

    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      reject(new Error(`EML header lookup timed out after ${Math.round(timeoutMs / 1000)} seconds.`));
    }, timeoutMs);

    loadedItem.getAsFileAsync((result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);

      if (result.status === Office.AsyncResultStatus.Failed) {
        reject(new Error(result.error.message));
        return;
      }

      try {
        resolve(extractHeaderBlockFromBase64(String(result.value || "")));
      } catch (error) {
        reject(error);
      }
    });
  });
}

function extractHeaderBlockFromBase64(base64Value) {
  const clean = String(base64Value || "").replace(/\s+/g, "");
  if (!clean) return "";

  // MIME headers are at the beginning of the EML. Decode only the first
  // ~192 KiB of message data so a large attachment doesn't waste memory.
  const maxBase64Chars = 262144; // must remain divisible by 4
  const sliceLength = Math.min(clean.length, maxBase64Chars - (maxBase64Chars % 4));
  const binary = atob(clean.slice(0, sliceLength - (sliceLength % 4)));

  const crlfBoundary = binary.indexOf("\r\n\r\n");
  const lfBoundary = binary.indexOf("\n\n");
  let boundary = -1;
  if (crlfBoundary >= 0) boundary = crlfBoundary;
  else if (lfBoundary >= 0) boundary = lfBoundary;

  const headerBytes = boundary >= 0 ? binary.slice(0, boundary) : binary;
  const bytes = Uint8Array.from(headerBytes, (char) => char.charCodeAt(0));

  try {
    return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
  } catch (_error) {
    return headerBytes;
  }
}

function getMailboxProfileAddress() {
  const profile = Office.context?.mailbox?.userProfile;
  if (!profile) return emptyAddress();
  return {
    email: String(profile.emailAddress || "").trim(),
    name: String(profile.displayName || "").trim(),
  };
}

function normaliseAddress(entry) {
  if (!entry) return emptyAddress();
  return {
    email: String(entry.emailAddress || "").trim(),
    name: String(entry.displayName || "").trim(),
  };
}

function normaliseAddressArray(entries) {
  return Array.isArray(entries)
    ? entries.map(normaliseAddress).filter(hasAddressData)
    : [];
}

function emptyAddress() {
  return { email: "", name: "" };
}

function hasAddressData(address) {
  return Boolean(address && (address.email || address.name));
}

function mergeHeaderAddress(headerAddress, structuredAddress) {
  if (!hasAddressData(headerAddress)) return structuredAddress || emptyAddress();

  return {
    email: headerAddress.email || structuredAddress?.email || "",
    name: headerAddress.name || findStructuredNameForEmail(headerAddress.email, [structuredAddress]) || structuredAddress?.name || "",
  };
}

function findStructuredNameForEmail(email, structuredAddresses) {
  const target = String(email || "").trim().toLowerCase();
  if (!target) return "";
  const match = (structuredAddresses || []).find(
    (address) => String(address?.email || "").trim().toLowerCase() === target
  );
  return String(match?.name || "").trim();
}

function mergeHeaderRecipients(headerRecipients, structuredRecipients) {
  return dedupeAddresses((headerRecipients || []).map((address) => ({
    email: address.email || "",
    name: address.name || findStructuredNameForEmail(address.email, structuredRecipients),
  })).filter(hasAddressData));
}

function resolveToRecipients(rawTo, structuredTo, headers) {
  const mimeRecipients = mergeHeaderRecipients(rawTo, structuredTo);
  if (mimeRecipients.length) {
    return { recipients: mimeRecipients, source: "MIME To" };
  }

  // If the normal MIME To header is absent, delivery headers are the next-best
  // source for spam and forwarded mail. These usually contain an email address
  // but not a display name.
  const deliveryCandidates = [
    ["x-original-to", "X-Original-To"],
    ["delivered-to", "Delivered-To"],
    ["envelope-to", "Envelope-To"],
    ["x-envelope-to", "X-Envelope-To"],
  ];

  for (const [headerName, source] of deliveryCandidates) {
    const recipients = getAllHeaders(headers, headerName).flatMap(parseAddressList).filter(hasAddressData);
    if (recipients.length) {
      return { recipients: mergeHeaderRecipients(recipients, structuredTo), source };
    }
  }

  if (structuredTo.length) {
    return { recipients: structuredTo, source: "Outlook To fallback" };
  }

  return { recipients: [], source: "" };
}

function parseInternetHeaders(rawHeaders) {
  const map = new Map();
  const unfolded = String(rawHeaders || "").replace(/\r?\n[ \t]+/g, " ");
  for (const line of unfolded.split(/\r?\n/)) {
    const match = line.match(/^([^:]+):\s*(.*)$/);
    if (!match) continue;
    const name = match[1].trim().toLowerCase();
    const value = match[2].trim();
    if (!map.has(name)) map.set(name, []);
    map.get(name).push(value);
  }
  return map;
}

function getFirstHeader(headers, name) {
  return getAllHeaders(headers, name)[0] || "";
}

function getAllHeaders(headers, name) {
  return headers.get(String(name).toLowerCase()) || [];
}

function parseAddressList(value) {
  const text = cleanHeaderValue(value);
  if (!text) return [];

  const parts = splitAddressList(text);
  const addresses = [];

  for (let part of parts) {
    part = part.trim();
    if (!part) continue;

    // Remove RFC group labels while retaining any addresses after the colon.
    if (part.includes(":")) {
      const colon = part.indexOf(":");
      const before = part.slice(0, colon);
      const after = part.slice(colon + 1);
      if (!before.includes("@") && after) part = after;
    }
    part = part.replace(/;\s*$/, "").trim();

    const angle = part.match(/^(.*)<\s*([^<>\s]+@[^<>\s]+)\s*>\s*$/);
    if (angle) {
      addresses.push({
        email: stripAddressPunctuation(angle[2]),
        name: cleanDisplayName(angle[1]),
      });
      continue;
    }

    const emailMatch = part.match(/[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    if (emailMatch) {
      const email = stripAddressPunctuation(emailMatch[0]);
      const name = cleanDisplayName(part.replace(emailMatch[0], "").replace(/[<>]/g, " "));
      addresses.push({ email, name });
    }
  }

  return dedupeAddresses(addresses.filter(hasAddressData));
}

function splitAddressList(text) {
  const parts = [];
  let current = "";
  let quoted = false;
  let angleDepth = 0;
  let escaped = false;

  for (const ch of text) {
    if (escaped) {
      current += ch;
      escaped = false;
      continue;
    }
    if (ch === "\\" && quoted) {
      current += ch;
      escaped = true;
      continue;
    }
    if (ch === '"') quoted = !quoted;
    if (!quoted && ch === "<") angleDepth += 1;
    if (!quoted && ch === ">" && angleDepth > 0) angleDepth -= 1;

    if (ch === "," && !quoted && angleDepth === 0) {
      parts.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  if (current) parts.push(current);
  return parts;
}

function cleanDisplayName(value) {
  let name = String(value || "").trim();
  name = name.replace(/^"|"$/g, "").trim();
  name = decodeMimeWords(name);
  return name;
}

function decodeMimeWords(value) {
  return String(value || "").replace(/=\?([^?]+)\?([bqBQ])\?([^?]*)\?=/g, (_match, charset, encoding, data) => {
    try {
      let bytes;
      if (encoding.toUpperCase() === "B") {
        const binary = atob(data);
        bytes = Uint8Array.from(binary, (char) => char.charCodeAt(0));
      } else {
        const qp = data.replace(/_/g, " ").replace(/=([A-Fa-f0-9]{2})/g, (_m, hex) =>
          String.fromCharCode(parseInt(hex, 16))
        );
        bytes = Uint8Array.from(qp, (char) => char.charCodeAt(0));
      }

      const normalisedCharset = charset.toLowerCase().replace("us-ascii", "utf-8");
      return new TextDecoder(normalisedCharset).decode(bytes);
    } catch (_error) {
      return data;
    }
  });
}

function stripAddressPunctuation(value) {
  return String(value || "").trim().replace(/^[<\s]+|[>;,\s]+$/g, "");
}

function dedupeAddresses(addresses) {
  const seen = new Set();
  const result = [];
  for (const address of addresses) {
    const key = `${String(address.email || "").toLowerCase()}|${String(address.name || "").toLowerCase()}`;
    if (seen.has(key)) continue;
    seen.add(key);
    result.push(address);
  }
  return result;
}

function cleanHeaderValue(value) {
  return decodeMimeWords(String(value || "").replace(/\s+/g, " ").trim());
}

function formatAddressList(addresses) {
  return addresses
    .map((address) => {
      if (address.name && address.email) return `${address.name} <${address.email}>`;
      return address.email || address.name;
    })
    .filter(Boolean)
    .join("; ");
}

function sortMessages(messages) {
  return [...messages].sort((a, b) => {
    const fromCompare = (a.from.email || "").localeCompare(b.from.email || "", undefined, { sensitivity: "base" });
    if (fromCompare !== 0) return fromCompare;

    const senderCompare = (a.sender.email || "").localeCompare(b.sender.email || "", undefined, { sensitivity: "base" });
    if (senderCompare !== 0) return senderCompare;

    return (a.subject || "").localeCompare(b.subject || "", undefined, { sensitivity: "base" });
  });
}

function buildExportRows(messages, layout, selectedFields) {
  const rows = [];

  for (const message of messages) {
    const toRecipients = message.to.length ? message.to : [emptyAddress()];

    if (layout === "recipient") {
      for (const recipient of toRecipients) {
        rows.push(selectedFields.map((field) => fieldValue(message, field.key, recipient, true)));
      }
    } else {
      rows.push(selectedFields.map((field) => fieldValue(message, field.key, null, false)));
    }
  }

  return rows;
}

function fieldValue(message, key, recipient, perRecipient) {
  switch (key) {
    case "fromEmail": return message.from.email;
    case "fromName": return message.from.name;
    case "senderEmail": return message.sender.email;
    case "senderName": return message.sender.name;
    case "toEmail":
      return perRecipient
        ? recipient?.email || ""
        : message.to.map((entry) => entry.email).filter(Boolean).join("; ");
    case "toName":
      return perRecipient
        ? recipient?.name || ""
        : message.to.map((entry) => entry.name).filter(Boolean).join("; ");
    case "recipientSource": return message.toSource;
    case "subject": return message.subject;
    case "replyTo": return message.replyTo;
    case "returnPath": return message.returnPath;
    case "internetMessageId": return message.internetMessageId;
    case "dateHeader": return message.dateHeader;
    default: return "";
  }
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function renderErrors() {
  clearErrors();
  if (!state.failures.length) return;

  els.errorSummary.textContent = `${state.failures.length} selected message(s) could not be included.`;
  for (const failure of state.failures.slice(0, 20)) {
    const li = document.createElement("li");
    li.textContent = `${failure.subject}: ${failure.message}`;
    els.errorList.appendChild(li);
  }
  if (state.failures.length > 20) {
    const li = document.createElement("li");
    li.textContent = `…and ${state.failures.length - 20} more.`;
    els.errorList.appendChild(li);
  }
  els.errorCard.classList.remove("hidden");
}

function clearErrors() {
  els.errorList.replaceChildren();
  els.errorCard.classList.add("hidden");
}

function clearCompletion() {
  els.completionText.textContent = "";
  els.completionText.classList.add("hidden");
}

function timestampForFilename() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

function setSelectionUi(count, enabled, hint) {
  els.selectionCount.textContent = count === 1 ? "1 message" : `${count} messages`;
  els.selectionHint.textContent = hint;
  state.selectedItems = enabled ? state.selectedItems : [];
  updateExportButton();
}

function updateExportButton() {
  if (!els.exportButton) return;
  els.exportButton.disabled = state.busy || state.selectedItems.length === 0 || getSelectedFields().length === 0;
}

function setBusyUi(busy) {
  document.querySelectorAll('input[name="exportField"], input[name="recipientLayout"]').forEach((input) => {
    input.disabled = busy;
  });
  [els.defaultFieldsButton, els.selectAllFieldsButton, els.clearFieldsButton].forEach((button) => {
    button.disabled = busy;
  });

  els.progressWrap.classList.toggle("hidden", !busy);
  if (!busy) {
    els.progressBar.value = 0;
    els.progressText.textContent = "";
  }
  updateExportButton();
}

function updateProgress(done, total, text) {
  els.progressBar.max = total || 1;
  els.progressBar.value = done;
  els.progressText.textContent = text;
}

function friendlyError(error) {
  if (!error) return "Unknown error";
  return error.message || String(error);
}
