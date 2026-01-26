/* global Office, document, console, localStorage */

const ATTR_KEY = "inboxAgent.attributes";
const LOCAL_ATTR_KEY = "inboxAgent.local.";
const NOTIFICATION_ID = "AportioAttributes";
const CATEGORY_OPTIONS = ["Product", "Support", "Sales", "Billing"];
const PRIORITY_OPTIONS = ["Low", "Normal", "High"];
const DEFAULT_SENTIMENT_OPTIONS = ["Negative", "Neutral", "Positive"];
const NOTIFICATION_ICON = "icon16";
const API_BASE_URL = "https://attorney-healthy-weddings-attacked.trycloudflare.com";
const ANALYZE_URL = "https://attorney-healthy-weddings-attacked.trycloudflare.com/analyze";
const API_KEY = "24324ddadasadasdasdawqeqwewewqqwewqzx";
const REPORTS_ENDPOINT = "/status";
//mock backend
const DEFAULT_REPORTS_BASE_URL =
  "https://attorney-healthy-weddings-attacked.trycloudflare.com";
const REPORTS_BEARER_TOKEN = process.env.REPORTS_BEARER_TOKEN || "";
const INBOXAGENT_BASE_URL = "https://alignment-usage-prescribed-airline.trycloudflare.com";
const INBOXAGENT_CUSTOMER_ID = "K7BE";
const INBOXAGENT_API_KEY_ID = process.env.INBOXAGENT_API_KEY_ID || "";
const INBOXAGENT_API_KEY_SECRET = process.env.INBOXAGENT_API_KEY_SECRET || "";
const TEMPLATE_HTML = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title></title>
  <style>
    body, p {font-family: Arial, Helvetica, sans-serif;}
  </style>
</head>
<body>
  {{content}}
</body>
</html>`;

function getItem() {
  return Office.context?.mailbox?.item;
}

function isComposeItem(item) {
  return Boolean(item?.body?.setAsync);
}

function updateStatus(text) {
  const status = document.getElementById("statusText");
  if (status) {
    status.textContent = text;
  }
}

function setAnalysisSummary(text) {
  const summary = document.getElementById("summary");
  if (summary) {
    summary.textContent = text;
  }
}

function setSuggestedActions(actions) {
  const container = document.getElementById("suggestedActions");
  if (!container) return;
  container.innerHTML = "";
  if (!actions || actions.length === 0) {
    container.textContent = "No suggested actions.";
    return;
  }
  actions.forEach((action) => {
    const chip = document.createElement("span");
    chip.className = "chip";
    chip.textContent = action;
    container.appendChild(chip);
  });
}

function setSuggestedActionsStatus(text) {
  const container = document.getElementById("suggestedActions");
  if (container) {
    container.textContent = text;
  }
}

function normalizeSuggestedActions(value) {
  if (!value) return [];
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry)).filter(Boolean);
  }
  if (typeof value === "string") {
    return value
      .split(/[,;]\s*/g)
      .map((entry) => entry.trim())
      .filter(Boolean);
  }
  return [];
}

function setFormValues(attrs) {
  const { category, sentiment, priority, assignee } = attrs || {};
  if (category !== undefined) document.getElementById("category").value = category || "";
  if (sentiment !== undefined) document.getElementById("sentiment").value = sentiment || "";
  if (priority !== undefined) document.getElementById("priority").value = priority || "";
  if (assignee !== undefined) document.getElementById("assignee").value = assignee || "";
}

function normalizeSelectValue(selectId, value) {
  if (!value) return "";
  const select = document.getElementById(selectId);
  if (!select) return value;
  const lower = String(value).toLowerCase();
  const match = Array.from(select.options).find(
    (opt) => String(opt.value).toLowerCase() === lower
  );
  return match ? match.value : value;
}

function ensureSelectOption(selectId, value) {
  if (!value) return;
  const select = document.getElementById(selectId);
  if (!select) return;
  const exists = Array.from(select.options).some((opt) => opt.value === value);
  if (!exists) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  }
}

function applyAnalysisToForm(data) {
  const category = normalizeSelectValue("category", data.category || "");
  const sentiment = normalizeSelectValue("sentiment", data.sentiment || "");
  const priority = normalizeSelectValue("priority", data.priority || "");
  const assignee = data.assignee || "";

  ensureSelectOption("category", category);
  ensureSelectOption("sentiment", sentiment);
  ensureSelectOption("priority", priority);

  setFormValues({ category, sentiment, priority, assignee });
}

function readFormValues() {
  return {
    category: document.getElementById("category").value,
    sentiment: document.getElementById("sentiment").value,
    priority: document.getElementById("priority").value,
    assignee: document.getElementById("assignee").value.trim(),
  };
}

function buildSummary(attrs) {
  const parts = [];
  if (attrs.category) parts.push(`Category: ${attrs.category}`);
  if (attrs.sentiment) parts.push(`Sentiment: ${attrs.sentiment}`);
  if (attrs.priority) parts.push(`Priority: ${attrs.priority}`);
  if (attrs.assignee) parts.push(`Assignee: ${attrs.assignee}`);
  return parts.length ? parts.join(" | ") : "No attributes set";
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function renderTemplate(attrs) {
  const summary = buildSummary(attrs);
  const includeSummary = summary !== "No attributes set";
  const summaryBlock = includeSummary ? `<p>${escapeHtml(summary)}</p>` : "";
  const content = `<p>Hi</p><p>This is a template reply.</p>${summaryBlock}<p>Thanks</p>`;
  return TEMPLATE_HTML.replace("{{content}}", content).replace("{{footer}}", "");
}

function renderTemplateById(templateId, attrs) {
  if (templateId === "processing") {
    const summary = buildSummary(attrs);
    const includeSummary = summary !== "No attributes set";
    const summaryBlock = includeSummary ? `<p>${escapeHtml(summary)}</p>` : "";
    const content = `<p>Hi</p><p>I am working on this and will reply within 3 business days.</p>${summaryBlock}<p>Thanks</p>`;
    return TEMPLATE_HTML.replace("{{content}}", content).replace("{{footer}}", "");
  }

  return renderTemplate(attrs);
}

function getLocalKey() {
  const item = getItem();
  if (item?.itemId) {
    return `${LOCAL_ATTR_KEY}${item.itemId}`;
  }
  return null;
}

function saveLocalAttributes(attrs) {
  const key = getLocalKey();
  if (key) {
    localStorage.setItem(key, JSON.stringify(attrs));
  }
}

function loadLocalAttributes() {
  const key = getLocalKey();
  if (!key) return null;
  const raw = localStorage.getItem(key);
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function clearLocalAttributes() {
  const key = getLocalKey();
  if (key) {
    localStorage.removeItem(key);
  }
}

function withCustomProperties(handler) {
  return new Promise((resolve, reject) => {
    const item = getItem();
    if (!item) {
      reject(new Error("Mailbox item is not available."));
      return;
    }

    item.loadCustomPropertiesAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error || new Error("Unable to load custom properties."));
        return;
      }

      handler(result.value, resolve, reject);
    });
  });
}

function loadAttributes() {
  return withCustomProperties((props, resolve, reject) => {
    try {
      const raw = props.get(ATTR_KEY);
      if (!raw) {
        resolve(null);
        return;
      }
      resolve(JSON.parse(raw));
    } catch (err) {
      reject(err);
    }
  });
}

function saveAttributes(attrs) {
  return withCustomProperties((props, resolve, reject) => {
    props.set(ATTR_KEY, JSON.stringify(attrs));
    props.saveAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }
      reject(result.error || new Error("Unable to save attributes."));
    });
  });
}

function clearAttributes() {
  return withCustomProperties((props, resolve, reject) => {
    try {
      if (typeof props.remove === "function") {
        props.remove(ATTR_KEY);
      } else {
        props.set(ATTR_KEY, "null");
      }
      props.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        reject(result.error || new Error("Unable to clear attributes."));
      });
    } catch (err) {
      reject(err);
    }
  });
}

function pushNotification(message, type = "informationalMessage") {
  return new Promise((resolve) => {
    const item = getItem();
    if (!item?.notificationMessages?.replaceAsync) {
      resolve();
      return;
    }

    const details = { type, message };
    if (type === "informationalMessage") {
      details.icon = NOTIFICATION_ICON;
      details.persistent = true;
    }

    try {
      item.notificationMessages.replaceAsync(NOTIFICATION_ID, details, () => resolve());
    } catch (e) {
      console.warn("Notification not shown (ignored):", e);
      resolve();
    }
  });
}

function setSelectOptions(select, options) {
  if (!select) {
    return;
  }
  const placeholder = select.querySelector('option[value=""]');
  select.innerHTML = "";
  if (placeholder) {
    select.appendChild(placeholder);
  } else {
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "—";
    select.appendChild(blank);
  }
  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option;
    opt.textContent = option;
    select.appendChild(opt);
  });
}

function setSelectOptionsWithValue(select, options) {
  if (!select) {
    return;
  }
  const placeholder = select.querySelector('option[value=""]');
  select.innerHTML = "";
  if (placeholder) {
    select.appendChild(placeholder);
  } else {
    const blank = document.createElement("option");
    blank.value = "";
    blank.textContent = "—";
    select.appendChild(blank);
  }
  options.forEach((option) => {
    const opt = document.createElement("option");
    opt.value = option.value;
    opt.textContent = option.label;
    select.appendChild(opt);
  });
}

function getReportsFormValues() {
  const baseUrl = DEFAULT_REPORTS_BASE_URL;
  const token = REPORTS_BEARER_TOKEN;
  return { baseUrl, token };
}

function buildReportsUrl(baseUrl) {
  if (!baseUrl) return "";
  return `${baseUrl.replace(/\/+$/, "")}${REPORTS_ENDPOINT}`;
}


function formatReportLabel(item, index) {
  const subject = item?.subject || item?.email_subject || item?.title || item?.name;
  const id = item?.id || item?.pk || item?.uuid;
  if (subject && id) return `${subject} (#${id})`;
  if (subject) return subject;
  if (id) return `Report ${id}`;
  return `Report ${index + 1}`;
}

function normalizeReportItems(payload) {
  if (Array.isArray(payload?.data?.reporter)) return payload.data.reporter;
  return [];
}

function setLatestEmailSummary(text) {
  const el = document.getElementById("latestEmailSummary");
  if (!el) return;
  el.textContent = text || "—";
}

function formatEmailFrom(value) {
  if (!value) return "";
  if (typeof value === "string") return value;
  const name = value.name || value.display_name || value.full_name || "";
  const email = value.email || value.address || value.email_address || "";
  if (name && email) return `${name} <${email}>`;
  return name || email || "";
}

function formatEmailDate(value) {
  if (!value) return "";
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return date.toLocaleString();
  }
  return String(value);
}

function formatEmailSummary(item) {
  if (!item) return "—";
  const subject = item?.subject || item?.email_subject || item?.title || "";
  const from = formatEmailFrom(item?.from || item?.from_address || item?.sender || item?.email_from);
  const date = formatEmailDate(item?.date || item?.sent_at || item?.created_at || item?.received_at);
  const parts = [];
  if (subject) parts.push(subject);
  if (from) parts.push(`from ${from}`);
  if (date) parts.push(date);
  return parts.length ? parts.join(" · ") : "—";
}

function getFromAddress(item) {
  const from = item?.from;
  return normalizeEmailAddress(from);
}

function getMessageId(item) {
  const raw = item?.internetMessageId || item?.messageId || "";
  return String(raw || "");
}

function getDateSent(item) {
  const candidate = item?.dateTimeSent || item?.dateTimeCreated || new Date().toISOString();
  try {
    return new Date(candidate).toISOString();
  } catch {
    return new Date().toISOString();
  }
}

async function loadReportEmails() {
  const select = document.getElementById("reportEmailsSelect");
  if (select) {
    setSelectOptionsWithValue(select, []);
  }

  const { baseUrl, token } = getReportsFormValues();
  if (!baseUrl || !token) {
    updateStatus("reports: base url and token required");
    return;
  }

  updateStatus("reports: loading...");

  const url = buildReportsUrl(baseUrl);
  try {
    const response = await fetch(url, {
      headers: {
        Accept: "application/json",
        Authorization: `Bearer ${token}`,
      },
    });

    if (!response.ok) {
      updateStatus(`reports: request failed (${response.status})`);
      return;
    }

    const data = await response.json();
    const items = normalizeReportItems(data);
    const options = items.map((item, index) => {
      const id = String(index + 1);
      return {
        value: String(id),
        label: String(item),
      };
    });
    if (select) {
      setSelectOptionsWithValue(select, options);
    }
    updateStatus(`reports: loaded ${items.length}`);
  } catch (err) {
    console.error("Report fetch failed", err);
    updateStatus("reports: request error");
  }
}

function updateItemCategories(selected, managed) {
  return new Promise((resolve, reject) => {
    const item = getItem();
    if (!item?.categories?.getAsync) {
      resolve();
      return;
    }

    item.categories.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error || new Error("Unable to read categories."));
        return;
      }

      const normalizeCategory = (name) => String(name || "").trim().toLowerCase();
      const current = (result.value || []).map((name) => String(name || ""));
      const selectedSet = new Set(selected.filter(Boolean).map(normalizeCategory));
      const managedSet = new Set(managed.map(normalizeCategory));
      const remaining = current.filter((name) => {
        const normalized = normalizeCategory(name);
        if (!managedSet.has(normalized)) {
          return true;
        }
        return selectedSet.has(normalized);
      });
      const toRemove = current.filter(
        (name) => managedSet.has(normalizeCategory(name)) && !selectedSet.has(normalizeCategory(name))
      );
      const toAdd = selected.filter(
        (name) =>
          name &&
          !current.some((existing) => normalizeCategory(existing) === normalizeCategory(name))
      );

      const finish = () => resolve();
      const addCategories = () => {
        if (!toAdd.length) {
          finish();
          return;
        }
        item.categories.addAsync(toAdd, (addResult) => {
          if (addResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(addResult.error || new Error("Unable to add category."));
            return;
          }
          finish();
        });
      };

      if (!toRemove.length) {
        addCategories();
        return;
      }

      const removeThenAdd = () => {
        item.categories.removeAsync(toRemove, (removeResult) => {
          if (removeResult.status !== Office.AsyncResultStatus.Succeeded) {
            reject(removeResult.error || new Error("Unable to remove categories."));
            return;
          }
          addCategories();
        });
      };

      if (typeof item.categories.setAsync === "function") {
        item.categories.setAsync(remaining, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
            return;
          }
          // Fallback for hosts that report setAsync but do not allow it.
          removeThenAdd();
        });
        return;
      }

      removeThenAdd();
    });
  });
}

function clearAllCategories() {
  return new Promise((resolve, reject) => {
    const item = getItem();
    if (!item?.categories?.getAsync) {
      resolve();
      return;
    }

    item.categories.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(result.error || new Error("Unable to read categories."));
        return;
      }

      const current = (result.value || [])
        .map((name) => String(name || "").trim())
        .filter(Boolean);
      if (!current.length) {
        resolve();
        return;
      }

      item.categories.removeAsync(current, (removeResult) => {
        if (removeResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }

        if (typeof item.categories.setAsync === "function") {
          item.categories.setAsync([], (setResult) => {
            if (setResult.status === Office.AsyncResultStatus.Succeeded) {
              resolve();
              return;
            }
            reject(setResult.error || new Error("Unable to clear categories."));
          });
          return;
        }

        reject(removeResult.error || new Error("Unable to clear categories."));
      });
    });
  });
}

function updateItemImportance(priority) {
  return new Promise((resolve, reject) => {
    const item = getItem();
    if (!isComposeItem(item)) {
      resolve();
      return;
    }
    if (!item?.importance?.setAsync) {
      resolve();
      return;
    }

    const value = priority ? priority.toLowerCase() : "normal";
    item.importance.setAsync(value, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
        return;
      }
      reject(result.error || new Error("Unable to set importance."));
    });
  });
}

function isUnsupportedHostError(err) {
  if (!err) return false;
  const message = String(err.message || "").toLowerCase();
  return err.code === 5000 || message.includes("not supported");
}

function clearNotification() {
  return new Promise((resolve) => {
    const item = getItem();
    if (!item?.notificationMessages) {
      resolve();
      return;
    }
    const notifications = item.notificationMessages;
    const fallbackReplace = () => {
      if (typeof notifications.replaceAsync !== "function") {
        resolve();
        return;
      }
      try {
        notifications.replaceAsync(
          NOTIFICATION_ID,
          { type: "informationalMessage", message: "" },
          () => resolve()
        );
      } catch (err) {
        console.warn("Notification clear fallback failed (ignored):", err);
        resolve();
      }
    };
    if (typeof notifications.removeAsync === "function") {
      notifications.removeAsync(NOTIFICATION_ID, (result) => {
        if (result?.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
          return;
        }
        fallbackReplace();
      });
      return;
    }
    fallbackReplace();
  });
}

async function hydrateForm(options = {}) {
  const { preserveForm = false, allowLocalFallback = true, notify = true } = options;
  const item = getItem();
  if (item?.subject) {
    setAnalysisSummary(`Analyzing: ${item.subject}`);
  } else {
    setAnalysisSummary("Analyzing...");
  }
  setSuggestedActionsStatus("Analyzing...");
  runAiAnalysis();

  try {
    const sentimentSelect = document.getElementById("sentiment");
    setSelectOptions(sentimentSelect, DEFAULT_SENTIMENT_OPTIONS);

    updateStatus("loading saved attributes...");
    let attrs = await loadAttributes();
    if (!attrs && allowLocalFallback) {
      attrs = loadLocalAttributes();
      if (attrs) {
        updateStatus("loaded local attributes");
        if (notify) {
          await pushNotification(`Restored local attributes • ${buildSummary(attrs)}`);
        }
      }
    }
    if (attrs) {
      if (!preserveForm) {
        setFormValues(attrs);
      }
      updateStatus("loaded saved attributes");
      if (notify) {
        await pushNotification(`Restored attributes • ${buildSummary(attrs)}`);
      }
    } else {
      updateStatus("ready");
    }
  } catch (err) {
    console.error("Failed to load saved attributes", err);
    updateStatus("ready (attributes not loaded)");
  }
}

async function handleApply() {
  const attrs = readFormValues();
  updateStatus("applying attributes...");

  try {
    let savedToMailbox = true;
    try {
      await saveAttributes(attrs);
    } catch (err) {
      if (isUnsupportedHostError(err)) {
        savedToMailbox = false;
      } else {
        throw err;
      }
    }
    const selectedCategories = [];
    if (attrs.category) selectedCategories.push(attrs.category);
    const managedCategories = CATEGORY_OPTIONS;
    try {
      await updateItemCategories(selectedCategories, managedCategories);
    } catch (err) {
      if (!isUnsupportedHostError(err)) {
        throw err;
      }
    }
    try {
      await updateItemImportance(attrs.priority);
    } catch (err) {
      console.warn("Failed to set importance.", err);
      await pushNotification(
        "Priority not updated in Outlook (importance update failed).",
        "errorMessage"
      );
    }
    saveLocalAttributes(attrs);
    const summary = buildSummary(attrs);
    updateStatus(
      savedToMailbox
        ? "attributes saved to this item"
        : "attributes saved locally (mailbox properties not supported)"
    );
    await pushNotification(`Aportio saved • ${summary}`);
  } catch (err) {
    console.error("Failed to apply attributes", err);
    saveLocalAttributes(attrs);
    const errorText = err && err.message ? err.message : String(err);
    updateStatus("saved locally (server save failed)");
    await pushNotification(
      `Saved locally only (could not save to mailbox): ${errorText}`,
      "errorMessage"
    );
  }
}

async function handleClear() {
  updateStatus("clearing attributes...");
  const errors = [];
  const recordError = (label, err) => {
    errors.push(label);
    console.error(`Failed to clear ${label}`, err);
  };

  try {
    await clearAttributes();
  } catch (err) {
    recordError("stored attributes", err);
  }

  try {
    clearLocalAttributes();
  } catch (err) {
    recordError("local attributes", err);
  }

  try {
    await clearAllCategories();
  } catch (err) {
    recordError("categories", err);
  }

  try {
    await updateItemImportance("");
  } catch (err) {
    recordError("importance", err);
  }

  setFormValues({
    category: "",
    sentiment: "",
    priority: "",
    assignee: "",
  });
  await clearNotification();

  if (errors.length) {
    updateStatus(`clear done (failed: ${errors.join(", ")})`);
  } else {
    updateStatus("attributes cleared");
  }
}

function handleSnooze() {
  updateStatus("snoozed (demo)");
}

function handleQuickAction(text) {
  updateStatus(text);
}

function getItemBodyText() {
  return new Promise((resolve) => {
    const item = getItem();
    if (!item?.body?.getAsync) {
      resolve("");
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }
      resolve("");
    });
  });
}


function getItemBodyHtml() {
  return new Promise((resolve) => {
    const item = getItem();
    if (!item?.body?.getAsync) {
      resolve("");
      return;
    }
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
        return;
      }
      resolve("");
    });
  });
}

function buildInboxagentUrl(path) {
  return `${INBOXAGENT_BASE_URL.replace(/\/+$/, "")}${path}`;
}

function buildBasicAuthHeader(id, secret) {
  if (!id || !secret) return "";
  return `Basic ${btoa(`${id}:${secret}`)}`;
}

function normalizeEmailAddress(entry) {
  if (!entry) return "";
  if (typeof entry === "string") return entry;
  return entry.emailAddress || entry.address || entry.name || "";
}

function joinEmailAddresses(entries) {
  if (!entries) return "";
  return entries
    .map((entry) => normalizeEmailAddress(entry))
    .map((value) => String(value || "").trim())
    .filter(Boolean)
    .join(", ");
}

function readRecipientsAsync(item, field) {
  return new Promise((resolve) => {
    const value = item?.[field];
    if (!value) {
      resolve("");
      return;
    }
    if (Array.isArray(value)) {
      resolve(joinEmailAddresses(value));
      return;
    }
    if (typeof value.getAsync === "function") {
      value.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(joinEmailAddresses(result.value || []));
          return;
        }
        resolve("");
      });
      return;
    }
    resolve("");
  });
}

async function getEmails() {
  if (!INBOXAGENT_BASE_URL || !INBOXAGENT_CUSTOMER_ID) {
    updateStatus("emails: base url or customer id missing");
    return;
  }
  const authHeader = buildBasicAuthHeader(INBOXAGENT_API_KEY_ID, INBOXAGENT_API_KEY_SECRET);
  if (!authHeader) {
    updateStatus("emails: API key missing");
    return;
  }

  updateStatus("emails: fetching...");

  try {
    const response = await fetch(
      "https://alignment-usage-prescribed-airline.trycloudflare.com/api/v2/reports/report_emails/?limit=50&offset=0&ordering=desc",
      {
        method: "GET",
        headers: {
          Authorization: authHeader,
        },
      }
    );

    if (!response.ok) {
      updateStatus(`emails failed (${response.status})`);
      return;
    }

    const data = await response.json();
    const count = Array.isArray(data?.results) ? data.results.length : "";
    const latest = Array.isArray(data?.results) ? data.results[0] : null;
    setLatestEmailSummary(formatEmailSummary(latest));
    updateStatus(count ? `emails ok (fetched: ${count})` : "emails ok");
    await pushNotification("Emails fetched", "informationalMessage");
  } catch (err) {
    console.error("Email fetch failed", err);
    setLatestEmailSummary("—");
    updateStatus("emails failed (network error)");
  }
}

async function syncToInboxagent() {
  const item = getItem();
  if (!item) {
    updateStatus("sync: no item selected");
    return;
  }
  if (!INBOXAGENT_BASE_URL || !INBOXAGENT_CUSTOMER_ID) {
    updateStatus("sync: base url or customer id missing");
    return;
  }
  const authHeader = buildBasicAuthHeader(INBOXAGENT_API_KEY_ID, INBOXAGENT_API_KEY_SECRET);
  if (!authHeader) {
    updateStatus("sync: API key missing");
    return;
  }

  updateStatus("sync: sending to InboxAgent...");

  const [plain, html, to, cc] = await Promise.all([
    getItemBodyText(),
    getItemBodyHtml(),
    readRecipientsAsync(item, "to"),
    readRecipientsAsync(item, "cc"),
  ]);

  const payload = {
    plain,
    html,
    headers: {
      date: getDateSent(item),
      from: getFromAddress(item),
      to,
      cc,
      subject: item.subject || "",
      message_id: getMessageId(item),
    },
  };

  try {
    const response = await fetch(
      buildInboxagentUrl(`/api/v4/features/email/${encodeURIComponent(INBOXAGENT_CUSTOMER_ID)}/classify/`),
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: authHeader,
        },
        body: JSON.stringify(payload),
      }
    );

    if (!response.ok) {
      updateStatus(`sync failed (${response.status})`);
      return;
    }

    const data = await response.json();
    const requestId = data?.id || data?.metadata?.request_id || "";
    updateStatus(requestId ? `sync ok (id: ${requestId})` : "sync ok");
    await pushNotification("InboxAgent synced", "informationalMessage");
  } catch (err) {
    console.error("InboxAgent sync failed", err);
    updateStatus("sync failed (network error)");
  }
}

async function runAiAnalysis(options = {}) {
  const { applyToForm = false } = options;
  const item = getItem();
  if (!item) {
    setAnalysisSummary("AI analysis not available.");
    setSuggestedActions([]);
    return;
  }

  setAnalysisSummary("Analyzing...");
  setSuggestedActionsStatus("Analyzing...");

  const subject = item.subject || "";
  const body = await getItemBodyText();

  try {
    const response = await fetch(ANALYZE_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, body, api_key: API_KEY }),
    });
    if (!response.ok) {
      throw new Error(`Analyze request failed (${response.status})`);
    }

    const data = await response.json();
    const summaryText =
      data.rationale ||
      data.reply_suggestion ||
      "No analysis returned.";
    const suggestedActions = normalizeSuggestedActions(
      data.actions && Array.isArray(data.actions) && data.actions.length
        ? data.actions
        : data.reply_suggestion
    );

    setAnalysisSummary(summaryText);
    setSuggestedActions(suggestedActions);

    if (applyToForm) {
      applyAnalysisToForm(data);
      updateStatus("analysis applied to attributes");
    }
  } catch (err) {
    console.error("AI analysis failed", err);
    setAnalysisSummary("AI analysis failed.");
    setSuggestedActions([]);
  }
}

async function handleAiReply() {
  const item = getItem();
  if (!item) {
    updateStatus("ai reply not available");
    return;
  }

  updateStatus("generating ai reply...");
  const subject = item.subject || "";
  const body = await getItemBodyText();

  try {
    const response = await fetch(`${API_BASE_URL}/reply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        subject,
        body,
        tone: "professional",
        api_key: API_KEY,
      }),
    });
    if (!response.ok) {
      throw new Error(`Reply request failed (${response.status})`);
    }
    const data = await response.json();
    const replyText = data.reply_text || "";
    if (!replyText) {
      throw new Error("Empty reply from server");
    }

    if (item.body?.setAsync) {
      const replyHtml = escapeHtml(replyText).replace(/\n/g, "<br>");
      item.body.setAsync(replyHtml, { coercionType: Office.CoercionType.Html }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateStatus("ai reply inserted");
          return;
        }
        updateStatus("ai reply insert failed");
      });
      return;
    }

    if (item.displayReplyForm) {
      item.displayReplyForm(replyText);
      updateStatus("ai reply draft opened");
      return;
    }

    updateStatus("ai reply not available");
  } catch (err) {
    console.error("AI reply failed", err);
    updateStatus("ai reply failed");
  }
}

function handleInsertTemplate() {
  const mailbox = Office.context?.mailbox;
  const attrs = readFormValues();
  const templateId = document.getElementById("templateSelect")?.value || "default";
  const html = renderTemplateById(templateId, attrs);

  if (mailbox?.displayNewMessageForm) {
    try {
      mailbox.displayNewMessageForm({ htmlBody: html });
      updateStatus("new email opened");
      return;
    } catch (err) {
      console.error("New email create failed", err);
      updateStatus("new email open failed");
      return;
    }
  }

  const item = getItem();
  if (!item?.body?.setAsync) {
    updateStatus("template insert not available");
    return;
  }

  const itemClass = item.itemClass || "";
  const shouldReplace =
    Boolean(item.inReplyTo) || /reply|forward/i.test(itemClass);

  const done = (result) => {
    if (result?.status === Office.AsyncResultStatus.Succeeded) {
      updateStatus("template inserted");
      return;
    }
    const err = result?.error || new Error("Unable to insert template.");
    console.error("Template insert failed", err);
    updateStatus("template insert failed");
  };

  if (shouldReplace) {
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, done);
    return;
  }

  if (!item.body.appendAsync) {
    item.body.setAsync(html, { coercionType: Office.CoercionType.Html }, done);
    return;
  }

  item.body.appendAsync(html, { coercionType: Office.CoercionType.Html }, done);
}

function handleReplyWithTemplate() {
  const item = getItem();
  if (!item?.displayReplyForm) {
    updateStatus("reply with template not available");
    return;
  }

  const attrs = readFormValues();
  const templateId = document.getElementById("templateSelect")?.value || "default";
  const html = renderTemplateById(templateId, attrs);
  try {
    item.displayReplyForm(html);
    updateStatus("reply draft opened");
  } catch (err) {
    console.error("Reply with template failed", err);
    updateStatus("reply template failed");
  }
}

Office.onReady(() => {
  hydrateForm();
  loadReportEmails();

  const bindClick = (id, handler) => {
    const el = document.getElementById(id);
    if (!el) {
      return;
    }
    el.addEventListener("click", handler);
  };

  bindClick("btnApply", () => handleApply());
  bindClick("btnSync", () => syncToInboxagent());
  bindClick("btnGetEmails", () => getEmails());
  bindClick("btnSnooze", () => handleSnooze());
  bindClick("btnRefresh", () => {
    hydrateForm({ preserveForm: true, allowLocalFallback: false, notify: true });
  });
  bindClick("btnClear", () => handleClear());
  bindClick("btnWorkflow", () => handleQuickAction("workflow started (demo)"));
  bindClick("btnAnalyze", () => runAiAnalysis({ applyToForm: true }));
  bindClick("btnTemplate", () => handleInsertTemplate());
  bindClick("btnReplyTemplate", () => handleReplyWithTemplate());
  bindClick("btnReply", () => handleAiReply());
});
