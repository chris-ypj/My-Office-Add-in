/* global Office, window */

Office.onReady(() => {
  if (Office.actions && typeof Office.actions.associate === "function") {
    Office.actions.associate("onMessageReadHandler", onMessageReadHandler);
  }
});

/**
 * ExecuteFunction handler for Analyze command (V1_0 compatible).
 * Shows a notification banner as proof of execution.
 */
function onAnalyzeCommand(event) {
  const mailbox = Office.context && Office.context.mailbox;
  const item = mailbox && mailbox.item;
  if (!item || !item.notificationMessages) {
    event.completed();
    return;
  }

  const subject = item.subject || "(No subject)";
  item.notificationMessages.replaceAsync(
    "aportioAnalyzeCommand",
    {
      type: "informationalMessage",
      message: `Analyze triggered for: ${subject}`,
      icon: "icon16",
      persistent: false,
    },
    () => event.completed()
  );
}

/**
 * Event-based activation handler for OnMessageRead.
 * Shows a lightweight notification banner in Outlook.
 */
function onMessageReadHandler(event) {
  const mailbox = Office.context && Office.context.mailbox;
  const item = mailbox && mailbox.item;
  if (!item || !item.notificationMessages) {
    event.completed();
    return;
  }

  const subject = item.subject || "(No subject)";
  const message = `Aportio saw: ${subject}`;
  item.notificationMessages.replaceAsync(
    "aportioOnRead",
    {
      type: "informationalMessage",
      message,
      icon: "icon16",
      persistent: false,
    },
    () => event.completed()
  );
}

/**
 * Called via ExecuteFunction from the manifest.
 * Shows a dialog listing subject(s) of the selected email(s).
 */
function showSubjectPopup(event) {
  getSelectedSubjects()
    .then((subjects) => openSubjectsDialog(subjects))
    .catch((err) =>
      openSubjectsDialog([`Error: ${err && err.message ? err.message : String(err)}`])
    )
    .finally(() => event.completed());
}

function getSelectedSubjects() {
  return new Promise((resolve, reject) => {
    const mailbox = Office.context && Office.context.mailbox;
    if (!mailbox) {
      reject(new Error("Mailbox context not available."));
      return;
    }

    // Preferred: multi-select (Mailbox 1.13+)
    if (typeof mailbox.getSelectedItemsAsync === "function") {
      mailbox.getSelectedItemsAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(result.error || new Error("getSelectedItemsAsync failed."));
          return;
        }
        const items = Array.isArray(result.value) ? result.value : [];
        const subjects = items.map((i) => i && i.subject).filter(Boolean);

        // Fallback: current item subject
        if (subjects.length === 0 && mailbox.item && mailbox.item.subject) {
          resolve([mailbox.item.subject]);
          return;
        }

        resolve(subjects.length ? subjects : ["(No subject available)"]);
      });
      return;
    }

    // Fallback for clients without multi-select
    if (mailbox.item && mailbox.item.subject) {
      resolve([mailbox.item.subject]);
      return;
    }

    resolve(["(No subject available)"]);
  });
}

function openSubjectsDialog(subjects) {
  const base = `${window.location.origin}/dialog/dialog.html`;
  const payload = encodeURIComponent(JSON.stringify(subjects));
  const url = `${base}?subjects=${payload}`;

  Office.context.ui.displayDialogAsync(
    url,
    { height: 35, width: 30, displayInIframe: true },
    () => {}
  );
}

// Expose for Office to call
if (typeof window !== "undefined") {
  window.onMessageReadHandler = onMessageReadHandler;
  window.onAnalyzeCommand = onAnalyzeCommand;
  window.showSubjectPopup = showSubjectPopup;
}
