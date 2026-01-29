/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { airtableClient } from "../services/airtableClient";
import {
  AirtableAttachmentInput,
  AirtableDocumentPayload,
  AirtableProjectOption,
  AirtableTaskPayload,
  CollaboratorOption,
  OutlookAttachmentPreview,
  OutlookMessageMetadata,
} from "../types/airtable";

const attachmentGroups = [{ containerId: "task-attachment-choices", checkboxClass: "task-attachment-checkbox" }];
const MAX_ATTACHMENT_SIZE_BYTES = 5 * 1024 * 1024;
const projectInputs = [
  { inputId: "task-project-input", datalistId: "task-project-datalist" },
  { inputId: "document-project-input", datalistId: "document-project-datalist" },
  { inputId: "note-project-input", datalistId: "note-project-datalist" },
];
const LINK_TITLE_PROXY = (process.env.TITLE_PROXY_URL || "").trim();
const GRAPH_CLIENT_ID = (process.env.GRAPH_CLIENT_ID || "").trim();
const GRAPH_TENANT_ID = (process.env.GRAPH_TENANT_ID || "common").trim();
const GRAPH_REDIRECT_URI = (process.env.GRAPH_REDIRECT_URI || "").trim();

let attachments: OutlookAttachmentPreview[] = [];
let detectedLinks: string[] = [];
let messageMetadata: OutlookMessageMetadata | null = null;
let projectOptions: AirtableProjectOption[] = [];
let collaboratorOptions: CollaboratorOption[] = [];
let externalOptions: AirtableProjectOption[] = [];
let senderEmail: string | undefined;
let cachedGraphToken: string | null = null;
let notePersonTokens: string[] = [];
let msalInstance: any | null = null;
const linkTitleCache: Record<string, string> = {};

function getEligibleAttachments(): OutlookAttachmentPreview[] {
  return attachments.filter((att) => !att.isInline);
}

Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "block";
    await initializePane();
  }
});

async function initializePane() {
  wireUpForms();
  setupTabs();
  await Promise.all([hydrateContext(), loadProjects(), loadCollaborators(), loadExternalPersons()]);
}

function wireUpForms() {
  const taskForm = document.getElementById("task-form");
  if (taskForm) {
    taskForm.addEventListener("submit", handleTaskSubmit);
  }

  const documentForm = document.getElementById("document-form");
  if (documentForm) {
    documentForm.addEventListener("submit", handleDocumentSubmit);
  }

  const noteForm = document.getElementById("note-form");
  if (noteForm) {
    noteForm.addEventListener("submit", handleNoteSubmit);
  }

  const notePersonsInput = document.getElementById("note-persons") as HTMLInputElement | null;
  if (notePersonsInput) {
    const commitNotePersonInput = () => addNotePersonToken(notePersonsInput.value);
    notePersonsInput.addEventListener("change", commitNotePersonInput);
    notePersonsInput.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === "," || ev.key === ";") {
        ev.preventDefault();
        commitNotePersonInput();
      }
    });
  }

  const documentSource = document.getElementById("document-source");
  if (documentSource) {
    documentSource.addEventListener("change", toggleDocumentSource);
  }
  toggleDocumentSource();
}

function setupTabs() {
  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>(".tab-btn"));
  const panels = Array.from(document.querySelectorAll<HTMLElement>(".tab-panel"));

  const activate = (targetId: string) => {
    buttons.forEach((btn) => btn.classList.toggle("active", btn.dataset.tabTarget === targetId));
    panels.forEach((panel) => panel.classList.toggle("active", panel.id === targetId));
  };

  buttons.forEach((btn) => {
    btn.addEventListener("click", () => {
      const target = btn.dataset.tabTarget;
      if (target) {
        activate(target);
      }
    });
  });
}

async function hydrateContext() {
  const mailboxItem = Office.context.mailbox.item as Office.MessageRead;
  messageMetadata = buildMetadata(mailboxItem);
  senderEmail = mailboxItem.from?.emailAddress || undefined;
  attachments = extractAttachments(mailboxItem);
  detectedLinks = await getLinksFromBody(mailboxItem);
  await fetchLinkTitles(detectedLinks);

  renderMailHeader(messageMetadata);
  prefillFormDefaults(messageMetadata);
  renderAttachmentGroups();
  renderDocumentAttachmentSelect();
  renderLinkOptions();
  const bodyText = await getBodyAsText(mailboxItem);
  prefillBodyIntoDescription(bodyText);
  prefillNoteDefaults(bodyText);
}

async function loadProjects() {
  setProjectStatus("Projekte werden geladen …", "pending");
  try {
    projectOptions = await airtableClient.fetchProjects();
    // Filter Status und Sortierung in der Service-Schicht nicht möglich -> bereits erfolgt via API-Filter/Sort.
    if (projectOptions.length === 0) {
      setProjectStatus("Keine Projekte gefunden. Bitte bei Bedarf manuell eingeben.", "info");
      renderProjectSelects(true);
    } else {
      renderProjectSelects(false);
      setProjectStatus(`${projectOptions.length} Projekte geladen.`, "success");
    }
  } catch (error) {
    console.error(error);
    setProjectStatus(`Projekte konnten nicht geladen werden: ${(error as Error).message}`, "error");
    renderProjectSelects(true);
  }
}

async function loadCollaborators() {
  const ownerDatalist = document.getElementById("task-owner-datalist") as HTMLDataListElement | null;
  if (!ownerDatalist) {
    return;
  }
  try {
    const collaborators = await airtableClient.fetchCollaborators();
    collaboratorOptions = collaborators;
    ownerDatalist.innerHTML = "";
    collaborators.forEach((c) => {
      const option = document.createElement("option");
      option.value = c.name ?? c.email ?? "";
      option.label = "";
      if (c.id) {
        option.dataset.id = c.id;
      }
      if (c.email) {
        option.dataset.email = c.email;
      }
      ownerDatalist.appendChild(option);
    });
  } catch (error) {
    console.error(error);
    ownerDatalist.innerHTML = "";
  }
}

async function loadExternalPersons() {
  const externalInput = document.getElementById("task-external") as HTMLInputElement | null;
  const notePersonDatalist = document.getElementById("note-person-datalist") as HTMLDataListElement | null;
  if (!externalInput) {
    return;
  }
  try {
    externalOptions = await airtableClient.fetchExternalPersons();
    renderExternalOptions();
    prefillSenderAsExternal();
    if (notePersonDatalist) {
      notePersonDatalist.innerHTML = "";
      externalOptions.forEach((person) => {
        const option = document.createElement("option");
        option.value = person.name;
        option.label = person.email ?? "";
        option.dataset.id = person.id;
        if (person.email) option.dataset.email = person.email;
        notePersonDatalist.appendChild(option);
      });
    }
  } catch (error) {
    console.error("Externe Personen konnten nicht geladen werden:", error);
    externalOptions = [];
    renderExternalOptions();
  }
}

function buildMetadata(item: Office.MessageRead): OutlookMessageMetadata {
  const sender = item.from ? `${item.from.displayName ?? ""} <${item.from.emailAddress ?? ""}>`.trim() : "";
  return {
    itemId: item.itemId ?? "",
    subject: item.subject ?? "",
    from: sender,
    receivedDate: item.dateTimeCreated ? new Date(item.dateTimeCreated) : null,
    webLink: buildOutlookWebLink(item.itemId ?? ""),
  };
}

function renderMailHeader(metadata: OutlookMessageMetadata) {
  const subjectElement = document.getElementById("mail-subject");
  const fromElement = document.getElementById("mail-from");
  const dateElement = document.getElementById("mail-date");

  if (subjectElement) {
    subjectElement.textContent = metadata.subject || "Kein Betreff";
  }
  if (fromElement) {
    fromElement.textContent = metadata.from || "Unbekannter Absender";
  }
  if (dateElement) {
    dateElement.textContent = metadata.receivedDate
      ? metadata.receivedDate.toLocaleString()
      : "Kein Datum verfügbar";
  }
}

function extractAttachments(item: Office.MessageRead): OutlookAttachmentPreview[] {
  const rawAttachments = item.attachments ?? [];
  return rawAttachments.map((att) => ({
    id: att.id,
    name: att.name,
    contentType: att.contentType,
    size: att.size,
    isInline: att.isInline,
  }));
}

function renderAttachmentGroups() {
  attachmentGroups.forEach(({ containerId, checkboxClass }) => {
    const container = document.getElementById(containerId);
    if (!container) {
      return;
    }

    const eligible = getEligibleAttachments();

    if (eligible.length === 0) {
      container.innerHTML = "<p class=\"hint\">Keine Anhaenge gefunden (Inline-Bilder ausgeblendet).</p>";
      return;
    }

    container.innerHTML = "";
    eligible.forEach((attachment) => {
      const label = document.createElement("label");
      label.className = "choice attachment-item doc-row";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.className = checkboxClass;
      checkbox.value = attachment.id;
      checkbox.checked = true;

      const icon = document.createElement("span");
      icon.className = `ms-Icon ${getAttachmentIconClass(attachment.contentType, attachment.name)} att-icon`;

      const span = document.createElement("span");
      span.className = "doc-text";
      span.textContent = `${attachment.name} (${Math.round(attachment.size / 1024)} KB)`;

      label.appendChild(checkbox);
      label.appendChild(icon);
      label.appendChild(span);
      container.appendChild(label);
    });
  });
}

function renderDocumentAttachmentSelect() {
  const container = document.getElementById("document-attachment-choices");
  if (!container) {
    return;
  }
  const eligible = getEligibleAttachments();

  if (eligible.length === 0) {
    container.innerHTML = "<p class=\"hint\">Keine Anhaenge verfuegbar.</p>";
    return;
  }

  container.innerHTML = "";
  eligible.forEach((att) => {
    const label = document.createElement("label");
    label.className = "choice attachment-item doc-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.className = "doc-attachment-checkbox";
    checkbox.value = att.id;
    checkbox.checked = false;

    const icon = document.createElement("span");
    icon.className = `ms-Icon ${getAttachmentIconClass(att.contentType, att.name)} att-icon`;

    const span = document.createElement("span");
    span.className = "doc-text";
    span.textContent = `${att.name} (${Math.round(att.size / 1024)} KB)`;

    label.appendChild(checkbox);
    label.appendChild(icon);
    label.appendChild(span);
    container.appendChild(label);
  });
}

async function getLinksFromBody(item: Office.MessageRead): Promise<string[]> {
  const body = limitBodyText(await getBodyAsText(item), 12000);
  const matches = body.match(/https?:\/\/\S+/gim) ?? [];

  const counts: Record<string, number> = {};
  matches.forEach((m) => {
    const cleaned = m.replace(/[).,]+$/, "");
    counts[cleaned] = (counts[cleaned] || 0) + 1;
  });

  const cleaned = matches
    .map((match) => match.replace(/[).,]+$/, ""))
    .filter((url) => filterLink(url, counts[url] || 1));

  return Array.from(new Set(cleaned));
}

function renderLinkOptions() {
  const listElement = document.getElementById("link-preview");

  if (listElement) {
    if (detectedLinks.length === 0) {
      listElement.innerHTML = "<p class=\"hint\">Keine Links im Text gefunden.</p>";
    } else {
      listElement.innerHTML = "";
      detectedLinks.slice(0, 10).forEach((link) => {
        const li = document.createElement("div");
        li.className = "link-chip";
        li.textContent = getLinkTitle(link);
        listElement.appendChild(li);
      });
    }
  }

  const docLinkContainer = document.getElementById("document-link-choices");
  if (docLinkContainer) {
    if (detectedLinks.length === 0) {
      docLinkContainer.innerHTML = "<p class=\"hint\">Keine Links verfuegbar.</p>";
    } else {
      docLinkContainer.innerHTML = "";
      detectedLinks.slice(0, 10).forEach((link) => {
        const label = document.createElement("label");
        label.className = "choice doc-row";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "doc-link-checkbox";
        checkbox.value = link;

        const icon = document.createElement("span");
        icon.className = "ms-Icon ms-Icon--Link att-icon";

        const span = document.createElement("span");
        span.className = "doc-text";
        span.textContent = getLinkTitle(link);

        label.appendChild(checkbox);
        label.appendChild(icon);
        label.appendChild(span);
        docLinkContainer.appendChild(label);
      });
    }
  }
}

function filterLink(url: string, count: number): boolean {
  const lower = url.toLowerCase();
  if (lower.includes("safelinks.protection.outlook.com")) return false;
  if (lower.includes("cid:")) return false;
  if (lower.endsWith(".png") || lower.endsWith(".jpg") || lower.endsWith(".jpeg") || lower.endsWith(".gif")) {
    if (lower.includes("signature") || lower.includes("logo")) return false;
  }
  const sigDomains = ["linkedin.com", "facebook.com", "instagram.com", "twitter.com", "youtube.com", "vimeo.com"];
  if (count > 1 && sigDomains.some((d) => lower.includes(d))) return false;
  return true;
}

function getLinkTitle(link: string): string {
  if (linkTitleCache[link]) return linkTitleCache[link];
  try {
    const u = new URL(link);
    const path = u.pathname && u.pathname !== "/" ? u.pathname : "";
    return `${u.hostname}${path}`;
  } catch (e) {
    return link;
  }
}

async function fetchLinkTitles(links: string[]): Promise<void> {
  if (!LINK_TITLE_PROXY || !links.length) {
    return;
  }
  const unique = Array.from(new Set(links)).filter((l) => !linkTitleCache[l]);
  const tasks = unique.map(async (link) => {
    try {
      const resp = await fetch(`${LINK_TITLE_PROXY}?url=${encodeURIComponent(link)}`);
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      const title = (data?.title || "").trim();
      if (title) {
        linkTitleCache[link] = title;
      }
    } catch (err) {
      // still fallback to hostname/path via getLinkTitle
    }
  });
  await Promise.all(tasks);
}

async function handleTaskSubmit(event: Event) {
  event.preventDefault();
  if (!messageMetadata) {
    return;
  }

  await executeWithStatus("task-status", async () => {
    const attachmentInputs = await prepareAirtableAttachments(getSelectedAttachments("task-attachment-checkbox"));
    console.info("Attachments an Airtable-Payload:", attachmentInputs);
    const payload: AirtableTaskPayload = {
      title: getInputValue("task-title") || messageMetadata.subject,
      description: truncateForAirtable(sanitizeForAirtableText(getInputValue("task-description"))),
      projectRecordId: getProjectRecordId("task"),
      start: convertDateToIso(document.getElementById("task-start") as HTMLInputElement, true),
      end: convertDateToIso(document.getElementById("task-end") as HTMLInputElement, true),
      internalOwnerId: getSelectedInternalOwnerId(),
      internalOwnerEmail: getSelectedInternalOwnerEmail(),
      externalAssigneeIds: resolveExternalAssignees(getRecordIdList("task-external")),
      priority: getInputValue("task-priority") || undefined,
      category: getInputValue("task-category") || undefined,
      status: getInputValue("task-status-select") || undefined,
      art: (getInputValue("task-art") as AirtableTaskPayload["art"]) || "Task",
      attachments: attachmentInputs,
      message: messageMetadata,
    };

    await airtableClient.createTask(payload);
  });
}

async function handleDocumentSubmit(event: Event) {
  event.preventDefault();
  if (!messageMetadata) {
    return;
  }

  const documentSource = document.getElementById("document-source") as HTMLSelectElement;
  const source = documentSource && documentSource.value ? documentSource.value : "link";
  const project = getInputValue("document-project-input");
  const label = truncateForAirtable(sanitizeForAirtableText(getInputValue("document-label")));

  const payload: AirtableDocumentPayload = {
    project,
    projectRecordId: getProjectRecordId("document"),
    label,
    type: source as AirtableDocumentPayload["type"],
    message: messageMetadata,
  };

  if (source === "attachment") {
    const checkboxes = document.querySelectorAll<HTMLInputElement>(".doc-attachment-checkbox:checked");
    const selectedIds = Array.from(checkboxes).map((c) => c.value);
    const selected = getEligibleAttachments().filter((att) => selectedIds.includes(att.id));
    if (selected.length) {
      payload.attachments = await prepareAirtableAttachments(selected);
    }
  } else {
    const linkChecks = document.querySelectorAll<HTMLInputElement>(".doc-link-checkbox:checked");
    const selectedLinks = Array.from(linkChecks).map((c) => c.value).filter(Boolean);
    if (selectedLinks.length) {
      payload.url = selectedLinks[0];
    }
  }

  await executeWithStatus("document-status", () => airtableClient.createDocument(payload));
}

async function handleNoteSubmit(event: Event) {
  event.preventDefault();
  if (!messageMetadata) {
    return;
  }

  const title = getInputValue("note-title") || messageMetadata.subject;
  const noteText = truncateForAirtable(sanitizeForAirtableText(getInputValue("note-body")));
  const artSelect = document.getElementById("note-art") as HTMLSelectElement | null;
  const art = artSelect?.value ?? "E-Mail";
  const rawPersons = notePersonTokens.length ? notePersonTokens : getRecordIdList("note-persons");
  const personIds = resolveExternalAssignees(rawPersons);

  const payload: AirtableNotePayload = {
    title,
    note: noteText,
    projectRecordId: getProjectRecordId("note"),
    art,
    personRecordIds: personIds,
    date: messageMetadata.receivedDate ? messageMetadata.receivedDate.toISOString() : undefined,
    message: messageMetadata,
  };

  await executeWithStatus("note-status", () => airtableClient.createNote(payload));
}

async function executeWithStatus(
  elementId: string,
  action: () => Promise<unknown> | unknown
) {
  setStatus(elementId, "Wird gespeichert …", "pending");
  try {
    await action();
    setStatus(elementId, "Erfolgreich an Airtable übertragen.", "success");
  } catch (error) {
    console.error(error);
    setStatus(elementId, `Fehler beim Speichern: ${(error as Error).message}`, "error");
  }
}

function setStatus(elementId: string, message: string, type: "pending" | "success" | "error") {
  const element = document.getElementById(elementId);
  if (!element) {
    return;
  }
  element.textContent = message;
  element.className = `status ${type}`;
}

function getInputValue(elementId: string): string {
  const element = document.getElementById(elementId) as HTMLInputElement | HTMLTextAreaElement;
  if (!element || !element.value) {
    return "";
  }
  return element.value.trim();
}

function convertDateToIso(input: HTMLInputElement, isDateTime = false): string | undefined {
  if (!input || !input.value) {
    return undefined;
  }

  if (isDateTime) {
    return new Date(input.value).toISOString();
  }

  return new Date(`${input.value}T00:00:00`).toISOString();
}

function getSelectedAttachments(checkboxClass: string): OutlookAttachmentPreview[] {
  const checkboxes = document.querySelectorAll<HTMLInputElement>(`.${checkboxClass}`);
  const selectedIds = Array.from(checkboxes)
    .filter((checkbox) => checkbox.checked)
    .map((checkbox) => checkbox.value);
  const eligible = getEligibleAttachments();
  return eligible.filter((att) => selectedIds.includes(att.id));
}

function getAttachmentIconClass(contentType?: string, name?: string): string {
  const lower = (contentType || "").toLowerCase();
  const lowerName = (name || "").toLowerCase();
  if (lower.startsWith("image/")) return "ms-Icon--Photo2";
  if (lower.includes("pdf") || lowerName.endsWith(".pdf")) return "ms-Icon--PDF";
  if (lower.includes("word") || lowerName.endsWith(".doc") || lowerName.endsWith(".docx")) return "ms-Icon--WordDocument";
  if (lower.includes("excel") || lowerName.endsWith(".xls") || lowerName.endsWith(".xlsx")) return "ms-Icon--ExcelDocument";
  if (lower.includes("powerpoint") || lowerName.endsWith(".ppt") || lowerName.endsWith(".pptx")) return "ms-Icon--PowerPointDocument";
  return "ms-Icon--Document";
}

function toggleDocumentSource() {
  const documentSource = document.getElementById("document-source") as HTMLSelectElement;
  const source = documentSource && documentSource.value ? documentSource.value : "link";
  const attachmentGroup = document.getElementById("document-attachment-group");
  const linkGroup = document.getElementById("document-link-group");

  if (source === "attachment") {
    if (attachmentGroup) {
      attachmentGroup.classList.remove("hidden");
    }
    if (linkGroup) {
      linkGroup.classList.add("hidden");
    }
  } else {
    if (linkGroup) {
      linkGroup.classList.remove("hidden");
    }
    if (attachmentGroup) {
      attachmentGroup.classList.add("hidden");
    }
  }
}

async function getBodyAsText(item: Office.MessageRead): Promise<string> {
  return new Promise((resolve, reject) => {
    const messageRead = item as Office.MessageRead;
    const composeItem = item as unknown as Office.MessageCompose;

    if (typeof (messageRead as any).getBodyAsync === "function") {
      (messageRead as any).getBodyAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
      return;
    }

    if (composeItem && composeItem.body && typeof composeItem.body.getAsync === "function") {
      composeItem.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
      return;
    }

    reject(new Error("Body-API ist in diesem Kontext nicht verfügbar."));
  });
}

function buildOutlookWebLink(itemId: string): string | undefined {
  if (!itemId) {
    return undefined;
  }
  const encoded = encodeURIComponent(itemId);
  return `https://outlook.office.com/owa/?ItemID=${encoded}&exvsurl=1&viewmodel=ReadMessageItem`;
}

function prefillFormDefaults(metadata: OutlookMessageMetadata) {
  setIfEmpty("task-title", metadata.subject);

  if (metadata.receivedDate) {
    setDateTimeInput("task-start", metadata.receivedDate);
    const plusOneHour = new Date(metadata.receivedDate.getTime() + 60 * 60 * 1000);
    setDateTimeInput("task-end", plusOneHour);
  }
}

function setIfEmpty(elementId: string, value?: string) {
  const element = document.getElementById(elementId) as HTMLInputElement | HTMLTextAreaElement | null;
  if (!element || !value) {
    return;
  }
  if (!element.value) {
    element.value = value;
  }
}

function setDateInput(elementId: string, date: Date) {
  const element = document.getElementById(elementId) as HTMLInputElement | null;
  if (!element) {
    return;
  }
  element.value = formatDate(date);
}

function setDateTimeInput(elementId: string, date: Date) {
  const element = document.getElementById(elementId) as HTMLInputElement | null;
  if (!element) {
    return;
  }
  element.value = formatDateTime(date);
}

function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDateTime(date: Date): string {
  const year = date.getFullYear();
  const month = `${date.getMonth() + 1}`.padStart(2, "0");
  const day = `${date.getDate()}`.padStart(2, "0");
  const hours = `${date.getHours()}`.padStart(2, "0");
  const minutes = `${date.getMinutes()}`.padStart(2, "0");
  return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function getRecordIdList(elementId: string): string[] {
  const value = getInputValue(elementId);
  if (!value) {
    return [];
  }
  return value
    .split(/[,;\n]/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function resolveExternalAssignees(rawEntries: string[]): string[] {
  if (!rawEntries.length) {
    return [];
  }

  const ids: string[] = [];
  rawEntries.forEach((entry) => {
    if (entry.startsWith("rec")) {
      ids.push(entry);
      return;
    }
    const option = externalOptions.find(
      (opt) =>
        opt.id === entry ||
        opt.name.toLowerCase() === entry.toLowerCase() ||
        (opt.email && opt.email.toLowerCase() === entry.toLowerCase())
    );
    if (option) {
      ids.push(option.id);
    }
  });
  return Array.from(new Set(ids)).filter((id) => id && id.startsWith("rec"));
}

function renderExternalOptions() {
  const datalist = document.getElementById("task-external-datalist") as HTMLDataListElement | null;
  if (!datalist) {
    return;
  }
  datalist.innerHTML = "";
  externalOptions.forEach((person) => {
    const option = document.createElement("option");
    option.value = person.name;
    option.label = "";
    option.dataset.id = person.id;
    if (person.email) {
      option.dataset.email = person.email;
    }
    datalist.appendChild(option);
  });
}

function prefillSenderAsExternal() {
  const input = document.getElementById("task-external") as HTMLInputElement | null;
  if (!input || !senderEmail) {
    return;
  }
  const match = externalOptions.find(
    (opt) => opt.email && opt.email.toLowerCase() === senderEmail.toLowerCase()
  );
  if (match) {
    input.value = match.name;
  }
}

function addNotePersonToken(raw: string) {
  const tokens = raw
    .split(/[;,\n]+/)
    .map((t) => t.trim())
    .filter(Boolean);
  if (!tokens.length) {
    return;
  }
  tokens.forEach((token) => {
    const exists = notePersonTokens.some((t) => t.toLowerCase() === token.toLowerCase());
    if (!exists) {
      notePersonTokens.push(token);
    }
  });
  const input = document.getElementById("note-persons") as HTMLInputElement | null;
    if (input) {
      input.value = "";
    }
  renderNotePersonTokens();
}

function renderNotePersonTokens() {
  const container = document.getElementById("note-persons-selected");
  if (!container) {
    return;
  }
  container.innerHTML = "";
  if (!notePersonTokens.length) {
    return;
  }
  notePersonTokens.forEach((token) => {
    const matchById = externalOptions.find((o) => o.id === token);
    const matchByName = externalOptions.find((o) => o.name?.toLowerCase() === token.toLowerCase());
    const matchByEmail = externalOptions.find((o) => o.email?.toLowerCase() === token.toLowerCase());
    const display =
      matchById?.name ??
      matchByName?.name ??
      matchByEmail?.name ??
      matchById?.email ??
      matchByEmail?.email ??
      token;
    const pill = document.createElement("span");
    pill.className = "token-pill";
    pill.textContent = display;
    container.appendChild(pill);
  });
}

function renderProjectSelects(forceManualOnly: boolean) {
  projectInputs.forEach(({ inputId, datalistId }) => {
    const input = document.getElementById(inputId) as HTMLInputElement | null;
    const datalist = document.getElementById(datalistId) as HTMLDataListElement | null;
    if (!input || !datalist) {
      return;
    }

    datalist.innerHTML = "";
    input.disabled = false;

    if (forceManualOnly) {
      const option = document.createElement("option");
      option.value = "Keine Projekte geladen - bitte recID eingeben";
      datalist.appendChild(option);
      return;
    }

    projectOptions.forEach((project) => {
      const option = document.createElement("option");
      option.value = project.name;
      option.label = "";
      option.dataset.id = project.id;
      datalist.appendChild(option);
    });
  });
}

function getProjectRecordId(prefix: "task" | "event" | "document" | "note"): string {
  const input = document.getElementById(`${prefix}-project-input`) as HTMLInputElement | null;
  const value = input?.value?.trim() ?? "";
  if (!value) {
    return "";
  }

  // recID direkt verwenden
  if (value.startsWith("rec")) {
    return value;
  }

  // Name → ID auflösen
  const match = projectOptions.find((p) => p.name.toLowerCase() === value.toLowerCase());
  return match?.id ?? "";
}

function setProjectStatus(message: string, type: "pending" | "success" | "error" | "info") {
  const element = document.getElementById("project-select-status");
  if (!element) {
    return;
  }
  element.textContent = message;
  element.className = `status ${type}`;
}

function prefillBodyIntoDescription(body: string) {
  const description = document.getElementById("task-description") as HTMLTextAreaElement | null;
  if (description && !description.value) {
    description.value = normalizeBodyText(limitBodyText(body || "", 12000));
  }
}

function prefillNoteDefaults(body: string) {
  const title = document.getElementById("note-title") as HTMLInputElement | null;
  const note = document.getElementById("note-body") as HTMLTextAreaElement | null;
  if (title && !title.value && messageMetadata?.subject) {
    title.value = messageMetadata.subject;
  }
  if (note && !note.value) {
    note.value = normalizeBodyText(limitBodyText(body || "", 12000));
  }
}

function limitBodyText(text: string, maxLength = 12000): string {
  if (!text) return "";
  if (text.length <= maxLength) return text;
  return text.substring(0, maxLength);
}

function normalizeBodyText(text: string): string {
  if (!text) return "";
  return text.replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
}

function getSelectedInternalOwnerId(): string | undefined {
  const input = document.getElementById("task-owner-input") as HTMLInputElement | null;
  const value = input?.value?.trim() ?? "";
  if (!value) {
    return undefined;
  }
  const match = collaboratorOptions.find(
    (c) =>
      c.id === value ||
      (c.email && c.email.toLowerCase() === value.toLowerCase()) ||
      (c.name && c.name.toLowerCase() === value.toLowerCase())
  );
  if (match?.id) {
    return match.id;
  }
  if (value.startsWith("usr")) {
    return value;
  }
  return undefined;
}

function getSelectedInternalOwnerEmail(): string | undefined {
  const input = document.getElementById("task-owner-input") as HTMLInputElement | null;
  const value = input?.value?.trim() ?? "";
  if (!value) {
    return undefined;
  }
  const match = collaboratorOptions.find(
    (c) =>
      (c.email && c.email.toLowerCase() === value.toLowerCase()) ||
      (c.name && c.name.toLowerCase() === value.toLowerCase()) ||
      c.id === value
  );
  if (match?.email) {
    return match.email;
  }
  if (value.includes("@")) {
    return value;
  }
  return undefined;
}

async function prepareAirtableAttachments(selected: OutlookAttachmentPreview[]): Promise<AirtableAttachmentInput[]> {
  if (!selected.length) {
    console.info("Keine Anhänge ausgewählt.");
    return [];
  }

  const mailboxItem = Office.context.mailbox.item as Office.MessageRead;
  const results: AirtableAttachmentInput[] = [];
  const graphToken = await getGraphToken();
  if (!graphToken) {
    console.warn("Kein Graph-Token – Anhänge werden ausgelassen.");
    return [];
  }

  console.info(`Anhänge zum Upload ausgewählt: ${selected.length}`);

  for (const attachment of selected) {
    const content = await new Promise<
      | { type: "url"; value: string }
      | { type: "base64"; value: string }
      | { type: "unsupported"; error?: unknown }
    >((resolve) => {
      mailboxItem.getAttachmentContentAsync(attachment.id, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url && result.value.content) {
            resolve({ type: "url", value: result.value.content });
          } else if (
            result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64 &&
            result.value.content
          ) {
            resolve({ type: "base64", value: result.value.content });
          } else {
            resolve({ type: "unsupported" });
          }
        } else {
          resolve({ type: "unsupported", error: result.error });
        }
      });
    });

    if (content.type === "url") {
      results.push({ filename: attachment.name, url: content.value });
      continue;
    }

    if (content.type === "base64") {
      try {
        const uploaded = await uploadToOneDriveAndShare(attachment.name, content.value, graphToken);
        if (uploaded) {
          console.info("Upload erfolgreich für", attachment.name);
          results.push(uploaded);
        } else {
          console.warn("Upload/Share ergab keine URL, übersprungen:", attachment.name);
        }
      } catch (error) {
        console.warn("Attachment-Upload fehlgeschlagen, übersprungen:", attachment.name, error);
      }
      continue;
    }

    console.warn(
      "Attachment wurde übersprungen (kein öffentlich erreichbarer Link, Airtable erfordert URL):",
      attachment.name,
      content.type === "unsupported" ? content.error : "Base64"
    );
  }

  return results;
}

async function getAttachmentBase64(attachmentId: string): Promise<string> {
  const mailboxItem = Office.context.mailbox.item as Office.MessageRead;
  return new Promise((resolve, reject) => {
    mailboxItem.getAttachmentContentAsync(attachmentId, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
          resolve(result.value.content);
        } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url && result.value.content) {
          // Fallback: fetch the URL, then convert to base64.
          fetch(result.value.content)
            .then((response) => response.blob())
            .then((blob) => blob.arrayBuffer())
            .then((buffer) => {
              const base64String = arrayBufferToBase64(buffer);
              resolve(base64String);
            })
            .catch((error) => reject(error));
        } else {
          reject(new Error("Unbekanntes Attachment-Format."));
        }
      } else {
        reject(result.error);
      }
    });
  });
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const slice = bytes.subarray(offset, offset + chunkSize);
    binary += String.fromCharCode.apply(null, Array.from(slice));
  }
  return btoa(binary);
}

async function getGraphToken(): Promise<string | null> {
  if (cachedGraphToken) return cachedGraphToken;
  if (!GRAPH_CLIENT_ID || !GRAPH_REDIRECT_URI || !GRAPH_TENANT_ID) {
    console.warn("Graph-Konfiguration fehlt (GRAPH_CLIENT_ID / GRAPH_TENANT_ID / GRAPH_REDIRECT_URI).");
    return null;
  }
  try {
    if (!msalInstance) {
      msalInstance = new (window as any).msal.PublicClientApplication({
        auth: {
          clientId: GRAPH_CLIENT_ID,
          authority: `https://login.microsoftonline.com/${GRAPH_TENANT_ID}`,
          redirectUri: GRAPH_REDIRECT_URI,
        },
        cache: {
          cacheLocation: "localStorage",
          storeAuthStateInCookie: true,
        },
      });
    }

    // Falls wir aus einem Redirect zurückkommen
    const redirectResult = await msalInstance.handleRedirectPromise();
    if (redirectResult?.accessToken) {
      cachedGraphToken = redirectResult.accessToken;
      return cachedGraphToken;
    }

    const scopes = ["Files.ReadWrite.All", "User.Read"];
    const request = { scopes };

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      const silentResult = await msalInstance.acquireTokenSilent({ ...request, account: accounts[0] });
      if (silentResult?.accessToken) {
        cachedGraphToken = silentResult.accessToken;
        return cachedGraphToken;
      }
    }

    // Popup-Fallback
    const popupResult = await msalInstance.acquireTokenPopup(request);
    if (popupResult?.accessToken) {
      cachedGraphToken = popupResult.accessToken;
      return cachedGraphToken;
    }

    return null;
  } catch (error) {
    console.error("Konnte kein Graph-Token via MSAL beziehen:", error);
    return null;
  }
}

async function uploadToOneDriveAndShare(
  filename: string,
  base64: string,
  graphToken: string
): Promise<AirtableAttachmentInput | null> {
  const safeName = filename || `upload-${Date.now()}`;
  const uploadSessionUrl = `https://graph.microsoft.com/v1.0/me/drive/special/approot:/OutlookAirtableUploads/${encodeURIComponent(
    safeName
  )}:/createUploadSession`;

  const sessionResp = await fetch(uploadSessionUrl, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${graphToken}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": "replace",
        name: safeName,
      },
    }),
  });

  if (!sessionResp.ok) {
    console.warn("Upload-Session fehlgeschlagen für", filename, await sessionResp.text());
    return null;
  }

  const { uploadUrl } = await sessionResp.json();
  const buffer = base64ToArrayBuffer(base64);

  const uploadResp = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Length": buffer.byteLength.toString(),
      "Content-Range": `bytes 0-${buffer.byteLength - 1}/${buffer.byteLength}`,
    },
    body: buffer,
  });

  if (!uploadResp.ok) {
    console.warn("Upload fehlgeschlagen für", filename, await uploadResp.text());
    return null;
  }

  const uploaded = await uploadResp.json();
  const itemId = uploaded?.id;

  if (!itemId) {
    console.warn("Kein Item-ID nach Upload erhalten für", filename);
    return null;
  }

  // Versuche direkten Download-Link (prä-authentifizierter, kurzlebiger Link)
  try {
    const dlResp = await fetch(
      `https://graph.microsoft.com/v1.0/me/drive/items/${itemId}?select=name,@microsoft.graph.downloadUrl`,
      {
        method: "GET",
        headers: {
          Authorization: `Bearer ${graphToken}`,
        },
      }
    );
    if (dlResp.ok) {
      const dlBody = await dlResp.json();
      const dlUrl = dlBody?.["@microsoft.graph.downloadUrl"];
      if (dlUrl) {
        console.info("Download-URL verwendet für", filename, dlUrl);
        return { filename, url: dlUrl };
      }
    } else {
      console.warn("Download-URL Abfrage fehlgeschlagen:", await dlResp.text());
    }
  } catch (e) {
    console.warn("Download-URL nicht abrufbar:", e);
  }

  const shareScopes = ["anonymous", "organization"];
  for (const scope of shareScopes) {
    const linkResp = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${itemId}/createLink`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${graphToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ type: "view", scope }),
    });

    if (!linkResp.ok) {
      console.warn(`Share-Link fehlgeschlagen (${scope}) für`, filename, await linkResp.text());
      continue;
    }

    const linkBody = await linkResp.json();
    const url = linkBody?.link?.webUrl || uploaded?.webUrl;
    if (url) {
      console.info(`Share-Link (${scope}) erhalten für`, filename, url);
      return { filename, url };
    }
  }

  if (uploaded?.webUrl) {
    console.warn("Falle zurück auf webUrl (möglicherweise eingeschränkt) für", filename, uploaded.webUrl);
    return { filename, url: uploaded.webUrl };
  }

  console.warn("Kein Share-Link erhalten für", filename);
  return null;
}

function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) {
    bytes[i] = binaryString.charCodeAt(i);
  }
  return bytes.buffer;
}


function sanitizeForAirtableText(value: string): string {
  if (!value) return "";
  return value.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "").trim();
}

function truncateForAirtable(value: string, max = 50000): string {
  if (!value) return "";
  return value.length > max ? value.slice(0, max) : value;
}
