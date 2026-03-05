/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import { airtableClient } from "../services/airtableClient";
import {
  AirtableAttachmentInput,
  AirtableDocumentPayload,
  AirtableFinancePayload,
  AirtablePersonPayload,
  AirtableCompanyPayload,
  AirtableCompanyOption,
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
  { inputId: "finance-project-input", datalistId: "finance-project-datalist" },
];
const LINK_TITLE_PROXY = (process.env.TITLE_PROXY_URL || "").trim();
const GRAPH_CLIENT_ID = (process.env.GRAPH_CLIENT_ID || "").trim();
const GRAPH_TENANT_ID = (process.env.GRAPH_TENANT_ID || "common").trim();
const GRAPH_REDIRECT_URI = (process.env.GRAPH_REDIRECT_URI || "").trim();
const LOOKUP_REFRESH_THROTTLE_MS = 20000;

let attachments: OutlookAttachmentPreview[] = [];
let detectedLinks: string[] = [];
let messageMetadata: OutlookMessageMetadata | null = null;
let messageBodyText = "";
let projectOptions: AirtableProjectOption[] = [];
let collaboratorOptions: CollaboratorOption[] = [];
let externalOptions: AirtableProjectOption[] = [];
let companyOptions: AirtableCompanyOption[] = [];
let personRoleOptions: string[] = [];
let companyCategoryOptions: string[] = [];
let senderEmail: string | undefined;
let cachedGraphToken: string | null = null;
let notePersonTokens: string[] = [];
let companyCategoryTokens: string[] = [];
let msalInstance: any | null = null;
const linkTitleCache: Record<string, string> = {};
let lastLookupRefreshAt = 0;
let personNameCheckTimer: number | undefined;
let personEmailCheckTimer: number | undefined;
let companyNameCheckTimer: number | undefined;
let personNameCheckSeq = 0;
let personEmailCheckSeq = 0;
let companyNameCheckSeq = 0;
let duplicatePersonMatchByNameId: string | null = null;
let duplicatePersonMatchByEmailId: string | null = null;

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
  setUiVersion();
  wireUpForms();
  refreshPersonActionButton();
  setupTabs();
  await Promise.all([
    hydrateContext(),
    loadProjects(),
    loadCollaborators(),
    loadExternalPersons(),
    loadCompanies(),
    loadPersonRoles(),
    loadCompanyCategories(),
  ]);
  triggerPrefilledPersonDuplicateChecks();
}

function setUiVersion() {
  const el = document.getElementById("ui-version");
  if (!el) return;
  const params = new URLSearchParams(window.location.search);
  const version = params.get("v") || params.get("version");
  el.textContent = version ? `v${version}` : "local";
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

  const financeForm = document.getElementById("finance-form");
  if (financeForm) {
    financeForm.addEventListener("submit", handleFinanceSubmit);
  }

  const createPersonBtn = document.getElementById("create-person-btn");
  if (createPersonBtn) {
    createPersonBtn.addEventListener("click", handleCreatePersonFromSender);
  }

  const createCompanyBtn = document.getElementById("create-company-btn");
  if (createCompanyBtn) {
    createCompanyBtn.addEventListener("click", handleCreateCompanyFromForm);
  }

  const personNameInput = document.getElementById("person-name") as HTMLInputElement | null;
  if (personNameInput) {
    personNameInput.addEventListener("input", () => schedulePersonNameCheck(personNameInput.value));
  }

  const personEmailInput = document.getElementById("person-email") as HTMLInputElement | null;
  if (personEmailInput) {
    personEmailInput.addEventListener("input", () => schedulePersonEmailCheck(personEmailInput.value));
  }

  const companyNameInput = document.getElementById("company-name") as HTMLInputElement | null;
  if (companyNameInput) {
    companyNameInput.addEventListener("input", () => scheduleCompanyNameCheck(companyNameInput.value));
  }

  const personToggle = document.getElementById("person-toggle");
  if (personToggle) {
    personToggle.addEventListener("click", togglePersonForm);
  }

  const companyToggle = document.getElementById("company-toggle");
  if (companyToggle) {
    companyToggle.addEventListener("click", toggleCompanyForm);
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

  const categoryInput = document.getElementById("company-category") as HTMLInputElement | null;
  if (categoryInput) {
    const commitCategoryInput = () => addCompanyCategoryToken(categoryInput.value);
    categoryInput.addEventListener("change", commitCategoryInput);
    categoryInput.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter" || ev.key === "," || ev.key === ";") {
        ev.preventDefault();
        commitCategoryInput();
      }
    });
  }
}

function refreshPersonActionButton() {
  const button = document.getElementById("create-person-btn") as HTMLButtonElement | null;
  if (!button) {
    return;
  }
  const existingId = duplicatePersonMatchByEmailId || duplicatePersonMatchByNameId;
  button.textContent = existingId ? "Person aus Absender aktualisieren" : "Person aus Absender anlegen";
}

function setPersonDuplicateMatch(source: "name" | "email", recordId: string | null) {
  if (source === "name") {
    duplicatePersonMatchByNameId = recordId;
  } else {
    duplicatePersonMatchByEmailId = recordId;
  }
  refreshPersonActionButton();
}

function setupTabs() {
  const panels = Array.from(document.querySelectorAll<HTMLElement>(".tab-panel"));
  const select = document.getElementById("section-select") as HTMLSelectElement | null;
  const buttons = Array.from(document.querySelectorAll<HTMLButtonElement>(".tab-btn"));

  const activate = (targetId: string) => {
    panels.forEach((panel) => panel.classList.toggle("active", panel.id === targetId));
    buttons.forEach((btn) => btn.classList.toggle("active", btn.dataset.tabTarget === targetId));
    refreshLookupData();
  };

  if (select) {
    select.addEventListener("change", () => {
      if (select.value) {
        activate(select.value);
      }
    });
    if (select.value) {
      activate(select.value);
    }
    return;
  }

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
  messageBodyText = await getBodyAsText(mailboxItem);
  prefillBodyIntoDescription(messageBodyText);
  prefillNoteDefaults(messageBodyText);
  prefillPersonDefaults(mailboxItem);
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

async function loadCompanies() {
  const datalist = document.getElementById("person-company-datalist") as HTMLDataListElement | null;
  const status = document.getElementById("company-select-status");
  if (!datalist) {
    return;
  }
  if (status) {
    status.textContent = "Firmen werden geladen â€¦";
    status.className = "status pending";
  }
  console.info("Lade Firmen aus Airtable â€¦", {
    base: process.env.AIRTABLE_BASE_COMPANIES,
    table: process.env.AIRTABLE_TABLE_COMPANIES,
    view: process.env.AIRTABLE_VIEW_COMPANIES,
  });
  try {
    companyOptions = await airtableClient.fetchCompanies();
    datalist.innerHTML = "";
    companyOptions.forEach((company) => {
      const option = document.createElement("option");
      option.value = company.name;
      option.label = "";
      option.dataset.id = company.id;
      if (company.email) {
        option.dataset.email = company.email;
      }
      if (company.website) {
        option.dataset.website = company.website;
      }
      datalist.appendChild(option);
    });
    if (status) {
      status.textContent = companyOptions.length
        ? `${companyOptions.length} Firmen geladen.`
        : "Keine Firmen gefunden. Bitte recID manuell eingeben.";
      status.className = companyOptions.length ? "status success" : "status info";
    }
    console.info(`Firmen geladen: ${companyOptions.length}`);
    prefillCompanyFromSender();
  } catch (error) {
    console.error("Firmen konnten nicht geladen werden:", error);
    companyOptions = [];
    datalist.innerHTML = "";
    if (status) {
      status.textContent = `Firmen konnten nicht geladen werden: ${(error as Error).message}`;
      status.className = "status error";
    }
  }
}

async function loadPersonRoles() {
  const datalist = document.getElementById("person-role-datalist") as HTMLDataListElement | null;
  if (!datalist) {
    return;
  }
  try {
    personRoleOptions = await airtableClient.fetchPersonRoles();
    datalist.innerHTML = "";
    personRoleOptions.forEach((role) => {
      const option = document.createElement("option");
      option.value = role;
      datalist.appendChild(option);
    });
  } catch (error) {
    console.error("Rollen konnten nicht geladen werden:", error);
    personRoleOptions = [];
    datalist.innerHTML = "";
  }
}

async function loadCompanyCategories() {
  const datalist = document.getElementById("company-category-datalist") as HTMLDataListElement | null;
  if (!datalist) {
    return;
  }
  try {
    companyCategoryOptions = await airtableClient.fetchCompanyCategories();
    datalist.innerHTML = "";
    const list = companyCategoryOptions.length
      ? companyCategoryOptions
      : ["Kunde", "Supplier", "DIT", "VFX", "IT", "Unterhalt/Bau", "Reinigung", "Quality Control"];
    list.forEach((category) => {
      const option = document.createElement("option");
      option.value = category;
      datalist.appendChild(option);
    });
  } catch (error) {
    console.error("Firmkategorien konnten nicht geladen werden:", error);
    companyCategoryOptions = [];
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

async function handleFinanceSubmit(event: Event) {
  event.preventDefault();
  if (!messageMetadata) {
    return;
  }

  await executeWithStatus("finance-status", async () => {
    const titleInput = getInputValue("finance-title");
    const offerStatus = getInputValue("finance-offer-status") || undefined;
    const channel = getInputValue("finance-channel") || "Mail";
    const amountRaw = getInputValue("finance-amount");
    const amount = amountRaw ? Number(amountRaw.replace(",", ".")) : undefined;
    const dateIso = convertDateToIso(document.getElementById("finance-date") as HTMLInputElement, true);

    const payload: AirtableFinancePayload = {
      title: titleInput || messageMetadata.subject,
      projectRecordId: getProjectRecordId("finance"),
      type: "Offerte",
      date: dateIso,
      amount: Number.isFinite(amount as number) ? (amount as number) : undefined,
      description: truncateForAirtable(sanitizeForAirtableText(buildMailDescription(messageBodyText))),
      offerStatus,
      channel,
      message: messageMetadata,
    };

    await airtableClient.createFinance(payload);
  });
}

async function handleCreatePersonFromSender() {
  if (!messageMetadata) {
    return;
  }

  setStatus("person-status", "Person wird erstellt ...", "pending");
  try {
    const emailInput = document.getElementById("person-email") as HTMLInputElement | null;
    const nameInput = document.getElementById("person-name") as HTMLInputElement | null;
    const roleInput = document.getElementById("person-role-input") as HTMLInputElement | null;
    const mobileInput = document.getElementById("person-phone-mobile") as HTMLInputElement | null;
    const phoneInput = document.getElementById("person-phone") as HTMLInputElement | null;
    const companyInput = document.getElementById("person-company-input") as HTMLInputElement | null;
    const email = emailInput?.value?.trim() || senderEmail || "";
    const signatureName = extractSignatureName(messageBodyText);
    const name = nameInput?.value?.trim() || signatureName || (email ? email.split("@")[0] : "Unbekannt");
    const roles = roleInput?.value
      ? roleInput.value
          .split(/[,;\n]/)
          .map((entry) => entry.trim())
          .filter(Boolean)
      : [];
    const roleValues = normalizeRoleValues(roles);
    const companyRecordIds = resolveCompanyRecordIds(companyInput?.value || "");

    if (!roleValues.length) {
      setStatus("person-status", "Bitte mindestens eine Rolle angeben.", "error");
      return;
    }

    const signatureInfo = extractSignatureInfo(extractPrimaryMessageBody(messageBodyText));
    const payload: AirtablePersonPayload = {
      name,
      email: email || undefined,
      phoneMobile: mobileInput?.value?.trim() || signatureInfo.mobile || undefined,
      phone: phoneInput?.value?.trim() || signatureInfo.phone || undefined,
      roleValues: roleValues.length ? roleValues : undefined,
      companyRecordIds: companyRecordIds.length ? companyRecordIds : undefined,
    };

    let existingId = "";
    if (email) {
      const existing = await airtableClient.findPersonByEmail(email);
      if (existing) {
        existingId = existing.id;
      }
    }
    if (!existingId && companyRecordIds.length && name) {
      const existing = await airtableClient.findPersonByNameAndCompany(name, companyRecordIds[0]);
      if (existing) {
        existingId = existing.id;
      }
    }
    if (!existingId && name) {
      const existing = await airtableClient.findPersonByName(name);
      if (existing) {
        existingId = existing.id;
      }
    }

    if (existingId) {
      setPersonDuplicateMatch("name", existingId);
      setPersonDuplicateMatch("email", existingId);
      const updateResult = await updateExistingPerson(existingId, payload);
      if (updateResult.updated) {
        const updated = updateResult.updatedFields.join(", ");
        const changed =
          updateResult.differentExistingFields.length > 0
            ? ` Anderer Wert existiert bei: ${updateResult.differentExistingFields.join(", ")}.`
            : "";
        setStatus("person-status", `Person existiert bereits.${changed} Datensatz aktualisiert: ${updated}.`, "success");
      } else {
        setStatus("person-status", "Person existiert bereits. Alle geprueften Werte sind bereits vorhanden.", "success");
      }
      await loadExistingPersonIntoForm(existingId);
      return;
    }

    await airtableClient.createPerson(payload);
    setStatus("person-status", "Person existiert nicht. Neuer Datensatz wird angelegt ...", "success");
    await loadExternalPersons();
    setPersonDuplicateMatch("name", null);
    setPersonDuplicateMatch("email", null);
    setStatus("person-status", "Person existiert nicht. Neuer Datensatz wurde angelegt.", "success");
    togglePersonForm(true);
  } catch (error) {
    console.error(error);
    setStatus("person-status", `Fehler beim Erstellen: ${(error as Error).message}`, "error");
  }
}

async function handleCreateCompanyFromForm() {
  setStatus("company-status", "Firma wird erstellt ...", "pending");
  try {
    const nameInput = document.getElementById("company-name") as HTMLInputElement | null;
    const emailInput = document.getElementById("company-email") as HTMLInputElement | null;
    const phoneInput = document.getElementById("company-phone") as HTMLInputElement | null;
    const websiteInput = document.getElementById("company-website") as HTMLInputElement | null;
    const streetInput = document.getElementById("company-street") as HTMLInputElement | null;
    const houseInput = document.getElementById("company-house-number") as HTMLInputElement | null;
    const zipInput = document.getElementById("company-zip") as HTMLInputElement | null;
    const cityInput = document.getElementById("company-city") as HTMLInputElement | null;
    const countryInput = document.getElementById("company-country") as HTMLSelectElement | null;
    const languageInput = document.getElementById("company-language") as HTMLSelectElement | null;
    const categoryInput = document.getElementById("company-category") as HTMLInputElement | null;
    const personCompanyInput = document.getElementById("person-company-input") as HTMLInputElement | null;

    const fallbackName = personCompanyInput?.value?.trim() || "";
    const name = nameInput?.value?.trim() || fallbackName;
    if (!name) {
      setStatus("company-status", "Bitte einen Firmennamen eingeben.", "error");
      return;
    }

    const existing =
      companyOptions.find((company) => company.name.toLowerCase() === name.toLowerCase()) ??
      (await airtableClient.findCompanyByName(name));

    const zipValue = zipInput?.value ? Number(zipInput.value) : undefined;
    if (categoryInput?.value?.trim()) {
      addCompanyCategoryToken(categoryInput.value);
    }
    const categories = companyCategoryTokens.slice();

    const payload: AirtableCompanyPayload = {
      name,
      email: emailInput?.value?.trim() || undefined,
      phone: phoneInput?.value?.trim() || undefined,
      website: websiteInput?.value?.trim() || undefined,
      street: streetInput?.value?.trim() || undefined,
      houseNumber: houseInput?.value?.trim() || undefined,
      zip: Number.isFinite(zipValue) ? zipValue : undefined,
      city: cityInput?.value?.trim() || undefined,
      country: countryInput?.value || undefined,
      language: languageInput?.value || undefined,
      categories: categories.length ? categories : undefined,
    };

    if (existing) {
      const updateResult = await updateExistingCompany(existing.id, payload);
      if (updateResult.updated) {
        const updated = updateResult.updatedFields.join(", ");
        const changed =
          updateResult.differentExistingFields.length > 0
            ? ` Anderer Wert existiert bei: ${updateResult.differentExistingFields.join(", ")}.`
            : "";
        setStatus("company-status", `Firma existiert bereits.${changed} Datensatz aktualisiert: ${updated}.`, "success");
      } else {
        setStatus("company-status", "Firma existiert bereits. Alle geprueften Werte sind bereits vorhanden.", "success");
      }
      await loadCompanies();
      return;
    }

    await airtableClient.createCompany(payload);
    setStatus("company-status", "Firma existiert nicht. Neuer Datensatz wird angelegt ...", "success");
    companyCategoryTokens = [];
    renderCompanyCategoryTokens();
    await loadCompanies();
    if (personCompanyInput) {
      personCompanyInput.value = name;
    }
    setStatus("company-status", "Firma existiert nicht. Neuer Datensatz wurde angelegt.", "success");
    toggleCompanyForm(true);
  } catch (error) {
    console.error(error);
    setStatus("company-status", `Fehler beim Erstellen: ${(error as Error).message}`, "error");
  }
}

function addCompanyCategoryToken(raw: string) {
  const tokens = raw
    .split(/[;,\n]+/)
    .map((t) => t.trim())
    .filter(Boolean);
  if (!tokens.length) {
    return;
  }
  tokens.forEach((token) => {
    const exists = companyCategoryTokens.some((t) => t.toLowerCase() === token.toLowerCase());
    if (!exists) {
      companyCategoryTokens.push(token);
    }
  });
  const input = document.getElementById("company-category") as HTMLInputElement | null;
  if (input) {
    input.value = "";
  }
  renderCompanyCategoryTokens();
}

function renderCompanyCategoryTokens() {
  const container = document.getElementById("company-category-selected");
  if (!container) {
    return;
  }
  container.innerHTML = "";
  if (!companyCategoryTokens.length) {
    return;
  }
  companyCategoryTokens.forEach((token) => {
    const pill = document.createElement("span");
    pill.className = "token-pill";

    const label = document.createElement("span");
    label.textContent = token;

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "token-remove";
    removeBtn.setAttribute("aria-label", `${token} entfernen`);
    removeBtn.textContent = "×";
    removeBtn.addEventListener("click", () => removeCompanyCategoryToken(token));

    pill.appendChild(label);
    pill.appendChild(removeBtn);
    container.appendChild(pill);
  });
}

async function refreshLookupData(force = false) {
  const now = Date.now();
  if (!force && now - lastLookupRefreshAt < LOOKUP_REFRESH_THROTTLE_MS) {
    return;
  }
  lastLookupRefreshAt = now;
  await Promise.all([
    loadProjects(),
    loadCompanies(),
    loadPersonRoles(),
    loadCompanyCategories(),
    loadExternalPersons(),
    loadCollaborators(),
  ]);
}

function removeCompanyCategoryToken(token: string) {
  companyCategoryTokens = companyCategoryTokens.filter((t) => t.toLowerCase() !== token.toLowerCase());
  renderCompanyCategoryTokens();
}

type EntitySyncResult = {
  updated: boolean;
  updatedFields: string[];
  differentExistingFields: string[];
  unchangedFields: string[];
};

async function updateExistingPerson(
  recordId: string,
  payload: AirtablePersonPayload
): Promise<EntitySyncResult> {
  const existingRecord = await airtableClient.getPersonRecord(recordId);
  if (!existingRecord) {
    return { updated: false, updatedFields: [], differentExistingFields: [], unchangedFields: [] };
  }
  const fields = existingRecord.fields as Record<string, unknown>;
  const existingName = typeof fields.Name === "string" ? fields.Name.trim() : "";
  const existingEmail = typeof fields["E-Mail"] === "string" ? fields["E-Mail"].trim() : "";
  const existingMobile = typeof fields["Telefon (Mobil)"] === "string" ? fields["Telefon (Mobil)"].trim() : "";
  const existingPhone = typeof fields.Telefon === "string" ? fields.Telefon.trim() : "";
  const existingRoles = Array.isArray(fields.Rolle) ? (fields.Rolle as string[]) : [];
  const existingCompanies = Array.isArray(fields.Firmen) ? (fields.Firmen as string[]) : [];

  const updates: Partial<AirtablePersonPayload> = {};
  const updatedFields: string[] = [];
  const differentExistingFields: string[] = [];
  const unchangedFields: string[] = [];

  if (payload.name) {
    if (!existingName) {
      updates.name = payload.name;
      updatedFields.push("Name");
    } else if (existingName.toLowerCase() !== payload.name.trim().toLowerCase()) {
      updates.name = payload.name.trim();
      updatedFields.push("Name");
      differentExistingFields.push("Name");
    } else {
      unchangedFields.push("Name");
    }
  }
  if (payload.email) {
    if (!existingEmail) {
      updates.email = payload.email;
      updatedFields.push("E-Mail");
    } else if (existingEmail.toLowerCase() !== payload.email.trim().toLowerCase()) {
      updates.email = payload.email.trim();
      updatedFields.push("E-Mail");
      differentExistingFields.push("E-Mail");
    } else {
      unchangedFields.push("E-Mail");
    }
  }
  if (payload.phoneMobile) {
    if (!existingMobile) {
      updates.phoneMobile = payload.phoneMobile;
      updatedFields.push("Telefon (Mobil)");
    } else if (existingMobile !== payload.phoneMobile.trim()) {
      updates.phoneMobile = payload.phoneMobile.trim();
      updatedFields.push("Telefon (Mobil)");
      differentExistingFields.push("Telefon (Mobil)");
    } else {
      unchangedFields.push("Telefon (Mobil)");
    }
  }
  if (payload.phone) {
    if (!existingPhone) {
      updates.phone = payload.phone;
      updatedFields.push("Telefon");
    } else if (existingPhone !== payload.phone.trim()) {
      updates.phone = payload.phone.trim();
      updatedFields.push("Telefon");
      differentExistingFields.push("Telefon");
    } else {
      unchangedFields.push("Telefon");
    }
  }

  const mergedRoles = mergeUnique(existingRoles, payload.roleValues ?? []);
  if (mergedRoles.changed) {
    updates.roleValues = mergedRoles.values;
    updatedFields.push("Rolle");
  } else if (payload.roleValues?.length) {
    unchangedFields.push("Rolle");
  }

  const mergedCompanies = mergeUnique(existingCompanies, payload.companyRecordIds ?? []);
  if (mergedCompanies.changed) {
    updates.companyRecordIds = mergedCompanies.values;
    updatedFields.push("Firma");
  } else if (payload.companyRecordIds?.length) {
    unchangedFields.push("Firma");
  }

  if (!Object.keys(updates).length) {
    return { updated: false, updatedFields: [], differentExistingFields, unchangedFields };
  }

  await airtableClient.updatePerson(recordId, updates);
  return { updated: true, updatedFields, differentExistingFields, unchangedFields };
}

async function updateExistingCompany(
  recordId: string,
  payload: AirtableCompanyPayload
): Promise<EntitySyncResult> {
  const existingRecord = await airtableClient.getCompanyRecord(recordId);
  if (!existingRecord) {
    return { updated: false, updatedFields: [], differentExistingFields: [], unchangedFields: [] };
  }

  const fields = existingRecord.fields as Record<string, unknown>;
  const existingName = typeof fields.Firmenname === "string" ? fields.Firmenname.trim() : "";
  const existingEmail = typeof fields["E-Mail"] === "string" ? fields["E-Mail"].trim() : "";
  const existingPhone = typeof fields.Telefon === "string" ? fields.Telefon.trim() : "";
  const existingWebsite = typeof fields.Webseite === "string" ? fields.Webseite.trim() : "";
  const existingStreet = typeof fields.Strasse === "string" ? fields.Strasse.trim() : "";
  const existingHouse = typeof fields.Hausnummer === "string" ? fields.Hausnummer.trim() : "";
  const existingZip = typeof fields.PLZ === "number" ? fields.PLZ : Number(fields.PLZ ?? NaN);
  const existingCity = typeof fields.Ort === "string" ? fields.Ort.trim() : "";
  const existingCountry = typeof fields.Land === "string" ? fields.Land.trim() : "";
  const existingLanguage = typeof fields.Sprache === "string" ? fields.Sprache.trim() : "";
  const existingCategories = Array.isArray(fields.Kategorie) ? (fields.Kategorie as string[]) : [];

  const updates: Partial<AirtableCompanyPayload> = {};
  const updatedFields: string[] = [];
  const differentExistingFields: string[] = [];
  const unchangedFields: string[] = [];

  const compareText = (
    label: string,
    incoming: string | undefined,
    current: string,
    assign: (value: string) => void
  ) => {
    if (!incoming) return;
    const value = incoming.trim();
    if (!value) return;
    if (!current) {
      assign(value);
      updatedFields.push(label);
      return;
    }
    if (current.toLowerCase() !== value.toLowerCase()) {
      assign(value);
      updatedFields.push(label);
      differentExistingFields.push(label);
      return;
    }
    unchangedFields.push(label);
  };

  compareText("Firmenname", payload.name, existingName, (value) => (updates.name = value));
  compareText("E-Mail", payload.email, existingEmail, (value) => (updates.email = value));
  compareText("Telefon", payload.phone, existingPhone, (value) => (updates.phone = value));
  compareText("Webseite", payload.website, existingWebsite, (value) => (updates.website = value));
  compareText("Strasse", payload.street, existingStreet, (value) => (updates.street = value));
  compareText("Hausnummer", payload.houseNumber, existingHouse, (value) => (updates.houseNumber = value));
  compareText("Ort", payload.city, existingCity, (value) => (updates.city = value));
  compareText("Land", payload.country, existingCountry, (value) => (updates.country = value));
  compareText("Sprache", payload.language, existingLanguage, (value) => (updates.language = value));

  if (payload.zip !== undefined && Number.isFinite(payload.zip)) {
    if (!Number.isFinite(existingZip)) {
      updates.zip = payload.zip;
      updatedFields.push("PLZ");
    } else if (existingZip !== payload.zip) {
      updates.zip = payload.zip;
      updatedFields.push("PLZ");
      differentExistingFields.push("PLZ");
    } else {
      unchangedFields.push("PLZ");
    }
  }

  const mergedCategories = mergeUnique(existingCategories, payload.categories ?? []);
  if (mergedCategories.changed) {
    updates.categories = mergedCategories.values;
    updatedFields.push("Kategorie");
  } else if (payload.categories?.length) {
    unchangedFields.push("Kategorie");
  }

  if (!Object.keys(updates).length) {
    return { updated: false, updatedFields: [], differentExistingFields, unchangedFields };
  }

  await airtableClient.updateCompany(recordId, updates);
  return { updated: true, updatedFields, differentExistingFields, unchangedFields };
}

function mergeUnique(existing: string[], incoming: string[]): { values: string[]; changed: boolean } {
  if (!incoming.length) {
    return { values: existing, changed: false };
  }
  const set = new Set(existing);
  let changed = false;
  incoming.forEach((value) => {
    if (value && !set.has(value)) {
      set.add(value);
      changed = true;
    }
  });
  return { values: Array.from(set), changed };
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

function setHint(elementId: string, message: string, type: "pending" | "success" | "error" | "info") {
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
  setIfEmpty("finance-title", metadata.subject);

  if (metadata.receivedDate) {
    setDateTimeInput("task-start", metadata.receivedDate);
    const plusOneHour = new Date(metadata.receivedDate.getTime() + 60 * 60 * 1000);
    setDateTimeInput("task-end", plusOneHour);
    setDateTimeInput("finance-date", metadata.receivedDate);
  }
}

function prefillPersonDefaults(item: Office.MessageRead) {
  const nameInput = document.getElementById("person-name") as HTMLInputElement | null;
  const emailInput = document.getElementById("person-email") as HTMLInputElement | null;
  const mobileInput = document.getElementById("person-phone-mobile") as HTMLInputElement | null;
  const phoneInput = document.getElementById("person-phone") as HTMLInputElement | null;
  const displayName = item.from?.displayName?.trim() || "";
  const email = item.from?.emailAddress?.trim() || senderEmail || "";
  const signatureName = extractSignatureName(extractPrimaryMessageBody(messageBodyText));
  const signatureInfo = extractSignatureInfo(extractPrimaryMessageBody(messageBodyText));
  if (nameInput && !nameInput.value) {
    nameInput.value = displayName || signatureName || (email ? email.split("@")[0] : "");
  }
  if (emailInput && !emailInput.value) {
    emailInput.value = email;
  }
  if (mobileInput && !mobileInput.value && signatureInfo.mobile) {
    mobileInput.value = signatureInfo.mobile;
  }
  if (phoneInput && !phoneInput.value && signatureInfo.phone) {
    phoneInput.value = signatureInfo.phone;
  }
}

function triggerPrefilledPersonDuplicateChecks() {
  const nameInput = document.getElementById("person-name") as HTMLInputElement | null;
  const emailInput = document.getElementById("person-email") as HTMLInputElement | null;
  const name = nameInput?.value?.trim() ?? "";
  const email = emailInput?.value?.trim() ?? "";

  if (name) {
    const seq = ++personNameCheckSeq;
    void checkPersonNameDuplicate(name, seq);
  } else {
    setPersonDuplicateMatch("name", null);
  }
  if (email) {
    const seq = ++personEmailCheckSeq;
    void checkPersonEmailDuplicate(email, seq);
  } else {
    setPersonDuplicateMatch("email", null);
  }
}

async function loadExistingPersonIntoForm(recordId: string) {
  if (!recordId) {
    return;
  }
  try {
    const record = await airtableClient.getPersonRecord(recordId);
    if (!record) {
      return;
    }
    const fields = record.fields as Record<string, unknown>;
    const name = typeof fields.Name === "string" ? fields.Name.trim() : "";
    const email = typeof fields["E-Mail"] === "string" ? fields["E-Mail"].trim() : "";
    const phoneMobile = typeof fields["Telefon (Mobil)"] === "string" ? fields["Telefon (Mobil)"].trim() : "";
    const phone = typeof fields.Telefon === "string" ? fields.Telefon.trim() : "";
    const roles = Array.isArray(fields.Rolle) ? (fields.Rolle as string[]) : [];
    const companyIds = Array.isArray(fields.Firmen) ? (fields.Firmen as string[]) : [];

    const nameInput = document.getElementById("person-name") as HTMLInputElement | null;
    const emailInput = document.getElementById("person-email") as HTMLInputElement | null;
    const roleInput = document.getElementById("person-role-input") as HTMLInputElement | null;
    const mobileInput = document.getElementById("person-phone-mobile") as HTMLInputElement | null;
    const phoneInput = document.getElementById("person-phone") as HTMLInputElement | null;
    const companyInput = document.getElementById("person-company-input") as HTMLInputElement | null;

    if (nameInput) nameInput.value = name || nameInput.value;
    if (emailInput) emailInput.value = email || emailInput.value;
    if (mobileInput) mobileInput.value = phoneMobile || mobileInput.value;
    if (phoneInput) phoneInput.value = phone || phoneInput.value;
    if (roleInput && roles.length) roleInput.value = roles.join(", ");

    if (companyInput && companyIds.length) {
      const firstId = companyIds[0];
      const match = companyOptions.find((company) => company.id === firstId);
      companyInput.value = match?.name || firstId;
    }
  } catch (error) {
    console.warn("Bestehende Person konnte nicht in Formular geladen werden:", error);
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

function resolveCompanyRecordIds(rawEntry: string): string[] {
  const value = rawEntry.trim();
  if (!value) {
    return [];
  }
  if (value.startsWith("rec")) {
    return [value];
  }
  const match = companyOptions.find((company) => company.name.toLowerCase() === value.toLowerCase());
  return match ? [match.id] : [];
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

function getProjectRecordId(prefix: "task" | "event" | "document" | "note" | "finance"): string {
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
    description.value = buildMailDescription(body);
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

function buildMailDescription(body: string): string {
  const headerLines: string[] = [];
  if (messageMetadata?.from) {
    headerLines.push(`Von: ${messageMetadata.from}`);
  }
  if (messageMetadata?.receivedDate) {
    headerLines.push(`Datum: ${messageMetadata.receivedDate.toLocaleString()}`);
  }
  if (messageMetadata?.subject) {
    headerLines.push(`Betreff: ${messageMetadata.subject}`);
  }

  const header = headerLines.join("\n");
  const content = normalizeBodyText(limitBodyText(body || "", 12000));

  if (header && content) {
    return `${header}\n\n${content}`;
  }
  return header || content;
}

function prefillCompanyFromSender() {
  const input = document.getElementById("person-company-input") as HTMLInputElement | null;
  if (!input || !senderEmail || !companyOptions.length) {
    return;
  }
  const domain = senderEmail.split("@")[1]?.toLowerCase() || "";
  if (!domain) {
    return;
  }
  const match = companyOptions.find((company) => {
    const emailDomain = company.email?.split("@")[1]?.toLowerCase();
    if (emailDomain && emailDomain === domain) {
      return true;
    }
    const websiteDomain = company.website ? extractDomain(company.website) : "";
    return websiteDomain && websiteDomain === domain;
  });
  if (match) {
    input.value = match.name;
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

function extractPrimaryMessageBody(text: string): string {
  if (!text) {
    return "";
  }
  const separators = [
    /^Am\s.+schrieb.+:$/i,
    /^Le\s.+a écrit\s*:$/i,
    /^On\s.+wrote:$/i,
    /^On\s.+,\s+at\s.+,\s+.+wrote:$/i,
    /^From:\s.+$/i,
    /^Von:\s.+$/i,
    /^Sent:\s.+$/i,
    /^Gesendet:\s.+$/i,
    /^-----Original Message-----$/i,
  ];
  const lines = text.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line) {
      continue;
    }
    if (separators.some((pattern) => pattern.test(line))) {
      return lines.slice(0, i).join("\n").trim();
    }
  }
  return text;
}

function togglePersonForm(forceCollapsed?: boolean) {
  const container = document.getElementById("person-form");
  const toggle = document.getElementById("person-toggle");
  if (!container || !toggle) {
    return;
  }
  const isCollapsed =
    typeof forceCollapsed === "boolean" ? forceCollapsed : container.classList.toggle("collapsed");
  container.classList.toggle("collapsed", isCollapsed);
  toggle.textContent = isCollapsed ? "Person erfassen" : "Person ausblenden";
  void refreshLookupData();
}

function toggleCompanyForm(forceCollapsed?: boolean) {
  const container = document.getElementById("company-form");
  const toggle = document.getElementById("company-toggle");
  if (!container || !toggle) {
    return;
  }
  const isCollapsed =
    typeof forceCollapsed === "boolean" ? forceCollapsed : container.classList.toggle("collapsed");
  container.classList.toggle("collapsed", isCollapsed);
  toggle.textContent = isCollapsed ? "Firma erfassen" : "Firma ausblenden";
  void refreshLookupData();
}

function normalizeRoleValues(values: string[]): string[] {
  if (!values.length) {
    return [];
  }
  if (!personRoleOptions.length) {
    return values;
  }
  return values.filter((value) =>
    personRoleOptions.some((role) => role.toLowerCase() === value.toLowerCase())
  );
}

function extractDomain(raw: string): string {
  try {
    const url = raw.includes("://") ? raw : `https://${raw}`;
    return new URL(url).hostname.replace(/^www\./, "").toLowerCase();
  } catch (error) {
    return raw.replace(/^www\./, "").toLowerCase();
  }
}

function schedulePersonNameCheck(value: string) {
  if (personNameCheckTimer) {
    window.clearTimeout(personNameCheckTimer);
  }
  const seq = ++personNameCheckSeq;
  personNameCheckTimer = window.setTimeout(() => checkPersonNameDuplicate(value, seq), 400);
}

function schedulePersonEmailCheck(value: string) {
  if (personEmailCheckTimer) {
    window.clearTimeout(personEmailCheckTimer);
  }
  const seq = ++personEmailCheckSeq;
  personEmailCheckTimer = window.setTimeout(() => checkPersonEmailDuplicate(value, seq), 400);
}

function scheduleCompanyNameCheck(value: string) {
  if (companyNameCheckTimer) {
    window.clearTimeout(companyNameCheckTimer);
  }
  const seq = ++companyNameCheckSeq;
  companyNameCheckTimer = window.setTimeout(() => checkCompanyNameDuplicate(value, seq), 400);
}

async function checkPersonNameDuplicate(raw: string, seq: number) {
  const name = raw.trim();
  if (!name) {
    setHint("person-duplicate-hint", "", "info");
    setPersonDuplicateMatch("name", null);
    return;
  }
  setHint("person-duplicate-hint", "Pruefe Name...", "info");
  const match = externalOptions.find((person) => person.name?.toLowerCase() === name.toLowerCase());
  if (match) {
    const extra = match.email ? ` (${match.email})` : "";
    setHint("person-duplicate-hint", `Person existiert bereits: ${match.name}${extra}`, "success");
    setPersonDuplicateMatch("name", match.id);
    await loadExistingPersonIntoForm(match.id);
    return;
  }
  try {
    const remoteMatch = await airtableClient.findPersonByName(name);
    if (seq !== personNameCheckSeq) return;
    if (remoteMatch) {
      const extra = remoteMatch.email ? ` (${remoteMatch.email})` : "";
      setHint("person-duplicate-hint", `Person existiert bereits: ${remoteMatch.name}${extra}`, "success");
      setPersonDuplicateMatch("name", remoteMatch.id);
      await loadExistingPersonIntoForm(remoteMatch.id);
    } else {
      setHint("person-duplicate-hint", "Keine bestehende Person gefunden.", "info");
      setPersonDuplicateMatch("name", null);
    }
  } catch (error) {
    if (seq !== personNameCheckSeq) return;
    setHint("person-duplicate-hint", "Pruefung fehlgeschlagen.", "warn");
  }
}

async function checkPersonEmailDuplicate(raw: string, seq: number) {
  const email = raw.trim().toLowerCase();
  if (!email) {
    setHint("person-email-duplicate-hint", "", "info");
    setPersonDuplicateMatch("email", null);
    return;
  }
  setHint("person-email-duplicate-hint", "Pruefe E-Mail...", "info");
  const match = externalOptions.find((person) => person.email?.toLowerCase() === email);
  if (match) {
    setHint(
      "person-email-duplicate-hint",
      `E-Mail existiert bereits: ${match.name ?? match.email}`,
      "success"
    );
    setPersonDuplicateMatch("email", match.id);
    await loadExistingPersonIntoForm(match.id);
    return;
  }
  try {
    const remoteMatch = await airtableClient.findPersonByEmail(email);
    if (seq !== personEmailCheckSeq) return;
    if (remoteMatch) {
      setHint(
        "person-email-duplicate-hint",
        `E-Mail existiert bereits: ${remoteMatch.name ?? remoteMatch.email}`,
        "success"
      );
      setPersonDuplicateMatch("email", remoteMatch.id);
      await loadExistingPersonIntoForm(remoteMatch.id);
    } else {
      setHint("person-email-duplicate-hint", "E-Mail noch nicht vorhanden.", "info");
      setPersonDuplicateMatch("email", null);
    }
  } catch (error) {
    if (seq !== personEmailCheckSeq) return;
    setHint("person-email-duplicate-hint", "Pruefung fehlgeschlagen.", "warn");
  }
}

async function checkCompanyNameDuplicate(raw: string, seq: number) {
  const name = raw.trim();
  if (!name) {
    setHint("company-duplicate-hint", "", "info");
    return;
  }
  setHint("company-duplicate-hint", "Pruefe Firma...", "info");
  const match = companyOptions.find((company) => company.name?.toLowerCase() === name.toLowerCase());
  if (match) {
    setHint("company-duplicate-hint", `Firma existiert bereits: ${match.name}`, "success");
    return;
  }
  try {
    const remoteMatch = await airtableClient.findCompanyByName(name);
    if (seq !== companyNameCheckSeq) return;
    if (remoteMatch) {
      setHint("company-duplicate-hint", `Firma existiert bereits: ${remoteMatch.name}`, "success");
    } else {
      setHint("company-duplicate-hint", "Keine bestehende Firma gefunden.", "info");
    }
  } catch (error) {
    if (seq !== companyNameCheckSeq) return;
    setHint("company-duplicate-hint", "Pruefung fehlgeschlagen.", "warn");
  }
}

function extractSignatureInfo(text: string): { mobile?: string; phone?: string } {
  if (!text) {
    return {};
  }
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const tail = lines.slice(-24);
  const phoneMatches = tail.flatMap((line) => {
    const matches = line.match(/(\+?\d[\d\s().-]{6,}\d)/g);
    if (!matches) return [];
    return matches.map((match) => ({ line, value: match }));
  });

  const cleaned = phoneMatches.map((entry) => ({
    line: entry.line.toLowerCase(),
    value: entry.value.replace(/\s+/g, " ").trim(),
  }));

  let mobile: string | undefined;
  let phone: string | undefined;

  cleaned.forEach((entry) => {
    if (!mobile && /(mobile|mobil|handy|cell|mobi)/i.test(entry.line)) {
      mobile = entry.value;
      return;
    }
    if (!phone && /(tel|telefon|phone|office|fon)/i.test(entry.line)) {
      phone = entry.value;
      return;
    }
    if (!phone) {
      phone = entry.value;
      return;
    }
    if (!mobile) {
      mobile = entry.value;
    }
  });

  return { mobile, phone };
}

function extractSignatureName(text: string): string {
  if (!text) {
    return "";
  }
  const lines = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  const tail = lines.slice(-20);
  for (const line of tail) {
    if (/@/.test(line)) continue;
    if (/\+?\d[\d\s().-]{6,}\d/.test(line)) continue;
    if (/^(telefon|phone|tel|mobile|mobil|handy|fax|www\.|http|https)/i.test(line)) continue;
    const clean = line.replace(/[|•]/g, " ").replace(/\s+/g, " ").trim();
    if (clean.split(" ").length >= 2 && /[A-Za-zÀ-ÿ]/.test(clean)) {
      return clean;
    }
  }
  return "";
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

  // Download-Link nur als Fallback vormerken.
  let fallbackDownloadUrl: string | null = null;
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
        fallbackDownloadUrl = dlUrl;
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
    const baseUrl = linkBody?.link?.webUrl || uploaded?.webUrl;
    const url = baseUrl ? `${baseUrl}${baseUrl.includes("?") ? "&" : "?"}download=1` : "";
    if (url) {
      console.info(`Share-Link (${scope}) erhalten für`, filename, url);
      return { filename, url };
    }
  }

  if (fallbackDownloadUrl) {
    console.info("Download-URL verwendet für", filename, fallbackDownloadUrl);
    return { filename, url: fallbackDownloadUrl };
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
