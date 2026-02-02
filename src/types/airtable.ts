export interface OutlookAttachmentPreview {
  id: string;
  name: string;
  contentType: string;
  size: number;
  isInline: boolean;
}

export interface AirtableAttachmentInput {
  filename: string;
  url: string;
}

export interface OutlookMessageMetadata {
  itemId: string;
  subject: string;
  from: string;
  receivedDate: Date | null;
  webLink?: string;
}

export interface AirtableTaskPayload {
  title: string;
  projectRecordId?: string;
  start?: string;
  end?: string;
  internalOwnerEmail?: string;
  internalOwnerId?: string;
  externalAssigneeIds?: string[];
  priority?: string;
  status?: string;
  category?: string;
  art?: "Task" | "Termin";
  description?: string;
  attachments?: AirtableAttachmentInput[];
  message: OutlookMessageMetadata;
}

export interface AirtableEventPayload {
  title: string;
  projectRecordId?: string;
  start: string;
  end: string;
  category?: string;
  location?: string;
  description?: string;
  participantRecordIds?: string[];
  companyRecordIds?: string[];
  attachments?: AirtableAttachmentInput[];
  message: OutlookMessageMetadata;
}

export interface AirtableDocumentPayload {
  project?: string;
  projectRecordId?: string;
  type: "attachment" | "link";
  label: string;
  url?: string;
  attachment?: OutlookAttachmentPreview;
  attachments?: AirtableAttachmentInput[];
  message: OutlookMessageMetadata;
}

export interface AirtableNotePayload {
  title: string;
  note: string;
  projectRecordId?: string;
  art?: string;
  personRecordIds?: string[];
  date?: string;
  message: OutlookMessageMetadata;
}

export interface AirtablePersonPayload {
  name: string;
  email?: string;
  phoneMobile?: string;
  phone?: string;
  roleValues?: string[];
  position?: string;
  companyRecordIds?: string[];
}

export interface AirtableCompanyPayload {
  name: string;
  email?: string;
  phone?: string;
  website?: string;
  street?: string;
  houseNumber?: string;
  zip?: number;
  city?: string;
  country?: string;
  language?: string;
  categories?: string[];
}

export interface AirtableCompanyOption {
  id: string;
  name: string;
  email?: string;
  website?: string;
}

export interface AirtableRecordResponse {
  id: string;
  createdTime: string;
}

export interface AirtableProjectOption {
  id: string;
  name: string;
  email?: string;
}

export interface CollaboratorOption {
  id?: string;
  email?: string;
  name?: string;
}
