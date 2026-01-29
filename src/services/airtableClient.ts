import { defaultAirtableConfig } from "../config/airtableConfig";
import {
  AirtableDocumentPayload,
  AirtableProjectOption,
  AirtableRecordResponse,
  AirtableNotePayload,
  AirtableTaskPayload,
  CollaboratorOption,
} from "../types/airtable";

const AIRTABLE_REST_BASE = "https://api.airtable.com/v0";

export class AirtableClient {
  constructor(private readonly config = defaultAirtableConfig) {}

  private hasValidToken(): boolean {
    if (this.config.proxyUrl && this.config.proxyUrl.trim().length > 0) {
      return true;
    }
    const token = this.config.personalAccessToken?.trim();
    return Boolean(token && !token.startsWith("YOUR_"));
  }

  private getBaseUrl(): string {
    const proxy = this.config.proxyUrl?.trim();
    if (proxy) {
      return proxy.replace(/\/+$/, "");
    }
    return AIRTABLE_REST_BASE;
  }

  private buildHeaders(): Record<string, string> {
    if (this.config.proxyUrl && this.config.proxyUrl.trim().length > 0) {
      return {
        "Content-Type": "application/json",
      };
    }
    return {
      Authorization: `Bearer ${this.config.personalAccessToken}`,
      "Content-Type": "application/json",
    };
  }

  private async createRecord(
    baseId: string,
    tableName: string,
    fields: Record<string, unknown>
  ): Promise<AirtableRecordResponse | null> {
    if (!this.hasValidToken()) {
      console.warn(
        "Airtable token not configured. Please provide AIRTABLE_* values in a local .env file or inject secrets at runtime."
      );
      return null;
    }

    const response = await fetch(`${this.getBaseUrl()}/${baseId}/${buildTablePath(tableName)}`, {
      method: "POST",
      headers: this.buildHeaders(),
      body: JSON.stringify({ fields }),
    });

    if (!response.ok) {
      const errorBody = await response.text();
      throw new Error(`Airtable error ${response.status}: ${errorBody}`);
    }

    return (await response.json()) as AirtableRecordResponse;
  }

  private async listRecords<TFields = Record<string, unknown>>(
    baseId: string,
    tableName: string,
    options?: { fields?: string[]; filterByFormula?: string; sort?: Array<{ field: string; direction?: "asc" | "desc" }> }
  ): Promise<Array<{ id: string; fields: TFields }>> {
    if (!this.hasValidToken()) {
      console.warn(
        "Airtable token not configured. Please provide AIRTABLE_* values in a local .env file or inject secrets at runtime."
      );
      return [];
    }

    const collected: Array<{ id: string; fields: TFields }> = [];
    let offset: string | undefined;

    do {
      const params = new URLSearchParams();
      if (offset) {
        params.append("offset", offset);
      }
      if (options?.fields?.length) {
        for (const field of options.fields) {
          params.append("fields[]", field);
        }
      }
      if (options?.filterByFormula) {
        params.append("filterByFormula", options.filterByFormula);
      }
      if (options?.sort?.length) {
        options.sort.forEach((entry, index) => {
          params.append(`sort[${index}][field]`, entry.field);
          if (entry.direction) {
            params.append(`sort[${index}][direction]`, entry.direction);
          }
        });
      }

      const query = params.toString();
      const url = `${this.getBaseUrl()}/${baseId}/${buildTablePath(tableName)}${query ? `?${query}` : ""}`;
      const response = await fetch(url, {
        method: "GET",
        headers: this.config.proxyUrl ? undefined : { Authorization: `Bearer ${this.config.personalAccessToken}` },
      });

      if (!response.ok) {
        const errorBody = await response.text();
        throw new Error(`Airtable list error ${response.status}: ${errorBody}`);
      }

      const body = (await response.json()) as {
        records: Array<{ id: string; fields: TFields }>;
        offset?: string;
      };

      collected.push(...body.records);
      offset = body.offset;
    } while (offset);

    return collected;
  }

  async createTask(payload: AirtableTaskPayload) {
    // Use field IDs to avoid umlaut/label issues.
    const fields: Record<string, unknown> = {
      fldTO8z4W11jV8OhT: payload.title || payload.message.subject, // Titel
      fldux74lImjEo3619: payload.description ?? "", // Beschreibung
      ...(payload.status ? { fldRlnk0GJrnwRg6t: payload.status } : {}), // Status
      ...(payload.priority ? { fldxLsL8R48p1k5dQ: payload.priority } : {}), // Priorität
      ...(payload.category ? { fldZnMF441uknZeB6: payload.category } : {}), // Kategorie
      fldnat5YVKgTg0JlR: payload.start ?? null, // Start
      fldnSnu1YjOSIa8V1: payload.end ?? null, // Ende
      fldJ9gcK3Twv8Qq5Z: payload.art ?? "Task", // Art
    };

    if (payload.projectRecordId && payload.projectRecordId.startsWith("rec")) {
      fields.fldffhQ9PYUGWbhqe = [payload.projectRecordId]; // Projekt (linked record IDs)
    }

    if (payload.internalOwnerId) {
      fields.fldphfDf5RxQ9zOjt = { id: payload.internalOwnerId }; // Zuständig intern
    } else if (payload.internalOwnerEmail) {
      fields.fldphfDf5RxQ9zOjt = { email: payload.internalOwnerEmail };
    }

    if (payload.externalAssigneeIds?.length) {
      const validExternal = payload.externalAssigneeIds.filter((id) => id.startsWith("rec"));
      if (validExternal.length) {
        fields.fldSWTH4tAqe7wPTz = validExternal; // Zuständig extern (linked record IDs)
      }
    }

    if (payload.attachments?.length) {
      fields.fld1mSTdRf0OAgK71 = payload.attachments; // Attachments (expects URL + filename)
    }

    return this.createRecord(this.config.baseIds.tasks, this.config.tableNames.tasks, fields);
  }
  async createDocument(payload: AirtableDocumentPayload) {
    const fields: Record<string, unknown> = {
      fldetAgRaOKc7Smlr: payload.label ?? "", // Dokumententitel
    };

    if (payload.projectRecordId && payload.projectRecordId.startsWith("rec")) {
      fields.fldWYtTP4z7UQl704 = [payload.projectRecordId]; // Projekt (linked record)
    }

    if (payload.type === "attachment" && payload.attachments?.length) {
      fields.fldFFlZfxmURDHJJE = payload.attachments; // Datei (Attachments)
    }

    if (payload.type === "link" && payload.url) {
      fields.fldJJxnpWs4OsNHSq = payload.url; // Link URL
    }

    return this.createRecord(this.config.baseIds.documents, this.config.tableNames.documents, fields);
  }

  async createNote(payload: AirtableNotePayload) {
    const fields: Record<string, unknown> = {
      fldyfl5z48QM8kcRE: payload.title ?? payload.message.subject, // Titel
      fldnRlEsEgAhmGMrQ: payload.note ?? "", // Notiz
      fldMbjwoxt02Rxzj1: payload.art ?? "E-Mail", // Art
      fld0AEgf6HqaODqKU: payload.date ?? payload.message.receivedDate?.toISOString() ?? null, // Datum
    };

    if (payload.projectRecordId && payload.projectRecordId.startsWith("rec")) {
      fields.fldkYXvwccvXE11hR = [payload.projectRecordId]; // Projekt (linked record IDs)
    }

    if (payload.personRecordIds?.length) {
      const valid = payload.personRecordIds.filter((id) => id.startsWith("rec"));
      if (valid.length) {
        fields.fldZFkB4H5rUJYTOl = valid; // Person(en) (linked record IDs)
      }
    }

    return this.createRecord(this.config.baseIds.notes || this.config.baseIds.tasks, this.config.tableNames.notes, fields);
  }

  async fetchProjects(): Promise<AirtableProjectOption[]> {
    const baseId = this.config.baseIds.projects || this.config.baseIds.tasks;
    const tableName = this.config.tableNames.projects;

    const records = await this.listRecords<Record<string, unknown>>(baseId, tableName, {
      filterByFormula:
        "AND(NOT({Status}='Abgeschlossen'),NOT({Status}='On Hold'),NOT({Status}='Abbruch'))",
    });

    return records.map((record) => ({
      id: record.id,
      name: pickProjectName(record.fields) ?? record.id,
    }));
  }

  async fetchCollaborators(): Promise<CollaboratorOption[]> {
    const baseId = this.config.baseIds.tasks;
    const tableName = this.config.tableNames.tasks;
    const targetFieldName = "Zuständig intern";

    try {
      const metadataUrl = `https://api.airtable.com/v0/meta/bases/${baseId}/tables`;
      const response = await fetch(metadataUrl, {
        headers: {
          Authorization: `Bearer ${this.config.personalAccessToken}`,
        },
      });

      if (!response.ok) {
        throw new Error(`Airtable metadata error ${response.status}: ${await response.text()}`);
      }

      const body = (await response.json()) as {
        tables: Array<{
          name: string;
          fields: Array<{ name: string; type: string; options?: { choices?: Array<{ email?: string; id?: string; name?: string }> } }>;
        }>;
      };

      const table = body.tables.find((t) => t.name === tableName);
      if (!table) {
        return fallbackCollaborators();
      }
      const field = table.fields.find((f) => f.name === targetFieldName && f.type === "singleCollaborator");
      if (!field?.options?.choices) {
        return fallbackCollaborators();
      }

      return field.options.choices.map((choice) => ({
        id: choice.id,
        email: choice.email,
        name: choice.name ?? choice.email ?? choice.id,
      }));
    } catch (error) {
      console.warn("Collaborators fallback used due to error:", error);
      return fallbackCollaborators();
    }
  }

  async fetchExternalPersons(): Promise<AirtableProjectOption[]> {
    const baseId = this.config.baseIds.persons || this.config.baseIds.tasks;
    const tableName = this.config.tableNames.persons;
    if (!baseId || baseId.startsWith("AIRTABLE_BASE_ID")) {
      return [];
    }
    const records = await this.listRecords<Record<string, unknown>>(baseId, tableName);
    return records.map((record) => ({
      id: record.id,
      name: pickProjectName(record.fields) ?? record.id,
      email: pickEmail(record.fields),
    }));
  }
}

function fallbackCollaborators(): CollaboratorOption[] {
  // Minimal Fallback basierend auf bekannten Einträgen
  return [
    { id: "usrxc0jGBYawmG5lG", email: "luz@unserefarben.ch", name: "Luzius Müller" },
    { id: "usrauUprjIiYxbOmv", email: "patrischa@unserefarben.ch", name: "Patrischa Freuler" },
    { id: "usrAISERVICE00000", email: "aiassist@noreply.airtable.com", name: "AI" },
    { id: "usrWORKFLOWEXESVC", email: "automations@noreply.airtable.com", name: "Automations" },
  ];
}

function buildTablePath(nameOrId: string): string {
  if (nameOrId.startsWith("tbl")) {
    return nameOrId;
  }
  return encodeURIComponent(nameOrId);
}

function pickProjectName(fields: Record<string, unknown>): string | undefined {
  // Bevorzugte Felder
  if (typeof fields.Projekt === "string" && fields.Projekt.trim()) {
    return fields.Projekt.trim();
  }
  if (typeof fields.Name === "string" && fields.Name.trim()) {
    return fields.Name.trim();
  }
  if (typeof fields.Titel === "string" && fields.Titel.trim()) {
    return fields.Titel.trim();
  }

  // Fallback: erstes string-Feld im Record
  for (const value of Object.values(fields)) {
    if (typeof value === "string" && value.trim()) {
      return value.trim();
    }
  }
  return undefined;
}

function pickEmail(fields: Record<string, unknown>): string | undefined {
  if (typeof fields.Email === "string" && fields.Email.trim()) {
    return fields.Email.trim();
  }
  if (typeof fields.email === "string" && fields.email.trim()) {
    return fields.email.trim();
  }
  return undefined;
}

export const airtableClient = new AirtableClient();
