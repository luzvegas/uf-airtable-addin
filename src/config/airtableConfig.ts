export interface AirtableEnvironmentConfig {
  /**
   * Personal Access Token (PAT) that can access all tables you want to use.
   * Replace the placeholder before shipping or read it from a secure storage location.
   */
  personalAccessToken: string;
  /**
   * Base IDs per logical module. They can point to the same Airtable base if desired.
   */
  baseIds: {
    tasks: string;
    events: string;
    documents: string;
    projects: string;
    persons: string;
    notes: string;
  };
  /**
   * Table names that will receive the records.
   */
  tableNames: {
    tasks: string;
    events: string;
    documents: string;
    projects: string;
    persons: string;
    notes: string;
  };
  /**
   * Optional proxy URL to avoid exposing PATs in the frontend (Azure Function).
   */
  proxyUrl?: string;
}

/**
 * Default configuration used by the in-app client. All values are placeholders to prevent
 * unintentional uploads. Replace them manually or load them dynamically (e.g. from a secure API).
 */
const FALLBACK_TOKEN = "YOUR_AIRTABLE_PERSONAL_ACCESS_TOKEN";
const FALLBACK_BASES = {
  tasks: "AIRTABLE_BASE_ID_TASKS",
  events: "AIRTABLE_BASE_ID_EVENTS",
  documents: "AIRTABLE_BASE_ID_DOCUMENTS",
  projects: "AIRTABLE_BASE_ID_TASKS",
  persons: "AIRTABLE_BASE_ID_TASKS",
  notes: "AIRTABLE_BASE_ID_TASKS",
};

const FALLBACK_TABLES = {
  tasks: "Tasks&Termine",
  events: "Tasks&Termine",
  documents: "Dokumente und Links",
  projects: "Projects",
  persons: "Personen",
  notes: "Gesprächsnotizen",
};

function readEnv(key: string): string | undefined {
  try {
    if (typeof process !== "undefined" && (process as any).env && typeof (process as any).env[key] === "string") {
      return (process as any).env[key] as string;
    }
  } catch (e) {
    // ignore
  }
  return undefined;
}

function envOrFallback(value: string | undefined, fallback: string) {
  return value && value.trim().length > 0 ? value : fallback;
}

export const defaultAirtableConfig: AirtableEnvironmentConfig = {
  personalAccessToken: envOrFallback(readEnv("AIRTABLE_PAT"), FALLBACK_TOKEN),
  baseIds: {
    tasks: envOrFallback(readEnv("AIRTABLE_BASE_TASKS"), FALLBACK_BASES.tasks),
    events: envOrFallback(readEnv("AIRTABLE_BASE_EVENTS"), FALLBACK_BASES.events),
    documents: envOrFallback(readEnv("AIRTABLE_BASE_DOCUMENTS"), FALLBACK_BASES.documents),
    projects: envOrFallback(readEnv("AIRTABLE_BASE_PROJECTS"), FALLBACK_BASES.projects),
    persons: envOrFallback(readEnv("AIRTABLE_BASE_PERSONS"), FALLBACK_BASES.persons),
    notes: envOrFallback(readEnv("AIRTABLE_BASE_NOTES"), FALLBACK_BASES.notes),
  },
  tableNames: {
    tasks: envOrFallback(readEnv("AIRTABLE_TABLE_TASKS"), FALLBACK_TABLES.tasks),
    events: envOrFallback(readEnv("AIRTABLE_TABLE_EVENTS"), FALLBACK_TABLES.events),
    documents: envOrFallback(readEnv("AIRTABLE_TABLE_DOCUMENTS"), FALLBACK_TABLES.documents),
    projects: envOrFallback(readEnv("AIRTABLE_TABLE_PROJECTS"), FALLBACK_TABLES.projects),
    persons: envOrFallback(readEnv("AIRTABLE_TABLE_PERSONS"), FALLBACK_TABLES.persons),
    notes: envOrFallback(readEnv("AIRTABLE_TABLE_NOTES"), FALLBACK_TABLES.notes),
  },
  proxyUrl: readEnv("AIRTABLE_PROXY_URL"),
};
