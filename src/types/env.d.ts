declare const process: {
  env: {
    AIRTABLE_PAT?: string;
    AIRTABLE_BASE_TASKS?: string;
    AIRTABLE_BASE_EVENTS?: string;
    AIRTABLE_BASE_DOCUMENTS?: string;
    AIRTABLE_BASE_PROJECTS?: string;
    AIRTABLE_BASE_PERSONS?: string;
    AIRTABLE_TABLE_TASKS?: string;
    AIRTABLE_TABLE_EVENTS?: string;
    AIRTABLE_TABLE_DOCUMENTS?: string;
    AIRTABLE_TABLE_PROJECTS?: string;
    AIRTABLE_TABLE_PERSONS?: string;
    NODE_ENV?: string;
    [key: string]: string | undefined;
  };
};

export {};
