/**
 * Values come from window.__ENV__ (written into /env.js by the Docker container
 * at start-up) and fall back to the .env file read by Vite at build time.
 */
const runtime = window.__ENV__ ?? {};
const build = import.meta.env;

function read(key: keyof ImportMetaEnv): string {
  return (runtime[key] || build[key] || "").trim();
}

export const config = {
  clientId: read("VITE_AZURE_CLIENT_ID"),
  tenantId: read("VITE_AZURE_TENANT_ID"),
  site: read("VITE_SP_SITE") || "semantica.sharepoint.com",
  list: read("VITE_SP_LIST"),
  operatorsList: read("VITE_SP_OPERATORS_LIST") || "Operatori NC",
  actionsList: read("VITE_SP_ACTIONS_LIST") || "Azioni Correttive/Miglioramenti",
  activitiesList: read("VITE_SP_ACTIVITIES_LIST") || "Attività proposte",
  complaintsList: read("VITE_SP_COMPLAINTS_LIST") || "Reclami commerciali",
  fields: read("VITE_SP_FIELDS")
    .split(",")
    .map((f) => f.trim())
    .filter(Boolean),
};

/** Names of the .env keys that still need a value before the app can sign in. */
export function missingConfig(): string[] {
  const missing: string[] = [];
  if (!config.clientId) missing.push("VITE_AZURE_CLIENT_ID");
  if (!config.tenantId) missing.push("VITE_AZURE_TENANT_ID");
  if (!config.list) missing.push("VITE_SP_LIST");
  return missing;
}
