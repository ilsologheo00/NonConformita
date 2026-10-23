/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_AZURE_CLIENT_ID?: string;
  readonly VITE_AZURE_TENANT_ID?: string;
  readonly VITE_SP_SITE?: string;
  readonly VITE_SP_LIST?: string;
  readonly VITE_SP_FIELDS?: string;
  readonly VITE_SP_OPERATORS_LIST?: string;
  readonly VITE_SP_ACTIONS_LIST?: string;
  readonly VITE_SP_ACTIVITIES_LIST?: string;
  readonly VITE_SP_COMPLAINTS_LIST?: string;
}

interface Window {
  /** Runtime configuration injected by the Docker container (see public/env.js). */
  __ENV__?: Partial<Record<keyof ImportMetaEnv, string>>;
}
