import { config } from "../config";
import { graph, graphAll, GraphError } from "./client";

/* Subset of Graph's columnDefinition that the form needs. */
interface ColumnDefinition {
  id: string;
  name: string;
  displayName: string;
  description?: string;
  hidden?: boolean;
  readOnly?: boolean;
  required?: boolean;
  defaultValue?: { value?: string } | null;
  text?: { allowMultipleLines?: boolean; maxLength?: number; linesForEditing?: number };
  choice?: { choices: string[]; displayAs?: "checkBoxes" | "dropDownMenu" | "radioButtons"; allowTextEntry?: boolean };
  number?: { minimum?: number; maximum?: number; decimalPlaces?: string };
  currency?: { locale?: string };
  dateTime?: { format?: "dateOnly" | "dateTime" };
  boolean?: object;
  personOrGroup?: { allowMultipleSelection?: boolean; chooseFromType?: string };
  lookup?: { listId: string; columnName: string; allowMultipleValues?: boolean };
  hyperlinkOrPicture?: { isPicture?: boolean };
  calculated?: object;
}

export type FieldKind =
  | "text"
  | "note"
  | "choice"
  | "multichoice"
  | "number"
  | "date"
  | "datetime"
  | "boolean"
  | "person"
  | "lookup"
  | "url";

export interface Field {
  name: string; // internal name, the key Graph writes to
  label: string; // display name, exactly as in SharePoint
  description?: string;
  required: boolean;
  kind: FieldKind;
  choices?: string[];
  allowTextEntry?: boolean;
  radio?: boolean;
  multiple?: boolean;
  maxLength?: number;
  min?: number;
  max?: number;
  lookupListId?: string;
  lookupColumn?: string;
  defaultValue?: string;
}

export interface ListSchema {
  siteId: string;
  listId: string;
  listName: string;
  listUrl: string;
  fields: Field[];
  /** Columns the form cannot write (managed metadata, location...). */
  skipped: string[];
}

// System columns that are editable on paper but never appear on a form.
const SYSTEM = new Set(["ContentType", "Attachments", "_ColorTag", "ComplianceAssetId", "Edit", "LinkTitle", "LinkTitleNoMenu", "DocIcon", "ItemChildCount", "FolderChildCount", "AppAuthor", "AppEditor"]);

function toField(c: ColumnDefinition): Field | null {
  const base = {
    name: c.name,
    label: c.displayName,
    description: c.description || undefined,
    required: !!c.required,
    defaultValue: c.defaultValue?.value ?? undefined,
  };
  if (c.text) return { ...base, kind: c.text.allowMultipleLines ? "note" : "text", maxLength: c.text.maxLength };
  if (c.choice) {
    const multi = c.choice.displayAs === "checkBoxes";
    return {
      ...base,
      kind: multi ? "multichoice" : "choice",
      choices: c.choice.choices,
      allowTextEntry: c.choice.allowTextEntry,
      radio: c.choice.displayAs === "radioButtons",
    };
  }
  if (c.number) return { ...base, kind: "number", min: c.number.minimum, max: c.number.maximum };
  if (c.currency) return { ...base, kind: "number" };
  if (c.dateTime) return { ...base, kind: c.dateTime.format === "dateTime" ? "datetime" : "date" };
  if (c.boolean) return { ...base, kind: "boolean" };
  if (c.personOrGroup) return { ...base, kind: "person", multiple: !!c.personOrGroup.allowMultipleSelection };
  if (c.lookup)
    return {
      ...base,
      kind: "lookup",
      multiple: !!c.lookup.allowMultipleValues,
      lookupListId: c.lookup.listId,
      lookupColumn: c.lookup.columnName,
    };
  if (c.hyperlinkOrPicture && !c.hyperlinkOrPicture.isPicture) return { ...base, kind: "url" };
  return null;
}

export function siteRef(): string {
  // "host" or "host:/sites/x" -> Graph site address
  const [host, path] = config.site.split(":");
  return path ? `/sites/${host}:${path}` : `/sites/${host}`;
}

/** VITE_SP_LIST does not match any list on the site; carries the names that do exist. */
export class ListNotFoundError extends Error {
  constructor(public available: string[]) {
    super(`Lista “${config.list}” non trovata`);
  }
}

export async function loadSchema(token: string): Promise<ListSchema> {
  const site = await graph<{ id: string }>(token, `${siteRef()}?$select=id`);
  const list = await graph<{ id: string; displayName: string; webUrl: string }>(
    token,
    `/sites/${site.id}/lists/${encodeURIComponent(config.list)}?$select=id,displayName,webUrl`
  ).catch(async (e) => {
    if (!(e instanceof GraphError) || e.status !== 404) throw e;
    const lists = await graphAll<{ displayName: string; list?: { hidden?: boolean } }>(
      token,
      `/sites/${site.id}/lists?$select=displayName,list`
    );
    throw new ListNotFoundError(lists.filter((l) => !l.list?.hidden).map((l) => l.displayName).sort());
  });

  // The default content type's columns come back in form order; the list's
  // own column collection does not.
  const types = await graphAll<{ id: string; hidden?: boolean; name: string }>(
    token,
    `/sites/${site.id}/lists/${list.id}/contentTypes?$select=id,name,hidden`
  );
  const ct = types.find((t) => !t.hidden && !t.id.startsWith("0x0120")) ?? types[0];
  const columns = await graphAll<ColumnDefinition>(token, `/sites/${site.id}/lists/${list.id}/contentTypes/${ct.id}/columns`);

  let candidates = columns.filter((c) => !c.hidden && !c.readOnly && !c.calculated && !SYSTEM.has(c.name));
  if (config.fields.length) {
    const byName = new Map(columns.map((c) => [c.name, c]));
    candidates = config.fields.map((n) => byName.get(n)).filter((c): c is ColumnDefinition => !!c);
  }

  const fields: Field[] = [];
  const skipped: string[] = [];
  for (const c of candidates) {
    const f = toField(c);
    if (f) fields.push(f);
    else skipped.push(c.displayName);
  }

  return { siteId: site.id, listId: list.id, listName: list.displayName, listUrl: list.webUrl, fields, skipped };
}

/* ---- People -------------------------------------------------------------- */

export interface Person {
  id: string;
  displayName: string;
  mail: string;
}

export async function searchPeople(token: string, q: string): Promise<Person[]> {
  const term = q.replace(/"/g, "");
  const r = await graph<{ value: { id: string; displayName: string; mail: string | null; userPrincipalName: string }[] }>(
    token,
    `/users?$search="displayName:${term}" OR "mail:${term}"&$select=id,displayName,mail,userPrincipalName&$top=8`,
    { headers: { ConsistencyLevel: "eventual" } }
  );
  return r.value.map((u) => ({ id: u.id, displayName: u.displayName, mail: u.mail ?? u.userPrincipalName }));
}

/**
 * Person columns are written as the user's id in the site's hidden
 * "User Information List". A user who never opened the site is not there yet.
 */
async function personLookupId(token: string, siteId: string, mail: string): Promise<string> {
  const r = await graph<{ value: { id: string }[] }>(
    token,
    `/sites/${siteId}/lists/User Information List/items?$select=id&$filter=fields/EMail eq '${mail.replace(/'/g, "''")}'`,
    { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } }
  );
  if (!r.value.length) {
    throw new Error(`${mail} non è ancora registrato sul sito SharePoint: deve aprirlo almeno una volta prima di poter essere selezionato.`);
  }
  return r.value[0].id;
}

/* ---- Lookups ------------------------------------------------------------- */

export interface LookupOption {
  id: string;
  label: string;
}

export async function loadLookupOptions(token: string, siteId: string, listId: string, column: string): Promise<LookupOption[]> {
  const items = await graphAll<{ id: string; fields: Record<string, unknown> }>(
    token,
    `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=${column})&$top=999`
  );
  return items
    .map((i) => ({ id: i.id, label: String(i.fields[column] ?? `#${i.id}`) }))
    .sort((a, b) => a.label.localeCompare(b.label, "it"));
}

/* ---- Create -------------------------------------------------------------- */

export type FieldValue = string | string[] | boolean | Person[] | undefined;

export function isEmpty(v: FieldValue): boolean {
  return v === undefined || v === "" || (Array.isArray(v) && v.length === 0);
}

export async function createItem(
  token: string,
  schema: ListSchema,
  values: Record<string, FieldValue>
): Promise<{ id: string; webUrl: string }> {
  const fields: Record<string, unknown> = {};

  for (const f of schema.fields) {
    const v = values[f.name];
    if (f.kind === "boolean") {
      fields[f.name] = !!v;
      continue;
    }
    if (isEmpty(v)) continue;

    switch (f.kind) {
      case "number":
        fields[f.name] = Number(v);
        break;
      case "date":
        // Midday UTC keeps the same calendar day in the site's time zone.
        fields[f.name] = `${v}T12:00:00Z`;
        break;
      case "datetime":
        fields[f.name] = new Date(v as string).toISOString();
        break;
      case "multichoice":
        fields[`${f.name}@odata.type`] = "Collection(Edm.String)";
        fields[f.name] = v;
        break;
      case "lookup":
        if (f.multiple) {
          fields[`${f.name}LookupId@odata.type`] = "Collection(Edm.Int32)";
          fields[`${f.name}LookupId`] = (v as string[]).map(Number);
        } else {
          fields[`${f.name}LookupId`] = Number(v);
        }
        break;
      case "person": {
        const ids = await Promise.all((v as Person[]).map((p) => personLookupId(token, schema.siteId, p.mail)));
        if (f.multiple) {
          fields[`${f.name}LookupId@odata.type`] = "Collection(Edm.String)";
          fields[`${f.name}LookupId`] = ids;
        } else {
          fields[`${f.name}LookupId`] = ids[0];
        }
        break;
      }
      default:
        fields[f.name] = v;
    }
  }

  return graph(token, `/sites/${schema.siteId}/lists/${schema.listId}/items`, {
    method: "POST",
    body: { fields },
  });
}
