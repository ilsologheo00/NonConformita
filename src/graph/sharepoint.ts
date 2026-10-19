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
  /** The site's hidden user list, where person columns point; its name depends on the site language. */
  userListId?: string;
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

/** The configured list name does not match any list on the site; carries the names that do exist. */
export class ListNotFoundError extends Error {
  constructor(public list: string, public available: string[]) {
    super(`Lista “${list}” non trovata`);
  }
}

/**
 * Lists are matched by display name (or GUID) among the site's lists rather than
 * addressed by name in the URL: names like "Azioni Correttive/Miglioramenti"
 * hold characters Graph cannot take in a path.
 */
export async function loadSchema(token: string, listName: string = config.list): Promise<ListSchema> {
  const site = await graph<{ id: string }>(token, `${siteRef()}?$select=id`);
  const lists = await graphAll<{ id: string; displayName: string; webUrl: string; list?: { hidden?: boolean } }>(
    token,
    `/sites/${site.id}/lists?$select=id,displayName,webUrl,list`
  );
  const wanted = listName.trim().toLowerCase();
  const list = lists.find((l) => l.id.toLowerCase() === wanted) ?? lists.find((l) => l.displayName.trim().toLowerCase() === wanted);
  if (!list) throw new ListNotFoundError(listName, lists.filter((l) => !l.list?.hidden).map((l) => l.displayName).sort());

  // The default content type's columns come back in form order; the list's
  // own column collection does not.
  const types = await graphAll<{ id: string; hidden?: boolean; name: string }>(
    token,
    `/sites/${site.id}/lists/${list.id}/contentTypes?$select=id,name,hidden`
  );
  const ct = types.find((t) => !t.hidden && !t.id.startsWith("0x0120")) ?? types[0];
  const [ctColumns, listColumns] = await Promise.all([
    graphAll<ColumnDefinition>(token, `/sites/${site.id}/lists/${list.id}/contentTypes/${ct.id}/columns`),
    graphAll<ColumnDefinition>(token, `/sites/${site.id}/lists/${list.id}/columns`),
  ]);
  // A column deleted and re-created on the list may be left out of the content
  // type: it still belongs on the form, after the content type's own columns.
  const inCt = new Set(ctColumns.map((c) => c.name));
  const columns = [...ctColumns, ...listColumns.filter((c) => !inCt.has(c.name))];

  let candidates = columns.filter((c) => !c.hidden && !c.readOnly && !c.calculated && !SYSTEM.has(c.name));
  // VITE_SP_FIELDS picks the columns of the non-conformity list only.
  if (config.fields.length && listName === config.list) {
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

  // Found by its fixed address, _catalogs/users; the display name is translated.
  const users =
    lists.find((l) => /\/_catalogs\/users\/?$/i.test(l.webUrl.replace(/\/[^/]*\.aspx$/i, ""))) ??
    lists.find((l) => /^(user information list|elenco informazioni utente)$/i.test(l.displayName.trim()));
  return { siteId: site.id, listId: list.id, listName: list.displayName, listUrl: list.webUrl, fields, skipped, userListId: users?.id };
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
 * user list (_catalogs/users). A user who never opened the site is not there yet.
 */
// Graph does not always list it among the site's lists, but takes its title in
// the path; the title is translated with the site language.
const USER_LIST_TITLES = ["User Information List", "Elenco informazioni utente", "Elenco informazioni utenti", "Lista informazioni utente"];
const userLists = new Map<string, Promise<string>>();

function userList(token: string, schema: ListSchema): Promise<string> {
  if (schema.userListId) return Promise.resolve(schema.userListId);
  let found = userLists.get(schema.siteId);
  if (!found) {
    found = (async () => {
      for (const title of USER_LIST_TITLES) {
        try {
          return (await graph<{ id: string }>(token, `/sites/${schema.siteId}/lists/${encodeURIComponent(title)}?$select=id`)).id;
        } catch (e) {
          if (!(e instanceof GraphError) || e.status !== 404) throw e;
        }
      }
      throw new Error(`Elenco utenti del sito (_catalogs/users) non trovato con i nomi ${USER_LIST_TITLES.join(", ")}: impossibile salvare le colonne Persona.`);
    })();
    found.catch(() => userLists.delete(schema.siteId));
    userLists.set(schema.siteId, found);
  }
  return found;
}

async function personLookupId(token: string, schema: ListSchema, mail: string): Promise<string> {
  const r = await graph<{ value: { id: string }[] }>(
    token,
    `/sites/${schema.siteId}/lists/${await userList(token, schema)}/items?$select=id&$filter=fields/EMail eq '${mail.replace(/'/g, "''")}'`,
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

/**
 * Values of a lookup column. byId puts the newest item first and prefixes its
 * ID, for lookups on records such as the non-conformities themselves.
 */
export async function loadLookupOptions(token: string, siteId: string, listId: string, column: string, byId = false): Promise<LookupOption[]> {
  const items = await graphAll<{ id: string; fields: Record<string, unknown> }>(
    token,
    `/sites/${siteId}/lists/${listId}/items?$expand=fields($select=${column})&$top=999`
  );
  const options = items.map((i) => {
    const value = i.fields[column] === undefined || i.fields[column] === null ? "" : String(i.fields[column]);
    return { id: i.id, label: byId ? (value && value !== i.id ? `${i.id} – ${value}` : i.id) : value || `#${i.id}` };
  });
  return byId ? options.sort((a, b) => Number(b.id) - Number(a.id)) : options.sort((a, b) => a.label.localeCompare(b.label, "it"));
}

/* ---- Create -------------------------------------------------------------- */

export type FieldValue = string | string[] | boolean | Person[] | undefined;

export function isEmpty(v: FieldValue): boolean {
  return v === undefined || v === "" || (Array.isArray(v) && v.length === 0);
}

/**
 * Form values as Graph's fields payload. Only the columns in only are written
 * (all when absent); an emptied column is cleared, which matters when editing.
 */
async function toGraphFields(token: string, schema: ListSchema, values: Record<string, FieldValue>, only?: Set<string>): Promise<Record<string, unknown>> {
  const fields: Record<string, unknown> = {};

  for (const f of schema.fields) {
    if (only && !only.has(f.name)) continue;
    const v = values[f.name];
    if (f.kind === "boolean") {
      fields[f.name] = !!v;
      continue;
    }
    if (isEmpty(v)) {
      if (!only) continue;
      if (f.kind === "multichoice") {
        fields[`${f.name}@odata.type`] = "Collection(Edm.String)";
        fields[f.name] = [];
      } else if ((f.kind === "lookup" || f.kind === "person") && f.multiple) {
        fields[`${f.name}LookupId@odata.type`] = f.kind === "lookup" ? "Collection(Edm.Int32)" : "Collection(Edm.String)";
        fields[`${f.name}LookupId`] = [];
      } else if (f.kind === "lookup" || f.kind === "person") {
        fields[`${f.name}LookupId`] = null;
      } else {
        fields[f.name] = null;
      }
      continue;
    }

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
        const ids = await Promise.all((v as Person[]).map((p) => personLookupId(token, schema, p.mail)));
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
  return fields;
}

export async function createItem(
  token: string,
  schema: ListSchema,
  values: Record<string, FieldValue>
): Promise<{ id: string; webUrl: string }> {
  return graph(token, `/sites/${schema.siteId}/lists/${schema.listId}/items`, {
    method: "POST",
    body: { fields: await toGraphFields(token, schema, values) },
  });
}

/* ---- Edit ---------------------------------------------------------------- */

/** One item as form values, with the eTag that tells whether someone changed it since. */
export interface EditableItem {
  id: string;
  eTag: string;
  webUrl: string;
  values: Record<string, FieldValue>;
}

/** The item was saved by someone else after it was opened (HTTP 412). */
export class ItemChangedError extends Error {
  constructor() {
    super("Nel frattempo qualcun altro ha modificato questo elemento.");
  }
}

const pad = (n: number) => String(n).padStart(2, "0");
const localDate = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;

/** Single-person columns only carry the user's id in the site's user list: read back name and e-mail. */
async function personById(token: string, schema: ListSchema, id: string): Promise<Person> {
  const u = await graph<{ fields: { Title?: string; EMail?: string } }>(
    token,
    `/sites/${schema.siteId}/lists/${await userList(token, schema)}/items/${id}?$expand=fields($select=Title,EMail)`
  );
  return { id: `sp:${id}`, displayName: u.fields.Title ?? `#${id}`, mail: u.fields.EMail ?? "" };
}

async function formValue(token: string, schema: ListSchema, f: Field, raw: Record<string, unknown>): Promise<FieldValue> {
  const v = raw[f.name];
  switch (f.kind) {
    case "boolean":
      return v === true;
    case "multichoice":
      return Array.isArray(v) ? v.map(String) : [];
    case "number":
      return typeof v === "number" ? String(v) : "";
    case "date":
      return typeof v === "string" ? localDate(new Date(v)) : "";
    case "datetime": {
      if (typeof v !== "string") return "";
      const d = new Date(v);
      return `${localDate(d)}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
    }
    case "lookup": {
      const ids = lookupIds(f, raw);
      return f.multiple ? ids : ids[0] ?? "";
    }
    case "person": {
      if (Array.isArray(v)) return (v as LookupValue[]).map((p) => ({ id: `sp:${p.LookupId}`, displayName: p.LookupValue ?? "", mail: p.Email ?? "" }));
      const id = raw[`${f.name}LookupId`];
      return id === undefined || id === null || id === "" ? [] : [await personById(token, schema, String(id))];
    }
    case "url": {
      const url = v && typeof v === "object" ? (v as { Url?: string }).Url : v;
      return typeof url === "string" ? url : "";
    }
    case "note":
      // Rich-text columns come back as HTML; the form edits plain text.
      if (typeof v !== "string") return "";
      return /^\s*</.test(v) ? new DOMParser().parseFromString(v, "text/html").body.textContent?.trim() ?? "" : v;
    default:
      return v === undefined || v === null ? "" : String(v);
  }
}

export async function loadItem(token: string, schema: ListSchema, id: string): Promise<EditableItem> {
  const item = await graph<{ id: string; eTag: string; webUrl: string; fields: Record<string, unknown> }>(
    token,
    `/sites/${schema.siteId}/lists/${schema.listId}/items/${id}?$expand=fields`
  );
  const values: Record<string, FieldValue> = {};
  await Promise.all(schema.fields.map(async (f) => (values[f.name] = await formValue(token, schema, f, item.fields))));
  return { id: item.id, eTag: item.eTag, webUrl: item.webUrl, values };
}

/** Columns whose value differs from the item as it was read. */
export function changedFields(schema: ListSchema, item: EditableItem, values: Record<string, FieldValue>): Set<string> {
  const same = (a: FieldValue, b: FieldValue) => JSON.stringify(a ?? "") === JSON.stringify(b ?? "");
  return new Set(schema.fields.filter((f) => !same(values[f.name], item.values[f.name])).map((f) => f.name));
}

/**
 * Writes only the columns that changed since the item was read, and only if
 * nobody saved it in the meantime: SharePoint checks the eTag (If-Match).
 */
export async function updateItem(token: string, schema: ListSchema, item: EditableItem, values: Record<string, FieldValue>): Promise<void> {
  const changed = changedFields(schema, item, values);
  if (!changed.size) return;
  try {
    await graph(token, `/sites/${schema.siteId}/lists/${schema.listId}/items/${item.id}/fields`, {
      method: "PATCH",
      body: await toGraphFields(token, schema, values, changed),
      headers: { "If-Match": item.eTag },
    });
  } catch (e) {
    if (e instanceof GraphError && e.status === 412) throw new ItemChangedError();
    throw e;
  }
}

/* ---- Read ---------------------------------------------------------------- */

export interface ListItem {
  id: string;
  created: string;
  createdBy: string;
  /** Display text per column, keyed by internal name. */
  text: Record<string, string>;
  /** Value to sort by per column: numbers and ISO dates sort as themselves. */
  sort: Record<string, string | number>;
  /** Item ids each lookup column points at, keyed by internal name. */
  lookupIds: Record<string, string[]>;
  title: string;
}

/** Graph writes list ids with or without braces, in either case. */
export const sameList = (a = "", b = "") => a.replace(/[{}]/g, "").toLowerCase() === b.replace(/[{}]/g, "").toLowerCase();

function lookupIds(f: Field, raw: Record<string, unknown>): string[] {
  const v = raw[f.name];
  if (Array.isArray(v)) return v.map((x) => (x as LookupValue).LookupId).filter((x) => x !== undefined && x !== null).map(String);
  const id = raw[`${f.name}LookupId`];
  return id === undefined || id === null || id === "" ? [] : [String(id)];
}

type LookupValue = { LookupId?: number | string; LookupValue?: string; Email?: string };

const dateFmt = new Intl.DateTimeFormat("it-IT", { dateStyle: "short" });
const dateTimeFmt = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" });

function cell(f: Field, raw: Record<string, unknown>, people: Map<string, string>): [string, string | number] {
  const v = raw[f.name];
  const lookupId = raw[`${f.name}LookupId`];
  if (Array.isArray(v)) {
    const text = v.map((x) => (typeof x === "object" && x ? (x as LookupValue).LookupValue ?? "" : String(x))).filter(Boolean).join(", ");
    return [text, text.toLowerCase()];
  }
  switch (f.kind) {
    case "boolean":
      return v === undefined || v === null ? ["", ""] : [v ? "Sì" : "No", v ? 1 : 0];
    case "number":
      return typeof v === "number" ? [v.toLocaleString("it-IT"), v] : ["", ""];
    case "date":
    case "datetime": {
      if (typeof v !== "string") return ["", ""];
      return [(f.kind === "date" ? dateFmt : dateTimeFmt).format(new Date(v)), v];
    }
    case "person": {
      const name = lookupId !== undefined ? people.get(String(lookupId)) ?? `#${lookupId}` : "";
      return [name, name.toLowerCase()];
    }
    case "lookup": {
      const text = typeof v === "string" ? v : lookupId !== undefined ? `#${lookupId}` : "";
      return [text, text.toLowerCase()];
    }
    case "url": {
      const url = v && typeof v === "object" ? (v as { Url?: string; Description?: string }) : null;
      const text = url ? url.Description || url.Url || "" : typeof v === "string" ? v : "";
      return [text, text.toLowerCase()];
    }
    case "note": {
      // Rich-text columns come back as HTML.
      const text = typeof v === "string" ? new DOMParser().parseFromString(v, "text/html").body.textContent?.trim() ?? "" : "";
      return [text, text.toLowerCase()];
    }
    default: {
      const text = v === undefined || v === null ? "" : String(v);
      return [text, text.toLowerCase()];
    }
  }
}

/** Every item of the list, with the form's columns turned into readable text. */
export async function loadItems(token: string, schema: ListSchema): Promise<ListItem[]> {
  const [items, people] = await Promise.all([
    graphAll<{ id: string; createdDateTime: string; createdBy?: { user?: { displayName?: string } }; fields: Record<string, unknown> }>(
      token,
      `/sites/${schema.siteId}/lists/${schema.listId}/items?$expand=fields&$top=500`
    ),
    // Single-person columns only carry the user's id in the site's user list.
    schema.fields.some((f) => f.kind === "person")
      ? userList(token, schema)
          .then((id) => graphAll<{ id: string; fields: { Title?: string } }>(token, `/sites/${schema.siteId}/lists/${id}/items?$expand=fields($select=Title)&$top=999`))
          .then((users) => new Map(users.map((u) => [u.id, u.fields.Title ?? `#${u.id}`])))
      : Promise.resolve(new Map<string, string>()),
  ]);

  return items.map((i) => {
    const text: Record<string, string> = {};
    const sort: Record<string, string | number> = {};
    const ids: Record<string, string[]> = {};
    for (const f of schema.fields) {
      [text[f.name], sort[f.name]] = cell(f, i.fields, people);
      if (f.kind === "lookup") ids[f.name] = lookupIds(f, i.fields);
    }
    const title = typeof i.fields.Title === "string" ? i.fields.Title : "";
    return { id: i.id, created: i.createdDateTime, createdBy: i.createdBy?.user?.displayName ?? "", text, sort, lookupIds: ids, title };
  });
}

/** Items of another list whose lookup column points at this list, grouped by the item they point at. */
export async function loadLinkedItems(token: string, schema: ListSchema, otherList: string): Promise<{ schema: ListSchema; byItem: Map<string, ListItem[]> } | null> {
  const other = await loadSchema(token, otherList);
  const field = other.fields.find((f) => f.kind === "lookup" && sameList(f.lookupListId, schema.listId));
  if (!field) return null;
  const byItem = new Map<string, ListItem[]>();
  for (const x of await loadItems(token, other)) {
    for (const id of x.lookupIds[field.name] ?? []) byItem.set(id, [...(byItem.get(id) ?? []), x]);
  }
  for (const xs of byItem.values()) xs.sort((a, b) => Number(a.id) - Number(b.id));
  return { schema: other, byItem };
}
