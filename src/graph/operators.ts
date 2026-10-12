import { config } from "../config";
import { graph, graphAll, GraphError } from "./client";
import { siteRef } from "./sharepoint";

/** One row of the Operatori list: Title holds the name, Email the address. */
export interface OperatorItem {
  id: string;
  name: string;
  email: string;
}

export class OperatorsListMissingError extends Error {
  constructor() {
    super(`La lista “${config.operatorsList}” non esiste sul sito “${config.site}”.`);
  }
}

export interface OperatorsList {
  siteId: string;
  listId: string;
  webUrl: string;
}

export async function openOperatorsList(token: string): Promise<OperatorsList> {
  const site = await graph<{ id: string }>(token, `${siteRef()}?$select=id`);
  try {
    const list = await graph<{ id: string; webUrl: string }>(token, `/sites/${site.id}/lists/${encodeURIComponent(config.operatorsList)}?$select=id,webUrl`);
    return { siteId: site.id, listId: list.id, webUrl: list.webUrl };
  } catch (e) {
    if (e instanceof GraphError && e.status === 404) throw new OperatorsListMissingError();
    throw e;
  }
}

const items = (l: OperatorsList) => `/sites/${l.siteId}/lists/${l.listId}/items`;

export async function loadOperators(token: string, l: OperatorsList): Promise<OperatorItem[]> {
  const rows = await graphAll<{ id: string; fields: { Title?: string; Email?: string } }>(token, `${items(l)}?$expand=fields($select=Title,Email)&$top=500`);
  return rows
    .filter((r) => r.fields.Title?.trim() && r.fields.Email?.trim())
    .map((r) => ({ id: r.id, name: r.fields.Title!.trim(), email: r.fields.Email!.trim() }));
}

export async function createOperator(token: string, l: OperatorsList, name: string, email: string): Promise<string> {
  const r = await graph<{ id: string }>(token, items(l), { method: "POST", body: { fields: { Title: name, Email: email } } });
  return r.id;
}

export function renameOperator(token: string, l: OperatorsList, id: string, name: string): Promise<unknown> {
  return graph(token, `${items(l)}/${id}/fields`, { method: "PATCH", body: { Title: name } });
}

export function deleteOperator(token: string, l: OperatorsList, id: string): Promise<unknown> {
  return graph(token, `${items(l)}/${id}`, { method: "DELETE" });
}
