const GRAPH = "https://graph.microsoft.com/v1.0";

export class GraphError extends Error {
  constructor(public status: number, public code: string, message: string) {
    super(message);
  }
}

export async function graph<T>(
  token: string,
  path: string,
  init: { method?: string; body?: unknown; headers?: Record<string, string> } = {}
): Promise<T> {
  const res = await fetch(path.startsWith("http") ? path : GRAPH + path, {
    method: init.method ?? "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      ...(init.body !== undefined ? { "Content-Type": "application/json" } : {}),
      ...init.headers,
    },
    body: init.body !== undefined ? JSON.stringify(init.body) : undefined,
  });
  if (!res.ok) {
    let code = String(res.status);
    let message = res.statusText;
    try {
      const j = await res.json();
      code = j.error?.code ?? code;
      message = j.error?.message ?? message;
    } catch {
      /* body was not JSON */
    }
    throw new GraphError(res.status, code, message);
  }
  return res.status === 204 ? (undefined as T) : res.json();
}

/** Follows @odata.nextLink until the collection is exhausted. */
export async function graphAll<T>(token: string, path: string, headers?: Record<string, string>): Promise<T[]> {
  const out: T[] = [];
  let next: string | undefined = path;
  while (next) {
    const page: { value: T[]; "@odata.nextLink"?: string } = await graph(token, next, { headers });
    out.push(...page.value);
    next = page["@odata.nextLink"];
  }
  return out;
}
