import { useEffect, useMemo, useState } from "react";
import { loadItems, loadSchema, type ListItem, type ListSchema } from "../graph/sharepoint";
import { ExternalIcon, RefreshIcon } from "./icons";

type Status = { kind: "loading" } | { kind: "error"; message: string } | { kind: "ready"; schema: ListSchema; items: ListItem[] };
type Column = { key: string; label: string; text: (i: ListItem) => string; sort: (i: ListItem) => string | number };

const dateTimeFmt = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" });

function columns(schema: ListSchema): Column[] {
  return [
    { key: "ID", label: "ID", text: (i) => i.id, sort: (i) => Number(i.id) },
    ...schema.fields.map((f) => ({ key: f.name, label: f.label, text: (i: ListItem) => i.text[f.name], sort: (i: ListItem) => i.sort[f.name] })),
    { key: "Created", label: "Creato", text: (i) => dateTimeFmt.format(new Date(i.created)), sort: (i) => i.created },
    { key: "Author", label: "Creato da", text: (i) => i.createdBy, sort: (i) => i.createdBy.toLowerCase() },
  ];
}

/**
 * Read-only view of the SharePoint list. SharePoint refuses to be shown inside
 * another site's frame, so the items are read through Graph and drawn here.
 */
export function NcList({ getToken, active, explain }: { getToken: () => Promise<string>; active: boolean; explain: (e: unknown) => string }) {
  const [status, setStatus] = useState<Status>({ kind: "loading" });
  const [reload, setReload] = useState(0);
  const [query, setQuery] = useState("");
  const [sort, setSort] = useState<{ key: string; desc: boolean }>({ key: "ID", desc: true });

  // Reloaded every time the view is opened, so a just-sent NC shows up.
  useEffect(() => {
    if (!active) return;
    let alive = true;
    setStatus((s) => (s.kind === "ready" ? s : { kind: "loading" }));
    const load = import.meta.env.DEV && new URLSearchParams(window.location.search).has("anteprima")
      ? import("../preview").then((m) => ({ schema: m.previewSchema, items: m.previewItems }))
      : getToken().then(async (token) => {
          const schema = await loadSchema(token);
          return { schema, items: await loadItems(token, schema) };
        });
    load
      .then(({ schema, items }) => alive && setStatus({ kind: "ready", schema, items }))
      .catch((e) => alive && setStatus({ kind: "error", message: explain(e) }));
    return () => {
      alive = false;
    };
  }, [active, reload, getToken, explain]);

  const cols = useMemo(() => (status.kind === "ready" ? columns(status.schema) : []), [status]);
  const rows = useMemo(() => {
    if (status.kind !== "ready") return [];
    const q = query.trim().toLowerCase();
    const filtered = q ? status.items.filter((i) => cols.some((c) => c.text(i).toLowerCase().includes(q))) : status.items;
    const col = cols.find((c) => c.key === sort.key) ?? cols[0];
    return [...filtered].sort((a, b) => {
      const x = col.sort(a);
      const y = col.sort(b);
      // Empty cells always go last, whichever the direction.
      if (x === "" || y === "") return x === y ? 0 : x === "" ? 1 : -1;
      const c = typeof x === "number" && typeof y === "number" ? x - y : String(x).localeCompare(String(y), "it");
      return sort.desc ? -c : c;
    });
  }, [status, cols, query, sort]);

  if (status.kind === "loading") return <div className="page"><p className="muted">Lettura delle non conformità da SharePoint…</p></div>;
  if (status.kind === "error")
    return (
      <div className="page">
        <h1>Impossibile leggere la lista</h1>
        <div className="notice notice-error">
          <p>{status.message}</p>
          <button type="button" className="btn-secondary" onClick={() => setReload((n) => n + 1)}>
            Riprova
          </button>
        </div>
      </div>
    );

  const s = status.schema;
  const itemUrl = (id: string) => `${s.listUrl.replace(/\/$/, "")}/DispForm.aspx?ID=${id}`;

  return (
    <div className="page page-list">
      <header className="list-head">
        <div>
          <h1>Elenco non conformità</h1>
          <p className="muted">
            {status.items.length === 1 ? "1 elemento" : `${status.items.length} elementi`} nella lista “{s.listName}”. Fai clic sull'ID per aprire la scheda su SharePoint.
          </p>
        </div>
        <div className="list-actions">
          <button type="button" className="btn-secondary" onClick={() => setReload((n) => n + 1)}>
            <RefreshIcon />
            Aggiorna
          </button>
          <a className="btn-secondary" href={s.listUrl} target="_blank" rel="noreferrer">
            <ExternalIcon />
            Apri su SharePoint
          </a>
        </div>
      </header>

      <input type="search" className="list-search" placeholder="Cerca in tutte le colonne…" aria-label="Cerca" value={query} onChange={(e) => setQuery(e.target.value)} />

      <div className="list-table-wrap">
        <table className="list-table">
          <thead>
            <tr>
              {cols.map((c) => (
                <th key={c.key} aria-sort={sort.key === c.key ? (sort.desc ? "descending" : "ascending") : undefined}>
                  <button type="button" onClick={() => setSort((p) => ({ key: c.key, desc: p.key === c.key ? !p.desc : false }))}>
                    {c.label}
                    <span className="sort-mark" aria-hidden>
                      {sort.key === c.key ? (sort.desc ? "↓" : "↑") : ""}
                    </span>
                  </button>
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((i) => (
              <tr key={i.id}>
                {cols.map((c) => {
                  const text = c.text(i);
                  return c.key === "ID" ? (
                    <td key={c.key}>
                      <a href={itemUrl(i.id)} target="_blank" rel="noreferrer">
                        {i.id}
                      </a>
                    </td>
                  ) : (
                    <td key={c.key} title={text.length > 40 ? text : undefined}>
                      {text}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length === 0 && <p className="list-empty muted">{query ? "Nessuna non conformità corrisponde alla ricerca." : "La lista è ancora vuota."}</p>}
      </div>
    </div>
  );
}
