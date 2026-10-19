import { useEffect, useMemo, useState } from "react";
import { loadItems, loadLinkedItems, loadSchema, type ListItem, type ListSchema } from "../graph/sharepoint";
import { ExternalIcon, RefreshIcon } from "./icons";

type Status = { kind: "loading" } | { kind: "error"; message: string } | { kind: "ready"; schema: ListSchema; items: ListItem[]; related: Related[] };
type Column = { key: string; label: string; text: (i: ListItem) => string; sort: (i: ListItem) => string | number; links?: (i: ListItem) => { id: string; label: string; open: () => void }[] };

/** Items of another list whose lookup column points at this list's items ("NC collegata"). */
type Related = { key: string; list: string; label: string; byItem: Map<string, ListItem[]> };

const dateTimeFmt = new Intl.DateTimeFormat("it-IT", { dateStyle: "short", timeStyle: "short" });

function columns(schema: ListSchema, related: Related[], openRelated: (list: string, id: string) => void): Column[] {
  return [
    { key: "ID", label: "ID", text: (i) => i.id, sort: (i) => Number(i.id) },
    ...schema.fields.map((f) => ({ key: f.name, label: f.label, text: (i: ListItem) => i.text[f.name], sort: (i: ListItem) => i.sort[f.name] })),
    ...related.map((r) => {
      const of = (i: ListItem) => r.byItem.get(i.id) ?? [];
      const label = (x: ListItem) => (x.title ? `${x.id} – ${x.title}` : x.id);
      return {
        key: r.key,
        label: r.label,
        text: (i: ListItem) => of(i).map(label).join(", "),
        sort: (i: ListItem) => of(i).length || "",
        links: (i: ListItem) => of(i).map((x) => ({ id: x.id, label: label(x), open: () => openRelated(r.list, x.id) })),
      };
    }),
    { key: "Created", label: "Creato", text: (i) => dateTimeFmt.format(new Date(i.created)), sort: (i) => i.created },
    { key: "Author", label: "Creato da", text: (i) => i.createdBy, sort: (i) => i.createdBy.toLowerCase() },
  ];
}

/** Another list to show as a column: its items linked to each row. */
export interface RelatedDef {
  list: string;
  label: string;
}

/** What one list page shows; preview is the sample used with ?anteprima. */
export interface ListDef {
  list: string;
  title: string;
  /** Plural, lower case: "non conformità", "azioni". */
  items: string;
  related?: RelatedDef[];
  preview: () => Promise<{ schema: ListSchema; items: ListItem[] }>;
}

/** A related list's items for each row; one that cannot be read only loses its column, not the whole page. */
async function loadRelated(token: string, schema: ListSchema, def: RelatedDef): Promise<Related | null> {
  try {
    const linked = await loadLinkedItems(token, schema, def.list);
    return linked && { key: `related:${linked.schema.listId}`, list: def.list, label: def.label, byItem: linked.byItem };
  } catch (e) {
    console.warn(`Colonna “${def.label}” non caricata`, e);
    return null;
  }
}

/**
 * Read-only view of a SharePoint list. SharePoint refuses to be shown inside
 * another site's frame, so the items are read through Graph and drawn here;
 * a click on a row opens the item in the app (onOpen), a linked item opens in its own list (onOpenRelated).
 */
export function ListView({
  def,
  getToken,
  active,
  explain,
  onOpen,
  onOpenRelated,
}: {
  def: ListDef;
  getToken: () => Promise<string>;
  active: boolean;
  explain: (e: unknown) => string;
  onOpen: (id: string) => void;
  onOpenRelated: (list: string, id: string) => void;
}) {
  const [status, setStatus] = useState<Status>({ kind: "loading" });
  const [reload, setReload] = useState(0);
  const [query, setQuery] = useState("");
  const [sort, setSort] = useState<{ key: string; desc: boolean }>({ key: "ID", desc: true });

  // Reloaded every time the view is opened, so a just-sent item shows up.
  useEffect(() => {
    if (!active) return;
    let alive = true;
    setStatus((s) => (s.kind === "ready" ? s : { kind: "loading" }));
    const load = import.meta.env.DEV && new URLSearchParams(window.location.search).has("anteprima")
      ? def.preview().then((p) => ({ ...p, related: [] as Related[] }))
      : getToken().then(async (token) => {
          const schema = await loadSchema(token, def.list);
          const [items, related] = await Promise.all([loadItems(token, schema), Promise.all((def.related ?? []).map((r) => loadRelated(token, schema, r)))]);
          return { schema, items, related: related.filter((r): r is Related => !!r) };
        });
    load
      .then(({ schema, items, related }) => alive && setStatus({ kind: "ready", schema, items, related }))
      .catch((e) => alive && setStatus({ kind: "error", message: explain(e) }));
    return () => {
      alive = false;
    };
  }, [def, active, reload, getToken, explain]);

  const cols = useMemo(() => (status.kind === "ready" ? columns(status.schema, status.related, onOpenRelated) : []), [status, onOpenRelated]);
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

  if (status.kind === "loading") return <div className="page"><p className="muted">Lettura delle {def.items} da SharePoint…</p></div>;
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

  return (
    <div className="page page-list">
      <header className="list-head">
        <div>
          <h1>{def.title}</h1>
          <p className="muted">
            {status.items.length === 1 ? "1 elemento" : `${status.items.length} elementi`} nella lista “{s.listName}”. Fai clic su una riga per aprire la scheda e modificarla.
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
              <tr key={i.id} className="list-row" onClick={(e) => !(e.target as HTMLElement).closest("a, button") && onOpen(i.id)}>
                {cols.map((c) => {
                  const text = c.text(i);
                  return c.key === "ID" ? (
                    <td key={c.key}>
                      <button type="button" className="link-button" onClick={() => onOpen(i.id)} aria-label={`Apri l'elemento ${i.id}`}>
                        {i.id}
                      </button>
                    </td>
                  ) : c.links ? (
                    <td key={c.key} title={text.length > 40 ? text : undefined}>
                      {c.links(i).map((l, n) => (
                        <span key={l.id}>
                          {n > 0 && ", "}
                          <button type="button" className="link-button" onClick={l.open}>
                            {l.label}
                          </button>
                        </span>
                      ))}
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
        {rows.length === 0 && <p className="list-empty muted">{query ? "Nessun elemento corrisponde alla ricerca." : "La lista è ancora vuota."}</p>}
      </div>
    </div>
  );
}
