import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { config, missingConfig } from "./config";
import { getToken, graphScopes, signInError } from "./auth/msal";
import {
  changedFields,
  createItem,
  isEmpty,
  ItemChangedError,
  ListNotFoundError,
  loadItem,
  loadLinkedItems,
  loadSchema,
  sameList,
  updateItem,
  type EditableItem,
  type FieldValue,
  type ListItem,
  type ListSchema,
} from "./graph/sharepoint";
import { GraphError } from "./graph/client";
import { FieldInput } from "./components/FieldInput";
import { useTheme } from "./theme/ThemeContext";
import { ListView, type ListDef, type RelatedDef } from "./components/ListView";
import { ActionIcon, ActivityIcon, CheckIcon, FormIcon, ListIcon, MoonIcon, PeopleIcon, SignOutIcon, SunIcon } from "./components/icons";
import { OperatorsDialog } from "./operators/OperatorsDialog";
import { isOperatorField, useOperators } from "./operators/OperatorsContext";

const isPreview = import.meta.env.DEV && new URLSearchParams(window.location.search).has("anteprima");

type ListPage = "list" | "actionsList" | "activitiesList";
type View = "form" | "actions" | "activities" | ListPage;
type FormView = Exclude<View, ListPage>;

/** The SharePoint lists the app can add items to, one form each. */
interface FormDef {
  list: string;
  /** .env key that names the list, shown when it is not found. */
  envKey: keyof ImportMetaEnv;
  title: string;
  sent: string;
  /** Who the sign-in notice says the item is registered for. */
  noun: string;
  /** Heading of an existing item open for editing, followed by its ID. */
  itemTitle: string;
}

const forms: Record<FormView, FormDef> = {
  form: { list: config.list, envKey: "VITE_SP_LIST", title: "Nuova non conformità", sent: "Non conformità inviata", noun: "la non conformità", itemTitle: "Non conformità" },
  actions: { list: config.actionsList, envKey: "VITE_SP_ACTIONS_LIST", title: "Nuova azione correttiva / miglioramento", sent: "Azione inviata", noun: "l'azione", itemTitle: "Azione" },
  activities: { list: config.activitiesList, envKey: "VITE_SP_ACTIVITIES_LIST", title: "Nuova attività proposta", sent: "Attività inviata", noun: "l'attività", itemTitle: "Attività proposta" },
};

const sample = (key: FormView) => () => import("./preview").then((m) => ({ schema: m.previewSchemas[key], items: key === "form" ? m.previewItems : [] }));

/** The read-only list pages, in sidebar order. */
const listPages: Record<ListPage, ListDef & { nav: string }> = {
  list: {
    list: config.list,
    title: "Elenco non conformità",
    nav: "Elenco non conformità",
    items: "non conformità",
    related: [{ list: config.actionsList, label: "Azioni correttive" }],
    preview: sample("form"),
  },
  actionsList: {
    list: config.actionsList,
    title: "Elenco azioni correttive / miglioramenti",
    nav: "Elenco azioni correttive",
    items: "azioni",
    related: [{ list: config.activitiesList, label: "Attività proposte" }],
    preview: sample("actions"),
  },
  activitiesList: { list: config.activitiesList, title: "Elenco attività proposte", nav: "Elenco attività proposte", items: "attività", preview: sample("activities") },
};

/** The form each list page edits its items with. */
const formOf: Record<ListPage, FormView> = { list: "form", actionsList: "actions", activitiesList: "activities" };
const isListPage = (v: View): v is ListPage => v in formOf;
const pageOfList = (list: string) => (Object.keys(listPages) as ListPage[]).find((k) => listPages[k].list === list);

/** The item open on each list page; none shows the list itself. */
type Details = Partial<Record<ListPage, string>>;
type OpenItem = (page: ListPage, id: string | null) => void;

export function App() {
  const [view, setView] = useState<View>("form");
  const [details, setDetails] = useState<Details>({});
  const openItem = useCallback<OpenItem>((page, id) => {
    setDetails((d) => ({ ...d, [page]: id ?? undefined }));
    setView(page);
    window.scrollTo(0, 0);
  }, []);
  const missing = missingConfig();
  const nav = { view, onView: setView, details, openItem };
  return (
    // A list in the sidebar always shows the list, even with one of its items open.
    <Shell view={view} onView={(v) => (isListPage(v) ? openItem(v, null) : setView(v))}>
      {isPreview ? <Views {...nav} account={null} /> : missing.length ? <SetupNeeded missing={missing} /> : <Authed {...nav} />}
    </Shell>
  );
}

type Nav = { view: View; onView: (v: View) => void; details: Details; openItem: OpenItem };

type Account = ReturnType<typeof useMsal>["accounts"][number];

/**
 * Views are mounted the first time they are opened and then stay mounted,
 * so a half-filled form survives a look at another page.
 */
function Views({ view, account, onView, details, openItem }: Nav & { account: Account | null }) {
  const token = useCallback(() => (account ? getToken(account) : Promise.resolve("preview")), [account]);
  const [opened, setOpened] = useState<Set<View>>(() => new Set([view]));
  // The item just sent in one form, handed to the form that follows it (NC → action → activity).
  const [links, setLinks] = useState<Partial<Record<FormView, Link>>>({});
  const onCreated = useCallback(
    (from: FormView, id: string, schema: ListSchema, values: Record<string, FieldValue>) => {
      const next = followUps[from];
      if (!next || !switchedOn(schema, values, next.flag)) return;
      setLinks((l) => ({ ...l, [next.view]: { id, listId: schema.listId, noun: next.linkedNoun, seq: (l[next.view]?.seq ?? 0) + 1 } }));
      onView(next.view);
      window.scrollTo(0, 0);
    },
    [onView]
  );
  useEffect(() => {
    setOpened((o) => (o.has(view) ? o : new Set(o).add(view)));
  }, [view]);

  return (
    <>
      {(Object.keys(forms) as FormView[]).map(
        (key) =>
          opened.has(key) && (
            <div key={key} hidden={view !== key}>
              <ListForm
                account={account}
                def={forms[key]}
                preview={key}
                active={view === key}
                onCreated={(id, schema, values) => onCreated(key, id, schema, values)}
                link={links[key] ?? null}
              />
            </div>
          )
      )}
      {(Object.keys(listPages) as ListPage[]).map(
        (key) =>
          opened.has(key) && (
            <div key={key} hidden={view !== key}>
              <div hidden={!!details[key]}>
                <ListView
                  def={listPages[key]}
                  getToken={token}
                  active={view === key && !details[key]}
                  explain={explain}
                  onOpen={(id) => openItem(key, id)}
                  onOpenRelated={(list, id) => openRelated(openItem, list, id)}
                />
              </div>
              {details[key] && (
                <ListForm
                  key={details[key]}
                  account={account}
                  def={forms[formOf[key]]}
                  preview={formOf[key]}
                  active={view === key}
                  edit={{ id: details[key], related: listPages[key].related ?? [], onBack: () => openItem(key, null), onOpenRelated: (list, id) => openRelated(openItem, list, id) }}
                />
              )}
            </div>
          )
      )}
    </>
  );
}

/* ---- Shell ----------------------------------------------------------------- */

function Shell({ view, onView, children }: { view: View; onView: (v: View) => void; children: React.ReactNode }) {
  const { theme, toggle } = useTheme();
  const { instance, accounts } = useMsal();
  const account = accounts[0];
  const { operators } = useOperators();
  const [operatorsOpen, setOperatorsOpen] = useState(false);

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="sidebar-brand">
          <span className="sidebar-brand-mark">NC</span>
          <span className="sidebar-brand-name">Qualità</span>
        </div>
        <nav className="sidebar-nav">
          {(Object.keys(listPages) as ListPage[]).map((key) => (
            <button key={key} type="button" className={`sidebar-link${view === key ? " active" : ""}`} aria-current={view === key ? "page" : undefined} onClick={() => onView(key)}>
              <ListIcon />
              {listPages[key].nav}
            </button>
          ))}
          <hr className="sidebar-sep" />
          <button type="button" className={`sidebar-link${view === "form" ? " active" : ""}`} aria-current={view === "form" ? "page" : undefined} onClick={() => onView("form")}>
            <FormIcon />
            Nuova non conformità
          </button>
          <button type="button" className={`sidebar-link${view === "actions" ? " active" : ""}`} aria-current={view === "actions" ? "page" : undefined} onClick={() => onView("actions")}>
            <ActionIcon />
            Nuova azione correttiva
          </button>
          <button type="button" className={`sidebar-link${view === "activities" ? " active" : ""}`} aria-current={view === "activities" ? "page" : undefined} onClick={() => onView("activities")}>
            <ActivityIcon />
            Nuova attività proposta
          </button>
        </nav>
        <div className="sidebar-footer">
          {account && (
            <p className="sidebar-user" title={account.username}>
              {account.name ?? account.username}
            </p>
          )}
          <button type="button" className="sidebar-link" onClick={() => setOperatorsOpen(true)} aria-haspopup="dialog">
            <PeopleIcon />
            Operatori
            {operators.length > 0 && <span className="sidebar-count">{operators.length}</span>}
          </button>
          <button type="button" className="sidebar-link" onClick={toggle} aria-pressed={theme === "dark"}>
            {theme === "light" ? <MoonIcon /> : <SunIcon />}
            {theme === "light" ? "Tema scuro" : "Tema chiaro"}
          </button>
          {account && (
            <button type="button" className="sidebar-link" onClick={() => instance.logoutRedirect({ account })}>
              <SignOutIcon />
              Esci
            </button>
          )}
        </div>
      </aside>
      <main className="app-content">{children}</main>
      <OperatorsDialog open={operatorsOpen} onClose={() => setOperatorsOpen(false)} />
    </div>
  );
}

/* ---- States before the form ------------------------------------------------ */

function SetupNeeded({ missing }: { missing: string[] }) {
  return (
    <div className="page">
      <h1>Configurazione mancante</h1>
      <div className="notice">
        <p>
          Per collegarsi a SharePoint l'app ha bisogno di questi valori nel file <code>.env</code>:
        </p>
        <ul>
          {missing.map((m) => (
            <li key={m}>
              <code>{m}</code>
            </li>
          ))}
        </ul>
        <p>Copia <code>.env.example</code> in <code>.env</code>, compila i valori e riavvia <code>npm run dev</code>. Il README spiega come registrare l'app in Entra ID.</p>
      </div>
    </div>
  );
}

function Authed(nav: Nav) {
  const { view } = nav;
  const { instance, accounts, inProgress } = useMsal();
  const signedIn = useIsAuthenticated();

  if (inProgress !== InteractionStatus.None && !signedIn) return <div className="page"><p className="muted">Accesso in corso…</p></div>;

  if (!signedIn) {
    const def = forms[isListPage(view) ? formOf[view] : view];
    return (
      <div className="page">
        <h1>{def.title}</h1>
        <div className="notice">
          <p>Accedi con il tuo account Semantica: {def.noun} verrà registrata a tuo nome nella lista SharePoint.</p>
          {signInError && (
            <p className="status-error">
              L'accesso precedente non è riuscito. Microsoft ha risposto: {signInError}
            </p>
          )}
          <button type="button" className="btn-primary" onClick={() => instance.loginRedirect({ scopes: graphScopes })}>
            Accedi con Microsoft
          </button>
        </div>
      </div>
    );
  }

  return <Views {...nav} account={accounts[0]} />;
}

/* ---- The form -------------------------------------------------------------- */

type Status = { kind: "loading" } | { kind: "error"; message: string; lists?: string[] } | { kind: "ready"; schema: ListSchema };

function initialValues(schema: ListSchema): Record<string, FieldValue> {
  const v: Record<string, FieldValue> = {};
  for (const f of schema.fields) {
    if (f.kind === "boolean") v[f.name] = f.defaultValue === "1";
    else if (f.kind === "multichoice") v[f.name] = f.defaultValue ? [f.defaultValue] : [];
    else if (f.kind === "person") v[f.name] = [];
    else if (f.kind === "lookup" && f.multiple) v[f.name] = [];
    else if (f.defaultValue && !f.defaultValue.startsWith("[")) v[f.name] = f.defaultValue; // skip formulas like [today]
    else if ((f.kind === "date" || f.kind === "datetime") && f.defaultValue?.toLowerCase() === "[today]") v[f.name] = today(f.kind);
    else v[f.name] = "";
  }
  // A new NC always starts open, whatever default the list column has.
  const closing = closingFields(schema);
  if (closing) {
    v[closing.closed] = false;
    v[closing.date] = "";
  }
  return v;
}

function today(kind: "date" | "datetime") {
  const d = new Date();
  const pad = (n: number) => String(n).padStart(2, "0");
  const date = `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  return kind === "date" ? date : `${date}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

/**
 * The "Chiusa" switch, the closing-date column it fills in and, where the list
 * has one, the "Verificatore" who checks the closing; recognised by their
 * display names since the columns are read from the list at runtime.
 */
function closingFields(schema: ListSchema) {
  const norm = (s: string) => s.trim().toLowerCase();
  const closed = schema.fields.find((f) => f.kind === "boolean" && /^chius[ao]\??$/.test(norm(f.label)));
  const date = schema.fields.find((f) => (f.kind === "date" || f.kind === "datetime") && /chiusura/.test(norm(f.label)));
  const verifier = schema.fields.find((f) => /^verificator[ei]$/.test(norm(f.label)));
  return closed && date ? { closed: closed.name, date: date.name, dateKind: date.kind as "date" | "datetime", verifier: verifier?.name } : null;
}

/** The verifier is asked right after "Chiusa", before the closing date. */
function withClosingOrder(schema: ListSchema): ListSchema {
  const closing = closingFields(schema);
  if (!closing?.verifier) return schema;
  const group = [closing.closed, closing.verifier, closing.date];
  const at = schema.fields.findIndex((f) => f.name === closing.closed);
  const rest = schema.fields.filter((f) => !group.includes(f.name));
  const moved = group.map((name) => schema.fields.find((f) => f.name === name)!);
  const before = schema.fields.slice(0, at).filter((f) => !group.includes(f.name)).length;
  return { ...schema, fields: [...rest.slice(0, before), ...moved, ...rest.slice(before)] };
}

/**
 * "Tipo NC" and the two defect columns that depend on it: "Difetto Fornitore"
 * for an NC against a supplier, "Difetto Riscontrato" for every other type.
 */
const isTypeField = (f: ListSchema["fields"][number]) => /^tipo\s*(di\s*)?nc$/.test(f.label.trim().toLowerCase());

/**
 * "Tipo NC", "In gestione a" and "Prossimo operatore" must always be filled in,
 * even if the list columns are not marked required.
 */
function withRequiredFields(schema: ListSchema): ListSchema {
  return { ...schema, fields: schema.fields.map((f) => (isTypeField(f) || isOperatorField(f.label) ? { ...f, required: true } : f)) };
}

function defectFields(schema: ListSchema) {
  const norm = (s: string) => s.trim().toLowerCase();
  const type = schema.fields.find(isTypeField);
  const supplier = schema.fields.find((f) => /^difetto\s+fornitore$/.test(norm(f.label)));
  const found = schema.fields.find((f) => /^difetto\s+riscontrato$/.test(norm(f.label)));
  return type && supplier && found ? { type: type.name, supplier: supplier.name, found: found.name } : null;
}

function isSupplierNc(v: FieldValue | undefined): boolean {
  const all = Array.isArray(v) ? v : [v];
  return all.some((x) => typeof x === "string" && /fornitor/i.test(x));
}

/** The defect column that does not apply to the chosen "Tipo NC". */
function hiddenDefect(defect: NonNullable<ReturnType<typeof defectFields>>, values: Record<string, FieldValue>) {
  return isSupplierNc(values[defect.type]) ? defect.found : defect.supplier;
}

/**
 * The closing date and the verifier are shown (and required, if the column is) only once the item is marked closed;
 * only the defect column matching "Tipo NC" is shown.
 */
function isHidden(schema: ListSchema, values: Record<string, FieldValue>, name: string): boolean {
  const closing = closingFields(schema);
  if (closing && (name === closing.date || name === closing.verifier) && values[closing.closed] !== true) return true;
  const defect = defectFields(schema);
  return !!defect && name === hiddenDefect(defect, values);
}

/**
 * sending: the list was already read when the form opened, so a 404 is about
 * something else (a person, a lookup value...) and SharePoint's own words say what.
 */
function explain(e: unknown, def: Pick<FormDef, "list" | "envKey"> = forms.form, sending = false): string {
  if (e instanceof ListNotFoundError) return `La lista “${e.list}” non esiste sul sito “${config.site}”.`;
  if (e instanceof GraphError) {
    if (e.status === 403) return "Il tuo account non ha i permessi per scrivere su questa lista SharePoint.";
    if (e.status === 404 && sending) return `Invio non riuscito, SharePoint non trova un elemento (${e.code}): ${e.message}`;
    if (e.status === 404) return `Sito o lista non trovati: controlla VITE_SP_SITE (“${config.site}”) e ${def.envKey} (“${def.list}”).`;
    return `SharePoint ha risposto ${e.status}: ${e.message}`;
  }
  return (e as Error).message || String(e);
}

/**
 * The form that opens after an item is sent, linked to it. flag matches the
 * label of the yes/no column that asks for it ("Crea azione correttiva");
 * without one it always opens.
 */
const followUps: Partial<Record<FormView, { view: FormView; flag?: RegExp; linkedNoun: string }>> = {
  form: { view: "actions", flag: /azion[ei]\s+correttiv/, linkedNoun: "alla non conformità" },
  actions: { view: "activities", linkedNoun: "all'azione" },
};

function switchedOn(schema: ListSchema, values: Record<string, FieldValue>, label?: RegExp): boolean {
  if (!label) return true;
  const flag = schema.fields.find((f) => f.kind === "boolean" && label.test(f.label.trim().toLowerCase()));
  return !!flag && values[flag.name] === true;
}

/** The lookup column ("NC collegata"...) pointing at the list the item came from. */
const linkField = (schema: ListSchema, link: Link) => schema.fields.find((f) => f.kind === "lookup" && sameList(f.lookupListId, link.listId));

/** An item to link a new one to; seq tells two requests for the same item apart. */
type Link = { id: string; listId: string; noun: string; seq: number };

/** A linked item ("NC collegata" → its actions) opens in its own list page. */
function openRelated(openItem: OpenItem, list: string, id: string) {
  const page = pageOfList(list);
  if (page) openItem(page, id);
}

/** An existing item open for editing, and the other lists whose items link to it. */
type Edit = { id: string; related: RelatedDef[]; onBack: () => void; onOpenRelated: (list: string, id: string) => void };
type Linked = { list: string; label: string; items: ListItem[] };

function ListForm({
  account,
  def,
  preview,
  active,
  onCreated,
  link,
  edit,
}: {
  account: Account | null;
  def: FormDef;
  preview: FormView;
  active: boolean;
  onCreated?: (id: string, schema: ListSchema, values: Record<string, FieldValue>) => void;
  link?: Link | null;
  edit?: Edit;
}) {
  const token = useCallback(() => (account ? getToken(account) : Promise.resolve("preview")), [account]);
  const [status, setStatus] = useState<Status>({ kind: "loading" });
  const [values, setValues] = useState<Record<string, FieldValue>>({});
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [sending, setSending] = useState(false);
  const [sendError, setSendError] = useState("");
  const [created, setCreated] = useState<{ id: string; webUrl: string } | null>(null);
  const formRef = useRef<HTMLFormElement>(null);
  // Bumped each time the form is shown, so lookups pick up items added meanwhile (e.g. a new NC).
  const [shown, setShown] = useState(0);
  const wasActive = useRef(active);
  useEffect(() => {
    if (active && !wasActive.current) setShown((n) => n + 1);
    wasActive.current = active;
  }, [active]);

  // Editing: the item as last read from SharePoint, and what is linked to it.
  const editId = edit?.id;
  const related = useRef(edit?.related ?? []);
  const [item, setItem] = useState<EditableItem | null>(null);
  const [saved, setSaved] = useState(false);
  const [conflict, setConflict] = useState(false);
  const [linked, setLinked] = useState<Linked[]>([]);
  const [reload, setReload] = useState(0);

  useEffect(() => {
    let alive = true;
    const load = account ? token().then((t) => loadSchema(t, def.list)) : import("./preview").then((m) => m.previewSchemas[preview]);
    load
      .then(withRequiredFields)
      .then(withClosingOrder)
      .then(async (schema) => {
        if (!editId) return { schema, item: null };
        if (!account) return { schema, item: { id: editId, eTag: "", webUrl: "#", values: initialValues(schema) } };
        return { schema, item: await loadItem(await token(), schema, editId) };
      })
      .then(({ schema, item }) => {
        if (!alive) return;
        setStatus({ kind: "ready", schema });
        setValues(item ? item.values : initialValues(schema));
        setItem(item);
        setErrors({});
        setConflict(false);
      })
      .catch((e) => alive && setStatus({ kind: "error", message: explain(e, def), lists: e instanceof ListNotFoundError ? e.available : undefined }));
    return () => {
      alive = false;
    };
  }, [token, account, def, preview, editId, reload]);

  // Re-read each time the item is shown, so an action just linked to this NC shows up.
  const listId = status.kind === "ready" ? status.schema.listId : "";
  useEffect(() => {
    if (!editId || !account || status.kind !== "ready") return;
    let alive = true;
    const schema = status.schema;
    token()
      .then((t) =>
        Promise.all(
          related.current.map((r) =>
            loadLinkedItems(t, schema, r.list)
              .then((l) => ({ list: r.list, label: r.label, items: l?.byItem.get(editId) ?? [] }))
              .catch(() => ({ list: r.list, label: r.label, items: [] }))
          )
        )
      )
      .then((l) => alive && setLinked(l));
    return () => {
      alive = false;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- the schema only matters through its list id
  }, [editId, account, token, listId, shown]);

  // A fresh item for the one just created in the previous form, with its link column already set.
  const [linkedTo, setLinkedTo] = useState("");
  const appliedLink = useRef(0);
  useEffect(() => {
    if (status.kind !== "ready" || !link || appliedLink.current === link.seq) return;
    appliedLink.current = link.seq;
    const field = linkField(status.schema, link);
    const fresh = initialValues(status.schema);
    if (field) fresh[field.name] = field.multiple ? [link.id] : link.id;
    setValues(fresh);
    setErrors({});
    setSendError("");
    setCreated(null);
    setLinkedTo(field ? `${link.noun} n. ${link.id}` : "");
  }, [status, link]);

  const schema = status.kind === "ready" ? status.schema : null;
  const missingRequired = useMemo(() => (schema ? schema.fields.filter((f) => f.required && f.kind !== "boolean" && !isHidden(schema, values, f.name) && isEmpty(values[f.name])) : []), [schema, values]);
  const dirty = useMemo(() => (schema && item ? changedFields(schema, item, values).size > 0 : false), [schema, item, values]);

  if (status.kind === "loading") return <div className="page"><p className="muted">{editId ? `Lettura dell'elemento ${editId} da SharePoint…` : "Lettura dei campi dalla lista SharePoint…"}</p></div>;
  if (status.kind === "error")
    return (
      <div className="page">
        {edit && (
          <button type="button" className="btn-secondary back-button" onClick={edit.onBack}>
            ← Torna all'elenco
          </button>
        )}
        <h1>{editId ? `Impossibile aprire l'elemento ${editId}` : "Impossibile aprire il modulo"}</h1>
        <div className="notice notice-error">
          <p>{status.message}</p>
          {status.lists && (
            <>
              <p>
                Liste presenti su “{config.site}”. Copia il nome giusto in {def.envKey} nel file .env e riavvia l'app:
              </p>
              <ul>
                {status.lists.map((name) => (
                  <li key={name}>{name}</li>
                ))}
              </ul>
            </>
          )}
        </div>
      </div>
    );

  const s = status.schema;

  const closing = closingFields(s);
  const defect = defectFields(s);

  function set(name: string, v: FieldValue) {
    setValues((prev) => {
      const next = { ...prev, [name]: v };
      // Switching "Chiusa" on stamps the closing date with this moment; off clears it.
      if (closing && name === closing.closed) next[closing.date] = v === true ? today(closing.dateKind) : "";
      // Reopening clears the verifier too, so a hidden name is not sent.
      if (closing?.verifier && name === closing.closed && v !== true) next[closing.verifier] = initialValues(s)[closing.verifier];
      // Changing "Tipo NC" clears the defect column that no longer applies, so it is not sent.
      if (defect && name === defect.type) {
        const off = hiddenDefect(defect, next);
        next[off] = initialValues(s)[off];
      }
      return next;
    });
    setSaved(false);
    if (errors[name]) setErrors(({ [name]: _, ...rest }) => rest);
    if (closing && name === closing.closed && v === true && errors[closing.date]) setErrors(({ [closing.date]: _, ...rest }) => rest);
  }

  function focusField(name: string) {
    const el = document.getElementById(`f-${name}`) ?? document.querySelector<HTMLElement>(`[name="${name}"]`);
    el?.scrollIntoView({ behavior: matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth", block: "center" });
    el?.focus({ preventScroll: true });
  }

  async function submit(e: FormEvent) {
    e.preventDefault();
    setSendError("");
    if (missingRequired.length) {
      setErrors(Object.fromEntries(missingRequired.map((f) => [f.name, "Campo obbligatorio."])));
      focusField(missingRequired[0].name);
      return;
    }
    setSending(true);
    try {
      if (item) {
        if (account) {
          const t = await token();
          await updateItem(t, s, item, values);
          const fresh = await loadItem(t, s, item.id);
          setItem(fresh);
          setValues(fresh.values);
        } else {
          setItem({ ...item, values });
        }
        setSaved(true);
        return;
      }
      if (!account) {
        await new Promise((r) => setTimeout(r, 600));
        setCreated({ id: "0 (anteprima, nulla è stato inviato)", webUrl: "#" });
        onCreated?.("0", s, values);
        return;
      }
      const added = await createItem(await token(), s, values);
      setCreated(added);
      setLinkedTo("");
      onCreated?.(added.id, s, values);
    } catch (err) {
      setConflict(err instanceof ItemChangedError);
      setSendError(err instanceof ItemChangedError ? `${err.message} Ricarica per vedere la versione attuale: le tue modifiche andranno perse.` : explain(err, def, true));
    } finally {
      setSending(false);
    }
  }

  if (created) {
    return (
      <div className="page">
        <div className="sent">
          <span className="sent-mark" aria-hidden>
            <CheckIcon />
          </span>
          <div>
            <h1>{def.sent}</h1>
            <p className="muted">
              Registrata in “{s.listName}” con ID {created.id}.
            </p>
            <div className="sent-actions">
              <button
                type="button"
                className="btn-primary"
                onClick={() => {
                  setCreated(null);
                  setValues(initialValues(s));
                  setErrors({});
                  window.scrollTo(0, 0);
                }}
              >
                Inserisci un'altra
              </button>
              <a className="btn-secondary" href={s.listUrl} target="_blank" rel="noreferrer">
                <ListIcon />
                Apri la lista
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <form className="page page-form" ref={formRef} onSubmit={submit} noValidate>
      {edit && item ? (
        <header className="form-head">
          <button type="button" className="btn-secondary back-button" onClick={edit.onBack}>
            ← Torna all'elenco
          </button>
          <h1>
            {def.itemTitle} n. {item.id}
          </h1>
          <p className="muted">
            Dalla lista{" "}
            <a href={s.listUrl} target="_blank" rel="noreferrer">
              {s.listName}
            </a>
            .{" "}
            <a href={item.webUrl} target="_blank" rel="noreferrer">
              Apri su SharePoint
            </a>
            . I campi con <span className="req-mark">*</span> sono obbligatori.
          </p>
          {linked.map((l) => (
            <p key={l.list} className="linked">
              <span className="linked-label">{l.label} collegate:</span>{" "}
              {l.items.length === 0 ? (
                <span className="muted">nessuna</span>
              ) : (
                l.items.map((x, n) => (
                  <span key={x.id}>
                    {n > 0 && ", "}
                    <button type="button" className="link-button" onClick={() => edit.onOpenRelated(l.list, x.id)}>
                      {x.title ? `${x.id} – ${x.title}` : x.id}
                    </button>
                  </span>
                ))
              )}
            </p>
          ))}
        </header>
      ) : (
      <header className="form-head">
        <h1>{def.title}</h1>
        <p className="muted">
          Verrà creata nella lista{" "}
          <a href={s.listUrl} target="_blank" rel="noreferrer">
            {s.listName}
          </a>
          . I campi con <span className="req-mark">*</span> sono obbligatori.
        </p>
        {linkedTo && <p className="notice">Collegata {linkedTo} appena inserita.</p>}
      </header>
      )}

      <div className="sheet">
        {s.fields.filter((f) => !isHidden(s, values, f.name)).map((f) => (
          <FieldInput key={f.name} field={f} value={values[f.name]} error={errors[f.name]} siteId={s.siteId} getToken={token} refreshKey={shown} onChange={(v) => set(f.name, v)} />
        ))}
        {s.skipped.length > 0 && <p className="field-hint">Non compilabili da qui: {s.skipped.join(", ")}. Puoi completarli direttamente su SharePoint.</p>}
      </div>

      <div className="submit-bar">
        <div className="submit-status" aria-live="polite">
          {sendError ? (
            <span className="status-error">
              {sendError}
              {conflict && (
                <>
                  {" "}
                  <button
                    type="button"
                    className="link-button"
                    onClick={() => {
                      setSendError("");
                      setSaved(false);
                      setReload((n) => n + 1);
                    }}
                  >
                    Ricarica
                  </button>
                </>
              )}
            </span>
          ) : missingRequired.length ? (
            <button type="button" className="status-missing" onClick={() => focusField(missingRequired[0].name)}>
              {missingRequired.length === 1 ? `Manca 1 campo obbligatorio: ${missingRequired[0].label}` : `Mancano ${missingRequired.length} campi obbligatori`}
            </button>
          ) : item ? (
            <span className={dirty ? "status-pending" : "status-ok"}>{dirty ? "Modifiche non salvate" : saved ? "Modifiche salvate" : "Nessuna modifica"}</span>
          ) : (
            <span className="status-ok">Pronta per l'invio</span>
          )}
        </div>
        {item ? (
          <button type="submit" className="btn-primary" disabled={sending || !dirty}>
            {sending ? "Salvataggio…" : "Salva modifiche"}
          </button>
        ) : (
          <button type="submit" className="btn-primary" disabled={sending}>
            {sending ? "Invio in corso…" : "Invia"}
          </button>
        )}
      </div>
    </form>
  );
}
