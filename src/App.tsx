import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { config, missingConfig } from "./config";
import { getToken, graphScopes, signInError } from "./auth/msal";
import { createItem, isEmpty, ListNotFoundError, loadSchema, type FieldValue, type ListSchema } from "./graph/sharepoint";
import { GraphError } from "./graph/client";
import { FieldInput } from "./components/FieldInput";
import { useTheme } from "./theme/ThemeContext";
import { NcList } from "./components/NcList";
import { ActionIcon, ActivityIcon, CheckIcon, FormIcon, ListIcon, MoonIcon, PeopleIcon, SignOutIcon, SunIcon } from "./components/icons";
import { OperatorsDialog } from "./operators/OperatorsDialog";
import { useOperators } from "./operators/OperatorsContext";

const isPreview = import.meta.env.DEV && new URLSearchParams(window.location.search).has("anteprima");

type View = "form" | "list" | "actions" | "activities";
type FormView = Exclude<View, "list">;

/** The SharePoint lists the app can add items to, one form each. */
interface FormDef {
  list: string;
  /** .env key that names the list, shown when it is not found. */
  envKey: keyof ImportMetaEnv;
  title: string;
  sent: string;
  /** Who the sign-in notice says the item is registered for. */
  noun: string;
}

const forms: Record<FormView, FormDef> = {
  form: { list: config.list, envKey: "VITE_SP_LIST", title: "Nuova non conformità", sent: "Non conformità inviata", noun: "la non conformità" },
  actions: { list: config.actionsList, envKey: "VITE_SP_ACTIONS_LIST", title: "Nuova azione correttiva / miglioramento", sent: "Azione inviata", noun: "l'azione" },
  activities: { list: config.activitiesList, envKey: "VITE_SP_ACTIVITIES_LIST", title: "Nuova attività proposta", sent: "Attività inviata", noun: "l'attività" },
};

export function App() {
  const [view, setView] = useState<View>("form");
  const missing = missingConfig();
  return (
    <Shell view={view} onView={setView}>
      {isPreview ? <Views view={view} account={null} /> : missing.length ? <SetupNeeded missing={missing} /> : <Authed view={view} />}
    </Shell>
  );
}

type Account = ReturnType<typeof useMsal>["accounts"][number];

/**
 * Views are mounted the first time they are opened and then stay mounted,
 * so a half-filled form survives a look at another page.
 */
function Views({ view, account }: { view: View; account: Account | null }) {
  const token = useCallback(() => (account ? getToken(account) : Promise.resolve("preview")), [account]);
  const [opened, setOpened] = useState<Set<View>>(() => new Set([view]));
  useEffect(() => {
    setOpened((o) => (o.has(view) ? o : new Set(o).add(view)));
  }, [view]);
  const listOpened = opened.has("list");

  return (
    <>
      {(Object.keys(forms) as FormView[]).map(
        (key) =>
          opened.has(key) && (
            <div key={key} hidden={view !== key}>
              <ListForm account={account} def={forms[key]} preview={key} />
            </div>
          )
      )}
      {listOpened && (
        <div hidden={view !== "list"}>
          <NcList getToken={token} active={view === "list"} explain={explain} />
        </div>
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
          <button type="button" className={`sidebar-link${view === "form" ? " active" : ""}`} aria-current={view === "form" ? "page" : undefined} onClick={() => onView("form")}>
            <FormIcon />
            Nuova non conformità
          </button>
          <button type="button" className={`sidebar-link${view === "list" ? " active" : ""}`} aria-current={view === "list" ? "page" : undefined} onClick={() => onView("list")}>
            <ListIcon />
            Elenco non conformità
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

function Authed({ view }: { view: View }) {
  const { instance, accounts, inProgress } = useMsal();
  const signedIn = useIsAuthenticated();

  if (inProgress !== InteractionStatus.None && !signedIn) return <div className="page"><p className="muted">Accesso in corso…</p></div>;

  if (!signedIn) {
    const def = forms[view === "list" ? "form" : view];
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

  return <Views view={view} account={accounts[0]} />;
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
 * The "Chiusa" switch and the closing-date column it fills in, recognised by
 * their display names since the columns are read from the list at runtime.
 */
function closingFields(schema: ListSchema) {
  const norm = (s: string) => s.trim().toLowerCase();
  const closed = schema.fields.find((f) => f.kind === "boolean" && /^chius[ao]\??$/.test(norm(f.label)));
  const date = schema.fields.find((f) => (f.kind === "date" || f.kind === "datetime") && /chiusura/.test(norm(f.label)));
  return closed && date ? { closed: closed.name, date: date.name, dateKind: date.kind as "date" | "datetime" } : null;
}

/**
 * "Tipo NC" and the two defect columns that depend on it: "Difetto Fornitore"
 * for an NC against a supplier, "Difetto Riscontrato" for every other type.
 */
const isTypeField = (f: ListSchema["fields"][number]) => /^tipo\s*(di\s*)?nc$/.test(f.label.trim().toLowerCase());

/** "Tipo NC" must always be chosen, even if the list column is not marked required. */
function withRequiredType(schema: ListSchema): ListSchema {
  return { ...schema, fields: schema.fields.map((f) => (isTypeField(f) ? { ...f, required: true } : f)) };
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
 * The closing date is shown (and required, if the column is) only once the NC is marked closed;
 * only the defect column matching "Tipo NC" is shown.
 */
function isHidden(schema: ListSchema, values: Record<string, FieldValue>, name: string): boolean {
  const closing = closingFields(schema);
  if (closing && name === closing.date && values[closing.closed] !== true) return true;
  const defect = defectFields(schema);
  return !!defect && name === hiddenDefect(defect, values);
}

function explain(e: unknown, def: Pick<FormDef, "list" | "envKey"> = forms.form): string {
  if (e instanceof ListNotFoundError) return `La lista “${e.list}” non esiste sul sito “${config.site}”.`;
  if (e instanceof GraphError) {
    if (e.status === 403) return "Il tuo account non ha i permessi per scrivere su questa lista SharePoint.";
    if (e.status === 404) return `Sito o lista non trovati: controlla VITE_SP_SITE (“${config.site}”) e ${def.envKey} (“${def.list}”).`;
    return `SharePoint ha risposto ${e.status}: ${e.message}`;
  }
  return (e as Error).message || String(e);
}

function ListForm({ account, def, preview }: { account: Account | null; def: FormDef; preview: FormView }) {
  const token = useCallback(() => (account ? getToken(account) : Promise.resolve("preview")), [account]);
  const [status, setStatus] = useState<Status>({ kind: "loading" });
  const [values, setValues] = useState<Record<string, FieldValue>>({});
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [sending, setSending] = useState(false);
  const [sendError, setSendError] = useState("");
  const [created, setCreated] = useState<{ id: string; webUrl: string } | null>(null);
  const formRef = useRef<HTMLFormElement>(null);

  useEffect(() => {
    let alive = true;
    const load = account ? token().then((t) => loadSchema(t, def.list)) : import("./preview").then((m) => m.previewSchemas[preview]);
    load
      .then(withRequiredType)
      .then((schema) => {
        if (!alive) return;
        setStatus({ kind: "ready", schema });
        setValues(initialValues(schema));
      })
      .catch((e) => alive && setStatus({ kind: "error", message: explain(e, def), lists: e instanceof ListNotFoundError ? e.available : undefined }));
    return () => {
      alive = false;
    };
  }, [token, account, def, preview]);

  const schema = status.kind === "ready" ? status.schema : null;
  const missingRequired = useMemo(() => (schema ? schema.fields.filter((f) => f.required && f.kind !== "boolean" && !isHidden(schema, values, f.name) && isEmpty(values[f.name])) : []), [schema, values]);

  if (status.kind === "loading") return <div className="page"><p className="muted">Lettura dei campi dalla lista SharePoint…</p></div>;
  if (status.kind === "error")
    return (
      <div className="page">
        <h1>Impossibile aprire il modulo</h1>
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
      // Changing "Tipo NC" clears the defect column that no longer applies, so it is not sent.
      if (defect && name === defect.type) {
        const off = hiddenDefect(defect, next);
        next[off] = initialValues(s)[off];
      }
      return next;
    });
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
      if (!account) {
        await new Promise((r) => setTimeout(r, 600));
        setCreated({ id: "0 (anteprima, nulla è stato inviato)", webUrl: "#" });
        return;
      }
      setCreated(await createItem(await token(), s, values));
    } catch (err) {
      setSendError(explain(err, def));
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
      <header className="form-head">
        <h1>{def.title}</h1>
        <p className="muted">
          Verrà creata nella lista{" "}
          <a href={s.listUrl} target="_blank" rel="noreferrer">
            {s.listName}
          </a>
          . I campi con <span className="req-mark">*</span> sono obbligatori.
        </p>
      </header>

      <div className="sheet">
        {s.fields.filter((f) => !isHidden(s, values, f.name)).map((f) => (
          <FieldInput key={f.name} field={f} value={values[f.name]} error={errors[f.name]} siteId={s.siteId} getToken={token} onChange={(v) => set(f.name, v)} />
        ))}
        {s.skipped.length > 0 && <p className="field-hint">Non compilabili da qui: {s.skipped.join(", ")}. Puoi completarli direttamente su SharePoint.</p>}
      </div>

      <div className="submit-bar">
        <div className="submit-status" aria-live="polite">
          {sendError ? (
            <span className="status-error">{sendError}</span>
          ) : missingRequired.length ? (
            <button type="button" className="status-missing" onClick={() => focusField(missingRequired[0].name)}>
              {missingRequired.length === 1 ? `Manca 1 campo obbligatorio: ${missingRequired[0].label}` : `Mancano ${missingRequired.length} campi obbligatori`}
            </button>
          ) : (
            <span className="status-ok">Pronta per l'invio</span>
          )}
        </div>
        <button type="submit" className="btn-primary" disabled={sending}>
          {sending ? "Invio in corso…" : "Invia"}
        </button>
      </div>
    </form>
  );
}
