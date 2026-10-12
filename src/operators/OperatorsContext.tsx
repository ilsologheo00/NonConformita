import { createContext, useCallback, useContext, useEffect, useMemo, useState, type ReactNode } from "react";
import { useMsal } from "@azure/msal-react";
import { getToken } from "../auth/msal";
import { createOperator, deleteOperator, loadOperators, openOperatorsList, OperatorsListMissingError, renameOperator, type OperatorsList } from "../graph/operators";
import { GraphError } from "../graph/client";

/** A name shown in the form for an e-mail address that SharePoint stores. */
export interface Operator {
  name: string;
  email: string;
  id?: string; // SharePoint item id
}

type Status = { kind: "loading" } | { kind: "ready" } | { kind: "error"; message: string; missing?: boolean };

interface OperatorsContextValue {
  operators: Operator[];
  status: Status;
  /** Link to the SharePoint list, once it is open. */
  listUrl?: string;
  add: (op: Operator) => Promise<void>;
  remove: (email: string) => Promise<void>;
  reload: () => void;
  /** Name associated with an e-mail, or the e-mail itself when unknown. */
  nameFor: (email: string) => string;
  /** Entries saved in this browser by the earlier, local-only version. */
  legacy: Operator[];
  importLegacy: () => Promise<void>;
}

// The address book lived in this browser before it moved to SharePoint; it is
// still used, alone, when nobody is signed in (the ?anteprima preview).
const STORAGE_KEY = "nc-operators";

// Fields whose values are operator e-mails, recognised by display name.
const OPERATOR_FIELDS = [/^in gestione a$/, /^prossimo operatore$/];

export function isOperatorField(label: string): boolean {
  const l = label.trim().toLowerCase();
  return OPERATOR_FIELDS.some((re) => re.test(l));
}

const OperatorsContext = createContext<OperatorsContextValue | undefined>(undefined);

function loadLocal(): Operator[] {
  try {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? "[]");
    return Array.isArray(raw) ? raw.filter((o) => typeof o?.name === "string" && typeof o?.email === "string") : [];
  } catch {
    return [];
  }
}

function saveLocal(list: Operator[]) {
  try {
    if (list.length) localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
    else localStorage.removeItem(STORAGE_KEY);
  } catch {
    /* storage blocked */
  }
}

const sameEmail = (a: string, b: string) => a.trim().toLowerCase() === b.trim().toLowerCase();
const byName = (a: Operator, b: Operator) => a.name.localeCompare(b.name, "it");

function describe(e: unknown): Status {
  if (e instanceof OperatorsListMissingError) return { kind: "error", message: e.message, missing: true };
  if (e instanceof GraphError && e.status === 403) return { kind: "error", message: "Il tuo account non ha i permessi per modificare la lista degli operatori." };
  return { kind: "error", message: (e as Error)?.message || String(e) };
}

export function OperatorsProvider({ children }: { children: ReactNode }) {
  const { accounts } = useMsal();
  const account = accounts[0];
  const [operators, setOperators] = useState<Operator[]>(() => (account ? [] : loadLocal()));
  const [status, setStatus] = useState<Status>(account ? { kind: "loading" } : { kind: "ready" });
  const [list, setList] = useState<OperatorsList | null>(null);
  const [legacy, setLegacy] = useState<Operator[]>(() => (account ? loadLocal() : []));
  const [attempt, setAttempt] = useState(0);

  useEffect(() => {
    if (!account) return;
    let alive = true;
    setStatus({ kind: "loading" });
    (async () => {
      const token = await getToken(account);
      const l = await openOperatorsList(token);
      const rows = await loadOperators(token, l);
      if (!alive) return;
      setList(l);
      setOperators(rows.sort(byName));
      setStatus({ kind: "ready" });
    })().catch((e) => alive && setStatus(describe(e)));
    return () => {
      alive = false;
    };
  }, [account, attempt]);

  const add = useCallback(
    async (op: Operator) => {
      const entry = { name: op.name.trim(), email: op.email.trim() };
      const existing = operators.find((o) => sameEmail(o.email, entry.email));
      if (!account) {
        // Re-adding an e-mail renames it instead of duplicating it.
        const next = [...operators.filter((o) => o !== existing), entry].sort(byName);
        setOperators(next);
        saveLocal(next);
        return;
      }
      if (!list) throw new Error("La lista degli operatori non è disponibile.");
      const token = await getToken(account);
      if (existing?.id) {
        await renameOperator(token, list, existing.id, entry.name);
        setOperators((cur) => cur.map((o) => (o.id === existing.id ? { ...o, name: entry.name } : o)).sort(byName));
      } else {
        const id = await createOperator(token, list, entry.name, entry.email);
        setOperators((cur) => [...cur, { ...entry, id }].sort(byName));
      }
    },
    [account, list, operators]
  );

  const remove = useCallback(
    async (email: string) => {
      const target = operators.find((o) => sameEmail(o.email, email));
      if (!target) return;
      if (!account) {
        const next = operators.filter((o) => o !== target);
        setOperators(next);
        saveLocal(next);
        return;
      }
      if (!list || !target.id) throw new Error("La lista degli operatori non è disponibile.");
      await deleteOperator(await getToken(account), list, target.id);
      setOperators((cur) => cur.filter((o) => o.id !== target.id));
    },
    [account, list, operators]
  );

  const importLegacy = useCallback(async () => {
    for (const op of legacy) {
      if (!operators.some((o) => sameEmail(o.email, op.email))) await add(op);
    }
    saveLocal([]);
    setLegacy([]);
  }, [legacy, operators, add]);

  const nameFor = useCallback((email: string) => operators.find((o) => sameEmail(o.email, email))?.name ?? email, [operators]);
  const reload = useCallback(() => setAttempt((n) => n + 1), []);

  const value = useMemo(
    () => ({ operators, status, listUrl: list?.webUrl, add, remove, reload, nameFor, legacy, importLegacy }),
    [operators, status, list, add, remove, reload, nameFor, legacy, importLegacy]
  );
  return <OperatorsContext.Provider value={value}>{children}</OperatorsContext.Provider>;
}

export function useOperators(): OperatorsContextValue {
  const ctx = useContext(OperatorsContext);
  if (!ctx) throw new Error("useOperators must be used inside OperatorsProvider");
  return ctx;
}
