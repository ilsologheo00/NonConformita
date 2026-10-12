import { useEffect, useRef, useState, type FormEvent } from "react";
import { config } from "../config";
import { useOperators } from "./OperatorsContext";
import { TrashIcon } from "../components/icons";

/** Address book that maps the names shown in the form to the e-mails sent to SharePoint. */
export function OperatorsDialog({ open, onClose }: { open: boolean; onClose: () => void }) {
  const ref = useRef<HTMLDialogElement>(null);
  const { operators, status, listUrl, add, remove, reload, legacy, importLegacy } = useOperators();
  const [name, setName] = useState("");
  const [email, setEmail] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const nameRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    const d = ref.current;
    if (!d) return;
    if (open && !d.open) d.showModal();
    if (!open && d.open) d.close();
  }, [open]);

  async function run(action: () => Promise<void>) {
    setBusy(true);
    setError("");
    try {
      await action();
    } catch (e) {
      setError((e as Error)?.message || String(e));
    } finally {
      setBusy(false);
    }
  }

  function submit(e: FormEvent) {
    e.preventDefault();
    if (!name.trim()) return setError("Inserisci il nome.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim())) return setError("La mail non è valida.");
    run(async () => {
      await add({ name, email });
      setName("");
      setEmail("");
      nameRef.current?.focus();
    });
  }

  const ready = status.kind === "ready";

  return (
    <dialog ref={ref} className="dialog" onClose={onClose} aria-labelledby="operators-title">
      <div className="dialog-head">
        <h2 id="operators-title">Operatori</h2>
        <button type="button" className="btn-secondary" onClick={onClose}>
          Chiudi
        </button>
      </div>
      <p className="muted dialog-intro">
        Nei campi “In gestione a” e “Prossimo operatore” compare il nome; all'invio su SharePoint viene scritta la mail associata.
        {listUrl && (
          <>
            {" "}
            L'elenco è condiviso nella lista{" "}
            <a href={listUrl} target="_blank" rel="noreferrer">
              {config.operatorsList}
            </a>
            .
          </>
        )}
      </p>

      {status.kind === "loading" && <p className="muted">Lettura degli operatori da SharePoint…</p>}

      {status.kind === "error" && (
        <div className="notice notice-error">
          <p>{status.message}</p>
          {status.missing && (
            <p>
              Va creata su SharePoint una lista “{config.operatorsList}” con la colonna Titolo per il nome e una colonna di testo “Email”. I passaggi sono nel README.
            </p>
          )}
          <button type="button" className="btn-secondary" onClick={reload}>
            Riprova
          </button>
        </div>
      )}

      {ready && (
        <form className="operator-add" onSubmit={submit} noValidate>
          <label>
            <span className="field-label">Nome</span>
            <input ref={nameRef} type="text" value={name} onChange={(e) => setName(e.target.value)} autoComplete="off" />
          </label>
          <label>
            <span className="field-label">Mail</span>
            <input type="text" inputMode="email" value={email} onChange={(e) => setEmail(e.target.value)} autoComplete="off" />
          </label>
          <button type="submit" className="btn-primary" disabled={busy}>
            Aggiungi
          </button>
        </form>
      )}
      {error && <p className="field-error">{error}</p>}

      {ready && legacy.length > 0 && (
        <div className="notice operator-legacy">
          <p>
            In questo browser ci sono {legacy.length === 1 ? "1 operatore salvato" : `${legacy.length} operatori salvati`} prima del passaggio a SharePoint.
          </p>
          <button type="button" className="btn-secondary" disabled={busy} onClick={() => run(importLegacy)}>
            Copiali nella lista
          </button>
        </div>
      )}

      {ready &&
        (operators.length ? (
          <table className="operator-table">
            <thead>
              <tr>
                <th>Nome</th>
                <th>Mail</th>
                <th aria-label="Azioni" />
              </tr>
            </thead>
            <tbody>
              {operators.map((o) => (
                <tr key={o.email}>
                  <td>{o.name}</td>
                  <td className="muted">{o.email}</td>
                  <td>
                    <button type="button" className="icon-danger" disabled={busy} onClick={() => run(() => remove(o.email))} aria-label={`Rimuovi ${o.name}`} title="Rimuovi">
                      <TrashIcon />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        ) : (
          <p className="muted operator-empty">Nessun operatore. Aggiungi il primo con il modulo qui sopra.</p>
        ))}
    </dialog>
  );
}
