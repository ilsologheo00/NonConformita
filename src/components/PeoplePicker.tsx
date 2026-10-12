import { useEffect, useId, useRef, useState } from "react";
import { searchPeople, type Person } from "../graph/sharepoint";
import { CloseIcon } from "./icons";

interface Props {
  id: string;
  value: Person[];
  multiple: boolean;
  invalid: boolean;
  describedBy?: string;
  getToken: () => Promise<string>;
  onChange: (v: Person[]) => void;
}

export function PeoplePicker({ id, value, multiple, invalid, describedBy, getToken, onChange }: Props) {
  const [q, setQ] = useState("");
  const [results, setResults] = useState<Person[]>([]);
  const [open, setOpen] = useState(false);
  const [active, setActive] = useState(0);
  const [searchError, setSearchError] = useState("");
  const listId = useId();
  const seq = useRef(0);

  useEffect(() => {
    const term = q.trim();
    if (term.length < 2) {
      setResults([]);
      return;
    }
    const n = ++seq.current;
    const t = setTimeout(async () => {
      try {
        const r = await searchPeople(await getToken(), term);
        if (n !== seq.current) return;
        setResults(r.filter((p) => !value.some((v) => v.id === p.id)));
        setActive(0);
        setOpen(true);
        setSearchError("");
      } catch (e) {
        if (n === seq.current) setSearchError((e as Error).message);
      }
    }, 250);
    return () => clearTimeout(t);
  }, [q, getToken, value]);

  function pick(p: Person) {
    onChange(multiple ? [...value, p] : [p]);
    setQ("");
    setResults([]);
    setOpen(false);
  }

  const showInput = multiple || value.length === 0;

  return (
    <div className={`people${invalid ? " is-invalid" : ""}`}>
      {value.map((p) => (
        <span className="person-chip" key={p.id}>
          <span className="person-initials" aria-hidden>
            {p.displayName
              .split(" ")
              .map((w) => w[0])
              .slice(0, 2)
              .join("")}
          </span>
          {p.displayName}
          <button type="button" className="chip-remove" aria-label={`Rimuovi ${p.displayName}`} onClick={() => onChange(value.filter((v) => v.id !== p.id))}>
            <CloseIcon />
          </button>
        </span>
      ))}
      {showInput && (
        <div className="people-search">
          <input
            id={id}
            type="text"
            role="combobox"
            aria-expanded={open && results.length > 0}
            aria-controls={listId}
            aria-describedby={describedBy}
            aria-invalid={invalid || undefined}
            autoComplete="off"
            placeholder="Cerca per nome o email"
            value={q}
            onChange={(e) => setQ(e.target.value)}
            onBlur={() => setTimeout(() => setOpen(false), 120)}
            onFocus={() => results.length && setOpen(true)}
            onKeyDown={(e) => {
              if (!open || !results.length) return;
              if (e.key === "ArrowDown") {
                e.preventDefault();
                setActive((a) => Math.min(a + 1, results.length - 1));
              } else if (e.key === "ArrowUp") {
                e.preventDefault();
                setActive((a) => Math.max(a - 1, 0));
              } else if (e.key === "Enter") {
                e.preventDefault();
                pick(results[active]);
              } else if (e.key === "Escape") {
                setOpen(false);
              }
            }}
          />
          {open && results.length > 0 && (
            <ul className="people-results" id={listId} role="listbox">
              {results.map((p, i) => (
                <li
                  key={p.id}
                  role="option"
                  aria-selected={i === active}
                  className={i === active ? "active" : undefined}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    pick(p);
                  }}
                >
                  <strong>{p.displayName}</strong>
                  <span>{p.mail}</span>
                </li>
              ))}
            </ul>
          )}
          {open && q.trim().length >= 2 && results.length === 0 && !searchError && <p className="people-empty">Nessuna persona trovata con “{q.trim()}”.</p>}
          {searchError && <p className="people-empty">Ricerca non riuscita: {searchError}</p>}
        </div>
      )}
    </div>
  );
}
