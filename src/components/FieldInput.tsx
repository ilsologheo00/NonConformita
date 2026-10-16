import { useEffect, useState } from "react";
import { loadLookupOptions, type Field, type FieldValue, type LookupOption, type Person } from "../graph/sharepoint";
import { PeoplePicker } from "./PeoplePicker";
import { isOperatorField, useOperators } from "../operators/OperatorsContext";

interface Props {
  field: Field;
  value: FieldValue;
  error?: string;
  siteId: string;
  getToken: () => Promise<string>;
  /** Changes every time the form is shown again, so lookup values are re-read. */
  refreshKey?: number;
  onChange: (v: FieldValue) => void;
}

const OTHER = "\u0000altro";

export function FieldInput({ field: f, value, error, siteId, getToken, refreshKey, onChange }: Props) {
  const id = `f-${f.name}`;
  const hintId = f.description ? `${id}-hint` : undefined;
  const errId = error ? `${id}-err` : undefined;
  const describedBy = [hintId, errId].filter(Boolean).join(" ") || undefined;
  const common = { id, "aria-describedby": describedBy, "aria-invalid": error ? true : undefined, "aria-required": f.required || undefined };

  const isGroup = f.kind === "multichoice" || (f.kind === "choice" && f.radio) || f.kind === "boolean";

  return (
    <div className={`field${isWide(f) ? " field-wide" : ""}${f.required ? " is-required" : ""}${error ? " has-error" : ""}`}>
      {isGroup && f.kind !== "boolean" ? (
        <span className="field-label" id={`${id}-label`}>
          {f.label}
          {f.required && <span className="req-mark" aria-label="obbligatorio">*</span>}
        </span>
      ) : (
        <label className="field-label" htmlFor={id}>
          {f.label}
          {f.required && <span className="req-mark" aria-label="obbligatorio">*</span>}
        </label>
      )}
      {f.description && (
        <p className="field-hint" id={hintId}>
          {f.description}
        </p>
      )}

      <Control f={f} value={value} common={common} describedBy={describedBy} invalid={!!error} siteId={siteId} getToken={getToken} refreshKey={refreshKey} onChange={onChange} />

      {error && (
        <p className="field-error" id={errId}>
          {error}
        </p>
      )}
    </div>
  );
}

/** Fields that need a full row of the grid; the rest sit side by side. */
function isWide(f: Field): boolean {
  if (f.kind === "note" || f.kind === "multichoice") return true;
  if (f.kind === "choice" && f.radio) return (f.choices?.length ?? 0) > 4;
  if (f.kind === "person" || f.kind === "lookup") return !!f.multiple;
  return false;
}

function Control({
  f,
  value,
  common,
  describedBy,
  invalid,
  siteId,
  getToken,
  refreshKey,
  onChange,
}: {
  f: Field;
  value: FieldValue;
  common: Record<string, unknown>;
  describedBy?: string;
  invalid: boolean;
  siteId: string;
  getToken: () => Promise<string>;
  refreshKey?: number;
  onChange: (v: FieldValue) => void;
}) {
  const str = typeof value === "string" ? value : "";

  if ((f.kind === "text" || f.kind === "choice") && isOperatorField(f.label)) return <OperatorSelect f={f} value={str} common={common} onChange={onChange} />;

  switch (f.kind) {
    case "text":
    case "url":
      return <input {...common} type={f.kind === "url" ? "url" : "text"} maxLength={f.maxLength} value={str} onChange={(e) => onChange(e.target.value)} />;

    case "note":
      return <textarea {...common} rows={5} value={str} onChange={(e) => onChange(e.target.value)} />;

    case "number":
      return <input {...common} className="input-number" type="number" inputMode="decimal" step="any" min={f.min} max={f.max} value={str} onChange={(e) => onChange(e.target.value)} />;

    case "date":
      return <input {...common} className="input-date" type="date" value={str} onChange={(e) => onChange(e.target.value)} />;

    case "datetime":
      return <input {...common} className="input-date" type="datetime-local" value={str} onChange={(e) => onChange(e.target.value)} />;

    case "boolean":
      return (
        <label className="switch">
          <input {...common} type="checkbox" checked={value === true} onChange={(e) => onChange(e.target.checked)} />
          <span className="switch-track" aria-hidden />
          <span>{value === true ? "Sì" : "No"}</span>
        </label>
      );

    case "choice":
      return f.radio ? <RadioChoice f={f} value={str} describedBy={describedBy} onChange={onChange} /> : <SelectChoice f={f} value={str} common={common} onChange={onChange} />;

    case "multichoice": {
      const selected = Array.isArray(value) ? (value as string[]) : [];
      return (
        <div className="options" role="group" aria-labelledby={`f-${f.name}-label`} aria-describedby={describedBy}>
          {f.choices!.map((c) => (
            <label className="option" key={c}>
              <input type="checkbox" checked={selected.includes(c)} onChange={(e) => onChange(e.target.checked ? [...selected, c] : selected.filter((s) => s !== c))} />
              <span>{c}</span>
            </label>
          ))}
        </div>
      );
    }

    case "person":
      return (
        <PeoplePicker
          id={`f-${f.name}`}
          value={Array.isArray(value) ? (value as Person[]) : []}
          multiple={!!f.multiple}
          invalid={invalid}
          describedBy={describedBy}
          getToken={getToken}
          onChange={onChange}
        />
      );

    case "lookup":
      return <LookupControl f={f} value={value} common={common} siteId={siteId} getToken={getToken} refreshKey={refreshKey} onChange={onChange} />;
  }
}

/**
 * "In gestione a" / "Prossimo operatore": the value is the e-mail SharePoint
 * stores, the option text is the name from the Operatori address book.
 */
function OperatorSelect({ f, value, common, onChange }: { f: Field; value: string; common: Record<string, unknown>; onChange: (v: FieldValue) => void }) {
  const { operators, nameFor } = useOperators();
  // A choice column only accepts its own values; a text column takes any operator.
  const emails = f.kind === "choice" ? f.choices! : operators.map((o) => o.email);
  if (value && !emails.includes(value)) emails.push(value);
  const options = emails.map((email) => ({ email, name: nameFor(email) })).sort((a, b) => a.name.localeCompare(b.name, "it"));

  return (
    <>
      <select {...common} value={value} onChange={(e) => onChange(e.target.value)}>
        <option value="">Seleziona…</option>
        {options.map((o) => (
          <option key={o.email} value={o.email}>
            {o.name}
          </option>
        ))}
      </select>
      {options.length === 0 && <p className="field-hint operator-hint">Nessun operatore: aggiungili da “Operatori” in basso a sinistra.</p>}
    </>
  );
}

function SelectChoice({ f, value, common, onChange }: { f: Field; value: string; common: Record<string, unknown>; onChange: (v: FieldValue) => void }) {
  const known = value === "" || f.choices!.includes(value);
  const [custom, setCustom] = useState(!known);
  return (
    <>
      <select
        {...common}
        value={custom ? OTHER : value}
        onChange={(e) => {
          if (e.target.value === OTHER) {
            setCustom(true);
            onChange("");
          } else {
            setCustom(false);
            onChange(e.target.value);
          }
        }}
      >
        <option value="">Seleziona…</option>
        {f.choices!.map((c) => (
          <option key={c} value={c}>
            {c}
          </option>
        ))}
        {f.allowTextEntry && <option value={OTHER}>Altro valore…</option>}
      </select>
      {custom && <input className="input-other" type="text" aria-label={`${f.label}: altro valore`} placeholder="Scrivi il valore" value={value} onChange={(e) => onChange(e.target.value)} autoFocus />}
    </>
  );
}

function RadioChoice({ f, value, describedBy, onChange }: { f: Field; value: string; describedBy?: string; onChange: (v: FieldValue) => void }) {
  return (
    <div className="options" role="radiogroup" aria-labelledby={`f-${f.name}-label`} aria-describedby={describedBy}>
      {f.choices!.map((c, i) => (
        <label className="option" key={c}>
          <input id={i === 0 ? `f-${f.name}` : undefined} type="radio" name={f.name} checked={value === c} onChange={() => onChange(c)} />
          <span>{c}</span>
        </label>
      ))}
    </div>
  );
}

function LookupControl({
  f,
  value,
  common,
  siteId,
  getToken,
  refreshKey,
  onChange,
}: {
  f: Field;
  value: FieldValue;
  common: Record<string, unknown>;
  siteId: string;
  getToken: () => Promise<string>;
  refreshKey?: number;
  onChange: (v: FieldValue) => void;
}) {
  const [options, setOptions] = useState<LookupOption[] | null>(null);
  const [err, setErr] = useState("");

  // "NC collegata" and the like point at non-conformities: newest first, with their ID.
  const byId = /^nc\b/.test(f.label.trim().toLowerCase());

  useEffect(() => {
    let alive = true;
    getToken()
      .then((t) => loadLookupOptions(t, siteId, f.lookupListId!, f.lookupColumn!, byId))
      .then((o) => {
        if (!alive) return;
        setOptions(o);
        setErr("");
      })
      .catch((e) => alive && setErr((e as Error).message));
    return () => {
      alive = false;
    };
  }, [f.lookupListId, f.lookupColumn, byId, siteId, getToken, refreshKey]);

  if (err) return <p className="field-error">Valori non caricati: {err}</p>;

  if (f.multiple) {
    const selected = Array.isArray(value) ? (value as string[]) : [];
    return (
      <select {...common} multiple size={Math.min(8, options?.length || 3)} value={selected} onChange={(e) => onChange(Array.from(e.target.selectedOptions, (o) => o.value))} disabled={!options}>
        {options?.map((o) => (
          <option key={o.id} value={o.id}>
            {o.label}
          </option>
        ))}
      </select>
    );
  }

  return (
    <select {...common} value={typeof value === "string" ? value : ""} onChange={(e) => onChange(e.target.value)} disabled={!options}>
      <option value="">{options ? "Seleziona…" : "Caricamento valori…"}</option>
      {options?.map((o) => (
        <option key={o.id} value={o.id}>
          {o.label}
        </option>
      ))}
    </select>
  );
}
