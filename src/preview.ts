import type { ListItem, ListSchema } from "./graph/sharepoint";

// Dev-only sample (open http://localhost:5174/?anteprima) to check layout and
// themes without signing in. These are NOT the real SharePoint columns: the
// real form always reads its fields from the list.
export const previewSchema: ListSchema = {
  siteId: "preview",
  listId: "preview",
  listName: "Anteprima (campi di esempio)",
  listUrl: "#",
  skipped: [],
  fields: [
    { name: "Title", label: "Titolo", kind: "text", required: true, maxLength: 255 },
    { name: "Data", label: "Data rilevazione", kind: "date", required: true, defaultValue: "[today]" },
    { name: "TipoNC", label: "Tipo NC", kind: "choice", required: true, choices: ["NC interna", "NC a fornitore", "NC da cliente"] },
    { name: "DifettoFornitore", label: "Difetto Fornitore", kind: "choice", required: false, choices: ["Dimensionale", "Estetico", "Documentale"] },
    { name: "DifettoRiscontrato", label: "Difetto Riscontrato", kind: "choice", required: false, choices: ["Montaggio", "Lavorazione", "Collaudo"] },
    { name: "Gravita", label: "Gravità", kind: "choice", radio: true, required: false, choices: ["Bassa", "Media", "Alta"] },
    { name: "Reparti", label: "Reparti coinvolti", kind: "multichoice", required: false, choices: ["Produzione", "Magazzino", "Ufficio tecnico", "Acquisti"] },
    { name: "Descrizione", label: "Descrizione", kind: "note", required: true, description: "Cosa è stato rilevato, dove e su quale lotto." },
    { name: "Quantita", label: "Quantità non conforme", kind: "number", required: false },
    { name: "Bloccato", label: "Materiale bloccato", kind: "boolean", required: false },
    { name: "Lotto", label: "Lotto", kind: "text", required: false },
    { name: "Fornitore", label: "Fornitore", kind: "text", required: false },
    { name: "InGestione", label: "In gestione a", kind: "text", required: false },
    { name: "ProssimoOperatore", label: "Prossimo operatore", kind: "choice", required: false, choices: ["mario.rossi@example.com", "laura.bianchi@example.com"] },
    { name: "Chiusa", label: "Chiusa", kind: "boolean", required: false, defaultValue: "1" },
    { name: "DataChiusura", label: "Data chiusura", kind: "datetime", required: false },
    { name: "CreaAzione", label: "Crea azione correttiva", kind: "boolean", required: false },
  ],
};

const sampleList = (listName: string, fields: ListSchema["fields"]): ListSchema => ({ siteId: "preview", listId: "preview", listName, listUrl: "#", skipped: [], fields });

/** Sample fields per form, keyed like the app's views. */
export const previewSchemas: Record<"form" | "actions" | "activities", ListSchema> = {
  form: previewSchema,
  actions: sampleList("Anteprima azioni (campi di esempio)", [
    { name: "Title", label: "Titolo", kind: "text", required: true, maxLength: 255 },
    { name: "Tipo", label: "Tipo", kind: "choice", radio: true, required: true, choices: ["Azione correttiva", "Miglioramento"] },
    { name: "Descrizione", label: "Descrizione", kind: "note", required: true },
    { name: "Responsabile", label: "Responsabile", kind: "text", required: false },
    { name: "Scadenza", label: "Scadenza", kind: "date", required: false },
    { name: "Chiusa", label: "Chiusa", kind: "boolean", required: false },
    { name: "DataChiusura", label: "Data chiusura", kind: "date", required: false },
  ]),
  activities: sampleList("Anteprima attività (campi di esempio)", [
    { name: "Title", label: "Titolo", kind: "text", required: true, maxLength: 255 },
    { name: "Descrizione", label: "Descrizione", kind: "note", required: true },
    { name: "Priorita", label: "Priorità", kind: "choice", required: false, choices: ["Bassa", "Media", "Alta"] },
    { name: "DataProposta", label: "Data proposta", kind: "date", required: false, defaultValue: "[today]" },
  ]),
};

function sample(id: number, created: string, createdBy: string, values: Record<string, string>): ListItem {
  const text: Record<string, string> = {};
  const sort: Record<string, string | number> = {};
  for (const f of previewSchema.fields) {
    text[f.name] = values[f.name] ?? "";
    sort[f.name] = f.kind === "date" && values[f.name] ? values[f.name].split("/").reverse().join("-") : (values[f.name] ?? "").toLowerCase();
  }
  return { id: String(id), created, createdBy, text, sort, lookupIds: {}, title: values.Title ?? "" };
}

export const previewItems: ListItem[] = [
  sample(1, "2026-09-02T08:14:00Z", "Mario Rossi", { Title: "Graffi su carter", Data: "02/09/2026", TipoNC: "NC a fornitore", DifettoFornitore: "Estetico", Gravita: "Media", Descrizione: "Carter lotto 2231 con graffi profondi sul lato visibile.", Quantita: "12", Bloccato: "Sì", Lotto: "2231", Fornitore: "Metalli Srl", Chiusa: "Sì", DataChiusura: "10/09/2026 16:30" }),
  sample(2, "2026-09-18T13:40:00Z", "Laura Bianchi", { Title: "Coppia di serraggio errata", Data: "18/09/2026", TipoNC: "NC interna", DifettoRiscontrato: "Montaggio", Gravita: "Alta", Reparti: "Produzione", Descrizione: "Viti M6 serrate a 6 Nm invece di 9 Nm.", Quantita: "40", Bloccato: "Sì", Chiusa: "No" }),
  sample(3, "2026-10-07T09:05:00Z", "Mario Rossi", { Title: "Etichetta mancante", Data: "07/10/2026", TipoNC: "NC da cliente", DifettoRiscontrato: "Collaudo", Gravita: "Bassa", Reparti: "Magazzino", Descrizione: "Il cliente segnala 3 colli senza etichetta di lotto.", Quantita: "3", Bloccato: "No", Chiusa: "No" }),
];
