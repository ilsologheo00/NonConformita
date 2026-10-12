import type { ListSchema } from "./graph/sharepoint";

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
    { name: "Tipo", label: "Tipologia", kind: "choice", required: true, choices: ["Interna", "Fornitore", "Cliente"] },
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
  ],
};
