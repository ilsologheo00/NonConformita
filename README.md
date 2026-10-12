# Non conformità

Webapp interna per inserire una nuova non conformità nella lista SharePoint
del controllo qualità. Replica il form di Microsoft Lists e al clic su
**Invia** crea l'elemento nella lista, a nome dell'utente che ha effettuato
l'accesso.

I campi **non sono scritti nel codice**: all'avvio l'app legge da SharePoint
(Microsoft Graph) le colonne della lista, con nome, tipo, obbligatorietà,
descrizione e valori delle scelte. Se in SharePoint si aggiunge una scelta o
si rinomina un campo, il form si aggiorna da solo.

Stack: React + TypeScript (Vite), MSAL per l'accesso Microsoft, nessun backend.

## 1. Registrare l'app in Microsoft Entra ID

1. portal.azure.com → **Microsoft Entra ID** → **Registrazioni app** → **Nuova registrazione**.
2. Nome: `Non conformità`. Account: *solo questa directory organizzativa*.
3. URI di reindirizzamento: piattaforma **Applicazione a pagina singola (SPA)**,
   valore `http://localhost:5174` (aggiungere poi l'URL di produzione).
4. **Autorizzazioni API** → Microsoft Graph → *Autorizzazioni delegate*:
   `Sites.ReadWrite.All`, `User.ReadBasic.All`. Se il tenant non permette il
   consenso degli utenti, un amministratore deve premere *Concedi consenso amministratore*.
5. Dalla pagina *Panoramica* copiare **ID applicazione (client)** e **ID directory (tenant)**.

## 2. Configurare

```powershell
Copy-Item .env.example .env
```

| Variabile | Valore |
|---|---|
| `VITE_AZURE_CLIENT_ID` | ID applicazione (client) |
| `VITE_AZURE_TENANT_ID` | ID directory (tenant) |
| `VITE_SP_SITE` | `semantica.sharepoint.com` oppure `semantica.sharepoint.com:/sites/NomeSito` |
| `VITE_SP_LIST` | nome visualizzato della lista (o il suo GUID) |
| `VITE_SP_FIELDS` | opzionale: nomi interni delle colonne, separati da virgola, per mostrare solo quelle e in quell'ordine |

Senza `VITE_SP_FIELDS` il form mostra tutte le colonne modificabili del tipo di
contenuto predefinito della lista, nell'ordine del tipo di contenuto. Se il
form di Lists nasconde alcune colonne, elencare qui solo quelle visibili.

### Lista degli operatori

La rubrica “Operatori” (nomi mostrati in *In gestione a* e *Prossimo operatore*,
mail scritte su SharePoint) è salvata in una lista dello stesso sito, condivisa
da tutti. Va creata una volta:

1. Sul sito di `VITE_SP_SITE`: **Nuovo** → **Elenco** → **Elenco vuoto**, nome `Operatori NC`.
2. **Aggiungi colonna** → **Riga di testo singola**, nome `Email` (esattamente così).
3. La colonna *Titolo* contiene il nome: si può rinominarla “Nome” dalle
   impostazioni della colonna, il nome interno resta `Title`.

Gli operatori si aggiungono dall'app (in basso a sinistra) o direttamente nella
lista. Chi usa l'app deve poter modificare la lista. Per un nome diverso,
impostare `VITE_SP_OPERATORS_LIST`.

## 3. Avviare

```powershell
npm install
npm run dev      # http://localhost:5174
```

`http://localhost:5174/?anteprima` (solo in sviluppo) mostra il form con campi
di esempio, senza accesso e senza inviare nulla: serve solo a controllare la grafica.

## 4. Docker

L'immagine compila l'app e la serve con nginx. La configurazione non è
inserita nell'immagine: all'avvio il container scrive le variabili `VITE_*`
in `/env.js`, quindi la stessa immagine va bene per qualsiasi lista.

```powershell
docker compose up -d --build      # http://localhost:8080, legge .env
```

Oppure senza compose:

```powershell
docker build -t non-conformita .
docker run -d -p 8080:80 --env-file .env non-conformita
```

In Entra ID, tra i redirect URI della registrazione (tipo *Single-page
application*), va aggiunto l'indirizzo con cui si apre l'app dal container, per
esempio `http://localhost:8080` o `https://nonconformita.azienda.it`. Per un
indirizzo diverso da `localhost` serve HTTPS, che di solito si ottiene con un
reverse proxy davanti al container.

## Limiti noti (Microsoft Graph)

- **Allegati**: Graph non permette di caricare allegati sugli elementi di lista.
- **Persone**: una persona può essere selezionata solo se ha già aperto il sito
  SharePoint almeno una volta. In caso contrario l'app mostra un errore che lo spiega.
- **Metadati gestiti, posizione, immagini**: non compilabili da qui. Il form li
  elenca in fondo, da completare su SharePoint.
- **Collegamenti ipertestuali**: Graph ha un supporto limitato in scrittura e
  potrebbero essere rifiutati.
