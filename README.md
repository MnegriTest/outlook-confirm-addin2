# Outlook Send Confirmation Add-in

Questo add-in di Outlook richiede una conferma all'utente prima di inviare un'email a un indirizzo specifico (es. `github@mario.it`). È implementato come un add-in basato su "Launch Events" che intercetta l'evento `OnMessageSend` e presenta un dialog di conferma modale.

---

## Funzionalità

* Intercetta l'invio di email.
* Controlla se uno dei destinatari corrisponde a un indirizzo predefinito (`email@dominio.it`).
* Se l'indirizzo corrisponde, mostra un popup di conferma all'utente.
* Blocca l'invio dell'email fino a quando l'utente non conferma o annulla.

---

## Prerequisiti

Per sviluppare e testare questo add-in localmente, avrai bisogno di:

* Node.js (LTS consigliata) e npm
* Yeoman e il generatore per Office Add-ins (`npm install -g yo generator-office`)
* Un account Microsoft 365 con accesso a Outlook (Outlook sul web o client desktop)

---

## Installazione e Avvio (Sviluppo Locale)

1.  **Clona questo repository** (o crea il progetto con `yo office` e poi copia i file).
    ```bash
    git clone [https://github.com/tuo-utente/tuo-repo.git](https://github.com/tuo-utente/tuo-repo.git)
    cd tuo-repo
    ```
2.  **Installa le dipendenze:**
    ```bash
    npm install
    ```
3.  **Avvia il server di sviluppo e carica l'add-in in Outlook:**
    ```bash
    npm start
    ```
    Questo comando avvierà un server HTTPS locale e aprirà Outlook sul web o il client desktop (a seconda della tua configurazione), caricando automaticamente l'add-in.

---

## Configurazione e Deployment (Produzione)

1.  **Aggiorna il `manifest.xml`:**
    * Genera un **nuovo GUID** unico e sostituisci `YOUR_ADDIN_GUID` nel `manifest.xml`.
    * Aggiorna tutti i placeholder `https://YOUR_GITHUB_PAGES_URL` con l'URL effettivo dove ospiterai i file del tuo add-in (es. il tuo URL di GitHub Pages). Assicurati che sia **HTTPS**.
    * **Modifica l'indirizzo email da monitorare** in `src/commands/commands.js` alla riga `const targetRecipient = "github@mario.it";`.

2.  **Build del progetto:**
    ```bash
    npm run build
    ```
    Questo creerà i file ottimizzati per la produzione nella cartella `dist`.

3.  **Ospita i file:**
    Carica il contenuto della cartella `dist` su un server web accessibile pubblicamente (es. GitHub Pages, Azure App Service, un tuo server web). L'URL di questo server dovrà corrispondere a quello configurato nel `manifest.xml`.

4.  **Distribuisci tramite Microsoft 365 Admin Center:**
    * Accedi al [Microsoft 365 Admin Center](https://admin.microsoft.com/).
    * Vai su **Mostra tutto** > **Impostazioni** > **Add-in integrati**.
    * Clicca su **Distribuisci add-in**.
    * Scegli **Carica manifesto personalizzato** e fornisci l'**URL pubblico del tuo `manifest.xml`**.
    * Segui le istruzioni per assegnare l'add-in agli utenti o gruppi desiderati.

---

## Struttura del Progetto
