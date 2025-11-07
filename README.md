# Plugin Outlook per Invio Dati Email via HTTP POST

## Descrizione

Questo plugin per Outlook consente di raccogliere le informazioni di un messaggio email (mittente, destinatari, oggetto, contenuto) e inviarle in formato JSON tramite una richiesta HTTP POST a un URL specificato.

## Caratteristiche

- **Pulsante personalizzato** nella ribbon di Outlook durante la lettura delle email
- **Raccolta dati completa**: mittente, destinatari (To, CC), oggetto, corpo del messaggio, allegati (metadati)
- **Invio tramite HTTP POST** in formato JSON
- **Compatibilità** con Outlook Desktop (Windows, Mac) e Outlook Web
- **Notifiche** per confermare l'invio o segnalare errori
- **Supporto CORS** configurabile tramite manifest

## Struttura dei File

```
outlook-email-exporter/
├── manifest.xml          # Configurazione del plugin
├── commands.html         # Pagina host per i comandi
├── commands.js           # Logica JavaScript principale
├── assets/
│   ├── icon-16.png      # Icona 16x16 pixel
│   ├── icon-32.png      # Icona 32x32 pixel
│   └── icon-80.png      # Icona 80x80 pixel
└── README.md            # Questa guida
```

## Installazione

### Prerequisiti

1. **Server HTTPS**: Tutti i file devono essere ospitati su un server con certificato SSL valido
2. **Endpoint API**: Un endpoint che accetti richieste POST con JSON
3. **Icone**: Tre icone PNG nelle dimensioni specificate

### Configurazione

#### 1. Modifica il Manifest (manifest.xml)

Apri `manifest.xml` e modifica le seguenti voci:

- **ID univoco**: Genera un nuovo GUID e sostituiscilo a `12345678-1234-1234-1234-123456789012`
- **Domini**:
  - Sostituisci `https://yourdomain.com` con il tuo dominio dove sono ospitati i file
  - Sostituisci `https://your-api-endpoint.com` con il dominio del tuo endpoint API
- **URL delle icone**: Aggiorna i path delle icone con i tuoi URL reali

#### 2. Configura l'Endpoint API (commands.js)

Apri `commands.js` e modifica:

```javascript
const apiUrl = "https://your-api-endpoint.com/api/email-data";
```

Sostituisci con l'URL del tuo endpoint API.

**Opzionale - Autenticazione**: Se il tuo endpoint richiede autenticazione, decommenta e configura:

```javascript
xhr.setRequestHeader("Authorization", "Bearer YOUR_TOKEN_HERE");
// oppure
xhr.setRequestHeader("X-API-Key", "YOUR_API_KEY_HERE");
```

#### 3. Carica i File sul Server

Carica tutti i file sul tuo server HTTPS mantenendo la struttura delle directory.

### Installazione in Outlook

#### Outlook Web / New Outlook for Windows

1. Apri Outlook Web (outlook.office.com)
2. Vai su **Impostazioni** (icona ingranaggio) > **Visualizza tutte le impostazioni di Outlook**
3. Seleziona **Posta** > **Personalizza azioni**
4. Scorri fino a "Add-in personalizzati"
5. Clicca **Aggiungi add-in personalizzato** > **Aggiungi da file**
6. Carica il file `manifest.xml`
7. Conferma l'installazione

#### Outlook Desktop Windows (Classic)

1. Vai su **File** > **Informazioni** > **Gestisci componenti aggiuntivi**
2. Si aprirà il browser con Outlook Web
3. Segui i passaggi per Outlook Web sopra descritti

#### Outlook Desktop Mac

1. Clicca sull'icona **...** nella barra degli strumenti
2. Seleziona **Ottieni componenti aggiuntivi**
3. Vai su **Add-in personali**
4. Clicca **Aggiungi add-in personalizzato** > **Aggiungi da file**
5. Carica il file `manifest.xml`

## Utilizzo

1. Apri un'email in Outlook
2. Nella ribbon (barra multifunzione), cerca il gruppo "Email Exporter"
3. Clicca sul pulsante **Invia Dati**
4. Il plugin raccoglierà automaticamente i dati e li invierà all'endpoint configurato
5. Riceverai una notifica di conferma o di errore

## Formato dei Dati Inviati

Il plugin invia un oggetto JSON con la seguente struttura:

```json
{
  "sender": {
    "displayName": "Nome Mittente",
    "emailAddress": "[email protected]"
  },
  "recipients": [
    {
      "displayName": "Nome Destinatario",
      "emailAddress": "[email protected]",
      "recipientType": "to"
    }
  ],
  "cc": [
    {
      "displayName": "Nome CC",
      "emailAddress": "[email protected]",
      "recipientType": "cc"
    }
  ],
  "subject": "Oggetto dell'email",
  "body": "Contenuto del corpo dell'email",
  "bodyType": "text",
  "dateReceived": "2025-11-07T10:30:00.000Z",
  "dateCreated": "2025-11-07T10:30:00.000Z",
  "itemId": "AAMkAGI...",
  "conversationId": "AAQkAGI...",
  "attachments": [
    {
      "id": "attachment_id",
      "name": "documento.pdf",
      "size": 12345,
      "attachmentType": "file",
      "isInline": false
    }
  ]
}
```

## Configurazione CORS sul Server API

Il tuo endpoint API deve essere configurato per accettare richieste CORS. Aggiungi questi header nelle risposte:

```
Access-Control-Allow-Origin: https://yourdomain.com
Access-Control-Allow-Methods: POST, OPTIONS
Access-Control-Allow-Headers: Content-Type, Authorization
Access-Control-Allow-Credentials: true
```

### Esempio Node.js/Express

```javascript
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://yourdomain.com');
  res.header('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.header('Access-Control-Allow-Credentials', 'true');
  
  if (req.method === 'OPTIONS') {
    return res.sendStatus(200);
  }
  next();
});
```

## Debug e Troubleshooting

### Il pulsante non appare

- Verifica che il manifest.xml sia valido (usa un validatore XML)
- Controlla che tutte le URL nel manifest siano HTTPS
- Verifica che i file siano accessibili dal browser
- Prova a disinstallare e reinstallare l'add-in

### Errore CORS

- Verifica che il dominio dell'API sia in `<AppDomains>` nel manifest
- Controlla che il server API restituisca gli header CORS corretti
- Usa strumenti di debug del browser (F12) per vedere gli errori di rete

### La richiesta POST non arriva

- Verifica che l'URL dell'endpoint sia corretto in `commands.js`
- Controlla i log della console del browser (F12 > Console)
- Verifica che il server API sia raggiungibile
- Prova l'endpoint con Postman o curl per verificare che funzioni

### Debug in Outlook Desktop Windows

1. Chiudi Outlook completamente
2. Apri Registry Editor (Win+R, digita `regedit`)
3. Vai a: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`
4. Crea un valore DWORD: `EnableLogging` = `1`
5. Riavvia Outlook
6. I log saranno in: `%TEMP%\wef\`

### Debug in Outlook Web

1. Apri Developer Tools (F12)
2. Vai su **Console** per vedere log ed errori
3. Vai su **Network** per vedere le richieste HTTP

## Sicurezza

### Best Practices

- ✅ Usa sempre HTTPS per tutti i file e endpoint
- ✅ Implementa autenticazione sul tuo endpoint API
- ✅ Valida e sanitizza tutti i dati ricevuti sul server
- ✅ Non includere credenziali o token nel codice JavaScript
- ✅ Limita le origini CORS solo ai domini necessari
- ✅ Implementa rate limiting sull'endpoint API
- ✅ Registra tutte le richieste per audit
- ❌ Non memorizzare dati sensibili nel localStorage
- ❌ Non esporre endpoint API pubblicamente senza autenticazione

## Personalizzazioni

### Aggiungere campi personalizzati

In `commands.js`, nella funzione `sendEmailData`, aggiungi campi all'oggetto `emailData`:

```javascript
emailData.customField = "valore personalizzato";
emailData.timestamp = new Date().toISOString();
emailData.userId = getCurrentUserId(); // tua funzione
```

### Modificare il comportamento del pulsante

Modifica la funzione `sendEmailData` in `commands.js` per cambiare cosa succede quando si clicca il pulsante.

### Aggiungere più pulsanti

Nel `manifest.xml`, duplica la sezione `<Control>` e aggiungi nuove funzioni in `commands.js`.

## Supporto e Contributi

Per problemi o domande:
- Consulta la documentazione ufficiale Microsoft: https://learn.microsoft.com/office/dev/add-ins/outlook/
- Verifica i log di debug
- Testa l'endpoint API indipendentemente con Postman

## Licenza

[Specifica qui la tua licenza]

## Autore

[Il tuo nome o azienda]

## Changelog

### Versione 1.0.0 (2025-11-07)
- Rilascio iniziale
- Supporto per invio dati email via HTTP POST
- Compatibilità con Outlook Desktop e Web