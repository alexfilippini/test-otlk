/**
 * Outlook Add-in per inviare dati email tramite HTTP POST
 * 
 * Questo script raccoglie le informazioni di un'email (mittente, destinatari, 
 * oggetto, corpo) e le invia in formato JSON a un endpoint API tramite POST.
 */

// Inizializzazione dell'add-in Office
Office.initialize = function() {
    console.log("Add-in inizializzato");
};

// Registra la funzione con Office.js
Office.actions.associate("sendEmailData", sendEmailData);

/**
 * Funzione principale che raccoglie i dati dell'email e li invia al server
 * @param {Object} event - Evento di Office.js
 */
function sendEmailData(event) {
    console.log("Funzione sendEmailData chiamata");
    
    // Ottieni l'item corrente (email)
    const item = Office.context.mailbox.item;
    
    // Verifica che sia un messaggio
    if (!item) {
        console.error("Nessun messaggio selezionato");
        if (event) {
            event.completed({ allowEvent: false });
        }
        return;
    }
    
    // Crea l'oggetto per contenere i dati dell'email
    const emailData = {
        sender: null,
        recipients: [],
        cc: [],
        bcc: [],
        subject: null,
        body: null,
        bodyType: null,
        dateReceived: null,
        dateCreated: null,
        itemId: null,
        conversationId: null,
        attachments: []
    };
    
    // Raccogli i dati sincroni (disponibili immediatamente)
    
    // Mittente
    if (item.from) {
        emailData.sender = {
            displayName: item.from.displayName,
            emailAddress: item.from.emailAddress
        };
    }
    
    // Destinatari (To)
    if (item.to && item.to.length > 0) {
        emailData.recipients = item.to.map(recipient => ({
            displayName: recipient.displayName,
            emailAddress: recipient.emailAddress,
            recipientType: recipient.recipientType
        }));
    }
    
    // Destinatari in copia (CC)
    if (item.cc && item.cc.length > 0) {
        emailData.cc = item.cc.map(recipient => ({
            displayName: recipient.displayName,
            emailAddress: recipient.emailAddress,
            recipientType: recipient.recipientType
        }));
    }
    
    // Oggetto dell'email
    emailData.subject = item.subject || "";
    
    // Date
    emailData.dateReceived = item.dateTimeCreated ? item.dateTimeCreated.toISOString() : null;
    emailData.dateCreated = item.dateTimeCreated ? item.dateTimeCreated.toISOString() : null;
    
    // ID del messaggio
    emailData.itemId = item.itemId;
    
    // ID della conversazione
    emailData.conversationId = item.conversationId;
    
    // Allegati (solo metadati, non il contenuto)
    if (item.attachments && item.attachments.length > 0) {
        emailData.attachments = item.attachments.map(attachment => ({
            id: attachment.id,
            name: attachment.name,
            size: attachment.size,
            attachmentType: attachment.attachmentType,
            isInline: attachment.isInline
        }));
    }
    
    // Raccogli il corpo del messaggio (operazione asincrona)
    item.body.getAsync(
        Office.CoercionType.Text, 
        function(asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                emailData.body = asyncResult.value;
                emailData.bodyType = "text";
                
                console.log("Dati email raccolti:", emailData);
                
                // Invia i dati al server
                sendDataToServer(emailData, event);
                
            } else {
                console.error("Errore nel recupero del corpo dell'email:", asyncResult.error);
                
                // Prova a ottenere il corpo in formato HTML come fallback
                item.body.getAsync(
                    Office.CoercionType.Html,
                    function(htmlResult) {
                        if (htmlResult.status === Office.AsyncResultStatus.Succeeded) {
                            emailData.body = htmlResult.value;
                            emailData.bodyType = "html";
                            
                            console.log("Dati email raccolti (HTML):", emailData);
                            sendDataToServer(emailData, event);
                            
                        } else {
                            console.error("Impossibile recuperare il corpo dell'email");
                            emailData.body = "[Impossibile recuperare il contenuto]";
                            emailData.bodyType = "error";
                            
                            sendDataToServer(emailData, event);
                        }
                    }
                );
            }
        }
    );
}

/**
 * Invia i dati raccolti al server tramite HTTP POST
 * @param {Object} emailData - Dati dell'email da inviare
 * @param {Object} event - Evento di Office.js per completare l'operazione
 */
function sendDataToServer(emailData, event) {
    console.log("Invio dati al server...");
    
    // *** IMPORTANTE: Sostituisci questo URL con il tuo endpoint API ***
    const apiUrl = "https://webhook.site/fa0fcbe2-e92c-4b56-8b13-24db45d74159";
    
    // Converti i dati in JSON
    const jsonData = JSON.stringify(emailData);
    
    console.log("JSON da inviare:", jsonData);
    
    // Usa XMLHttpRequest per massima compatibilità con Outlook Desktop
    const xhr = new XMLHttpRequest();
    
    // Gestisci la risposta del server
    xhr.addEventListener("readystatechange", function() {
        if (this.readyState === 4) {
            
            if (this.status >= 200 && this.status < 300) {
                // Successo
                console.log("Dati inviati con successo al server");
                console.log("Risposta del server:", this.responseText);
                
                showNotification(
                    "Successo", 
                    "I dati dell'email sono stati inviati correttamente al server"
                );
                
            } else {
                // Errore
                console.error("Errore nell'invio dei dati. Status:", this.status);
                console.error("Risposta del server:", this.responseText);
                
                showNotification(
                    "Errore", 
                    "Impossibile inviare i dati al server. Codice errore: " + this.status
                );
            }
            
            // Completa l'evento
            if (event) {
                event.completed({ allowEvent: true });
            }
        }
    });
    
    // Gestisci errori di rete
    xhr.addEventListener("error", function() {
        console.error("Errore di rete durante l'invio dei dati");
        
        showNotification(
            "Errore di Rete", 
            "Impossibile contattare il server. Verifica la connessione."
        );
        
        if (event) {
            event.completed({ allowEvent: false });
        }
    });
    
    // Configura e invia la richiesta
    try {
        xhr.open("POST", apiUrl);
        
        // Imposta gli header necessari
        xhr.setRequestHeader("Content-Type", "application/json");
        
        // *** OPZIONALE: Aggiungi header di autenticazione se necessario ***
        // xhr.setRequestHeader("Authorization", "Bearer YOUR_TOKEN_HERE");
        // xhr.setRequestHeader("X-API-Key", "YOUR_API_KEY_HERE");
        
        // Invia i dati
        xhr.send(jsonData);
        
    } catch (error) {
        console.error("Eccezione durante l'invio:", error);
        
        showNotification(
            "Errore", 
            "Si è verificato un errore durante l'invio dei dati"
        );
        
        if (event) {
            event.completed({ allowEvent: false });
        }
    }
}

/**
 * Mostra una notifica all'utente nell'interfaccia di Outlook
 * @param {string} title - Titolo della notifica
 * @param {string} message - Messaggio della notifica
 */
function showNotification(title, message) {
    console.log("Notifica:", title, "-", message);
    
    try {
        // Verifica che l'API di notifica sia disponibile
        if (Office.context.mailbox.item.notificationMessages) {
            
            Office.context.mailbox.item.notificationMessages.replaceAsync(
                "emailExporter",
                {
                    type: "informationalMessage",
                    message: message,
                    icon: "Icon.80x80",
                    persistent: false
                },
                function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                        console.error("Errore nella visualizzazione della notifica:", asyncResult.error);
                    }
                }
            );
            
        } else {
            // Fallback: mostra alert se le notifiche non sono disponibili
            console.warn("API notifiche non disponibile, uso alert");
            alert(title + ": " + message);
        }
        
    } catch (error) {
        console.error("Errore nella funzione showNotification:", error);
        // Ultimo fallback
        alert(title + ": " + message);
    }
}

/**
 * Funzione di utility per generare un timestamp
 * @returns {string} Timestamp in formato ISO
 */
function getCurrentTimestamp() {
    return new Date().toISOString();
}

/**
 * Funzione di utility per validare l'email
 * @param {string} email - Indirizzo email da validare
 * @returns {boolean} True se l'email è valida
 */
function isValidEmail(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}