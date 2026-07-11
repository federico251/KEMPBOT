import os
import json
import base64
import re
import requests
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_KEY  = os.environ.get("ANTHROPIC_KEY")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
GOOGLE_CREDS   = os.environ.get("GOOGLE_CREDS")
TELEGRAM_API   = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

# Memoria conversazione per utente (in-memory, dura finché il server è up)
conversation_history = {}
MAX_HISTORY = 8  # ultimi 8 scambi

SYSTEM_PROMPT = """Sei l'assistente contabile di Kemp Studio SRL, studio italiano di video production.
Hai memoria della conversazione — se l'utente ha già dato informazioni prima, usale senza richiederle.

DATA OGGI: {oggi}

REGOLE IMPORTANTI:
- Inserisci SUBITO con i dati disponibili. Se manca solo il numero fattura, usa stringa vuota. Se manca la data usa oggi.
- Chiedi UNA SOLA cosa alla volta solo se manca qualcosa di essenziale (importo o tipo entrata/uscita).
- NON chiedere numero fattura se l'utente ha detto che non ce l'ha.
- NON chiedere metodo se non specificato: usa "Bonifico" per entrate, "Carta di credito" per uscite piccole.
- IVA default 22%. Collaboratori forfettari (Albi, Marco, Milo, Karim) = 0%. Pasti = 10%.
- Se l'utente dice "1500 + iva" → imponibile 1500, iva 22%.
- Se l'utente dice "pagato X€" → scorpora IVA al 22%.

CATEGORIE USCITE: Collaboratori, Consulenze Professionali, Attrezzatura, Software/Abbonamenti, Carburante, Trasferte, Ufficio/Utenze, Noleggi/Leasing, Assicurazioni, Pasti/Rappresentanza, Marketing, Altro
CATEGORIE ENTRATE: Video Production, Motion Design, Art Direction, Post-Production, Consulenza, Altro
METODI: Carta di credito, Bonifico, Addebito diretto, PayPal, Contanti

Rispondi SEMPRE e SOLO con JSON valido, nessun testo extra.

INSERISCI USCITA:
{"azione":"inserisci_uscita","data":"DD/MM/YYYY","descrizione":"...","fornitore":"...","categoria":"...","metodo":"...","imponibile":123.45,"iva_pct":0.22,"note":"..."}

INSERISCI ENTRATA (con IVA):
{"azione":"inserisci_entrata","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","n_fattura":"","categoria":"Video Production","metodo":"Bonifico","imponibile":1500.00,"iva_pct":0.22,"stato":"Da incassare","note":""}

INSERISCI RIMBORSO (entrata senza IVA, es. Berto, Giuse Barbieri, Repeople):
{"azione":"inserisci_rimborso","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","importo":123.45,"note":""}

MODIFICA:
{"azione":"modifica","foglio":"Uscite","cerca_descrizione":"...","cerca_cliente_o_fornitore":"...","campo":"imponibile","nuovo_valore":"1500"}

ELIMINA:
{"azione":"elimina","foglio":"Uscite","cerca_descrizione":"...","cerca_cliente_o_fornitore":"..."}

DOMANDA (solo se manca info essenziale — UNA sola domanda):
{"azione":"chiedi","testo":"Una sola domanda concisa"}

RISPOSTA TESTUALE:
{"azione":"risposta","testo":"..."}
"""

def get_sheets_service():
    import google.oauth2.service_account as sa
    from googleapiclient.discovery import build
    creds_dict = json.loads(GOOGLE_CREDS)
    creds = sa.Credentials.from_service_account_info(
        creds_dict, scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)

def append_row(sheet_name, values):
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:A"
    ).execute()
    rows = result.get("values", [])

    insert_row = len(rows) + 1
    for i, row in enumerate(rows):
        if row and "TOTALI" in str(row[0]).upper():
            insert_row = i + 1
            break

    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == sheet_name:
            sheet_id = s["properties"]["sheetId"]
            break
    if sheet_id is None:
        return

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"insertDimension": {"range": {
            "sheetId": sheet_id, "dimension": "ROWS",
            "startIndex": insert_row - 1, "endIndex": insert_row
        }, "inheritFromBefore": True}}]}
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A{insert_row}",
        valueInputOption="USER_ENTERED",
        body={"values": [values]}
    ).execute()

def get_sheet_data(sheet_name):
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:Z"
    ).execute()
    return result.get("values", [])

def find_row(sheet_name, cerca_desc=None, cerca_nome=None):
    data = get_sheet_data(sheet_name)
    for i, row in enumerate(data):
        if i == 0:
            continue
        desc = row[1].lower() if len(row) > 1 else ""
        nome = row[2].lower() if len(row) > 2 else ""
        match_desc = cerca_desc and cerca_desc.lower() in desc
        match_nome = cerca_nome and cerca_nome.lower() in nome
        if match_desc or match_nome:
            return i + 1, row
    return None, None

COL_MAP = {
    "Uscite": {"data":"A","descrizione":"B","fornitore":"C","categoria":"D","metodo":"E","imponibile":"F","iva_pct":"G","iva_eur":"H","totale":"I","note":"J"},
    "Entrate": {"data":"A","descrizione":"B","cliente":"C","n_fattura":"D","categoria":"E","metodo":"F","imponibile":"G","iva_pct":"H","iva_eur":"I","totale":"J","stato":"K","note":"L"},
    "Rimborsi Spese": {"data":"A","descrizione":"B","cliente":"C","importo":"D","note":"E"}
}

def update_cell(sheet_name, row_num, col_letter, value):
    service = get_sheets_service()
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!{col_letter}{row_num}",
        valueInputOption="USER_ENTERED",
        body={"values": [[value]]}
    ).execute()

def delete_row(sheet_name, row_num):
    service = get_sheets_service()
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    sheet_id = None
    for s in meta["sheets"]:
        if s["properties"]["title"] == sheet_name:
            sheet_id = s["properties"]["sheetId"]
            break
    if sheet_id is None:
        return False
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"deleteDimension": {"range": {
            "sheetId": sheet_id, "dimension": "ROWS",
            "startIndex": row_num - 1, "endIndex": row_num
        }}}]}
    ).execute()
    return True

def chiedi_claude(chat_id, testo=None, immagine_b64=None, mime_type="image/jpeg"):
    oggi = datetime.now().strftime("%d/%m/%Y")
    system = SYSTEM_PROMPT.replace("{oggi}", oggi)

    # Costruisci messaggio corrente
    if immagine_b64:
        if mime_type == "application/pdf":
            current_content = [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": immagine_b64}},
                {"type": "text", "text": testo or "Analizza questa fattura PDF e inseriscila nel foglio. Rispondi SOLO con JSON."}
            ]
        else:
            current_content = [
                {"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": immagine_b64}},
                {"type": "text", "text": testo or "Analizza questo scontrino e inseriscilo nel foglio. Rispondi SOLO con JSON."}
            ]
    else:
        current_content = testo

    # Recupera storia conversazione
    history = conversation_history.get(chat_id, [])
    messages = history + [{"role": "user", "content": current_content}]

    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        },
        json={"model": "claude-sonnet-4-6", "max_tokens": 800, "system": system, "messages": messages}
    )
    resp_json = resp.json()

    if "error" in resp_json:
        return None, {"azione": "risposta", "testo": f"Errore API: {resp_json['error'].get('message','')}"}

    if not resp_json.get("content"):
        return None, {"azione": "risposta", "testo": "Nessuna risposta. Riprova."}

    raw = resp_json["content"][0]["text"].strip()
    assistant_content = raw

    raw = raw.replace("```json","").replace("```","").strip()
    if not raw.startswith("{"):
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            raw = match.group(0)
        else:
            return assistant_content, {"azione": "risposta", "testo": raw[:500]}

    try:
        result = json.loads(raw)
        return assistant_content, result
    except:
        return assistant_content, {"azione": "risposta", "testo": "Errore parsing. Riprova."}

def update_history(chat_id, user_msg, assistant_msg):
    if chat_id not in conversation_history:
        conversation_history[chat_id] = []
    conversation_history[chat_id].append({"role": "user", "content": user_msg})
    conversation_history[chat_id].append({"role": "assistant", "content": assistant_msg})
    # Mantieni solo ultimi MAX_HISTORY scambi
    if len(conversation_history[chat_id]) > MAX_HISTORY * 2:
        conversation_history[chat_id] = conversation_history[chat_id][-(MAX_HISTORY * 2):]

def esegui_azione(risultato):
    azione = risultato.get("azione")

    if azione == "inserisci_uscita":
        d = risultato
        imp = float(d["imponibile"])
        iva = imp * float(d.get("iva_pct", 0.22))
        tot = imp + iva
        append_row("Uscite", [d["data"], d["descrizione"], d["fornitore"], d["categoria"], d["metodo"], imp, d.get("iva_pct",0.22), iva, tot, d.get("note","")])
        return f"✅ *Uscita inserita*\n📅 {d['data']} — {d['fornitore']}\n💸 €{imp:.2f} + IVA €{iva:.2f} = €{tot:.2f}\n📁 {d['categoria']}"

    elif azione == "inserisci_entrata":
        d = risultato
        imp = float(d["imponibile"])
        iva = imp * float(d.get("iva_pct", 0.22))
        tot = imp + iva
        append_row("Entrate", [d["data"], d["descrizione"], d["cliente"], d.get("n_fattura",""), d["categoria"], d["metodo"], imp, d.get("iva_pct",0.22), iva, tot, d.get("stato","Da incassare"), d.get("note","")])
        return f"✅ *Entrata inserita*\n📅 {d['data']} — {d['cliente']}\n💰 €{imp:.2f} + IVA €{iva:.2f} = €{tot:.2f}\n📁 {d['categoria']} — {d.get('stato','Da incassare')}"

    elif azione == "inserisci_rimborso":
        d = risultato
        imp = float(d["importo"])
        append_row("Rimborsi Spese", [d["data"], d["descrizione"], d["cliente"], imp, d.get("note","")])
        return f"✅ *Rimborso inserito*\n📅 {d['data']} — {d['cliente']}\n💵 €{imp:.2f}"

    elif azione == "modifica":
        d = risultato
        foglio = d.get("foglio","Uscite")
        row_num, _ = find_row(foglio, cerca_desc=d.get("cerca_descrizione"), cerca_nome=d.get("cerca_cliente_o_fornitore"))
        if row_num is None:
            return "⚠️ Riga non trovata. Specifica meglio descrizione o cliente/fornitore."
        campo = d.get("campo","").lower()
        col = COL_MAP.get(foglio,{}).get(campo)
        if not col:
            return f"⚠️ Campo '{campo}' non riconosciuto."
        update_cell(foglio, row_num, col, d.get("nuovo_valore"))
        return f"✏️ *Modifica effettuata*\n📋 {foglio} — {campo} → {d.get('nuovo_valore')}"

    elif azione == "elimina":
        d = risultato
        foglio = d.get("foglio","Uscite")
        row_num, row_data = find_row(foglio, cerca_desc=d.get("cerca_descrizione"), cerca_nome=d.get("cerca_cliente_o_fornitore"))
        if row_num is None:
            return "⚠️ Riga non trovata."
        delete_row(foglio, row_num)
        desc = row_data[1] if row_data and len(row_data) > 1 else "?"
        return f"🗑️ *Eliminata*: {desc}"

    elif azione in ("chiedi", "risposta"):
        return risultato.get("testo","")

    return "⚠️ Non ho capito. Riprova."

def send_message(chat_id, text):
    requests.post(f"{TELEGRAM_API}/sendMessage", json={"chat_id": chat_id, "text": text, "parse_mode": "Markdown"})

def get_file_url(file_id):
    r = requests.get(f"{TELEGRAM_API}/getFile", params={"file_id": file_id})
    path = r.json()["result"]["file_path"]
    return f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{path}"

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    message = data.get("message", {})
    chat_id = message.get("chat", {}).get("id")
    if not chat_id:
        return jsonify({"ok": True})

    try:
        immagine_b64 = None
        mime_type = "image/jpeg"
        user_text = None

        if "photo" in message or "document" in message:
            send_message(chat_id, "🔍 Analizzo...")
            if "photo" in message:
                file_id = message["photo"][-1]["file_id"]
            else:
                file_id = message["document"]["file_id"]
                mime_type = message["document"].get("mime_type", "image/jpeg")
            file_url = get_file_url(file_id)
            img_data = requests.get(file_url).content
            immagine_b64 = base64.b64encode(img_data).decode()
            user_text = message.get("caption", "Analizza e inserisci nel foglio.")

        elif "text" in message:
            user_text = message["text"]
            if user_text == "/start":
                conversation_history[chat_id] = []
                send_message(chat_id,
                    "👋 Ciao! Sono l'assistente contabile di *Kemp Studio*.\n\n"
                    "📸 Manda foto di scontrini\n"
                    "✍️ Scrivi *'pagato 80€ benzina carta'*\n"
                    "💰 O *'fattura Skeptical marzo 1500+iva bonifico'*\n"
                    "✏️ O *'modifica Berto giugno da 1000 a 1500'*\n\n"
                    "Ricordo il contesto della conversazione!"
                )
                return jsonify({"ok": True})
            if user_text == "/reset":
                conversation_history[chat_id] = []
                send_message(chat_id, "🔄 Memoria resettata.")
                return jsonify({"ok": True})
        else:
            send_message(chat_id, "Manda una foto o scrivi il movimento 👆")
            return jsonify({"ok": True})

        assistant_raw, risultato = chiedi_claude(chat_id, testo=user_text, immagine_b64=immagine_b64, mime_type=mime_type)
        risposta = esegui_azione(risultato)

        # Aggiorna storia solo per messaggi di testo (non immagini, troppo pesanti)
        if not immagine_b64 and assistant_raw:
            update_history(chat_id, user_text, assistant_raw)

        send_message(chat_id, risposta)

    except Exception as e:
        send_message(chat_id, f"❌ Errore: {str(e)}")

    return jsonify({"ok": True})

@app.route("/")
def health():
    return "Kemp Bot OK ✅"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
