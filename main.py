import os
import json
import base64
import requests
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_KEY    = os.environ.get("ANTHROPIC_KEY")
SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID")
GOOGLE_CREDS     = os.environ.get("GOOGLE_CREDS")
TELEGRAM_API     = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

SYSTEM_PROMPT = """Sei l'assistente contabile di Kemp Studio SRL, studio italiano di video production.
Analizzi messaggi di testo o immagini di scontrini/fatture e restituisci SOLO JSON valido, nessun testo aggiuntivo.

DATA DEFAULT: {oggi}

REGOLE:
- Se l'importo sembra IVA inclusa, scorporala al 22% (default). Se dici "imponibile" usalo diretto.
- Collaboratori forfettari (Albi, Marco, Milo, Karim) = IVA 0%
- Pasti/ristoranti = IVA 10%
- Carburante = IVA 22%

CATEGORIE USCITE: Collaboratori, Consulenze Professionali, Attrezzatura, Software/Abbonamenti, Carburante, Trasferte, Ufficio/Utenze, Noleggi/Leasing, Assicurazioni, Pasti/Rappresentanza, Marketing, Altro
CATEGORIE ENTRATE: Video Production, Motion Design, Art Direction, Post-Production, Consulenza, Altro
METODI: Carta di credito, Bonifico, Addebito diretto, PayPal, Contanti

AZIONI DISPONIBILI:

1. INSERISCI USCITA:
{"azione":"inserisci_uscita","data":"DD/MM/YYYY","descrizione":"...","fornitore":"...","categoria":"...","metodo":"...","imponibile":123.45,"iva_pct":0.22,"note":"..."}

2. INSERISCI ENTRATA (con IVA):
{"azione":"inserisci_entrata","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","n_fattura":"...","categoria":"Video Production","metodo":"Bonifico","imponibile":123.45,"iva_pct":0.22,"stato":"Da incassare","note":"..."}

3. INSERISCI RIMBORSO (entrata senza IVA, es. Berto, Giuse Barbieri):
{"azione":"inserisci_rimborso","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","importo":123.45,"note":"..."}

4. MODIFICA RIGA - quando l'utente vuole correggere/aggiornare un dato esistente:
{"azione":"modifica","foglio":"Uscite o Entrate o Rimborsi Spese","cerca_descrizione":"testo da cercare nella colonna descrizione","cerca_cliente_o_fornitore":"nome da cercare","campo":"nome colonna da modificare","nuovo_valore":"nuovo valore"}
Campi modificabili: data, descrizione, fornitore, cliente, categoria, metodo, imponibile, iva_pct, stato, note, importo

5. ELIMINA RIGA:
{"azione":"elimina","foglio":"Uscite o Entrate o Rimborsi Spese","cerca_descrizione":"testo da cercare","cerca_cliente_o_fornitore":"nome da cercare"}

6. DOMANDA O CHIARIMENTO:
{"azione":"chiedi","testo":"La tua domanda"}

7. RIEPILOGO:
{"azione":"riepilogo","testo":"Risposta con calcoli"}
"""

def get_sheets_service():
    import google.oauth2.service_account as sa
    from googleapiclient.discovery import build
    creds_dict = json.loads(GOOGLE_CREDS)
    creds = sa.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=creds)

def append_row(sheet_name, values):
    service = get_sheets_service()
    body = {"values": [values]}
    service.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:Z",
        valueInputOption="USER_ENTERED",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

def get_sheet_data(sheet_name):
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!A:Z"
    ).execute()
    return result.get("values", [])

def find_row(sheet_name, cerca_desc=None, cerca_nome=None):
    """Trova la riga che corrisponde alla ricerca. Restituisce (indice_riga, dati_riga) o None."""
    data = get_sheet_data(sheet_name)
    # Colonne: Uscite: A=data, B=desc, C=fornitore | Entrate: A=data, B=desc, C=cliente | Rimborsi: A=data, B=desc, C=cliente
    for i, row in enumerate(data):
        if i == 0:
            continue  # salta intestazione
        desc = row[1].lower() if len(row) > 1 else ""
        nome = row[2].lower() if len(row) > 2 else ""
        match_desc = cerca_desc and cerca_desc.lower() in desc
        match_nome = cerca_nome and cerca_nome.lower() in nome
        if match_desc or match_nome:
            return i + 1, row  # +1 perché Sheets è 1-indexed
    return None, None

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
    # Ottieni sheet ID
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
            "sheetId": sheet_id,
            "dimension": "ROWS",
            "startIndex": row_num - 1,
            "endIndex": row_num
        }}}]}
    ).execute()
    return True

# Mappa campo → lettera colonna
COL_MAP = {
    "Uscite": {
        "data": "A", "descrizione": "B", "fornitore": "C", "categoria": "D",
        "metodo": "E", "imponibile": "F", "iva_pct": "G", "iva_eur": "H",
        "totale": "I", "note": "J"
    },
    "Entrate": {
        "data": "A", "descrizione": "B", "cliente": "C", "n_fattura": "D",
        "categoria": "E", "metodo": "F", "imponibile": "G", "iva_pct": "H",
        "iva_eur": "I", "totale": "J", "stato": "K", "note": "L"
    },
    "Rimborsi Spese": {
        "data": "A", "descrizione": "B", "cliente": "C", "importo": "D", "note": "E"
    }
}

def chiedi_claude(testo=None, immagine_b64=None, mime_type="image/jpeg"):
    oggi = datetime.now().strftime("%d/%m/%Y")
    system = SYSTEM_PROMPT.replace("{oggi}", oggi)

    if immagine_b64:
        if mime_type == "application/pdf":
            content = [
                {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": immagine_b64}},
                {"type": "text", "text": testo or "Analizza questa fattura/documento e inseriscila nel foglio contabilità di Kemp Studio."}
            ]
        else:
            content = [
                {"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": immagine_b64}},
                {"type": "text", "text": testo or "Analizza questo scontrino/fattura e inseriscilo nel foglio contabilità."}
            ]
    else:
        content = testo

    resp = requests.post(
        "https://api.anthropic.com/v1/messages",
        headers={
            "x-api-key": ANTHROPIC_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json"
        },
        json={
            "model": "claude-sonnet-4-6",
            "max_tokens": 600,
            "system": system,
            "messages": [{"role": "user", "content": content}]
        }
    )
    raw = resp.json()["content"][0]["text"].strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

def esegui_azione(risultato):
    azione = risultato.get("azione")

    if azione == "inserisci_uscita":
        d = risultato
        imp = float(d["imponibile"])
        iva = imp * float(d["iva_pct"])
        tot = imp + iva
        append_row("Uscite", [
            d["data"], d["descrizione"], d["fornitore"], d["categoria"],
            d["metodo"], imp, d["iva_pct"], iva, tot, d.get("note","")
        ])
        return (
            f"✅ *Uscita inserita*\n"
            f"📅 {d['data']}\n"
            f"🏷️ {d['fornitore']}\n"
            f"💸 Imp: €{imp:.2f} | IVA: €{iva:.2f} | Tot: €{tot:.2f}\n"
            f"📁 {d['categoria']} — {d['metodo']}"
        )

    elif azione == "inserisci_entrata":
        d = risultato
        imp = float(d["imponibile"])
        iva = imp * float(d["iva_pct"])
        tot = imp + iva
        append_row("Entrate", [
            d["data"], d["descrizione"], d["cliente"], d.get("n_fattura",""),
            d["categoria"], d["metodo"], imp, d["iva_pct"], iva, tot,
            d.get("stato","Da incassare"), d.get("note","")
        ])
        return (
            f"✅ *Entrata inserita*\n"
            f"📅 {d['data']}\n"
            f"🏷️ {d['cliente']}\n"
            f"💰 Imp: €{imp:.2f} | IVA: €{iva:.2f} | Tot: €{tot:.2f}\n"
            f"📁 {d['categoria']} — {d.get('stato','Da incassare')}"
        )

    elif azione == "inserisci_rimborso":
        d = risultato
        imp = float(d["importo"])
        append_row("Rimborsi Spese", [
            d["data"], d["descrizione"], d["cliente"], imp, d.get("note","")
        ])
        return (
            f"✅ *Rimborso inserito*\n"
            f"📅 {d['data']}\n"
            f"🏷️ {d['cliente']}\n"
            f"💵 €{imp:.2f}"
        )

    elif azione == "modifica":
        d = risultato
        foglio = d.get("foglio", "Uscite")
        row_num, row_data = find_row(
            foglio,
            cerca_desc=d.get("cerca_descrizione"),
            cerca_nome=d.get("cerca_cliente_o_fornitore")
        )
        if row_num is None:
            return "⚠️ Riga non trovata. Prova a essere più specifico sulla descrizione o il nome."

        campo = d.get("campo", "").lower()
        nuovo = d.get("nuovo_valore")
        col_map = COL_MAP.get(foglio, {})
        col = col_map.get(campo)

        if not col:
            return f"⚠️ Campo '{campo}' non riconosciuto."

        update_cell(foglio, row_num, col, nuovo)

        # Se modifico imponibile o iva_pct, ricalcola iva_eur e totale
        if campo in ("imponibile", "iva_pct") and foglio in ("Uscite", "Entrate"):
            try:
                imp_col = col_map.get("imponibile")
                iva_col = col_map.get("iva_pct")
                data = get_sheet_data(foglio)
                row = data[row_num - 1]
                imp = float(row[ord(imp_col)-ord('A')])
                iva_pct = float(row[ord(iva_col)-ord('A')])
                iva_eur = imp * iva_pct
                tot = imp + iva_eur
                update_cell(foglio, row_num, col_map["iva_eur"], iva_eur)
                update_cell(foglio, row_num, col_map["totale"], tot)
            except:
                pass

        return (
            f"✏️ *Modifica effettuata*\n"
            f"📋 Foglio: {foglio}\n"
            f"🔧 Campo: {campo} → {nuovo}"
        )

    elif azione == "elimina":
        d = risultato
        foglio = d.get("foglio", "Uscite")
        row_num, row_data = find_row(
            foglio,
            cerca_desc=d.get("cerca_descrizione"),
            cerca_nome=d.get("cerca_cliente_o_fornitore")
        )
        if row_num is None:
            return "⚠️ Riga non trovata."
        desc = row_data[1] if len(row_data) > 1 else "?"
        delete_row(foglio, row_num)
        return f"🗑️ *Riga eliminata*\n📋 {foglio}: {desc}"

    elif azione in ("chiedi", "riepilogo"):
        return risultato.get("testo", "")

    return "⚠️ Azione non riconosciuta"

def send_message(chat_id, text):
    requests.post(f"{TELEGRAM_API}/sendMessage", json={
        "chat_id": chat_id,
        "text": text,
        "parse_mode": "Markdown"
    })

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
        if "photo" in message or "document" in message:
            send_message(chat_id, "🔍 Analizzo il documento...")
            if "photo" in message:
                file_id = message["photo"][-1]["file_id"]
                mime = "image/jpeg"
                file_url = get_file_url(file_id)
                img_data = requests.get(file_url).content
                img_b64 = base64.b64encode(img_data).decode()
            else:
                file_id = message["document"]["file_id"]
                mime = message["document"].get("mime_type", "image/jpeg")
                file_url = get_file_url(file_id)
                img_data = requests.get(file_url).content
                # Se è un PDF, mandalo come documento base64 a Claude
                if mime == "application/pdf":
                    img_b64 = base64.b64encode(img_data).decode()
                    mime = "application/pdf"
                else:
                    img_b64 = base64.b64encode(img_data).decode()
            caption = message.get("caption", "")
            risultato = chiedi_claude(testo=caption, immagine_b64=img_b64, mime_type=mime)

        elif "text" in message:
            testo = message["text"]
            if testo == "/start":
                send_message(chat_id,
                    "👋 Ciao! Sono l'assistente contabile di *Kemp Studio*.\n\n"
                    "Puoi:\n"
                    "📸 Mandarmi la *foto di uno scontrino*\n"
                    "✍️ Scrivere *'pagato 80€ benzina carta'*\n"
                    "💰 O *'fattura Skeptical agosto 1500€'*\n"
                    "✏️ O *'modifica fattura Berto giugno da 1000 a 1500'*\n"
                    "🗑️ O *'elimina uscita Deliveroo del 10 maggio'*\n\n"
                    "Inserisco e aggiorno tutto nel foglio automaticamente!"
                )
                return jsonify({"ok": True})
            risultato = chiedi_claude(testo=testo)
        else:
            send_message(chat_id, "Manda una foto o scrivi il movimento 👆")
            return jsonify({"ok": True})

        risposta = esegui_azione(risultato)
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
