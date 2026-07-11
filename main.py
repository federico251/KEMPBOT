import os
import json
import base64
import requests
from flask import Flask, request, jsonify
from datetime import datetime

app = Flask(__name__)

# ── CONFIGURAZIONE ────────────────────────────────────────────────────────────
TELEGRAM_TOKEN   = os.environ.get("TELEGRAM_TOKEN")
ANTHROPIC_KEY    = os.environ.get("ANTHROPIC_KEY")
SPREADSHEET_ID   = os.environ.get("SPREADSHEET_ID")
GOOGLE_CREDS     = os.environ.get("GOOGLE_CREDS")  # JSON credenziali service account

TELEGRAM_API     = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"

# ── SYSTEM PROMPT ─────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """Sei l'assistente contabile di Kemp Studio SRL, studio italiano di video production.
Analizzi messaggi di testo o immagini di scontrini/fatture e restituisci SOLO JSON valido.

DATA DEFAULT: {oggi}

REGOLE:
- Se l'importo sembra IVA inclusa, scorporala al 22% (default). Se dici "imponibile" usalo diretto.
- Collaboratori forfettari (Albi, Marco, Milo, Karim) = IVA 0%
- Pasti/ristoranti = IVA 10%
- Carburante = IVA 22%
- Se non capisci qualcosa chiedi conferma con azione "chiedi"

CATEGORIE USCITE: Collaboratori, Consulenze Professionali, Attrezzatura, Software/Abbonamenti, Carburante, Trasferte, Ufficio/Utenze, Noleggi/Leasing, Assicurazioni, Pasti/Rappresentanza, Marketing, Altro
CATEGORIE ENTRATE: Video Production, Motion Design, Art Direction, Post-Production, Consulenza, Altro
METODI: Carta di credito, Bonifico, Addebito diretto, PayPal, Contanti

FORMATO USCITA:
{"azione":"inserisci_uscita","data":"DD/MM/YYYY","descrizione":"...","fornitore":"...","categoria":"...","metodo":"...","imponibile":123.45,"iva_pct":0.22,"note":"..."}

FORMATO ENTRATA:
{"azione":"inserisci_entrata","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","n_fattura":"...","categoria":"Video Production","metodo":"Bonifico","imponibile":123.45,"iva_pct":0.22,"stato":"Da incassare","note":"..."}

FORMATO RIMBORSO (entrata senza IVA da clienti come Berto, Giuse Barbieri):
{"azione":"inserisci_rimborso","data":"DD/MM/YYYY","descrizione":"...","cliente":"...","importo":123.45,"note":"..."}

FORMATO DOMANDA/CHIARIMENTO:
{"azione":"chiedi","testo":"La tua domanda qui"}

FORMATO RIEPILOGO:
{"azione":"riepilogo","testo":"Risposta con dati dal foglio"}
"""

# ── GOOGLE SHEETS ─────────────────────────────────────────────────────────────
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

def get_sheet_data(sheet_name, range_str="A:Z"):
    service = get_sheets_service()
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{sheet_name}!{range_str}"
    ).execute()
    return result.get("values", [])

# ── CLAUDE API ────────────────────────────────────────────────────────────────
def chiedi_claude(testo=None, immagine_b64=None, mime_type="image/jpeg"):
    oggi = datetime.now().strftime("%d/%m/%Y")
    system = SYSTEM_PROMPT.replace("{oggi}", oggi)

    if immagine_b64:
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
            "max_tokens": 500,
            "system": system,
            "messages": [{"role": "user", "content": content}]
        }
    )
    raw = resp.json()["content"][0]["text"].strip()
    raw = raw.replace("```json", "").replace("```", "").strip()
    return json.loads(raw)

# ── ESEGUI AZIONE ─────────────────────────────────────────────────────────────
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
            f"💸 Imponibile: €{imp:.2f} | IVA: €{iva:.2f} | Totale: €{tot:.2f}\n"
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
            f"💰 Imponibile: €{imp:.2f} | IVA: €{iva:.2f} | Totale: €{tot:.2f}\n"
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

    elif azione in ("chiedi", "riepilogo"):
        return risultato.get("testo", "")

    return "⚠️ Azione non riconosciuta"

# ── TELEGRAM ──────────────────────────────────────────────────────────────────
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

# ── WEBHOOK ───────────────────────────────────────────────────────────────────
@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.json
    message = data.get("message", {})
    chat_id = message.get("chat", {}).get("id")
    if not chat_id:
        return jsonify({"ok": True})

    try:
        # FOTO / DOCUMENTO
        if "photo" in message or "document" in message:
            send_message(chat_id, "🔍 Analizzo lo scontrino...")
            if "photo" in message:
                file_id = message["photo"][-1]["file_id"]
                mime = "image/jpeg"
            else:
                file_id = message["document"]["file_id"]
                mime = message["document"].get("mime_type", "image/jpeg")

            file_url = get_file_url(file_id)
            img_data = requests.get(file_url).content
            img_b64 = base64.b64encode(img_data).decode()
            caption = message.get("caption", "")
            risultato = chiedi_claude(testo=caption, immagine_b64=img_b64, mime_type=mime)

        # TESTO
        elif "text" in message:
            testo = message["text"]
            if testo == "/start":
                send_message(chat_id,
                    "👋 Ciao! Sono l'assistente contabile di *Kemp Studio*.\n\n"
                    "Puoi:\n"
                    "📸 Mandarmi la *foto di uno scontrino*\n"
                    "✍️ Scrivere *'pagato 80€ benzina carta'*\n"
                    "💰 O *'fattura Skeptical agosto 1500€'*\n\n"
                    "Inserisco tutto nel foglio automaticamente!"
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
