"""
MENGHINI Bot per Telegram - basato su DETTATURA CODICI CATALOGO
Funzioni: dettatura vocale, cerca codice, cerca cliente, netto 3 fasce, lista, Excel con solo Codice e Quantita
NON mettere il token nel codice. Mettilo come variabile d'ambiente BOT_TOKEN o in .env
Esegui: pip install -r requirements.txt  poi  python telegram_bot_menghini.py
"""

import os
import csv
import io
import re
import logging
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO)

# --- Web server finto per Render Web Service Free ---
# Render Web Service si aspetta che il codice apra una porta web,
# altrimenti resta in "In progress" per sempre. Questo server finto
# tiene viva la pagina, mentre il bot Telegram gira in polling.
from http.server import BaseHTTPRequestHandler, HTTPServer
import threading

class _Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type","text/plain")
        self.end_headers()
        self.wfile.write(b"MENGHINI Bot Telegram is running")
    def log_message(self, *a, **k):
        pass

def avvia_web_server():
    port = int(os.getenv("PORT", "10000"))
    try:
        srv = HTTPServer(("0.0.0.0", port), _Handler)
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        logger.info(f"Web server finto su porta {port}")
    except Exception as e:
        logger.info(f"Web server non avviato: {e}")

logger = logging.getLogger("menghini_bot")

BOT_TOKEN = os.getenv("BOT_TOKEN", "INSERISCI_IL_TUO_TOKEN_QUI")

# --- DATI: sostituisci questi 3 file con i tuoi completi ---
# catalogo.csv colonne: Codice,Descrizione,Prezzo
# clienti.csv colonne: NumCliente,Nome,Localita,Indirizzo,SalesOffice
# netti.csv colonne: Codice,Descrizione,Fascia0,Fascia1,Fascia2

CATALOGO = {}  # codice -> descrizione
CLIENTI = []   # lista dict
NETTI = {}     # codice -> {desc, f0,f1,f2}

def carica_csv(nome, func):
    if not os.path.exists(nome):
        logger.warning(f"File {nome} non trovato, uso dati esempio")
        return
    with open(nome, newline='', encoding='utf-8') as f:
        r = csv.DictReader(f)
        for row in r:
            func(row)

def _load_catalogo(row):
    codice = (row.get("Codice") or "").strip().upper()
    if codice:
        CATALOGO[codice] = (row.get("Descrizione") or "").strip()

def _load_clienti(row):
    CLIENTI.append({
        "num": (row.get("NumCliente") or "").strip(),
        "nome": (row.get("Nome") or "").strip(),
        "localita": (row.get("Localita") or "").strip(),
        "indirizzo": (row.get("Indirizzo") or "").strip(),
        "office": (row.get("SalesOffice") or "").strip(),
    })

def _load_netti(row):
    codice = (row.get("Codice") or "").strip().upper()
    if codice:
        NETTI[codice] = {
            "desc": (row.get("Descrizione") or "").strip(),
            "f0": row.get("Fascia0") or "",
            "f1": row.get("Fascia1") or "",
            "f2": row.get("Fascia2") or "",
        }

# Esempi per test - cancellali quando carichi i file completi 2373 / 692 / 507
CATALOGO.update({
    "DCG418T2-QW": "SMERIGLIATRICE ANGOLARE 125MM 54V XR FLEXVOLT 2 BATTERIE",
    "DCD996P3-QW": "TRAPANO AVVITATORE A PERCUSSIONE 18V XR 3 BATTERIE",
})
NETTI.update({
    "DCG418T2-QW": {"desc": "SMERIGLIATRICE 125MM 54V", "f0": "312,00", "f1": "295,00", "f2": "278,00"},
})
CLIENTI.extend([
    {"num": "1001", "nome": "CAMBIELLI REGGIO", "localita": "REGGIO EMILIA", "indirizzo": "VIA ESEMPIO 1", "office": "REGGIO"},
])

carica_csv("catalogo.csv", _load_catalogo)
carica_csv("clienti.csv", _load_clienti)
carica_csv("netti.csv", _load_netti)

# Stato per chat: cliente scelto + lista codici
def get_state(context):
    if "cliente" not in context.user_data:
        context.user_data["cliente"] = None
    if "lista" not in context.user_data:
        context.user_data["lista"] = []  # lista di (codice, quantita)
    return context.user_data

TASTIERA = ReplyKeyboardMarkup([
    [KeyboardButton("🎤 Dettatura"), KeyboardButton("🔍 Cerca codice")],
    [KeyboardButton("👤 Cliente"), KeyboardButton("💶 Netto 3 fasce")],
    [KeyboardButton("📋 Lista"), KeyboardButton("📥 Excel")],
    [KeyboardButton("🗑️ Pulisci lista")],
], resize_keyboard=True)

def normalizza_testo(testo):
    # da dettatura: "DCG 418 T2 QW 2 pezzi" -> codice + quantita
    t = testo.upper().strip()
    t = re.sub(r'\s+', ' ', t)
    return t

def estrai_codice_quantita(testo):
    """Ritorna (codice, quantita). Quantita di default 1."""
    t = normalizza_testo(testo)
    # cerca pattern codice tipo DCG418T2-QW oppure DCG 418 T2 QW
    # rimuovi spazi dentro al codice
    t_nospazi = t.replace(" ", "")
    # cerca quantita alla fine: "x2", "2 PEZZI", "QT 2"
    qta = 1
    m_qta = re.search(r'(\d+)\s*(PEZZI|PZ|QT|QUANTITA)?\s*$', t)
    if m_qta:
        # se il testo finisce con numero e non è parte del codice, prendilo come quantita
        # solo se c'è uno spazio prima o parola PEZZI
        if "PEZZI" in t or "PZ" in t or " " in t.strip():
            try:
                cand = int(m_qta.group(1))
                # evita di confondere codice che finisce con numero
                if cand < 500:
                    qta = cand
                    # togli quantita dal testo per isolare codice
                    t = t[:m_qta.start()].strip()
                    t_nospazi = t.replace(" ", "")
            except Exception:
                pass
    # codice = sequenza alfanumerica con - , lunga 5-20
    m = re.search(r'([A-Z0-9]{3,}[-]?[A-Z0-9\-]{3,})', t_nospazi)
    if m:
        codice = m.group(1)
        # normalizza: inserisci - se manca e codice noto
        if codice in CATALOGO or codice in NETTI:
            return codice, qta
        # prova a trovare codice noto che contiene questo
        for k in list(CATALOGO.keys()) + list(NETTI.keys()):
            if k.replace("-", "") == codice:
                return k, qta
        return codice, qta
    return None, qta

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_state(context)
    await update.message.reply_text(
        "Ciao, sono *MENGHINI Bot* - DETTATURA CODICI CATALOGO\n\n"
        "1. Scegli il cliente con 👤 Cliente\n"
        "2. Detta i codici con 🎤 Dettatura o scrivili: es. `DCG418T2-QW 2`\n"
        "3. Con 📥 Excel ti mando il file con solo *Codice* e *Quantità*\n\n"
        "Catalogo sempre salvato. Pulisci lista cancella solo la lista, mai catalogo/clienti.",
        parse_mode="Markdown", reply_markup=TASTIERA)

async def cerca_codice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = update.message.text or ""
    # se arriva da bottone, chiedi codice
    if testo.strip() == "🔍 Cerca codice":
        await update.message.reply_text("Scrivimi il codice o il nome articolo. Es. DCG418T2-QW oppure SMERIGLIATRICE 125")
        return
    q = testo.strip().upper()
    if q in CATALOGO:
        await update.message.reply_text(f"✅ *{q}*\n{CATALOGO[q]}", parse_mode="Markdown")
        return
    # ricerca per descrizione
    trovati = [(k, v) for k, v in CATALOGO.items() if q in v.upper() or q in k][:10]
    if trovati:
        txt = "Trovati:\n" + "\n".join([f"• *{k}* - {v}" for k, v in trovati])
        await update.message.reply_text(txt, parse_mode="Markdown")
    else:
        await update.message.reply_text(f"Codice {q} non trovato in catalogo. Controlla il codice e riprova.")

async def cliente(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "👤 Cliente":
        await update.message.reply_text("Scrivimi nome cliente o località. Es. CAMBIELLI REGGIO")
        return
    q = testo.upper()
    # se testo è "CAMBIELLI..." o simile, cerca
    trovati = [c for c in CLIENTI if q in c["nome"].upper() or q in c["localita"].upper() or q == c["num"]]
    if not trovati:
        await update.message.reply_text("Cliente non trovato. Riprova con nome o località.")
        return
    if len(trovati) == 1:
        c = trovati[0]
        get_state(context)["cliente"] = c
        await update.message.reply_text(f"✅ Cliente scelto:\n*{c['nome']}*\n{c['localita']} - {c['indirizzo']}\nNum. {c['num']}", parse_mode="Markdown")
        return
    txt = "Trovati più clienti, riscrivi nome completo:\n" + "\n".join([f"• {c['nome']} - {c['localita']} (Num. {c['num']})" for c in trovati[:10]])
    await update.message.reply_text(txt)

async def netto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "💶 Netto 3 fasce":
        await update.message.reply_text("Scrivimi il codice per il netto. Es. DCG418T2-QW")
        return
    q = testo.strip().upper().replace(" ", "")
    # normalizza con -
    for k in NETTI:
        if k.replace("-", "") == q:
            q = k
            break
    if q in NETTI:
        n = NETTI[q]
        await update.message.reply_text(
            f"💶 *{q}*\n{n['desc']}\n\nFascia 0: {n['f0']}\nFascia 1: {n['f1']}\nFascia 2: {n['f2']}",
            parse_mode="Markdown")
    else:
        await update.message.reply_text("Codice non presente nei Netti 3 Fasce. Controlla il codice.")

async def aggiungi_codice(update: Update, context: ContextTypes.DEFAULT_TYPE, testo_raw: str):
    codice, qta = estrai_codice_quantita(testo_raw)
    if not codice:
        await update.message.reply_text("Non ho capito il codice. Riprova: es. DCG418T2-QW 2")
        return
    st = get_state(context)
    # aggiorna se già in lista
    for i, (c, q) in enumerate(st["lista"]):
        if c == codice:
            st["lista"][i] = (c, q + qta)
            break
    else:
        st["lista"].append((codice, qta))
    desc = CATALOGO.get(codice, NETTI.get(codice, {}).get("desc", ""))
    await update.message.reply_text(f"✅ Aggiunto: *{codice}* x {qta}\n{desc}\nTotale righe in lista: {len(st['lista'])}", parse_mode="Markdown")

async def gestisci_testo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if not testo:
        return
    if testo in ["🎤 Dettatura", "🔍 Cerca codice", "👤 Cliente", "💶 Netto 3 fasce", "📋 Lista", "📥 Excel", "🗑️ Pulisci lista"]:
        if testo == "🎤 Dettatura":
            await update.message.reply_text("Parla ora: tieni premuto il microfono di Telegram e detta il codice, es. DCG418T2-QW 2 pezzi. Oppure scrivilo qui.")
        elif testo == "📋 Lista":
            await mostra_lista(update, context)
        elif testo == "📥 Excel":
            await manda_excel(update, context)
        elif testo == "🗑️ Pulisci lista":
            get_state(context)["lista"] = []
            await update.message.reply_text("Lista svuotata. Catalogo e clienti restano salvati.")
        elif testo == "🔍 Cerca codice":
            await cerca_codice(update, context)
        elif testo == "👤 Cliente":
            await cliente(update, context)
        elif testo == "💶 Netto 3 fasce":
            await netto(update, context)
        return
    # se sembra un codice (contiene lettere+ numeri) -> aggiungi alla lista
    # se il contesto è ricerca cliente/netto, il bot usa l'ultimo comando? Per semplicità: se testo è solo nome cliente e non sembra codice, prova cliente
    # Euristica: se contiene almeno 3 lettere e un numero e lunghezza < 20 -> codice
    if re.search(r'[A-Z]', testo.upper()) and re.search(r'\d', testo):
        # se inizia con nome cliente noto, trattalo come cliente
        for c in CLIENTI:
            if testo.upper() in c["nome"].upper():
                await cliente(update, context)
                return
        await aggiungi_codice(update, context, testo)
        return
    # altrimenti prova cliente o cerca
    await cliente(update, context)

async def mostra_lista(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = get_state(context)
    if not st["lista"]:
        await update.message.reply_text("Lista vuota. Detta un codice per iniziare.")
        return
    cliente_txt = ""
    if st["cliente"]:
        cliente_txt = f"Cliente: {st['cliente']['nome']} - {st['cliente']['localita']}\n\n"
    txt = cliente_txt + "📋 Lista:\n" + "\n".join([f"• {c} x {q}" for c, q in st["lista"]])
    await update.message.reply_text(txt)

async def manda_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = get_state(context)
    if not st["lista"]:
        await update.message.reply_text("Lista vuota, niente da esportare.")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Ordine"
    ws.append(["Codice", "Quantità"])  # SOLO 2 colonne, come richiesto
    for c, q in st["lista"]:
        ws.append([c, q])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    nome = "Ordine_MENGHINI.xlsx"
    if st["cliente"]:
        nome = f"Ordine_{st['cliente']['nome'].replace(' ', '_')}.xlsx"
    await update.message.reply_document(document=buf, filename=nome, caption=f"✅ Excel pronto - {len(st['lista'])} righe, solo Codice e Quantità\nCliente: {st['cliente']['nome'] if st['cliente'] else 'Non scelto'}")
    # NON cancella catalogo/clienti, lista resta per controllo. Se vuoi svuotare, usa 🗑️ Pulisci lista

async def gestisci_vocale(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Telegram manda voice come file .ogg. Qui lo scarichiamo e proviamo trascrizione con Whisper se installato.
    # Se Whisper non è disponibile, chiediamo di scrivere il codice.
    try:
        voice = update.message.voice
        f = await voice.get_file()
        data = await f.download_as_bytearray()
        # Prova Whisper (openai-whisper) se installato
        try:
            import whisper
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".ogg", delete=False) as tmp:
                tmp.write(data)
                tmp_path = tmp.name
            model = whisper.load_model("tiny")
            res = model.transcribe(tmp_path, language="it")
            testo = res.get("text", "").strip()
            os.unlink(tmp_path)
            if testo:
                await update.message.reply_text(f"Ho capito: _{testo}_")
                await aggiungi_codice(update, context, testo)
                return
        except Exception as e:
            logger.info(f"Whisper non disponibile: {e}")
        await update.message.reply_text("Ho ricevuto il vocale, ma la trascrizione automatica non è attiva su questo server. Scrivimi il codice qui, es. DCG418T2-QW 2, e lo aggiungo subito.")
    except Exception as e:
        logger.error(e)
        await update.message.reply_text("Vocale non letto. Scrivimi il codice qui.")

def main():
    if BOT_TOKEN == "INSERISCI_IL_TUO_TOKEN_QUI":
        print("ERRORE: metti il token in variabile BOT_TOKEN o nel file .env")
        print("Esempio: BOT_TOKEN=92427...  python telegram_bot_menghini.py")
        print("Non mettere mai il token in pubblico.")
        return
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("lista", mostra_lista))
    app.add_handler(CommandHandler("excel", manda_excel))
    app.add_handler(CommandHandler("cerca", cerca_codice))
    app.add_handler(MessageHandler(filters.VOICE, gestisci_vocale))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, gestisci_testo))
    avvia_web_server()
    print("MENGHINI Bot Telegram avviato - DETTATURA CODICI CATALOGO")
    app.run_polling()

if __name__ == "__main__":
    main()
