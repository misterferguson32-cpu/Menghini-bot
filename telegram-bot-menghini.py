"""
MENGHINI Bot per Telegram - DETTATURA CODICI CATALOGO - NUOVO FLUSSO
1. Metti codice -> bot dice cos'e' + 3 netti -> chiede se aggiungere e quantita'
2. Rispondi SI + quantita' (es. SI 2) oppure NO
3. Lista utenti con codici e nomi con /utenti o bottone Lista utenti
NON mettere il token nel codice. Usa variabile d'ambiente BOT_TOKEN.
"""

import os, csv, io, re, logging, threading, http.server, socketserver
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, ContextTypes, filters
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("menghini_bot")
BOT_TOKEN = os.getenv("BOT_TOKEN", "INSERISCI_IL_TUO_TOKEN_QUI")

CATALOGO = {}
CLIENTI = []
NETTI = {}

def carica_csv(nome, func):
    if not os.path.exists(nome):
        logger.warning(f"File {nome} non trovato")
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
        "num": (row.get("NumCliente") or row.get("Codice") or "").strip(),
        "nome": (row.get("Nome") or row.get("NomeCliente") or "").strip(),
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

CATALOGO.update({
    "DCG418T2-QW": "SMERIGLIATRICE ANGOLARE 125MM 54V XR FLEXVOLT 2 BATTERIE",
    "DCD996P3-QW": "TRAPANO AVVITATORE A PERCUSSIONE 18V XR 3 BATTERIE",
})
NETTI.update({
    "DCG418T2-QW": {"desc": "SMERIGLIATRICE 125MM 54V", "f0": "312,00", "f1": "295,00", "f2": "278,00"},
})
CLIENTI.extend([
    {"num": "1001", "nome": "CAMBIELLI REGGIO", "localita": "REGGIO EMILIA", "indirizzo": "VIA ESEMPIO 1", "office": "REGGIO"},
    {"num": "1002", "nome": "ROSSI FERRAMENTA", "localita": "MILANO", "indirizzo": "VIA ESEMPIO 2", "office": "MILANO"},
])
carica_csv("catalogo.csv", _load_catalogo)
carica_csv("clienti.csv", _load_clienti)
carica_csv("netti.csv", _load_netti)

def get_state(context):
    if "cliente" not in context.user_data:
        context.user_data["cliente"] = None
    if "lista" not in context.user_data:
        context.user_data["lista"] = []
    if "in_attesa" not in context.user_data:
        context.user_data["in_attesa"] = None  # {codice, qta_proposta}
    if "modalita" not in context.user_data:
        context.user_data["modalita"] = None  # None oppure "cerca_utente"
    return context.user_data

TASTIERA = ReplyKeyboardMarkup([
    [KeyboardButton("🎤 Dettatura"), KeyboardButton("🔍 Cerca codice")],
    [KeyboardButton("👤 Cliente"), KeyboardButton("💶 Netto 3 fasce")],
    [KeyboardButton("👥 Lista utenti"), KeyboardButton("🔍 Cerca utente")],
    [KeyboardButton("📋 Lista"), KeyboardButton("📥 Excel")],
    [KeyboardButton("🗑️ Pulisci lista")],
], resize_keyboard=True)

def normalizza_codice(q):
    q = q.strip().upper().replace(" ", "")
    for k in list(CATALOGO.keys()) + list(NETTI.keys()):
        if k.replace("-", "") == q:
            return k
    return q

def estrai_codice_quantita(testo):
    t = testo.upper().strip()
    qta = 1
    m_qta = re.search(r'(\d+)\s*(PEZZI|PZ|QT|QUANTITA)?\s*$', t)
    if m_qta and ("PEZZI" in t or "PZ" in t or " " in t.strip()):
        try:
            cand = int(m_qta.group(1))
            if cand < 500:
                qta = cand
                t = t[:m_qta.start()].strip()
        except: pass
    t_nospazi = t.replace(" ", "")
    m = re.search(r'([A-Z0-9]{3,}[-]?[A-Z0-9\-]{3,})', t_nospazi)
    if m:
        codice = normalizza_codice(m.group(1))
        return codice, qta
    return None, qta

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_state(context)
    await update.message.reply_text(
        "Ciao, sono *MENGHINI Bot* - DETTATURA CODICI CATALOGO\n\n"
        "NUOVO FLUSSO:\n"
        "1. Metti un codice, es. `DCG418T2-QW`\n"
        "2. Ti dico cos'è + i 3 netti\n"
        "3. Ti chiedo: vuoi aggiungere alla lista? E quanti pezzi?\n"
        "4. Rispondi `SI 2` oppure `NO`\n\n"
        "Con 👥 Lista utenti vedi utenti con codici e nomi.\nCon 🔍 Cerca utente cerchi per codice o nome.\n"
        "Con 📥 Excel ti mando il file con solo *Codice* e *Quantità*",
        parse_mode="Markdown", reply_markup=TASTIERA)

# --- NUOVO: mostra info + netti e chiede conferma ---
async def mostra_info_codice(update: Update, context: ContextTypes.DEFAULT_TYPE, codice_raw: str, qta_proposta: int = 1):
    codice = normalizza_codice(codice_raw)
    if codice not in CATALOGO and codice not in NETTI:
        await update.message.reply_text(f"Codice {codice} non trovato. Controlla il codice e riprova.")
        return
    desc = CATALOGO.get(codice, NETTI.get(codice, {}).get("desc", ""))
    netto_txt = ""
    if codice in NETTI:
        n = NETTI[codice]
        netto_txt = f"\n\n💶 Netti 3 fasce:\nFascia 0: {n['f0']}\nFascia 1: {n['f1']}\nFascia 2: {n['f2']}"
    else:
        netto_txt = "\n\n💶 Netto non presente nei Netti 3 Fasce per questo codice."
    st = get_state(context)
    st["in_attesa"] = {"codice": codice, "qta_proposta": qta_proposta}
    await update.message.reply_text(
        f"📦 *{codice}*\n{desc}{netto_txt}\n\n"
        f"Vuoi aggiungere alla lista?\n"
        f"Quantità proposta: {qta_proposta}\n"
        f"Rispondi: *SI {qta_proposta}* per aggiungere {qta_proposta} pezzi, oppure *SI 5* per 5 pezzi, oppure *NO* per non aggiungere.",
        parse_mode="Markdown")

async def gestisci_conferma(update: Update, context: ContextTypes.DEFAULT_TYPE, testo: str):
    st = get_state(context)
    attesa = st.get("in_attesa")
    if not attesa:
        return False
    t = testo.strip().upper()
    # NO
    if t in ["NO", "N", "NO GRAZIE"]:
        cod = attesa["codice"]
        st["in_attesa"] = None
        await update.message.reply_text(f"Ok, {cod} non aggiunto. Dimmi un altro codice.")
        return True
    # SI / SÌ + quantità
    m = re.match(r'^(SI|SÌ|SI\B|SÌ\B|YES)\s*(\d+)?', t)
    if m:
        qta = int(m.group(2)) if m.group(2) else attesa["qta_proposta"]
        cod = attesa["codice"]
        # aggiungi
        for i, (c, q) in enumerate(st["lista"]):
            if c == cod:
                st["lista"][i] = (c, q + qta)
                break
        else:
            st["lista"].append((cod, qta))
        st["in_attesa"] = None
        await update.message.reply_text(f"✅ Aggiunto: *{cod}* x {qta}\nTotale righe in lista: {len(st['lista'])}\n\nDimmi un altro codice, oppure tocca 📥 Excel per il file.", parse_mode="Markdown")
        return True
    # solo numero = quantità per codice in attesa
    if re.match(r'^\d+$', t):
        qta = int(t)
        cod = attesa["codice"]
        for i, (c, q) in enumerate(st["lista"]):
            if c == cod:
                st["lista"][i] = (c, q + qta)
                break
        else:
            st["lista"].append((cod, qta))
        st["in_attesa"] = None
        await update.message.reply_text(f"✅ Aggiunto: *{cod}* x {qta}\nTotale righe: {len(st['lista'])}", parse_mode="Markdown")
        return True
    return False

async def cerca_codice(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "🔍 Cerca codice":
        await update.message.reply_text("Scrivimi il codice. Es. DCG418T2-QW")
        return
    codice, qta = estrai_codice_quantita(testo)
    if codice:
        await mostra_info_codice(update, context, codice, qta)
    else:
        await update.message.reply_text("Scrivimi il codice. Es. DCG418T2-QW")

async def cliente(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "👤 Cliente":
        await update.message.reply_text("Scrivimi nome cliente o località. Es. CAMBIELLI REGGIO")
        return
    q = testo.upper()
    trovati = [c for c in CLIENTI if q in c["nome"].upper() or q in c["localita"].upper() or q == c["num"]]
    if not trovati:
        await update.message.reply_text("Cliente non trovato. Riprova con nome o località.")
        return
    if len(trovati) == 1:
        c = trovati[0]
        get_state(context)["cliente"] = c
        await update.message.reply_text(f"✅ Cliente scelto:\n*{c['nome']}*\n{c['localita']} - {c['indirizzo']}\nCodice cliente: {c['num']}", parse_mode="Markdown")
        return
    txt = "Trovati più clienti, riscrivi nome completo:\n" + "\n".join([f"• *{c['num']}* - {c['nome']} - {c['localita']}" for c in trovati[:10]])
    await update.message.reply_text(txt, parse_mode="Markdown")

async def netto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "💶 Netto 3 fasce":
        await update.message.reply_text("Scrivimi il codice per il netto. Es. DCG418T2-QW")
        return
    codice, qta = estrai_codice_quantita(testo)
    if codice:
        await mostra_info_codice(update, context, codice, qta)

async def lista_utenti(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Mostra utenti con codici e nomi - primi 20 + invito a cercare
    if not CLIENTI:
        await update.message.reply_text("Lista utenti vuota. Carica clienti.csv con 692 utenti.")
        return
    txt = "👥 *Lista utenti - Codici e Nomi* (primi 20):\n\n" + "\n".join([f"• *{c['num']}* - {c['nome']} - {c['localita']}" for c in CLIENTI[:20]])
    txt += f"\n\nTotale utenti: {len(CLIENTI)}\nPer cercare un utente scrivi il nome o il codice, es. CAMBIELLI o 1001, oppure tocca 👤 Cliente."
    await update.message.reply_text(txt, parse_mode="Markdown")
    # Manda anche Excel lista utenti
    wb = Workbook()
    ws = wb.active
    ws.title = "Utenti"
    ws.append(["Codice Cliente", "Nome Cliente", "Località", "Indirizzo", "Sales Office"])
    for c in CLIENTI:
        ws.append([c["num"], c["nome"], c["localita"], c["indirizzo"], c["office"]])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    await update.message.reply_document(document=buf, filename="Lista_Utenti_Codici_Nomi.xlsx", caption=f"Lista utenti completa - {len(CLIENTI)} utenti con codici e nomi")

async def cerca_utente(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if testo == "🔍 Cerca utente":
        get_state(context)["modalita"] = "cerca_utente"
        await update.message.reply_text(
            "🔍 *Cerca utente*\n\nScrivimi il codice cliente o il nome.\nEs. `1001` oppure `CAMBIELLI` oppure `REGGIO`",
            parse_mode="Markdown")
        return
    # se siamo in modalita cerca_utente, cerca
    q = testo.upper()
    trovati = [c for c in CLIENTI if q in c["nome"].upper() or q in c["localita"].upper() or q == c["num"] or q in c["num"]]
    get_state(context)["modalita"] = None
    if not trovati:
        await update.message.reply_text(f"Nessun utente trovato per `{testo}`. Riprova con codice o nome.", parse_mode="Markdown")
        return
    if len(trovati) == 1:
        c = trovati[0]
        get_state(context)["cliente"] = c
        await update.message.reply_text(
            f"✅ Utente trovato:\n*Codice:* {c['num']}\n*Nome:* {c['nome']}\n*Località:* {c['localita']}\n*Indirizzo:* {c['indirizzo']}\n*Sales Office:* {c['office']}\n\nCliente scelto per l'ordine. Ora dimmi un codice articolo.",
            parse_mode="Markdown")
        return
    txt2 = f"Trovati {len(trovati)} utenti:\n\n" + "\n".join([f"• *{c['num']}* - {c['nome']} - {c['localita']}" for c in trovati[:15]])
    txt2 += "\n\nRiscrivi nome completo o codice per sceglierne uno."
    await update.message.reply_text(txt2, parse_mode="Markdown")

async def mostra_lista(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = get_state(context)
    if not st["lista"]:
        await update.message.reply_text("Lista vuota. Metti un codice per iniziare.")
        return
    cliente_txt = f"Cliente: {st['cliente']['nome']} (Cod. {st['cliente']['num']})\n\n" if st["cliente"] else ""
    txt = cliente_txt + "📋 Lista:\n" + "\n".join([f"• {c} x {q}" for c, q in st["lista"]])
    await update.message.reply_text(txt)

async def manda_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    st = get_state(context)
    if not st["lista"]:
        await update.message.reply_text("Lista vuota, niente da esportare.")
        return
    wb = Workbook(); ws = wb.active; ws.title = "Ordine"
    ws.append(["Codice", "Quantità"])
    for c, q in st["lista"]:
        ws.append([c, q])
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    nome = "Ordine_MENGHINI.xlsx"
    if st["cliente"]:
        nome = f"Ordine_{st['cliente']['nome'].replace(' ', '_')}.xlsx"
    await update.message.reply_document(document=buf, filename=nome, caption=f"✅ Excel pronto - {len(st['lista'])} righe, solo Codice e Quantità\nCliente: {st['cliente']['nome'] if st['cliente'] else 'Non scelto'}")

async def gestisci_testo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    testo = (update.message.text or "").strip()
    if not testo:
        return
    # Prima: se sei in ricerca utente, cerca
    st_tmp = get_state(context)
    if st_tmp.get("modalita") == "cerca_utente":
        await cerca_utente(update, context)
        return
    # Prima: se c'è un codice in attesa, gestisci SI/NO/quantità
    if await gestisci_conferma(update, context, testo):
        return
    if testo in ["🎤 Dettatura", "🔍 Cerca codice", "👤 Cliente", "💶 Netto 3 fasce", "👥 Lista utenti", "🔍 Cerca utente", "📋 Lista", "📥 Excel", "🗑️ Pulisci lista"]:
        if testo == "🎤 Dettatura":
            await update.message.reply_text("Dimmi il codice, es. DCG418T2-QW. Poi ti dico cos'è + i 3 netti e ti chiedo se aggiungerlo.")
        elif testo == "📋 Lista":
            await mostra_lista(update, context)
        elif testo == "📥 Excel":
            await manda_excel(update, context)
        elif testo == "🗑️ Pulisci lista":
            get_state(context)["lista"] = []
            get_state(context)["in_attesa"] = None
            await update.message.reply_text("Lista svuotata. Catalogo e clienti restano salvati.")
        elif testo == "🔍 Cerca codice":
            await cerca_codice(update, context)
        elif testo == "👤 Cliente":
            await cliente(update, context)
        elif testo == "💶 Netto 3 fasce":
            await netto(update, context)
        elif testo == "👥 Lista utenti":
            await lista_utenti(update, context)
        elif testo == "🔍 Cerca utente":
            await cerca_utente(update, context)
        return
    # Se sembra un codice -> mostra info + netti e chiede conferma (NON aggiunge subito)
    if re.search(r'[A-Z]', testo.upper()) and re.search(r'\d', testo):
        # se è un nome cliente noto, trattalo come cliente
        for c in CLIENTI:
            if testo.upper() in c["nome"].upper() and len(testo) > 4:
                await cliente(update, context)
                return
        codice, qta = estrai_codice_quantita(testo)
        if codice:
            await mostra_info_codice(update, context, codice, qta)
            return
    await cliente(update, context)

async def gestisci_vocale(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        voice = update.message.voice
        f = await voice.get_file()
        data = await f.download_as_bytearray()
        try:
            import whisper, tempfile
            with tempfile.NamedTemporaryFile(suffix=".ogg", delete=False) as tmp:
                tmp.write(data); tmp_path = tmp.name
            model = whisper.load_model("tiny")
            res = model.transcribe(tmp_path, language="it")
            testo = res.get("text", "").strip()
            os.unlink(tmp_path)
            if testo:
                await update.message.reply_text(f"Ho capito: _{testo}_")
                codice, qta = estrai_codice_quantita(testo)
                if codice:
                    await mostra_info_codice(update, context, codice, qta)
                return
        except Exception as e:
            logger.info(f"Whisper non disponibile: {e}")
        await update.message.reply_text("Ho ricevuto il vocale, ma la trascrizione non è attiva su questo server. Scrivimi il codice qui, es. DCG418T2-QW")
    except Exception as e:
        logger.error(e)
        await update.message.reply_text("Vocale non letto. Scrivimi il codice qui.")

def start_dummy_server():
    # Render Web Service vuole una porta aperta, altrimenti dice "No open ports detected"
    # e riavvia il bot in loop. Questo server finto tiene aperta la porta.
    port = int(os.getenv("PORT", 10000))
    try:
        handler = http.server.SimpleHTTPRequestHandler
        httpd = socketserver.TCPServer(("0.0.0.0", port), handler)
        t = threading.Thread(target=httpd.serve_forever, daemon=True)
        t.start()
        print(f"Dummy server acceso su porta {port} per Render")
    except Exception as e:
        print(f"Dummy server non partito: {e}")

def main():
    if BOT_TOKEN == "INSERISCI_IL_TUO_TOKEN_QUI":
        print("ERRORE: metti il token in variabile BOT_TOKEN")
        return
    start_dummy_server()
    app = ApplicationBuilder().token(BOT_TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("lista", mostra_lista))
    app.add_handler(CommandHandler("excel", manda_excel))
    app.add_handler(CommandHandler("utenti", lista_utenti))
    app.add_handler(CommandHandler("cerca", cerca_codice))
    app.add_handler(MessageHandler(filters.VOICE, gestisci_vocale))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, gestisci_testo))
    print("MENGHINI Bot Telegram avviato - NUOVO FLUSSO")
    # drop_pending_updates=True risolve il Conflict: terminated by other getUpdates
    # quando il bot era partito 2 volte. Ora riparte pulito.
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
