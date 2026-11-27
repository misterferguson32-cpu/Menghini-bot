"""
MENGHINI Bot per Telegram - DETTATURA CODICI CATALOGO - NUOVO FLUSSO
1. Metti codice -> bot dice cos'e' + 3 netti -> chiede se aggiungere e quantita'
2. Rispondi SI + quantita' (es. SI 2) oppure NO
3. Lista utenti con codici e nomi con /utenti o bottone Lista utenti
NON mettere il token nel codice. Usa variabile d'ambiente BOT_TOKEN.
"""

import os, csv, io, re, base64, logging, threading, http.server, socketserver
from telegram import Update, ReplyKeyboardMarkup, KeyboardButton
from telegram.error import Conflict
from telegram.ext import ApplicationBuilder, CommandHandler, MessageHandler, TypeHandler, ApplicationHandlerStop, ContextTypes, filters
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

carica_csv("catalogo.csv", _load_catalogo)
carica_csv("clienti.csv", _load_clienti)
carica_csv("netti.csv", _load_netti)

# --- SICUREZZA: allow-list di ID Telegram autorizzati ---
# Legge gli ID da (1) variabile d'ambiente AUTORIZZATI (separati da virgola/spazio)
# e (2) file autorizzati.txt (un ID per riga, righe con # ignorate).
def _carica_autorizzati():
    ids = set()
    for x in re.split(r"[,\s]+", os.getenv("AUTORIZZATI", "")):
        x = x.strip()
        if x.lstrip("-").isdigit():
            ids.add(int(x))
    if os.path.exists("autorizzati.txt"):
        with open("autorizzati.txt", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                tok = line.split()[0]
                if tok.lstrip("-").isdigit():
                    ids.add(int(tok))
    return ids

AUTORIZZATI = _carica_autorizzati()

async def check_autorizzazione(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Blocca chiunque non sia nella allow-list. Gira per primo (group=-1).
    user = update.effective_user
    if user is None:
        raise ApplicationHandlerStop
    if user.id not in AUTORIZZATI:
        logger.warning(f"Accesso NEGATO: id={user.id} username={user.username} nome={user.full_name}")
        if update.effective_message:
            await update.effective_message.reply_text(
                f"⛔ Accesso non autorizzato.\n\nIl tuo ID Telegram è: {user.id}\n"
                f"Invia questo numero all'amministratore per farti sbloccare."
            )
        raise ApplicationHandlerStop

def get_state(context):
    if "cliente" not in context.user_data:
        context.user_data["cliente"] = None
    if "lista" not in context.user_data:
        context.user_data["lista"] = []
    if "in_attesa" not in context.user_data:
        context.user_data["in_attesa"] = None  # {codice, qta_proposta}
    if "modalita" not in context.user_data:
        context.user_data["modalita"] = None  # None oppure "cerca_utente"
    if "foto_attesa" not in context.user_data:
        context.user_data["foto_attesa"] = None  # anteprima lettura foto
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
    """Estrae (codice, quantita) accettando vari formati:
    'CODICE 3', 'CODICE x3', '3 CODICE', 'CODICE 3 PZ', 'CODICE' (qta=1)."""
    t = testo.upper().strip()
    # 'x3' / 'x 3' -> ' 3 '
    t = re.sub(r'\bX\s*(\d{1,3})\b', r' \1 ', t)
    # '3 PZ' / '3 PEZZI' / '3 QT' / '3 QUANTITA' -> ' 3 '
    t = re.sub(r'\b(\d{1,3})\s*(PZ|PEZZI|PEZ|QT|QTA|QUANTITA)\b', r' \1 ', t)
    qta = 1
    codice = None
    for tok in re.split(r'[\s,;:]+', t):
        tok = tok.strip()
        if not tok:
            continue
        if tok.lstrip('-').isdigit():
            n = int(tok)
            if 0 < n <= 999:
                qta = n
        elif codice is None:
            codice = tok
    if codice is None:
        return None, qta
    return normalizza_codice(codice), qta

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    get_state(context)
    await update.message.reply_text(
        "Ciao, sono *MENGHINI Bot* - LISTA ACQUISTO\n\n"
        "*Percorso rapido:*\n"
        "1️⃣ Mandami il codice prodotto (scritto o 🎤 vocale), es. `DCG418T2-QW 3`\n"
        "2️⃣ Ti dico cos'è + i 3 netti e ti chiedo conferma\n"
        "3️⃣ Rispondi `SI` (usa la quantità proposta), `SI 5` per cambiarla, oppure `NO`\n"
        "4️⃣ Ripeti per ogni articolo\n"
        "5️⃣ Alla fine tocca 📥 *Excel* per il file (2 colonne: Codice, Quantità)\n\n"
        "Formati quantità: `CODICE 3`, `CODICE x3`, `3 CODICE`.\n"
        "📷 Puoi anche mandare la FOTO di un foglio scritto a mano (codici + quantità): li leggo e te li faccio confermare con `OK`.\n"
        "Utilità: 👤 Cliente, 👥 Lista utenti, 🔍 Cerca utente, 📋 Lista, 🗑️ Pulisci lista.",
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
    # Conferma della lettura foto (anteprima in attesa)
    st = get_state(context)
    if st.get("foto_attesa"):
        t = testo.upper()
        if t in ["OK", "SI", "SÌ", "CONFERMA", "AGGIUNGI"]:
            aggiunti = 0
            for p in st["foto_attesa"]:
                if not p["trovato"]:
                    continue
                for i, (c, q) in enumerate(st["lista"]):
                    if c == p["codice"]:
                        st["lista"][i] = (c, q + p["qta"]); break
                else:
                    st["lista"].append((p["codice"], p["qta"]))
                aggiunti += 1
            st["foto_attesa"] = None
            await update.message.reply_text(f"✅ Aggiunte {aggiunti} righe dalla foto.\nTotale lista: {len(st['lista'])} righe. Tocca 📥 Excel per il file.", parse_mode="Markdown")
            return
        if t in ["ANNULLA", "NO", "CANCELLA"]:
            st["foto_attesa"] = None
            await update.message.reply_text("Ok, lettura foto annullata. Nessuna riga aggiunta.")
            return
        st["foto_attesa"] = None  # altro testo: prosegui con il flusso normale
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
            await update.message.reply_text("Mandami un 🎤 vocale oppure scrivi il codice, es. `DCG418T2-QW 3`.\nTi dico cos'è + i 3 netti e ti chiedo se aggiungerlo.", parse_mode="Markdown")
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

def _ai_config():
    # Sceglie il provider AI: Groq (gratis) se GROQ_API_KEY è impostata, altrimenti OpenAI.
    # Restituisce (base_url, api_key, modello_stt, modello_vision).
    groq = os.getenv("GROQ_API_KEY", "").strip()
    if groq:
        return (
            "https://api.groq.com/openai/v1",
            groq,
            os.getenv("STT_MODEL", "whisper-large-v3-turbo"),
            os.getenv("VISION_MODEL", "qwen/qwen3.8-27b"),
        )
    oa = os.getenv("OPENAI_API_KEY", "").strip()
    if oa:
        return (
            "https://api.openai.com/v1",
            oa,
            os.getenv("STT_MODEL", "whisper-1"),
            os.getenv("VISION_MODEL", "gpt-4o-mini"),
        )
    raise RuntimeError("Nessuna chiave AI configurata (GROQ_API_KEY o OPENAI_API_KEY)")

async def trascrivi_vocale(update: Update) -> str:
    # Trascrive un vocale (Whisper). Usa Groq se configurata, altrimenti OpenAI.
    base, api_key, stt_model, _ = _ai_config()
    voice = update.message.voice
    f = await voice.get_file()
    data = await f.download_as_bytearray()
    import httpx
    async with httpx.AsyncClient(timeout=60) as client:
        r = await client.post(
            f"{base}/audio/transcriptions",
            headers={"Authorization": f"Bearer {api_key}"},
            data={"model": stt_model, "language": "it"},
            files={"file": ("vocale.ogg", bytes(data), "audio/ogg")},
        )
        r.raise_for_status()
        return (r.json().get("text") or "").strip()

async def gestisci_vocale(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        testo = await trascrivi_vocale(update)
    except RuntimeError:
        await update.message.reply_text("🎤 Trascrizione vocale non configurata (manca GROQ_API_KEY). Scrivimi il codice, es. DCG418T2-QW.")
        return
    except Exception as e:
        logger.error(f"Errore trascrizione vocale: {e}")
        await update.message.reply_text("Non sono riuscito a leggere il vocale. Riprova o scrivimi il codice.")
        return
    if not testo:
        await update.message.reply_text("Non ho capito l'audio. Riprova o scrivimi il codice.")
        return
    await update.message.reply_text(f"🎤 Ho capito: {testo}")
    # Il vocale può anche essere la conferma (SI/NO/quantità) di un codice in attesa
    if await gestisci_conferma(update, context, testo):
        return
    codice, qta = estrai_codice_quantita(testo)
    if codice:
        await mostra_info_codice(update, context, codice, qta)
    else:
        await update.message.reply_text("Non ho riconosciuto un codice. Riprova o scrivimi il codice.")

async def leggi_foglio_foto(update: Update):
    # Legge una foto di un foglio scritto a mano (codice + quantità) con un modello vision.
    # Restituisce una lista di dict: {"codice": str, "quantita": int}.
    base, api_key, _, vision_model = _ai_config()
    photo = update.message.photo[-1]  # file di risoluzione più alta
    f = await photo.get_file()
    data = await f.download_as_bytearray()
    b64 = base64.b64encode(bytes(data)).decode()
    prompt = (
        "Leggi questa foto di un foglio scritto a mano. Contiene righe con un CODICE PRODOTTO "
        "e una QUANTITÀ (numero di pezzi). I codici sono alfanumerici tipo DCG418T2-QW o D25133K-QS. "
        "Rispondi SOLO con un array JSON, senza altro testo, nel formato: "
        '[{"codice": "DCG418T2-QW", "quantita": 3}]. '
        "Se una quantità non è indicata usa 1. Ignora intestazioni e righe non di prodotto. "
        "Se non riconosci nessun codice rispondi con []."
    )
    import httpx, json as _json
    async with httpx.AsyncClient(timeout=90) as client:
        r = await client.post(
            f"{base}/chat/completions",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": vision_model,
                "temperature": 0,
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    ],
                }],
            },
        )
        r.raise_for_status()
        txt = (r.json()["choices"][0]["message"]["content"] or "").strip()
    # rimuovi eventuali fence markdown ```json ... ```
    if txt.startswith("```"):
        txt = txt.split("\n", 1)[1] if "\n" in txt else txt
        txt = txt.rsplit("```", 1)[0].strip()
    try:
        dati = _json.loads(txt)
    except Exception:
        raise RuntimeError("Risposta IA non leggibile")
    out = []
    for it in (dati if isinstance(dati, list) else []):
        if isinstance(it, dict) and it.get("codice"):
            try:
                q = int(it.get("quantita") or 1)
            except Exception:
                q = 1
            out.append({"codice": str(it["codice"]).strip().upper(), "quantita": max(1, q)})
    return out

async def gestisci_foto(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        await update.message.reply_text("📷 Leggo il foglio...")
        righe = await leggi_foglio_foto(update)
    except RuntimeError as e:
        await update.message.reply_text(f"📷 Lettura foto non disponibile ({e}).\nScrivimi i codici a mano oppure attiva GROQ_API_KEY su Render.")
        return
    except Exception as e:
        logger.error(f"Errore lettura foto: {e}")
        await update.message.reply_text("Non sono riuscito a leggere la foto. Riprova con una foto più nitida e ben illuminata.")
        return
    if not righe:
        await update.message.reply_text("Non ho trovato codici nella foto. Riprova con una foto più nitida.")
        return
    preview = []
    for it in righe:
        cod = normalizza_codice(it["codice"])
        trovato = cod in CATALOGO or cod in NETTI
        desc = ""
        if trovato:
            desc = CATALOGO.get(cod) or NETTI.get(cod, {}).get("desc", "")
        preview.append({"codice": cod, "qta": it["quantita"], "trovato": trovato, "letto": it["codice"], "desc": desc})
    st = get_state(context)
    st["foto_attesa"] = preview
    st["in_attesa"] = None
    ok = [p for p in preview if p["trovato"]]
    ko = [p for p in preview if not p["trovato"]]
    righe_txt = []
    for p in preview:
        if p["trovato"]:
            righe_txt.append(f"• ✅ `{p['codice']}` x {p['qta']}" + (f" - {p['desc'][:45]}" if p["desc"] else ""))
        else:
            righe_txt.append(f"• ❓ `{p['letto']}` x {p['qta']} (non trovato, letto: {p['letto']})")
    txt = "📷 *Ho letto dalla foto:*\n" + "\n".join(righe_txt)
    txt += f"\n\n{len(ok)} validi" + (f", {len(ko)} da verificare" if ko else "") + "."
    txt += "\nRispondi *OK* per aggiungere alla lista (solo i validi), oppure *ANNULLA*."
    await update.message.reply_text(txt, parse_mode="Markdown")

async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE):
    # Logga gli errori senza far riavviare il bot in loop.
    if isinstance(context.error, Conflict):
        logger.warning("Conflict getUpdates: un'altra istanza ha ripreso il polling (normale durante un deploy/riavvio).")
    else:
        logger.error("Errore durante la gestione di un update", exc_info=context.error)

class _HealthHandler(http.server.BaseHTTPRequestHandler):
    # Risponde solo OK, senza servire file della cartella (evita di esporre
    # clienti.csv, catalogo, netti, listino... sull'URL pubblico di Render).
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"MENGHINI bot attivo")
    def do_HEAD(self):
        self.send_response(200)
        self.end_headers()
    def log_message(self, *args):
        pass

def start_dummy_server():
    # Render Web Service vuole una porta aperta, altrimenti dice "No open ports detected"
    # e riavvia il bot in loop. Questo server finto tiene aperta la porta.
    port = int(os.getenv("PORT", 10000))
    try:
        httpd = socketserver.TCPServer(("0.0.0.0", port), _HealthHandler)
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

    async def _post_init(app):
        # Rimuove eventuali webhook residui che bloccano il polling.
        try:
            await app.bot.delete_webhook()
            logger.info("Webhook rimosso (o assente): polling attivo.")
        except Exception as e:
            logger.warning(f"delete_webhook non riuscito: {e}")

    app = ApplicationBuilder().token(BOT_TOKEN).post_init(_post_init).build()
    app.add_error_handler(on_error)
    # Controllo accesso PRIMA di tutto (group=-1)
    app.add_handler(TypeHandler(Update, check_autorizzazione), group=-1)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("lista", mostra_lista))
    app.add_handler(CommandHandler("excel", manda_excel))
    app.add_handler(CommandHandler("utenti", lista_utenti))
    app.add_handler(CommandHandler("cerca", cerca_codice))
    app.add_handler(MessageHandler(filters.PHOTO, gestisci_foto))
    app.add_handler(MessageHandler(filters.VOICE, gestisci_vocale))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, gestisci_testo))
    print("MENGHINI Bot Telegram avviato - NUOVO FLUSSO")
    # Nessun drop_pending_updates: i messaggi arrivati mentre il bot era offline
    # vengono elaborati e ricevono risposta al riavvio.
    app.run_polling()

if __name__ == "__main__":
    main()
