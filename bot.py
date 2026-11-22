import time, math, sqlite3, asyncio, json, os, re, tempfile, io, urllib.request, urllib.parse
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from xml.sax.saxutils import escape
from PIL import Image, ImageDraw, ImageFilter
from aiohttp import web

from aiogram import Bot, Dispatcher, F
from aiogram.types import Message, ReplyKeyboardMarkup, KeyboardButton, FSInputFile
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.context import FSMContext
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.enums import ParseMode

from google import genai
from google.genai import types

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml

from reportlab.lib.pagesizes import landscape, portrait, A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Image as RLImage
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 1. SOZLAMALAR VA KALITLAR (Environment Variables orqali olinadi)
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "").strip()
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "").strip()
ADMIN_CARD = os.environ.get("ADMIN_CARD", "8600 0000 0000 0000").strip()
ADMIN_OWNER = os.environ.get("ADMIN_OWNER", "Admin").strip()
INITIAL_BONUS_BALANCE = int(os.environ.get("INITIAL_BONUS_BALANCE", "50000"))

if not TELEGRAM_TOKEN or not GEMINI_API_KEY:
    raise ValueError("TELEGRAM_TOKEN yoki GEMINI_API_KEY kiritilmagan!")

client = genai.Client(api_key=GEMINI_API_KEY)
TEXT_MODELS = ["gemini-2.5-flash", "gemini-2.5-flash-lite"]
EXECUTOR = ThreadPoolExecutor(max_workers=6)

bot = Bot(token=TELEGRAM_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# 2. BAZA (Joriy papkada saqlanadi)
conn = sqlite3.connect("slayd_bot_pro_v7.db", check_same_thread=False)
cur = conn.cursor()
cur.execute("""
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    full_name TEXT,
    balance INTEGER DEFAULT 0,
    total_count INTEGER DEFAULT 0,
    paid_count INTEGER DEFAULT 0,
    ref_count INTEGER DEFAULT 0,
    invited_by INTEGER
)
""")
conn.commit()

def register_user(uid, name, inviter=None):
    cur.execute("SELECT user_id FROM users WHERE user_id=?", (uid,))
    if not cur.fetchone():
        valid_inv = None
        if inviter and inviter != uid:
            cur.execute("SELECT user_id FROM users WHERE user_id=?", (inviter,))
            if cur.fetchone():
                valid_inv = inviter
                cur.execute("UPDATE users SET balance=balance+1000, ref_count=ref_count+1 WHERE user_id=?", (valid_inv,))
        cur.execute("INSERT INTO users VALUES (?, ?, ?, 0, 0, 0, ?)", (uid, name, INITIAL_BONUS_BALANCE, valid_inv))
        conn.commit()
        return valid_inv
    return None

def get_user(uid):
    cur.execute("SELECT balance, total_count, paid_count, ref_count FROM users WHERE user_id=?", (uid,))
    row = cur.fetchone()
    if not row:
        register_user(uid, "Foydalanuvchi")
        return (INITIAL_BONUS_BALANCE, 0, 0, 0)
    return row

def update_bal(uid, amt):
    cur.execute("UPDATE users SET balance=balance+? WHERE user_id=?", (amt, uid))
    conn.commit()

def record_slide(uid, was_paid):
    if was_paid:
        cur.execute("UPDATE users SET total_count=total_count+1, paid_count=paid_count+1 WHERE user_id=?", (uid,))
    else:
        cur.execute("UPDATE users SET total_count=total_count+1 WHERE user_id=?", (uid,))
    conn.commit()

# 3. MENYULAR VA OCHROQ DIZAYNLAR
user_themes = {}
user_photo_batches = {}
PRICES = {"5": 3000, "10": 3000, "15": 5000, "20": 7000}

class BotState(StatesGroup):
    author = State()
    subject = State()
    topic = State()
    design_pick = State()
    count = State()
    crossword_topic = State()
    img_to_pdf = State()
    topup_amt = State()
    receipt = State()

main_menu = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="🎨 Slayd yaratish"), KeyboardButton(text="🧩 Krossvord yasash")],
    [KeyboardButton(text="🖼 Rasmlarni PDF qilish"), KeyboardButton(text="💎 Dizayn tanlash")],
    [KeyboardButton(text="💳 Hisobni to'ldirish"), KeyboardButton(text="👛 Balansim")],
    [KeyboardButton(text="👥 Do'stlarni taklif qilish"), KeyboardButton(text="💰 Tariflar & Aksiya")],
    [KeyboardButton(text="ℹ️ Yordam")]
], resize_keyboard=True)

design_menu = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="☀️ Ochroq Osmon"), KeyboardButton(text="🌿 Ochroq Yalpiz")],
    [KeyboardButton(text="🍑 Ochroq Shaftoli"), KeyboardButton(text="🌸 Ochroq Lavanda")],
    [KeyboardButton(text="🤖 AI avtomatik tanlasin"), KeyboardButton(text="⬅️ Asosiy menyu")]
], resize_keyboard=True)

count_menu = ReplyKeyboardMarkup(keyboard=[
    [KeyboardButton(text="5"), KeyboardButton(text="10")],
    [KeyboardButton(text="15"), KeyboardButton(text="20")],
    [KeyboardButton(text="⬅️ Asosiy menyu")]
], resize_keyboard=True)

THEMES = {
    "☀️ Ochroq Osmon": {"bg": (238, 245, 255), "c1": (37, 99, 235), "c2": (14, 165, 233), "accent": "1D4ED8", "text": "0F172A", "sub": "334155"},
    "🌿 Ochroq Yalpiz": {"bg": (238, 250, 245), "c1": (16, 185, 129), "c2": (20, 184, 166), "accent": "047857", "text": "062E22", "sub": "2B5447"},
    "🍑 Ochroq Shaftoli": {"bg": (255, 246, 238), "c1": (249, 115, 22), "c2": (245, 158, 11), "accent": "C2410C", "text": "2C1608", "sub": "5C3A21"},
    "🌸 Ochroq Lavanda": {"bg": (248, 242, 255), "c1": (139, 92, 246), "c2": (236, 72, 153), "accent": "6D28D9", "text": "1E1035", "sub": "473466"}
}

def get_theme(uid, topic="", explicit_choice=None):
    sel = explicit_choice or user_themes.get(uid, "🤖 AI avtomatik tanlasin")
    if sel == "🤖 AI avtomatik tanlasin":
        t = (topic or "").lower()
        if any(w in t for w in ["tabiat", "ekologiya", "biologiya", "tibbiyot"]): sel = "🌿 Ochroq Yalpiz"
        elif any(w in t for w in ["biznes", "iqtisod", "tarix", "madaniyat"]): sel = "🍑 Ochroq Shaftoli"
        elif any(w in t for w in ["adabiyot", "psixologiya", "dizayn", "til"]): sel = "🌸 Ochroq Lavanda"
        else: sel = "☀️ Ochroq Osmon"
    return THEMES.get(sel, THEMES["☀️ Ochroq Osmon"])

# 4. OCHROQ FON VA MAVZUGA MOS HAQIQIY RASMLAR
def make_clean_light_bg(theme, idx, total, bg_path):
    w, h = 1600, 900
    bg = Image.new("RGB", (w, h), theme["bg"])
    draw = ImageDraw.Draw(bg, "RGBA")
    ang = (idx / max(1, total)) * 2 * math.pi
    c1, c2 = theme["c1"], theme["c2"]
    for k in range(3):
        kx = int(320 + k * 480 + 80 * math.cos(ang + k))
        ky = int(190 + (k % 2) * 440 + 60 * math.sin(ang + k))
        for r in range(250, 30, -35):
            draw.ellipse((kx - r, ky - r, kx + r, ky + r), fill=(*(c1 if k % 2 == 0 else c2), int(15 * (1 - r / 250))))
    for line_i in range(4):
        pts = [(x, 765 + line_i * 20 + int(22 * math.sin(x * 0.012 + ang * 2 + line_i * 0.6))) for x in range(850, 1620, 40)]
        draw.line(pts, fill=(*c1, 45), width=3)
    bg.filter(ImageFilter.GaussianBlur(radius=1.0)).save(bg_path, "PNG")

def make_3d_styled_card(raw_img, theme, obj_path):
    w, h = 680, 500
    canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    shadow = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ImageDraw.Draw(shadow).rounded_rectangle((35, 35, w - 15, h - 15), radius=34, fill=(15, 23, 42, 65))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=14))
    canvas.paste(shadow, (0, 0), shadow)

    card_w, card_h = w - 50, h - 50
    ImageDraw.Draw(canvas).rounded_rectangle((20, 20, 20 + card_w, 20 + card_h), radius=32, fill=(255, 255, 255, 245), outline=(*theme["c1"], 180), width=4)

    inner_w, inner_h = card_w - 20, card_h - 20
    iw, ih = raw_img.size
    scale = max(inner_w / iw, inner_h / ih)
    nw, nh = int(iw * scale), int(ih * scale)
    resized = raw_img.resize((nw, nh), Image.LANCZOS)
    left, top = (nw - inner_w) // 2, (nh - inner_h) // 2
    cropped = resized.crop((left, top, left + inner_w, top + inner_h)).convert("RGBA")

    mask = Image.new("L", (inner_w, inner_h), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, inner_w, inner_h), radius=24, fill=255)
    canvas.paste(cropped, (30, 30), mask)
    canvas.save(obj_path, "PNG")

def fetch_topic_photo(queries, theme, idx, total, obj_path):
    headers = {"User-Agent": "SlaydProBot/1.0 (Educational Presentation Builder)"}
    for q in queries:
        if not q:
            continue
        try:
            wp_url = (
                "https://en.wikipedia.org/w/api.php?action=query&generator=search"
                f"&gsrsearch={urllib.parse.quote(q)}&gsrlimit=4&prop=pageimages&pithumbsize=900&format=json"
            )
            req = urllib.request.Request(wp_url, headers=headers)
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            pages = list(data.get("query", {}).get("pages", {}).values())
            pages.sort(key=lambda x: x.get("index", 99))
            for p in pages:
                thumb = p.get("thumbnail", {}).get("source")
                if thumb and not thumb.lower().endswith(".svg"):
                    img_req = urllib.request.Request(thumb, headers=headers)
                    with urllib.request.urlopen(img_req, timeout=8) as img_resp:
                        raw_im = Image.open(io.BytesIO(img_resp.read())).convert("RGB")
                        if raw_im.width >= 200 and raw_im.height >= 150:
                            make_3d_styled_card(raw_im, theme, obj_path)
                            return

            cm_url = (
                "https://commons.wikimedia.org/w/api.php?action=query&generator=search"
                f"&gsrsearch=filetype:bitmap+{urllib.parse.quote(q)}&gsrnamespace=6&gsrlimit=4"
                "&prop=imageinfo&iiprop=url&iiurlwidth=900&format=json"
            )
            req2 = urllib.request.Request(cm_url, headers=headers)
            with urllib.request.urlopen(req2, timeout=8) as resp2:
                data2 = json.loads(resp2.read().decode("utf-8"))
            for p in data2.get("query", {}).get("pages", {}).values():
                ii = p.get("imageinfo", [{}])[0]
                t_url = ii.get("thumburl") or ii.get("url")
                if t_url and not any(bad in t_url.lower() for bad in [".svg", ".pdf", ".tif"]):
                    img_req = urllib.request.Request(t_url, headers=headers)
                    with urllib.request.urlopen(img_req, timeout=8) as img_resp:
                        raw_im = Image.open(io.BytesIO(img_resp.read())).convert("RGB")
                        if raw_im.width >= 200 and raw_im.height >= 150:
                            make_3d_styled_card(raw_im, theme, obj_path)
                            return
        except Exception:
            continue

    ang = (idx / max(1, total)) * 2 * math.pi
    c1, c2 = theme["c1"], theme["c2"]
    obj = Image.new("RGBA", (640, 540), (0, 0, 0, 0))
    odraw = ImageDraw.Draw(obj, "RGBA")
    cx, cy = 320, 255
    size = 115
    dx, dy = int(42 * math.cos(ang)), int(26 * math.sin(ang))
    odraw.polygon([(cx - size - dx, cy + dy), (cx, cy + size - dy), (cx, cy + size + 75), (cx - size - dx, cy + 75 + dy)], fill=(*c1, 185))
    odraw.polygon([(cx, cy + size - dy), (cx + size + dx, cy - dy), (cx + size + dx, cy + 75 - dy), (cx, cy + size + 75)], fill=(*c2, 205))
    odraw.polygon([(cx, cy - size + dy), (cx + size + dx, cy - dy), (cx, cy + size - dy), (cx - size - dx, cy + dy)], fill=(*c1, 240))
    obj.save(obj_path, "PNG")

# 5. GEMINI AI
def call_gemini(prompt, system_inst, temp=0.35):
    last_err = None
    for m in TEXT_MODELS:
        for _ in range(2):
            try:
                res = client.models.generate_content(
                    model=m, contents=prompt,
                    config=types.GenerateContentConfig(system_instruction=system_inst, response_mime_type="application/json", temperature=temp)
                )
                raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", (res.text or "").strip())
                return json.loads(raw)
            except Exception as e:
                last_err = e
                time.sleep(0.7)
    raise last_err or ValueError("AI javob bera olmadi.")

def generate_presentation_data(author, subject, topic, count):
    prompt = f"""
    O'zbek tilida oliy o'quv yurti va ilmiy konferensiya darajasida JUDA BOY, BATAFSIL VA CHUQUR ma'lumotlarga ega taqdimot yarat.
    Muallif: {author} | Fan: {subject} | Mavzu: {topic} | Kontent slaydlar soni: {count - 1} ta.

    Faqat JSON formatida javob ber:
    {{
      "topic_en": "2-3 word exact English Wikipedia search term for '{topic}'",
      "slides": [
        {{
          "title": "Slaydning aniq ilmiy sarlavhasi",
          "search_keyword_en": "2-3 word exact English Wikipedia search keyword for this specific slide (different for each slide!)",
          "bullets": [
            "1-band: Batafsil ilmiy ta'rif, tarixiy sana yoki nazariy asos (kamida 18-25 so'zli to'liq ma'lumot)",
            "2-band: Aniq faktlar, statistik ko'rsatkichlar yoki muhim xususiyatlar tahlili",
            "3-band: Amaliy misol, qo'llanilish sohasi yoki jarayon mexanizmi",
            "4-band: Muammo, yechim yoki zamonaviy rivojlanish tendensiyalari",
            "5-band: Ushbu bo'lim bo'yicha asosiy ilmiy xulosa"
          ]
        }}
      ],
      "references": [
        "1. Muallif I.O. Kitob yoki darslik nomi. — Toshkent: Nashriyot, yil.",
        "2. Ikkinchi ilmiy adabiyot yoki monografiya nomi.",
        "3. Uchinchi darslik yoki xalqaro manba nomi.",
        "4. To'rtinchi ilmiy maqola yoki qo'llanma."
      ],
      "funny_joke": "Mavzuga mos qisqa kulgili hazil 😂"
    }}

    Qat'iy talablar:
    - "slides" ro'yxatida aynan {count - 1} ta slayd bo'lsin.
    - Har bir slaydda 5 ta mazmunli, uzun va faktlarga boy band bo'lsin.
    - Har bir slayd uchun "search_keyword_en" maydoniga Wikipedia'da rasmi aniq chiqadigan 2-3 so'zli inglizcha atama yoz.
    - "references" bo'limida "{subject}" fani va "{topic}" mavzusiga oid 4-5 ta real kitob yozilsin.
    """
    data = call_gemini(
        prompt,
        "Sen chuqur ilmiy ma'lumotlar yozuvchi professional taqdimot ekspertisan. Faqat to'g'ri JSON qaytar.",
        temp=0.4
    )
    if not data.get("slides"):
        raise ValueError("Slaydlar yaratilmadi.")
    return data

def generate_crossword_data(topic):
    prompt = f'O\'zbek tilida "{topic}" mavzusiga oid KROSSVORD tayyorla. JSON: {{"title": "Sarlavha", "across": [{{"num": 1, "clue": "Savol", "answer": "JAVOB"}}], "down": [{{"num": 2, "clue": "Savol", "answer": "JAVOB"}}], "funny_outro": "Hazil 😂"}} (5 gorizontal, 5 vertikal).'
    return call_gemini(prompt, "Sen krossvord tuzuvchi AI san. Faqat JSON qaytar.", temp=0.5)

async def verify_receipt_with_ai(photo_bytes, claimed_amount):
    try:
        img = Image.open(io.BytesIO(photo_bytes))
        res = await asyncio.to_thread(
            client.models.generate_content, model=TEXT_MODELS[0],
            contents=[img, f"To'lov cheki tekshiruvi ({claimed_amount} so'm). JSON: {{\"note\": \"qisqa izoh\"}}"],
            config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.1)
        )
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", (res.text or "").strip())
        return json.loads(raw)
    except Exception:
        return {"note": "Chek qabul qilindi"}

# 6. PPTX VA PDF
def create_pptx(data, author, subject, topic, theme, filename, work_dir):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    accent = RGBColor.from_string(theme["accent"])
    text_c = RGBColor.from_string(theme["text"])
    sub_c = RGBColor.from_string(theme["sub"])

    topic_en = data.get("topic_en", topic)
    all_slides = [{"type": "title", "queries": [topic_en, topic]}]
    for s in data["slides"]:
        kw = s.get("search_keyword_en", "")
        all_slides.append({"type": "content", "queries": [kw, topic_en], **s})
    if data.get("references"):
        all_slides.append({
            "type": "refs",
            "title": "📚 Foydalanilgan kitoblar va ilmiy adabiyotlar",
            "queries": ["Library books", "University library", topic_en],
            "bullets": data["references"]
        })
    total = len(all_slides)

    def build_assets(idx):
        bg_p = os.path.join(work_dir, f"bg_{idx}.png")
        obj_p = os.path.join(work_dir, f"obj_{idx}.png")
        make_clean_light_bg(theme, idx, total, bg_p)
        fetch_topic_photo(all_slides[idx]["queries"], theme, idx, total, obj_p)

    list(EXECUTOR.map(build_assets, range(total)))

    for idx, item in enumerate(all_slides):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide._element.append(parse_xml(
            '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" spd="fast">'
            '<p14:prism dir="l" isContent="1"/></p:transition>'
        ))
        slide.shapes.add_picture(os.path.join(work_dir, f"bg_{idx}.png"), Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))

        shift_y = 1.65 + (0.22 if idx % 2 == 0 else -0.15)
        pic = slide.shapes.add_picture(os.path.join(work_dir, f"obj_{idx}.png"), Inches(8.6), Inches(shift_y), width=Inches(4.35))
        if idx % 2 == 1:
            pic.rotation = 2.5

        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.4), Inches(0.45), Inches(0.08), Inches(6.55))
        bar.fill.solid(); bar.fill.fore_color.rgb = accent; bar.line.fill.background()

        if item["type"] == "title":
            p1 = slide.shapes.add_textbox(Inches(0.8), Inches(1.1), Inches(7.6), Inches(0.65)).text_frame.paragraphs[0]
            p1.text = f"FAN: {subject.upper()}"
            p1.font.size = Pt(17); p1.font.bold = True; p1.font.color.rgb = accent

            tf = slide.shapes.add_textbox(Inches(0.8), Inches(1.95), Inches(7.6), Inches(2.6)).text_frame
            tf.word_wrap = True
            p2 = tf.paragraphs[0]
            p2.text = topic
            p2.font.size = Pt(34); p2.font.bold = True; p2.font.color.rgb = text_c

            tf_a = slide.shapes.add_textbox(Inches(0.8), Inches(4.9), Inches(7.5), Inches(1.4)).text_frame
            tf_a.word_wrap = True
            p3 = tf_a.paragraphs[0]
            p3.text = f"👤 Tayyorladi: {author}\n📅 Sana: {datetime.now().strftime('%d.%m.%Y')}"
            p3.font.size = Pt(18); p3.font.color.rgb = sub_c
        else:
            tf_t = slide.shapes.add_textbox(Inches(0.75), Inches(0.4), Inches(7.7), Inches(0.95)).text_frame
            tf_t.word_wrap = True
            pt = tf_t.paragraphs[0]
            pt.text = str(item["title"])
            pt.font.size = Pt(22); pt.font.bold = True; pt.font.color.rgb = accent

            tf = slide.shapes.add_textbox(Inches(0.75), Inches(1.42), Inches(7.7), Inches(5.45)).text_frame
            tf.word_wrap = True
            for j, b in enumerate(item["bullets"]):
                p = tf.paragraphs[0] if j == 0 else tf.add_paragraph()
                p.text = ("" if item["type"] == "refs" else "▸  ") + str(b)
                p.font.size = Pt(14.5)
                p.font.color.rgb = text_c
                p.space_after = Pt(9)

        pn = slide.shapes.add_textbox(Inches(0.8), Inches(7.02), Inches(12.0), Inches(0.32)).text_frame.paragraphs[0]
        pn.text = f"{subject}  |  {author}  |  {idx+1}/{total}"
        pn.alignment = PP_ALIGN.RIGHT; pn.font.size = Pt(10.5); pn.font.color.rgb = sub_c

    prs.save(filename)
    return all_slides

def create_pdf(all_slides, author, subject, topic, theme, filename, work_dir):
    font_regular = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    font_bold = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"
    use_dejavu = os.path.exists(font_regular) and os.path.exists(font_bold)
    if use_dejavu:
        pdfmetrics.registerFont(TTFont("DejaVu", font_regular))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", font_bold))
        fn_reg, fn_bold = "DejaVu", "DejaVu-Bold"
    else:
        fn_reg, fn_bold = "Helvetica", "Helvetica-Bold"

    accent = colors.HexColor("#" + theme["accent"])
    text_c = colors.HexColor("#" + theme["text"])
    sub_c = colors.HexColor("#" + theme["sub"])
    ps = landscape(A4)
    doc = SimpleDocTemplate(filename, pagesize=ps, rightMargin=265, leftMargin=42, topMargin=34, bottomMargin=34)
    s_fan = ParagraphStyle("F", fontName=fn_bold, fontSize=13.5, leading=18, textColor=accent, spaceAfter=10)
    s_hero = ParagraphStyle("H", fontName=fn_bold, fontSize=24, leading=31, textColor=text_c, spaceAfter=18)
    s_meta = ParagraphStyle("M", fontName=fn_reg, fontSize=13.5, leading=20, textColor=sub_c)
    s_tit = ParagraphStyle("T", fontName=fn_bold, fontSize=17.5, leading=23, textColor=accent, spaceAfter=12)
    s_bod = ParagraphStyle("B", fontName=fn_reg, fontSize=11.2, leading=16.5, textColor=text_c, spaceAfter=7.5)

    story = []
    for idx, item in enumerate(all_slides):
        if item["type"] == "title":
            story.extend([
                Spacer(1, 45),
                Paragraph(f"FAN: {escape(subject.upper())}", s_fan),
                Paragraph(escape(topic), s_hero),
                Spacer(1, 22),
                Paragraph(f"👤 Tayyorladi: {escape(author)}<br/>📅 Sana: {datetime.now().strftime('%d.%m.%Y')}", s_meta)
            ])
        else:
            story.append(Paragraph(escape(str(item["title"])), s_tit))
            for b in item["bullets"]:
                story.append(Paragraph(("" if item["type"] == "refs" else "▸ ") + escape(str(b)), s_bod))
        if idx < len(all_slides) - 1:
            story.append(PageBreak())

    def paint(canvas, _):
        p_idx = min(canvas.getPageNumber() - 1, len(all_slides) - 1)
        bgp = os.path.join(work_dir, f"bg_{p_idx}.png")
        objp = os.path.join(work_dir, f"obj_{p_idx}.png")
        canvas.saveState()
        if os.path.exists(bgp):
            canvas.drawImage(bgp, 0, 0, width=ps[0], height=ps[1])
        canvas.setFillColor(accent)
        canvas.rect(20, 24, 4, ps[1] - 48, fill=1, stroke=0)
        if os.path.exists(objp):
            canvas.drawImage(objp, ps[0] - 255, 130, width=230, height=170, mask="auto")
        canvas.restoreState()

    doc.build(story, onFirstPage=paint, onLaterPages=paint)

def images_to_pdf(photo_paths, out_pdf):
    doc = SimpleDocTemplate(out_pdf, pagesize=portrait(A4), rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    max_w, max_h = portrait(A4)[0] - 50, portrait(A4)[1] - 50
    story = []
    for i, p in enumerate(photo_paths):
        with Image.open(p) as im: iw, ih = im.size
        scale = min(max_w / iw, max_h / ih)
        story.append(RLImage(p, width=iw * scale, height=ih * scale))
        if i < len(photo_paths) - 1: story.append(PageBreak())
    doc.build(story)

# 7. HANDLERLAR
@dp.message(F.text.startswith("/start"))
async def cmd_start(m: Message, state: FSMContext):
    await state.clear()
    user_photo_batches.pop(m.from_user.id, None)
    parts = m.text.split()
    inv = int(parts[1].replace("ref_", "")) if len(parts) > 1 and parts[1].startswith("ref_") and parts[1][4:].isdigit() else None
    rew = register_user(m.from_user.id, m.from_user.full_name, inv)
    user_themes.setdefault(m.from_user.id, "☀️ Ochroq Osmon")
    if rew:
        try: await bot.send_message(rew, f"🎉 {m.from_user.full_name} havolangiz orqali qo'shildi! +1 000 so'm bonus!")
        except Exception: pass
    bal, tot, _, _ = get_user(m.from_user.id)
    gift = "🎁 **Birinchi taqdimotingiz 100% BEPUL!**\n" if tot == 0 else ""
    await m.answer(f"Assalomu alaykum, {m.from_user.first_name}! 👋\n💎 **AI SLAYD PRO** botiga xush kelibsiz!\n{gift}👛 Balans: **{bal:,} so'm**\n\nKerakli bo'limni tanlang:", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

@dp.message(F.text == "⬅️ Asosiy menyu")
async def cmd_back(m: Message, state: FSMContext):
    await state.clear()
    await m.answer("Asosiy menyu:", reply_markup=main_menu)

@dp.message(F.text == "🎨 Slayd yaratish")
async def s_auth(m: Message, state: FSMContext):
    await state.clear()
    await state.set_state(BotState.author)
    await m.answer("👤 **1-qadam:** Ism va familiyangizni kiriting:", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.author)
async def s_subj(m: Message, state: FSMContext):
    if len((m.text or "").strip()) < 2: return await m.answer("Ism-familiyani kiriting.")
    await state.update_data(author=m.text.strip())
    await state.set_state(BotState.subject)
    await m.answer("📘 **2-qadam:** Qaysi **fan** bo'yicha slayd tayyorlaymiz?", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.subject)
async def s_top(m: Message, state: FSMContext):
    if len((m.text or "").strip()) < 2: return await m.answer("Fan nomini yozing.")
    await state.update_data(subject=m.text.strip())
    await state.set_state(BotState.topic)
    await m.answer("📝 **3-qadam:** Slayd **mavzusini** yozing:", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.topic)
async def s_des_pick(m: Message, state: FSMContext):
    if len((m.text or "").strip()) < 2: return await m.answer("Mavzuni yozing.")
    await state.update_data(topic=m.text.strip())
    await state.set_state(BotState.design_pick)
    await m.answer("🎨 **4-qadam:** Ochroq dizaynlardan birini tanlang:", reply_markup=design_menu)

@dp.message(BotState.design_pick)
async def s_cnt(m: Message, state: FSMContext):
    valid_d = list(THEMES.keys()) + ["🤖 AI avtomatik tanlasin"]
    chosen = m.text if m.text in valid_d else "🤖 AI avtomatik tanlasin"
    user_themes[m.from_user.id] = chosen
    await state.update_data(chosen_design=chosen)
    await state.set_state(BotState.count)
    _, tot, paid, _ = get_user(m.from_user.id)
    hdr = "🎁 **5-qadam (1-TAQDIMOT BEPUL!):** Nechta slayd kerak?" if tot == 0 else ("🎁 **5-qadam (4-BUYURTMA BEPUL!):** Nechta slayd kerak?" if (paid > 0 and paid % 3 == 0) else "📊 **5-qadam:** Nechta slayd kerak?\n• 5 yoki 10 ta — **3 000 so'm**\n• 15 ta — **5 000 so'm**\n• 20 ta — **7 000 so'm**")
    await m.answer(hdr, parse_mode=ParseMode.MARKDOWN, reply_markup=count_menu)

@dp.message(BotState.count)
async def s_gen(m: Message, state: FSMContext):
    if m.text not in PRICES: return await m.answer("5, 10, 15 yoki 20 sonidan birini tanlang.")
    count, price, uid = int(m.text), PRICES[m.text], m.from_user.id
    bal, tot, paid, _ = get_user(uid)
    is_free = (tot == 0) or (paid > 0 and paid % 3 == 0)
    if not is_free and bal < price:
        await state.clear()
        return await m.answer(f"❌ Balans yetarli emas (**{bal:,} so'm**). Kerak: **{price:,} so'm**.", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

    d = await state.get_data()
    author, subject, topic = d["author"], d["subject"], d["topic"]
    theme = get_theme(uid, topic, d.get("chosen_design"))
    await state.clear()

    promo = "🎁 **1-TAQDIMOT BEPUL!**\n" if tot == 0 else ("🎁 **4-BUYURTMA BEPUL!**\n" if is_free else f"💳 Narx: **{price:,} so'm**\n")
    st = await m.answer(
        f"{promo}⚡ **Chuqur ilmiy ma'lumotlar va mavzuga mos rasmlar yuklanmoqda...**\n"
        f"👤 {author} | 📘 {subject}\n📝 {topic} 🚀",
        parse_mode=ParseMode.MARKDOWN
    )
    try:
        ai_data = await asyncio.to_thread(generate_presentation_data, author, subject, topic, count)
        with tempfile.TemporaryDirectory() as folder:
            safe_name = re.sub(r"[^\w\s-]", "", topic)[:20].strip() or "Taqdimot"
            pptx_p, pdf_p = os.path.join(folder, f"{safe_name}_3D.pptx"), os.path.join(folder, f"{safe_name}_3D.pdf")
            all_s = await asyncio.to_thread(create_pptx, ai_data, author, subject, topic, theme, pptx_p, folder)
            await asyncio.to_thread(create_pdf, all_s, author, subject, topic, theme, pdf_p, folder)
            if not is_free: update_bal(uid, -price)
            record_slide(uid, was_paid=(not is_free))
            nb, _, np, _ = get_user(uid)
            joke = ai_data.get("funny_joke", "Slayd tayyor, endi faqat yaxshi baho olish qoldi! 😂🎓")
            await m.answer_document(FSInputFile(pptx_p), caption="📊 **3D PowerPoint (PPTX)** — mavzuga mos rasmlar va chuqur matnlar bilan!", parse_mode=ParseMode.MARKDOWN)
            await m.answer_document(FSInputFile(pdf_p), caption="📄 **PDF format** — kitoblar ro'yxati bilan!", parse_mode=ParseMode.MARKDOWN)
            await m.answer(f"🤣 **{joke}**", parse_mode=ParseMode.MARKDOWN)
        nxt = "🎉 Keyingi taqdimotingiz BEPUL!" if (np > 0 and np % 3 == 0) else f"Bepul taqdimotga {3 - (np % 3)} ta pullik buyurtma qoldi"
        await st.edit_text(f"✅ **Tayyor!** Balans: **{nb:,} so'm** | {nxt}", parse_mode=ParseMode.MARKDOWN)
    except Exception as e:
        print("XATOLIK:", repr(e))
        await st.edit_text(f"❌ Xatolik ({type(e).__name__}). Qayta urinib ko'ring.")
    await m.answer("Yana slayd yaratishingiz mumkin! 🎨😄", reply_markup=main_menu)

@dp.message(F.text == "💎 Dizayn tanlash")
async def cmd_des(m: Message): await m.answer("🎨 Ochroq dizaynni tanlang:", reply_markup=design_menu)

@dp.message(F.text.in_(list(THEMES.keys()) + ["🤖 AI avtomatik tanlasin"]))
async def save_des_standalone(m: Message):
    user_themes[m.from_user.id] = m.text
    await m.answer("✅ Asosiy dizayn saqlandi: " + m.text, reply_markup=main_menu)

@dp.message(F.text == "👛 Balansim")
async def cmd_bal(m: Message):
    bal, tot, paid, ref = get_user(m.from_user.id)
    st = "🎁 1-taqdimotingiz BEPUL!" if tot == 0 else ("🎉 Keyingi taqdimotingiz BEPUL!" if (paid > 0 and paid % 3 == 0) else f"Bepul taqdimotga {3 - (paid % 3)} ta qoldi")
    await m.answer(f"👛 **BALANS:** **{bal:,} so'm**\n📊 Jami: **{tot} ta** | 🎁 **{st}**\n👥 Takliflar: **{ref} kishi**", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

@dp.message(F.text == "👥 Do'stlarni taklif qilish")
async def cmd_ref(m: Message):
    me = await bot.get_me()
    link = f"https://t.me/{me.username}?start=ref_{m.from_user.id}"
    await m.answer(f"👥 Har bir do'stingiz uchun **+1 000 so'm** bonus!\n🔗 Havolangiz:\n`{link}`", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

@dp.message(F.text == "💰 Tariflar & Aksiya")
async def cmd_tariffs(m: Message):
    await m.answer("💰 **TARIFLAR:**\n🎁 1-taqdimot: **BEPUL!**\n• 5 yoki 10 ta — **3 000 so'm**\n• 15 ta — **5 000 so'm**\n• 20 ta — **7 000 so'm**\n🔥 Har 3 ta pullikdan keyin 4-si BEPUL!", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

@dp.message(F.text == "💳 Hisobni to'ldirish")
async def cmd_topup(m: Message, state: FSMContext):
    await state.set_state(BotState.topup_amt)
    await m.answer(f"💳 Karta: `{ADMIN_CARD}` ({ADMIN_OWNER})\nQancha summa o'tkazmoqchisiz? (Masalan: `10000`):", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.topup_amt)
async def step_topup_amt(m: Message, state: FSMContext):
    d = re.sub(r"\D", "", m.text or "")
    if not d or int(d) < 1000: return await m.answer("Kamida 1000 so'm kiriting.")
    await state.update_data(amt=int(d))
    await state.set_state(BotState.receipt)
    await m.answer(f"✅ **{int(d):,} so'm** uchun to'lov cheki rasmini (skrinshot) yuboring:", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.receipt)
async def step_receipt(m: Message, state: FSMContext):
    if not m.photo: return await m.answer("To'lov cheki rasmini yuboring.")
    amt = (await state.get_data())["amt"]
    await state.clear()
    buf = io.BytesIO()
    await bot.download(m.photo[-1], destination=buf)
    res = await verify_receipt_with_ai(buf.getvalue(), amt)
    update_bal(m.from_user.id, amt)
    bal, _, _, _ = get_user(m.from_user.id)
    await m.answer(f"✅ **Chek qabul qilindi!** ({res.get('note', 'OK')})\n➕ Qo'shildi: **{amt:,} so'm** | 👛 Balans: **{bal:,} so'm**", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

@dp.message(F.text == "🧩 Krossvord yasash")
async def cw_start(m: Message, state: FSMContext):
    await state.set_state(BotState.crossword_topic)
    await m.answer("🧩 Krossvord mavzusini yozing:")

@dp.message(BotState.crossword_topic)
async def cw_build(m: Message, state: FSMContext):
    topic = (m.text or "").strip()
    if len(topic) < 2: return await m.answer("Mavzuni yozing.")
    await state.clear()
    w = await m.answer("🧩 Krossvord tuzilmoqda...")
    try:
        cw = await asyncio.to_thread(generate_crossword_data, topic)
        lines = [f"🧩 **{cw.get('title', topic.upper())}**\n\n➡️ **Gorizontal:**"]
        for i in cw.get("across", []): lines.append(f"{i.get('num')}. {i.get('clue')} — ||`{i.get('answer')}`||")
        lines.append("\n⬇️ **Vertikal:**")
        for i in cw.get("down", []): lines.append(f"{i.get('num')}. {i.get('clue')} — ||`{i.get('answer')}`||")
        lines.append(f"\n😂 {cw.get('funny_outro', 'Ajoyib! 🧠😂')}")
        await w.delete()
        await m.answer("\n".join(lines), parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)
    except Exception as e:
        await w.edit_text(f"❌ Xatolik: {type(e).__name__}")

@dp.message(F.text == "🖼 Rasmlarni PDF qilish")
async def pdf_img_start(m: Message, state: FSMContext):
    user_photo_batches[m.from_user.id] = []
    await state.set_state(BotState.img_to_pdf)
    await m.answer("🖼 Rasmlarni yuboring va tugatgach **`/pdf_tayyor`** deb yozing:", parse_mode=ParseMode.MARKDOWN)

@dp.message(BotState.img_to_pdf, F.photo)
async def pdf_img_collect(m: Message):
    uid = m.from_user.id
    user_photo_batches.setdefault(uid, [])
    with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp:
        await bot.download(m.photo[-1], destination=tmp.name)
        user_photo_batches[uid].append(tmp.name)
    await m.answer(f"📥 Rasm qo'shildi ({len(user_photo_batches[uid])} ta). Yana tashlang yoki `/pdf_tayyor` yozing.")

@dp.message(BotState.img_to_pdf, F.text == "/pdf_tayyor")
async def pdf_img_finish(m: Message, state: FSMContext):
    uid = m.from_user.id
    paths = user_photo_batches.pop(uid, [])
    await state.clear()
    if not paths: return await m.answer("Rasm yuborilmadi.", reply_markup=main_menu)
    with tempfile.TemporaryDirectory() as folder:
        out_pdf = os.path.join(folder, f"rasmlar_{uid}.pdf")
        await asyncio.to_thread(images_to_pdf, paths, out_pdf)
        for p in paths:
            if os.path.exists(p): os.remove(p)
        await m.answer_document(FSInputFile(out_pdf), caption="📄 Rasmlaringiz PDF formatda tayyor! 😎📎", reply_markup=main_menu)

@dp.message(F.text == "ℹ️ Yordam")
async def cmd_help(m: Message):
    await m.answer("ℹ️ **Slayd yaratish** -> Ism -> Fan -> Mavzu -> Dizayn -> Slayd sonini tanlang!", parse_mode=ParseMode.MARKDOWN, reply_markup=main_menu)

async def handle_ping(request):
    return web.Response(text="AI SLAYD PRO ishlayapti!")

async def main():
    app = web.Application()
    app.router.add_get("/", handle_ping)
    runner = web.AppRunner(app)
    await runner.setup()
    port = int(os.environ.get("PORT", 10000))
    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()

    await bot.delete_webhook(drop_pending_updates=True)
    print(f"🚀 AI SLAYD PRO {port}-PORTDA VA 24/7 REJIMDA ISHGA TUSHDI!")
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
