import asyncio
import os
import sqlite3
from datetime import datetime
from typing import Optional

from aiogram import Bot, Dispatcher, F, Router
from aiogram.filters import Command, CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    CallbackQuery,
    FSInputFile,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    KeyboardButton,
    Message,
    ReplyKeyboardMarkup,
)
from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

# ================= SOZLAMALAR =================
BOT_TOKEN = os.getenv("BOT_TOKEN", "9220906070:AAFq_pHmFonr8lJh1V73_pQwaxnKuoOOSCQ")
ADMIN_IDS = [int(x) for x in os.getenv("ADMIN_IDS", "8052376816").split(",") if x.strip().isdigit()]
CARD_NUMBER = os.getenv("CARD_NUMBER", "4073 4200 7949 1878")
CARD_OWNER = os.getenv("CARD_OWNER", "NAZARBEK NORBOYEV.")
PRICE_PER_SLIDE = 350  # 1 slayd = 350 so'm

DB_PATH = "slayd_bot_manual_topup.db"
OUT_DIR = "generated_files"
os.makedirs(OUT_DIR, exist_ok=True)

# ================= BAZA (SQLite) =================
conn = sqlite3.connect(DB_PATH, check_same_thread=False)
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS users (
    user_id INTEGER PRIMARY KEY,
    full_name TEXT,
    username TEXT,
    balance INTEGER DEFAULT 0,
    referrer_id INTEGER,
    joined_at TEXT
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS works (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    work_type TEXT,
    title TEXT,
    language TEXT,
    pages_or_slides INTEGER,
    design TEXT,
    has_images INTEGER,
    ending_style TEXT,
    author_title TEXT,
    price INTEGER,
    file_path TEXT,
    created_at TEXT
)
""")

cur.execute("""
CREATE TABLE IF NOT EXISTS topup_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    claimed_amount INTEGER,
    status TEXT DEFAULT 'pending',
    admin_id INTEGER,
    credited_amount INTEGER DEFAULT 0,
    created_at TEXT
)
""")
conn.commit()


def ensure_user(user_id: int, full_name: str, username: Optional[str], referrer_id: Optional[int] = None) -> dict:
    row = cur.execute("SELECT user_id, full_name, username, balance, referrer_id, joined_at FROM users WHERE user_id = ?", (user_id,)).fetchone()
    if not row:
        valid_ref = referrer_id if (referrer_id and referrer_id != user_id) else None
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        cur.execute(
            "INSERT INTO users (user_id, full_name, username, balance, referrer_id, joined_at) VALUES (?, ?, ?, 0, ?, ?)",
            (user_id, full_name or "Foydalanuvchi", username or "", valid_ref, now),
        )
        conn.commit()
        row = cur.execute("SELECT user_id, full_name, username, balance, referrer_id, joined_at FROM users WHERE user_id = ?", (user_id,)).fetchone()
    return {
        "user_id": row[0],
        "full_name": row[1],
        "username": row[2],
        "balance": row[3],
        "referrer_id": row[4],
        "joined_at": row[5],
    }


def get_balance(user_id: int) -> int:
    row = cur.execute("SELECT balance FROM users WHERE user_id = ?", (user_id,)).fetchone()
    return int(row[0]) if row else 0


def add_balance(user_id: int, amount: int) -> int:
    ensure_user(user_id, "Foydalanuvchi", "")
    cur.execute("UPDATE users SET balance = balance + ? WHERE user_id = ?", (amount, user_id))
    conn.commit()
    return get_balance(user_id)


def deduct_balance(user_id: int, amount: int) -> bool:
    bal = get_balance(user_id)
    if bal < amount:
        return False
    cur.execute("UPDATE users SET balance = balance - ? WHERE user_id = ?", (amount, user_id))
    conn.commit()
    return True


def is_admin(user_id: int) -> bool:
    return user_id in ADMIN_IDS


# ================= TUGMALAR =================
def main_menu_kb(user_id: int) -> ReplyKeyboardMarkup:
    rows = [
        [KeyboardButton(text="📊 Taqdimot yaratish")],
        [KeyboardButton(text="📄 Referat"), KeyboardButton(text="📝 Mustaqil ish")],
        [KeyboardButton(text="📚 Kurs ishi"), KeyboardButton(text="📁 Ishlarim")],
        [KeyboardButton(text="💰 Hisobim"), KeyboardButton(text="👥 Taklif qilish")],
    ]
    if is_admin(user_id):
        rows.append([KeyboardButton(text="🛠 Admin panel")])
    return ReplyKeyboardMarkup(keyboard=rows, resize_keyboard=True)


def back_kb() -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(keyboard=[[KeyboardButton(text="⬅️ Orqaga")]], resize_keyboard=True)


def balance_inline_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="💳 Balans to‘ldirish", callback_data="topup:start")]
        ]
    )


def admin_topup_kb(req_id: int, target_user_id: int) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="+10 000", callback_data=f"adm_add:{req_id}:{target_user_id}:10000"),
                InlineKeyboardButton(text="+25 000", callback_data=f"adm_add:{req_id}:{target_user_id}:25000"),
            ],
            [
                InlineKeyboardButton(text="+50 000", callback_data=f"adm_add:{req_id}:{target_user_id}:50000"),
                InlineKeyboardButton(text="+100 000", callback_data=f"adm_add:{req_id}:{target_user_id}:100000"),
            ],
            [
                InlineKeyboardButton(text="✍️ Boshqa summa (qo‘lda)", callback_data=f"adm_custom:{req_id}:{target_user_id}"),
                InlineKeyboardButton(text="❌ Rad etish", callback_data=f"adm_reject:{req_id}:{target_user_id}"),
            ],
        ]
    )


def slide_count_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text=f"📦 8 · {8 * PRICE_PER_SLIDE:,}".replace(",", " "), callback_data="sl_cnt:8"),
                InlineKeyboardButton(text=f"📦 10 · {10 * PRICE_PER_SLIDE:,}".replace(",", " "), callback_data="sl_cnt:10"),
            ],
            [
                InlineKeyboardButton(text=f"📦 15 · {15 * PRICE_PER_SLIDE:,}".replace(",", " "), callback_data="sl_cnt:15"),
                InlineKeyboardButton(text=f"📦 20 · {20 * PRICE_PER_SLIDE:,}".replace(",", " "), callback_data="sl_cnt:20"),
            ],
            [InlineKeyboardButton(text="✏️ Boshqa son (5–30)", callback_data="sl_cnt:custom")],
            [
                InlineKeyboardButton(text="⬅️ Orqaga", callback_data="flow:back_topic"),
                InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel"),
            ],
        ]
    )


DESIGNS = [
    ("1. 💎 Klassik", (15, 23, 42), (37, 99, 235), (248, 250, 252)),
    ("2. 🌊 Okean", (8, 47, 73), (14, 165, 233), (240, 249, 255)),
    ("3. 🌅 Shafaq", (67, 20, 7), (234, 88, 12), (255, 247, 237)),
    ("4. 🌿 Tabiat", (20, 83, 45), (22, 163, 74), (240, 253, 244)),
    ("5. 🌙 Tungi", (9, 13, 22), (250, 204, 21), (241, 245, 249)),
    ("6. 👑 Shohona", (49, 24, 98), (234, 179, 8), (250, 245, 255)),
    ("7. ⚪ Minimal", (255, 255, 255), (30, 41, 59), (15, 23, 42)),
]


def design_kb() -> InlineKeyboardMarkup:
    rows = []
    for i in range(0, 6, 2):
        rows.append([
            InlineKeyboardButton(text=DESIGNS[i][0], callback_data=f"dsg:{i}"),
            InlineKeyboardButton(text=DESIGNS[i + 1][0], callback_data=f"dsg:{i + 1}"),
        ])
    rows.append([InlineKeyboardButton(text=DESIGNS[6][0], callback_data="dsg:6")])
    rows.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def images_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(text="🖼 Rasmli", callback_data="img:1"),
                InlineKeyboardButton(text="📄 Rasmsiz", callback_data="img:0"),
            ],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel")],
        ]
    )


ENDINGS = ["1. 🎉 Konfetti", "2. 🎆 Mushakbozlik", "3. 🎈 Sharlar", "4. ✨ Yulduzlar", "5. 💖 Yuraklar", "6. 🫧 Pufakchalar"]


def ending_kb() -> InlineKeyboardMarkup:
    rows = []
    for i in range(0, 6, 2):
        rows.append([
            InlineKeyboardButton(text=ENDINGS[i], callback_data=f"end:{ENDINGS[i]}"),
            InlineKeyboardButton(text=ENDINGS[i + 1], callback_data=f"end:{ENDINGS[i + 1]}"),
        ])
    rows.append([InlineKeyboardButton(text="➖ Kerak emas", callback_data="end:Kerak emas")])
    rows.append([InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def author_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="⏭ O‘tkazib yuborish", callback_data="auth:skip")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel")],
        ]
    )


def confirm_order_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text="✅ Tasdiqlash", callback_data="ord:confirm")],
            [InlineKeyboardButton(text="❌ Bekor qilish", callback_data="flow:cancel")],
        ]
    )


# ================= HOLATLAR (FSM) =================
class Flow(StatesGroup):
    waiting_topic = State()
    waiting_custom_slides = State()
    waiting_author = State()
    waiting_topup_amount = State()
    waiting_admin_custom_amount = State()


# ================= FAYL YARATISH =================
def build_pptx(data: dict, out_path: str) -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    d_idx = data.get("design_idx", 0)
    _, bg_rgb, acc_rgb, txt_rgb = DESIGNS[d_idx]
    bg_color = RGBColor(*bg_rgb)
    acc_color = RGBColor(*acc_rgb)
    txt_color = RGBColor(*txt_rgb)

    total_slides = int(data["pages_or_slides"])
    topic = data["topic"]
    author = data.get("author_title") or ""
    ending = data.get("ending_style") or "Kerak emas"

    for s_num in range(1, total_slides + 1):
        slide = prs.slides.add_slide(blank)
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = bg_color

        bar = slide.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.333), Inches(0.22))
        bar.fill.solid()
        bar.fill.fore_color.rgb = acc_color
        bar.line.fill.background()

        if s_num == 1:
            tb = slide.shapes.add_textbox(Inches(1.1), Inches(1.8), Inches(11.1), Inches(3.6))
            tf = tb.text_frame
            tf.word_wrap = True
            p1 = tf.paragraphs[0]
            p1.text = topic
            p1.font.size = Pt(42)
            p1.font.bold = True
            p1.font.color.rgb = txt_color

            p2 = tf.add_paragraph()
            p2.text = f"\nTil: {data['language']}  |  Slaydlar: {total_slides} ta"
            p2.font.size = Pt(20)
            p2.font.color.rgb = acc_color

            if author:
                p3 = tf.add_paragraph()
                p3.text = f"\nTayyorladi: {author}"
                p3.font.size = Pt(22)
                p3.font.color.rgb = txt_color
        elif s_num == total_slides and ending != "Kerak emas":
            tb = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(10.9), Inches(3.0))
            tf = tb.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.text = f"E’tiboringiz uchun rahmat!\n{ending}"
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = acc_color
        else:
            tb = slide.shapes.add_textbox(Inches(0.9), Inches(0.6), Inches(11.5), Inches(1.0))
            tf = tb.text_frame
            tf.word_wrap = True
            p_title = tf.paragraphs[0]
            p_title.text = f"{s_num}-slayd: {topic} — {s_num - 1}-bo‘lim"
            p_title.font.size = Pt(28)
            p_title.font.bold = True
            p_title.font.color.rgb = acc_color

            body_box = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(11.2), Inches(4.5))
            btf = body_box.text_frame
            btf.word_wrap = True
            bullets = [
                f"«{topic}» mavzusining {s_num - 1}-yo‘nalishi bo‘yicha asosiy tezislar.",
                "Amaliy misollar, tahliliy dalillar va xulosalar ketma-ketligi.",
                f"Vizual bezak: {'Rasmli blok' if data.get('has_images') else 'Matnli blok'}.",
            ]
            for idx_b, b_text in enumerate(bullets):
                bp = btf.paragraphs[0] if idx_b == 0 else btf.add_paragraph()
                bp.text = f"• {b_text}"
                bp.font.size = Pt(22)
                bp.font.color.rgb = txt_color
                bp.space_after = Pt(14)

        num_box = slide.shapes.add_textbox(Inches(11.8), Inches(6.85), Inches(1.2), Inches(0.45))
        ntf = num_box.text_frame
        np = ntf.paragraphs[0]
        np.text = f"{s_num} / {total_slides}"
        np.font.size = Pt(14)
        np.font.color.rgb = txt_color

    prs.save(out_path)


def build_docx(data: dict, out_path: str) -> None:
    doc = Document()
    doc.add_heading(data["topic"], level=0)
    doc.add_paragraph(f"Ish turi: {data['work_type']}")
    doc.add_paragraph(f"Til: {data['language']} | Hajmi: {data['pages_or_slides']} bet")
    if data.get("author_title"):
        doc.add_paragraph(f"Muallif: {data['author_title']}")
    doc.add_heading("Kirish", level=1)
    doc.add_paragraph(f"Ushbu ish «{data['topic']}» mavzusining nazariy va amaliy jihatlarini yoritadi.")
    doc.add_heading("Asosiy qism", level=1)
    doc.add_paragraph(f"«{data['topic']}» bo‘yicha tahlil, asosiy tushunchalar va amaliy tavsiyalar keltirilgan.")
    doc.add_heading("Xulosa va adabiyotlar", level=1)
    doc.add_paragraph("Mavzu yuzasidan umumiy xulosalar shakllantirildi.")
    doc.save(out_path)


# ================= ROUTER & HANDLERS =================
router = Router()

WELCOME_TEXT = (
    "👋 Assalomu alaykum, {name}!\n\n"
    "Men mavzu bo‘yicha tayyor ish yozib beraman:\n\n"
    "📊 Taqdimot — PPTX yoki PDF, 7 xil dizayn\n"
    "📄 Referat va 📝 mustaqil ish — 5–15 bet\n"
    "📚 Kurs ishi — 20–40 bet, boblar bilan\n\n"
    "Hujjatlar Word (DOCX) yoki PDF: titul varaq, mundarija va adabiyotlar bilan.\n"
    "🎬 Animatsiyali yakun — taqdimotingiz oxiriga harakatlanuvchi «Rahmat!» slaydi\n"
    "🌐 Til: o‘zbek, rus, ingliz\n\n"
    "Kerakli bo‘limni pastdagi tugmalardan tanlang."
)


@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    ref_id: Optional[int] = None
    if message.text:
        parts = message.text.split(maxsplit=1)
        if len(parts) == 2 and parts[1].startswith("ref_"):
            raw = parts[1].replace("ref_", "")
            if raw.isdigit():
                ref_id = int(raw)

    user = ensure_user(
        user_id=message.from_user.id,
        full_name=message.from_user.full_name,
        username=message.from_user.username,
        referrer_id=ref_id,
    )
    if ref_id and user["referrer_id"] == ref_id:
        add_balance(ref_id, 2000)

    await message.answer(
        WELCOME_TEXT.format(name=message.from_user.full_name or "Foydalanuvchi"),
        reply_markup=main_menu_kb(message.from_user.id),
    )


@router.message(F.text == "⬅️ Orqaga")
async def go_back(message: Message, state: FSMContext):
    await state.clear()
    await message.answer("Asosiy menyu:", reply_markup=main_menu_kb(message.from_user.id))


@router.message(F.text == "💰 Hisobim")
async def my_balance(message: Message):
    ensure_user(message.from_user.id, message.from_user.full_name, message.from_user.username)
    bal = get_balance(message.from_user.id)
    text = (
        "💰 <b>Hisobingiz</b>\n\n"
        f"💵 Joriy balans: <b>{bal:,} so‘m</b>\n"
        f"📊 1 slayd narxi: <b>{PRICE_PER_SLIDE} so‘m</b>\n\n"
        "Slayd yasash uchun balansingizni to‘ldiring."
    ).replace(",", " ")
    await message.answer(text, parse_mode="HTML", reply_markup=balance_inline_kb())


@router.callback_query(F.data == "topup:start")
async def topup_start(callback: CallbackQuery, state: FSMContext):
    await state.set_state(Flow.waiting_topup_amount)
    text = (
        "💳 <b>Balans to‘ldirish</b>\n\n"
        f"Karta raqami: <code>{CARD_NUMBER}</code>\n"
        f"Karta egasi: <b>{CARD_OWNER}</b>\n\n"
        "Kartaga pul o‘tkazganingizdan so‘ng <b>summani</b> yozib yuboring (masalan: <code>15000</code>).\n"
        "Admin to‘lovni tekshirib, balansingizni qo‘lda to‘ldirib beradi."
    )
    await callback.message.answer(text, parse_mode="HTML", reply_markup=back_kb())
    await callback.answer()


@router.message(Flow.waiting_topup_amount)
async def topup_receive_amount(message: Message, state: FSMContext, bot: Bot):
    cleaned = (message.text or "").replace(" ", "").replace("so‘m", "").replace("so'm", "").strip()
    if not cleaned.isdigit() or int(cleaned) <= 0:
        await message.answer("Iltimos, musbat summa kiriting (masalan: <code>10000</code>).", parse_mode="HTML")
        return

    amount = int(cleaned)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    cur.execute(
        "INSERT INTO topup_requests (user_id, claimed_amount, status, created_at) VALUES (?, ?, 'pending', ?)",
        (message.from_user.id, amount, now),
    )
    conn.commit()
    req_id = cur.lastrowid
    await state.clear()

    await message.answer(
        f"✅ So‘rov yuborildi (Summa: <b>{amount:,} so‘m</b>).\nAdmin tasdiqlashini kuting.".replace(",", " "),
        parse_mode="HTML",
        reply_markup=main_menu_kb(message.from_user.id),
    )

    admin_text = (
        f"💳 <b>Yangi balans to‘ldirish so‘rovi #{req_id}</b>\n\n"
        f"👤 Foydalanuvchi: <b>{message.from_user.full_name}</b>\n"
        f"🆔 ID: <code>{message.from_user.id}</code>\n"
        f"💰 Yozgan summa: <b>{amount:,} so‘m</b>\n"
        f"💵 Hozirgi balansi: <b>{get_balance(message.from_user.id):,} so‘m</b>\n\n"
        "Pastdagi tugmalardan qo‘lda balans qo‘shing:"
    ).replace(",", " ")

    for adm_id in ADMIN_IDS:
        try:
            await bot.send_message(
                adm_id,
                admin_text,
                parse_mode="HTML",
                reply_markup=admin_topup_kb(req_id, message.from_user.id),
            )
        except Exception:
            pass


@router.callback_query(F.data.startswith("adm_add:"))
async def admin_quick_add(callback: CallbackQuery, bot: Bot):
    if not is_admin(callback.from_user.id):
        await callback.answer("Faqat admin uchun!", show_alert=True)
        return

    _, req_id_s, uid_s, amt_s = callback.data.split(":")
    req_id, uid, amt = int(req_id_s), int(uid_s), int(amt_s)

    new_bal = add_balance(uid, amt)
    cur.execute(
        "UPDATE topup_requests SET status = 'approved', admin_id = ?, credited_amount = ? WHERE id = ?",
        (callback.from_user.id, amt, req_id),
    )
    conn.commit()

    await callback.message.edit_text(
        callback.message.text + f"\n\n✅ <b>+{amt:,} so‘m qo‘lda qo‘shildi</b> (Yangi balans: {new_bal:,} so‘m)".replace(",", " "),
        parse_mode="HTML",
    )
    try:
        await bot.send_message(
            uid,
            f"✅ Balansingiz admin tomonidan <b>+{amt:,} so‘m</b>ga to‘ldirildi!\n"
            f"💰 Joriy balans: <b>{new_bal:,} so‘m</b>".replace(",", " "),
            parse_mode="HTML",
        )
    except Exception:
        pass
    await callback.answer("Balans qo‘shildi!")


@router.callback_query(F.data.startswith("adm_custom:"))
async def admin_custom_amount_prompt(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("Faqat admin uchun!", show_alert=True)
        return

    _, req_id_s, uid_s = callback.data.split(":")
    await state.update_data(adm_req_id=int(req_id_s), adm_target_uid=int(uid_s))
    await state.set_state(Flow.waiting_admin_custom_amount)
    await callback.message.answer(
        f"✍️ Foydalanuvchi <code>{uid_s}</code> uchun qo‘lda qo‘shiladigan summani kiriting (masalan: <code>35000</code>):",
        parse_mode="HTML",
    )
    await callback.answer()


@router.message(Flow.waiting_admin_custom_amount)
async def admin_custom_amount_save(message: Message, state: FSMContext, bot: Bot):
    if not is_admin(message.from_user.id):
        return
    cleaned = (message.text or "").replace(" ", "").strip()
    if not cleaned.isdigit() or int(cleaned) <= 0:
        await message.answer("Faqat musbat son kiriting.")
        return

    amt = int(cleaned)
    st_data = await state.get_data()
    req_id = st_data["adm_req_id"]
    uid = st_data["adm_target_uid"]
    await state.clear()

    new_bal = add_balance(uid, amt)
    cur.execute(
        "UPDATE topup_requests SET status = 'approved', admin_id = ?, credited_amount = ? WHERE id = ?",
        (message.from_user.id, amt, req_id),
    )
    conn.commit()

    await message.answer(
        f"✅ Foydalanuvchi <code>{uid}</code> balansiga <b>+{amt:,} so‘m</b> qo‘shildi.\n"
        f"Yangi balans: <b>{new_bal:,} so‘m</b>".replace(",", " "),
        parse_mode="HTML",
        reply_markup=main_menu_kb(message.from_user.id),
    )
    try:
        await bot.send_message(
            uid,
            f"✅ Balansingizga <b>+{amt:,} so‘m</b> qo‘shildi!\n💰 Joriy balans: <b>{new_bal:,} so‘m</b>".replace(",", " "),
            parse_mode="HTML",
        )
    except Exception:
        pass


@router.callback_query(F.data.startswith("adm_reject:"))
async def admin_reject_topup(callback: CallbackQuery, bot: Bot):
    if not is_admin(callback.from_user.id):
        await callback.answer("Faqat admin uchun!", show_alert=True)
        return

    _, req_id_s, uid_s = callback.data.split(":")
    req_id, uid = int(req_id_s), int(uid_s)
    cur.execute("UPDATE topup_requests SET status = 'rejected', admin_id = ? WHERE id = ?", (callback.from_user.id, req_id))
    conn.commit()

    await callback.message.edit_text(callback.message.text + "\n\n❌ <b>Rad etildi</b>", parse_mode="HTML")
    try:
        await bot.send_message(uid, "❌ Balans to‘ldirish so‘rovingiz rad etildi. Qayta urinib ko‘ring.")
    except Exception:
        pass
    await callback.answer("Rad etildi")


@router.message(Command("addbal"))
async def cmd_addbal(message: Message, bot: Bot):
    if not is_admin(message.from_user.id):
        return
    parts = (message.text or "").split()
    if len(parts) != 3 or not parts[1].isdigit() or not parts[2].isdigit():
        await message.answer("Ishlatish: <code>/addbal USER_ID SUMMA</code>", parse_mode="HTML")
        return
    uid, amt = int(parts[1]), int(parts[2])
    new_bal = add_balance(uid, amt)
    await message.answer(
        f"✅ <code>{uid}</code> balansiga <b>+{amt:,} so‘m</b> qo‘shildi. Yangi balans: <b>{new_bal:,} so‘m</b>".replace(",", " "),
        parse_mode="HTML",
    )
    try:
        await bot.send_message(uid, f"✅ Balansingizga <b>+{amt:,} so‘m</b> qo‘shildi! Joriy balans: <b>{new_bal:,} so‘m</b>".replace(",", " "), parse_mode="HTML")
    except Exception:
        pass


@router.message(F.text == "🛠 Admin panel")
async def admin_panel_msg(message: Message):
    if not is_admin(message.from_user.id):
        return
    pending = cur.execute("SELECT id, user_id, claimed_amount, created_at FROM topup_requests WHERE status = 'pending' ORDER BY id DESC LIMIT 10").fetchall()
    if not pending:
        await message.answer("Hozircha kutilayotgan to‘lov so‘rovlari yo‘q.\nQo‘lda balans berish: <code>/addbal USER_ID SUMMA</code>", parse_mode="HTML")
        return
    for req_id, uid, amt, dt in pending:
        await message.answer(
            f"⏳ So‘rov #{req_id} | User: <code>{uid}</code> | Summa: <b>{amt:,} so‘m</b> | {dt}".replace(",", " "),
            parse_mode="HTML",
            reply_markup=admin_topup_kb(req_id, uid),
        )


@router.message(F.text == "👥 Taklif qilish")
async def invite_friends(message: Message, bot: Bot):
    me = await bot.get_me()
    link = f"https://t.me/{me.username}?start=ref_{message.from_user.id}"
    cnt = cur.execute("SELECT COUNT(*) FROM users WHERE referrer_id = ?", (message.from_user.id,)).fetchone()[0]
    await message.answer(
        f"👥 <b>Do‘stlarni taklif qilish</b>\n\n"
        f"Sizning havolangiz:\n<code>{link}</code>\n\n"
        f"Taklif qilinganlar soni: <b>{cnt} ta</b>",
        parse_mode="HTML",
    )


@router.message(F.text == "📁 Ishlarim")
async def my_works(message: Message):
    rows = cur.execute("SELECT id, work_type, title, pages_or_slides, created_at FROM works WHERE user_id = ? ORDER BY id DESC LIMIT 10", (message.from_user.id,)).fetchall()
    if not rows:
        await message.answer("📁 Sizda hali tayyorlangan ishlar yo‘q.")
        return
    lines = ["📁 <b>Oxirgi ishlaringiz:</b>\n"]
    for wid, wtype, title, cnt, dt in rows:
        lines.append(f"#{wid} | {wtype} — <b>{title}</b> ({cnt}) • {dt}")
    await message.answer("\n".join(lines), parse_mode="HTML")


# ================= TAQDIMOT / HUJJAT OQIMI =================
@router.message(F.text.in_({"📊 Taqdimot yaratish", "📄 Referat", "📝 Mustaqil ish", "📚 Kurs ishi"}))
async def start_work_flow(message: Message, state: FSMContext):
    ensure_user(message.from_user.id, message.from_user.full_name, message.from_user.username)
    wtype = message.text
    await state.clear()
    await state.update_data(work_type=wtype, language="O‘zbekcha")
    await state.set_state(Flow.waiting_topic)

    if wtype == "📊 Taqdimot yaratish":
        txt = (
            "✍️ <b>Taqdimot mavzusini yozing</b>\n\n"
            "<i>Masalan: Amir Temur davlati, Fotosintez jarayoni, Sun’iy intellekt va ta’lim</i>\n\n"
            "Mavzu 3–200 belgi bo‘lsin."
        )
    else:
        txt = f"✍️ <b>{wtype} mavzusini yozing</b>\n\nMavzu 3–200 belgi bo‘lsin."

    await message.answer(txt, parse_mode="HTML", reply_markup=back_kb())


@router.message(Flow.waiting_topic)
async def receive_topic(message: Message, state: FSMContext):
    topic = (message.text or "").strip()
    if len(topic) < 3 or len(topic) > 200:
        await message.answer("Mavzu uzunligi 3 dan 200 belgigacha bo‘lishi kerak. Qayta yozing:")
        return

    data = await state.get_data()
    wtype = data["work_type"]
    await state.update_data(topic=topic)

    if wtype == "📊 Taqdimot yaratish":
        info = (
            f"📌 <b>Mavzu:</b> {topic}\n"
            f"🌐 <b>Til:</b> 🇺🇿 O‘zbekcha\n"
            f"💵 <b>1 slayd — {PRICE_PER_SLIDE} so‘m</b>\n\n"
            "📦 <b>Slaydlar sonini tanlang yoki o‘zingiz yozing:</b>"
        )
        await message.answer(info, parse_mode="HTML", reply_markup=slide_count_kb())
    else:
        pages = 10 if wtype in ("📄 Referat", "📝 Mustaqil ish") else 25
        price = pages * 500
        await state.update_data(pages_or_slides=pages, price=price, design_idx=0, has_images=0, ending_style="Kerak emas", author_title="")
        await state.set_state(Flow.waiting_author)
        await message.answer(
            "✍️ <b>Titul uchun ism-familiyangizni yozing</b>\n<i>(Kerak bo‘lmasa «O‘tkazib yuborish» tugmasini bosing)</i>",
            parse_mode="HTML",
            reply_markup=author_kb(),
        )


@router.callback_query(F.data.startswith("sl_cnt:"))
async def pick_slide_count(callback: CallbackQuery, state: FSMContext):
    choice = callback.data.split(":")[1]
    if choice == "custom":
        await state.set_state(Flow.waiting_custom_slides)
        await callback.message.answer("✏️ Slaydlar sonini kiriting (5 dan 30 gacha):")
        await callback.answer()
        return

    cnt = int(choice)
    data = await state.get_data()
    price = cnt * PRICE_PER_SLIDE
    await state.update_data(pages_or_slides=cnt, price=price)

    summary = (
        f"📌 <b>Mavzu:</b> {data['topic']}\n"
        f"🌐 <b>Til:</b> 🇺🇿 O‘zbekcha\n"
        f"📦 <b>Slaydlar:</b> {cnt} ta ({price:,} so‘m)\n\n"
        "🎨 <b>Dizaynni tanlang:</b>"
    ).replace(",", " ")
    await callback.message.answer(summary, parse_mode="HTML", reply_markup=design_kb())
    await callback.answer()


@router.message(Flow.waiting_custom_slides)
async def receive_custom_slides(message: Message, state: FSMContext):
    raw = (message.text or "").strip()
    if not raw.isdigit() or not (5 <= int(raw) <= 30):
        await message.answer("5 dan 30 gacha son kiriting:")
        return

    cnt = int(raw)
    data = await state.get_data()
    price = cnt * PRICE_PER_SLIDE
    await state.update_data(pages_or_slides=cnt, price=price)

    summary = (
        f"📌 <b>Mavzu:</b> {data['topic']}\n"
        f"🌐 <b>Til:</b> 🇺🇿 O‘zbekcha\n"
        f"📦 <b>Slaydlar:</b> {cnt} ta ({price:,} so‘m)\n\n"
        "🎨 <b>Dizaynni tanlang:</b>"
    ).replace(",", " ")
    await message.answer(summary, parse_mode="HTML", reply_markup=design_kb())


@router.callback_query(F.data.startswith("dsg:"))
async def pick_design(callback: CallbackQuery, state: FSMContext):
    idx = int(callback.data.split(":")[1])
    await state.update_data(design_idx=idx)
    txt = (
        "🖼 <b>Slaydlarga mavzuga mos rasmlar qo‘shilsinmi?</b>\n\n"
        "<i>Rasmli taqdimot chiroyliroq chiqadi, tayyorlash biroz uzoqroq davom etadi.</i>"
    )
    await callback.message.answer(txt, parse_mode="HTML", reply_markup=images_kb())
    await callback.answer()


@router.callback_query(F.data.startswith("img:"))
async def pick_images(callback: CallbackQuery, state: FSMContext):
    has_img = int(callback.data.split(":")[1])
    await state.update_data(has_images=has_img)
    txt = (
        "🎬 <b>Taqdimot oxiriga animatsiyali yakun qo‘shilsinmi?</b>\n\n"
        "<i>«E’tiboringiz uchun rahmat!» slaydi. Uslubni tanlang 👇</i>"
    )
    await callback.message.answer(txt, parse_mode="HTML", reply_markup=ending_kb())
    await callback.answer()


@router.callback_query(F.data.startswith("end:"))
async def pick_ending(callback: CallbackQuery, state: FSMContext):
    ending = callback.data.split(":", 1)[1]
    await state.update_data(ending_style=ending)
    await state.set_state(Flow.waiting_author)
    txt = (
        "✍️ <b>Titul slayd uchun ism-familiyangizni yozing</b>\n\n"
        "<i>Masalan: Aliyev Vali, TATU 2-kurs talabasi</i>\n\n"
        "Kerak bo‘lmasa o‘tkazib yuboring."
    )
    await callback.message.answer(txt, parse_mode="HTML", reply_markup=author_kb())
    await callback.answer()


@router.message(Flow.waiting_author)
async def receive_author(message: Message, state: FSMContext):
    await state.update_data(author_title=(message.text or "").strip())
    await show_order_confirmation(message, state, message.from_user.id)


@router.callback_query(F.data == "auth:skip")
async def skip_author(callback: CallbackQuery, state: FSMContext):
    await state.update_data(author_title="")
    await show_order_confirmation(callback.message, state, callback.from_user.id)
    await callback.answer()


async def show_order_confirmation(target_msg: Message, state: FSMContext, user_id: int):
    data = await state.get_data()
    price = int(data.get("price", 2800))
    bal = get_balance(user_id)
    d_name = DESIGNS[data.get("design_idx", 0)][0]
    img_lbl = "bor" if data.get("has_images") else "yo‘q"
    end_lbl = data.get("ending_style", "Kerak emas")
    titul_lbl = data.get("author_title") or "Ko‘rsatilmagan"

    text = (
        "📋 <b>Buyurtmani tasdiqlang</b>\n\n"
        f"📌 <b>Mavzu:</b> {data['topic']}\n"
        f"🌐 <b>Til:</b> 🇺🇿 {data['language']}\n"
        f"📦 <b>Slaydlar:</b> {data['pages_or_slides']} ta\n"
        f"🎨 <b>Dizayn:</b> {d_name}\n"
        f"🖼 <b>Rasmlar:</b> {img_lbl}\n"
        f"🎬 <b>Yakun:</b> {end_lbl}\n"
        f"👤 <b>Titul:</b> {titul_lbl}\n\n"
        f"💵 <b>Narx:</b> {price:,} so‘m\n"
        f"💰 <b>Balans:</b> {bal:,} so‘m"
    ).replace(",", " ")

    if bal < price:
        text += "\n\n⚠️ <b>Balansingiz yetarli emas!</b> Avval «💰 Hisobim» -> «💳 Balans to‘ldirish» orqali balans to‘ldiring."

    await target_msg.answer(text, parse_mode="HTML", reply_markup=confirm_order_kb())


@router.callback_query(F.data == "ord:confirm")
async def confirm_order(callback: CallbackQuery, state: FSMContext):
    data = await state.get_data()
    if not data or "topic" not in data:
        await callback.answer("Buyurtma topilmadi.", show_alert=True)
        return

    price = int(data.get("price", 2800))
    uid = callback.from_user.id

    if not deduct_balance(uid, price):
        bal = get_balance(uid)
        await callback.message.answer(
            f"❌ <b>Balans yetarli emas!</b>\n"
            f"Kerakli summa: <b>{price:,} so‘m</b>\n"
            f"Joriy balans: <b>{bal:,} so‘m</b>\n\n"
            "Balansni to‘ldirish uchun «💰 Hisobim» bo‘limiga kiring.".replace(",", " "),
            parse_mode="HTML",
            reply_markup=balance_inline_kb(),
        )
        await callback.answer("Balans yetarli emas!", show_alert=True)
        return

    await state.clear()
    await callback.message.answer("⏳ Fayl tayyorlanmoqda, biroz kuting...", reply_markup=main_menu_kb(uid))

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    wtype = data["work_type"]

    if wtype == "📊 Taqdimot yaratish":
        file_path = os.path.join(OUT_DIR, f"taqdimot_{uid}_{ts}.pptx")
        build_pptx(data, file_path)
    else:
        file_path = os.path.join(OUT_DIR, f"hujjat_{uid}_{ts}.docx")
        build_docx(data, file_path)

    cur.execute(
        """
        INSERT INTO works (user_id, work_type, title, language, pages_or_slides, design, has_images, ending_style, author_title, price, file_path, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            uid,
            wtype,
            data["topic"],
            data["language"],
            int(data["pages_or_slides"]),
            DESIGNS[data.get("design_idx", 0)][0],
            int(data.get("has_images", 0)),
            data.get("ending_style", "Kerak emas"),
            data.get("author_title", ""),
            price,
            file_path,
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ),
    )
    conn.commit()

    await callback.message.answer_document(
        FSInputFile(file_path),
        caption=f"✅ Tayyor! <b>{data['topic']}</b>\n💰 Balansdan yechildi: <b>{price:,} so‘m</b> (Qoldiq: <b>{get_balance(uid):,} so‘m</b>)".replace(",", " "),
        parse_mode="HTML",
    )
    await callback.answer()


@router.callback_query(F.data.in_({"flow:cancel", "flow:back_topic"}))
async def flow_cancel_or_back(callback: CallbackQuery, state: FSMContext):
    if callback.data == "flow:back_topic":
        await state.set_state(Flow.waiting_topic)
        await callback.message.answer("✍️ Yangi mavzuni kiriting:", reply_markup=back_kb())
    else:
        await state.clear()
        await callback.message.answer("❌ Jarayon bekor qilindi.", reply_markup=main_menu_kb(callback.from_user.id))
    await callback.answer()


async def main():
    bot = Bot(token=BOT_TOKEN)
    dp = Dispatcher(storage=MemoryStorage())
    dp.include_router(router)
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
