import os
os.environ['IS_TELEGRAM_BOT'] = '1'

import asyncio
import time
import configparser
import logging
from typing import List, Dict, Any, Optional

from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import (
    FSInputFile, InlineKeyboardMarkup, InlineKeyboardButton, CallbackQuery,
    BotCommand, BotCommandScopeDefault, BotCommandScopeAllPrivateChats, BotCommandScopeChat
)
from aiogram.client.default import DefaultBotProperties
from aiogram.enums import ParseMode
from aiogram.filters.callback_data import CallbackData
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.utils.keyboard import InlineKeyboardBuilder

# Setup Logging
log_dir = "logs"
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(log_dir, "bot_activity.log"), encoding='utf-8'),
        logging.StreamHandler()
    ]
)

from modules.plugins.scraper import GoogleMapsScraperAsync
from modules.plugins.review_scraper import GoogleMapsReviewScraperAsync

def format_duration(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    h, m = divmod(m, 60)
    if h > 0:
        return f"{h} Jam {m} Menit {s} Detik"
    elif m > 0:
        return f"{m} Menit {s} Detik"
    else:
        return f"{s} Detik"

def make_progress_bar(current: int, total: int, bar_length: int = 10) -> str:
    if total <= 0:
        return "[░░░░░░░░░░] 0%"
    percent = min(max(current / total, 0.0), 1.0)
    filled = int(bar_length * percent)
    bar = '█' * filled + '░' * (bar_length - filled)
    return f"[{bar}] {int(percent * 100)}% ({current}/{total})"

class TelegramProgressTracker:
    def __init__(self, message: types.Message, target_type: str, query: str, limit: int, formats: List[str]):
        self.message = message
        self.target_type = target_type
        self.query = query
        self.limit = limit
        self.formats_str = ", ".join(formats) if formats else "CSV"
        self.last_update = 0.0
        self.min_interval = 2.0  # Limit edit_text calls to once per 2 seconds (Telegram rate limits)
        self.lock = asyncio.Lock()

    async def update(self, current: int, total: int, phase_text: str = ""):
        now = time.time()
        if (now - self.last_update < self.min_interval) and (0 < current < total):
            return

        async with self.lock:
            now = time.time()
            if (now - self.last_update < self.min_interval) and (0 < current < total):
                return
            self.last_update = now

            title_label = "Tempat" if self.target_type == "places" else "Ulasan"
            bar_str = make_progress_bar(current, total)
            text = (
                f"🎯 <b>Target {title_label}:</b> <code>{self.query}</code>\n"
                f"📊 <b>Limit:</b> <code>{self.limit or total}</code>\n"
                f"📄 <b>Format Output:</b> <code>{self.formats_str}</code>\n\n"
                f"⏳ <b>Progress:</b> <code>{bar_str}</code>"
            )
            if phase_text:
                text += f"\n<i>{phase_text}</i>"

            try:
                await self.message.edit_text(text)
            except Exception:
                pass

# Load Configuration
config = configparser.ConfigParser()
config.read('config.ini')

BOT_TOKEN = config.get('Telegram', 'BotToken', fallback=None)
ALLOWED_USERS_STR = config.get('Access', 'AllowedUsers', fallback="")
ALLOWED_USERS = [int(u.strip()) for u in ALLOWED_USERS_STR.split(',') if u.strip().isdigit()]

if not BOT_TOKEN:
    print("Error: BotToken tidak ditemukan di config.ini")
    exit(1)

# Initialize Bot & Dispatcher
bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

OUTPUT_DIR = os.path.join("output", "telegram")
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)
    
# Globals for Playwright & Queue
scrape_semaphore = None

AVAILABLE_FORMATS = {
    "csv": "CSV",
    "json": "JSON",
    "xlsx": "Excel (XLSX)",
    "excel": "Excel (XLSX)",
    "spreadsheet": "Excel (XLSX)",
    "docx": "Word (DOCX)",
    "word": "Word (DOCX)",
    "doc": "Word (DOCX)"
}

ALL_DISPLAY_FORMATS = ["CSV", "Excel (XLSX)", "JSON", "Word (DOCX)"]

# ----------------------------------------------------
# SQLite Database & Multi-Tier Management Engine
# ----------------------------------------------------
import sqlite3
import time
from datetime import date, datetime, timedelta
from typing import Tuple

DB_PATH = "bot_database.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabel Users: menyimpan status tier & tanggal expired
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY,
            tier TEXT DEFAULT 'free',
            expire_date TEXT
        )
    """)
    
    # Tabel Usage Logs: menyimpan riwayat penggunaan per hari & timestamp cooldown
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS usage_logs (
            user_id INTEGER,
            scrape_date TEXT,
            total_count INTEGER DEFAULT 0,
            count_51_200 INTEGER DEFAULT 0,
            count_201_500 INTEGER DEFAULT 0,
            count_500_plus INTEGER DEFAULT 0,
            last_timestamp REAL DEFAULT 0,
            PRIMARY KEY (user_id, scrape_date)
        )
    """)
    
    # Migrasi Otomatis: Tambahkan kolom yang belum ada jika file database lama digunakan
    cursor.execute("PRAGMA table_info(usage_logs)")
    existing_cols = [row[1] for row in cursor.fetchall()]
    
    cols_to_add = {
        "total_count": "INTEGER DEFAULT 0",
        "count_51_200": "INTEGER DEFAULT 0",
        "count_201_500": "INTEGER DEFAULT 0",
        "count_500_plus": "INTEGER DEFAULT 0",
        "last_timestamp": "REAL DEFAULT 0"
    }
    
    for col_name, col_type in cols_to_add.items():
        if col_name not in existing_cols:
            try:
                cursor.execute(f"ALTER TABLE usage_logs ADD COLUMN {col_name} {col_type}")
            except Exception:
                pass

    if "count" in existing_cols and "total_count" in existing_cols:
        try:
            cursor.execute("UPDATE usage_logs SET total_count = count WHERE (total_count IS NULL OR total_count = 0) AND count IS NOT NULL")
        except Exception:
            pass

    conn.commit()
    conn.close()

init_db()

def is_admin(user_id: int) -> bool:
    return user_id in ALLOWED_USERS

def check_access(user_id: int) -> bool:
    return True

def get_user_tier(user_id: int) -> str:
    if is_admin(user_id):
        return "admin"
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT tier, expire_date FROM users WHERE user_id = ?", (user_id,))
    row = cursor.fetchone()
    conn.close()
    
    if not row:
        return "free"
    
    tier, expire_str = row
    if expire_str:
        try:
            exp_date = date.fromisoformat(expire_str)
            if exp_date < date.today():
                return "free"
        except Exception:
            pass
            
    return tier if tier in ("starter", "pro", "admin") else "free"

def set_user_tier(user_id: int, tier: str, days: int = 30):
    exp_date = (date.today() + timedelta(days=days)).isoformat() if days > 0 else None
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO users (user_id, tier, expire_date)
        VALUES (?, ?, ?)
        ON CONFLICT(user_id) DO UPDATE SET tier = ?, expire_date = ?
    """, (user_id, tier, exp_date, tier, exp_date))
    conn.commit()
    conn.close()

def get_user_last_timestamp(user_id: int) -> float:
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT last_timestamp FROM usage_logs WHERE user_id = ? ORDER BY last_timestamp DESC LIMIT 1", (user_id,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row and row[0] else 0.0

def get_cooldown_seconds(user_id: int) -> int:
    tier = get_user_tier(user_id)
    if tier == "admin":
        return 0
    elif tier == "pro":
        return 300  # 5 Menit
    elif tier == "starter":
        return 600  # 10 Menit
    else:
        return 60   # 1 Menit (Free)

def check_cooldown(user_id: int) -> Tuple[bool, int]:
    required_cd = get_cooldown_seconds(user_id)
    if required_cd == 0:
        return True, 0
    
    last_ts = get_user_last_timestamp(user_id)
    if last_ts == 0:
        return True, 0
        
    elapsed = time.time() - last_ts
    remaining = int(required_cd - elapsed)
    if remaining > 0:
        return False, remaining
    return True, 0

def get_starter_quota_today(user_id: int) -> int:
    today_str = date.today().isoformat()
    yesterday_str = (date.today() - timedelta(days=1)).isoformat()
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    cursor.execute("SELECT total_count FROM usage_logs WHERE user_id = ? AND scrape_date = ?", (user_id, today_str))
    r_today = cursor.fetchone()
    today_used = r_today[0] if r_today else 0
    
    cursor.execute("SELECT total_count FROM usage_logs WHERE user_id = ? AND scrape_date = ?", (user_id, yesterday_str))
    r_yesterday = cursor.fetchone()
    yesterday_used = r_yesterday[0] if r_yesterday else 0
    
    conn.close()
    
    # Rollover sisa kuota kemarin max 5 (1 hari)
    unused_yesterday = max(0, 5 - yesterday_used)
    max_quota_today = 5 + unused_yesterday
    
    return max(0, max_quota_today - today_used)

def can_user_scrape(user_id: int, requested_limit: int) -> Tuple[bool, str]:
    tier = get_user_tier(user_id)
    if tier == "admin":
        return True, ""
        
    # 1. Cek Cooldown
    ok_cd, rem_cd = check_cooldown(user_id)
    if not ok_cd:
        m, s = divmod(rem_cd, 60)
        time_str = f"{m}m {s}s" if m > 0 else f"{s}s"
        return False, f"⏳ <b>Cooldown Aktif ({tier.capitalize()} Tier)</b>\n\nSilakan tunggu <b>{time_str}</b> lagi sebelum melakukan scraping berikutnya."
        
    today_str = date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total_count, count_51_200, count_201_500, count_500_plus FROM usage_logs WHERE user_id = ? AND scrape_date = ?", (user_id, today_str))
    row = cursor.fetchone()
    conn.close()
    
    tot_cnt = row[0] if row else 0
    cnt_51_200 = row[1] if row else 0
    cnt_201_500 = row[2] if row else 0
    cnt_500_plus = row[3] if row else 0
    
    if tier == "free":
        if tot_cnt >= 1:
            return False, "⚠️ <b>Batas Kuota Harian Tercapai (Tier Gratis)</b>\n\nAnda telah menggunakan 1x kuota gratis hari ini. Reset besok!"
        if requested_limit > 15:
            return False, "🔒 Tier Gratis maksimal 15 data per scraping."
        return True, ""
        
    elif tier == "starter":
        avail = get_starter_quota_today(user_id)
        if avail <= 0:
            return False, "⚠️ <b>Batas Kuota Harian Tercapai (Tier Starter)</b>\n\nKuota Starter Anda hari ini telah habis. Kuota akan di-reset besok!"
        if requested_limit > 50:
            return False, "🔒 Tier Starter maksimal 50 data per scraping."
        return True, ""
        
    elif tier == "pro":
        if requested_limit > 500:
            if cnt_500_plus >= 1:
                return False, "⚠️ <b>Batas Harian Terlampaui (Tier Pro)</b>\n\nScraping skala besar (>500 data) maksimal 1x per hari."
        elif requested_limit > 200:
            if cnt_201_500 >= 5:
                return False, "⚠️ <b>Batas Harian Terlampaui (Tier Pro)</b>\n\nScraping skala 201-500 data maksimal 5x per hari."
        elif requested_limit > 50:
            if cnt_51_200 >= 10:
                return False, "⚠️ <b>Batas Harian Terlampaui (Tier Pro)</b>\n\nScraping skala 51-200 data maksimal 10x per hari."
        return True, ""
        
    return True, ""

def increment_usage(user_id: int, scraped_limit: int):
    if is_admin(user_id):
        return
        
    today_str = date.today().isoformat()
    now_ts = time.time()
    
    c_51_200 = 1 if 51 <= scraped_limit <= 200 else 0
    c_201_500 = 1 if 201 <= scraped_limit <= 500 else 0
    c_500_plus = 1 if scraped_limit > 500 else 0
    
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO usage_logs (user_id, scrape_date, total_count, count_51_200, count_201_500, count_500_plus, last_timestamp)
        VALUES (?, ?, 1, ?, ?, ?, ?)
        ON CONFLICT(user_id, scrape_date) DO UPDATE SET
            total_count = total_count + 1,
            count_51_200 = count_51_200 + ?,
            count_201_500 = count_201_500 + ?,
            count_500_plus = count_500_plus + ?,
            last_timestamp = ?
    """, (user_id, today_str, c_51_200, c_201_500, c_500_plus, now_ts, c_51_200, c_201_500, c_500_plus, now_ts))
    conn.commit()
    conn.close()

def get_profile_text(user_id: int) -> str:
    tier = get_user_tier(user_id)
    ok_cd, rem_cd = check_cooldown(user_id)
    cd_str = "Tidak Ada (Siap Digunakan)" if ok_cd else f"{rem_cd // 60}m {rem_cd % 60}s"
    
    today_str = date.today().isoformat()
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("SELECT total_count FROM usage_logs WHERE user_id = ? AND scrape_date = ?", (user_id, today_str))
    r = cursor.fetchone()
    
    cursor.execute("SELECT expire_date FROM users WHERE user_id = ?", (user_id,))
    exp_row = cursor.fetchone()
    conn.close()
    
    used_today = r[0] if r else 0
    exp_date_str = exp_row[0] if exp_row and exp_row[0] else "Selamanya (Permanen)"

    tier_label = {
        "admin": "👑 Admin / VIP (Unlimited)",
        "pro": "⚡ Pro Tier",
        "starter": "⭐ Starter Tier",
        "free": "🆓 Free Tier"
    }.get(tier, "🆓 Free Tier")

    quota_info = ""
    if tier == "free":
        quota_info = f"{used_today} / 1x scraping (Max 15 Data/CSV)"
    elif tier == "starter":
        avail = get_starter_quota_today(user_id)
        quota_info = f"{used_today}x terpakai hari ini (Sisa Kuota: {avail}x / Max 50 Data)"
    elif tier == "pro":
        quota_info = "Sesuai Skala Data (s.d. 500+ Data)"
    else:
        quota_info = "Tanpa Batas (Unlimited)"

    return (
        f"👤 <b>Informasi Profil Akun Anda</b>\n\n"
        f"🆔 <b>User ID:</b> <code>{user_id}</code>\n"
        f"🎖️ <b>Status Tier:</b> <b>{tier_label}</b>\n"
        f"📅 <b>Masa Aktif:</b> <code>{exp_date_str}</code>\n"
        f"📊 <b>Pemakaian Kuota:</b> {quota_info}\n"
        f"⏳ <b>Status Cooldown:</b> {cd_str}"
    )

# ----------------------------------------------------
# FSM States & CallbackData Definitions
# ----------------------------------------------------
class ScrapeWizard(StatesGroup):
    waiting_for_places_query = State()
    waiting_for_reviews_url = State()
    selecting_limit = State()
    custom_limit_input = State()
    selecting_formats = State()

class AdminTierWizard(StatesGroup):
    waiting_for_user_id = State()
    selecting_tier = State()
    selecting_duration = State()
    custom_duration_input = State()

class MenuCB(CallbackData, prefix="menu"):
    action: str

class LimitCB(CallbackData, prefix="lim"):
    val: int

class FormatCB(CallbackData, prefix="fmt"):
    name: str

class ActionCB(CallbackData, prefix="act"):
    action: str

class AdminTierCB(CallbackData, prefix="atier"):
    action: str
    name: str

class AdminDurationCB(CallbackData, prefix="adur"):
    days: int

WELCOME_TEXT = (
    "👋 <b>Selamat Datang di Google Maps Pis Scraper Bot!</b>\n\n"
    "Gunakan menu interaktif di bawah ini, atau gunakan perintah langsung:\n\n"
    "<code>/places 10 csv,xlsx Kafe di Jogja</code>\n"
    "<code>/reviews 50 csv,json https://maps...</code>\n\n"
    "<i>Format yang didukung: csv, json, xlsx, docx</i>"
)

# ----------------------------------------------------
# Inline Keyboard Builders
# ----------------------------------------------------
def get_main_menu_keyboard(user_id: int = 0) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text="📍 Scrape Tempat", callback_data=MenuCB(action="places"), style="success")
    builder.button(text="⭐ Scrape Ulasan", callback_data=MenuCB(action="reviews"), style="success")
    builder.button(text="👤 Profil Akun", callback_data=MenuCB(action="profile"), style="primary")
    builder.button(text="⚙️ Info Format", callback_data=MenuCB(action="formats"))
    
    if is_admin(user_id):
        builder.button(text="👑 Manage Tier (Admin)", callback_data=MenuCB(action="admin_settier"), style="primary")
        builder.button(text="💡 Bantuan", callback_data=MenuCB(action="help"), style="danger")
        builder.adjust(2, 1, 1, 1, 1)
    else:
        builder.button(text="💡 Bantuan", callback_data=MenuCB(action="help"), style="danger")
        builder.adjust(2, 1, 1, 1)
        
    return builder.as_markup()

def get_sub_menu_back_keyboard() -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text="← Kembali", callback_data=MenuCB(action="main"))
    return builder.as_markup()

def get_limit_keyboard(user_id: int = 0) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    tier = get_user_tier(user_id)
    
    if tier in ("admin", "pro"):
        builder.button(text="10 Data", callback_data=LimitCB(val=10))
        builder.button(text="25 Data", callback_data=LimitCB(val=25))
        builder.button(text="50 Data", callback_data=LimitCB(val=50))
        builder.button(text="100 Data", callback_data=LimitCB(val=100))
        builder.button(text="✏️ Custom Limit", callback_data=LimitCB(val=-1))
        builder.adjust(2, 2, 1)
    elif tier == "starter":
        builder.button(text="10 Data", callback_data=LimitCB(val=10))
        builder.button(text="25 Data", callback_data=LimitCB(val=25))
        builder.button(text="50 Data (Max Starter)", callback_data=LimitCB(val=50))
        builder.button(text="🔒 100 Data (Pro)", callback_data=LimitCB(val=-99))
        builder.button(text="🔒 Custom (Pro)", callback_data=LimitCB(val=-99))
        builder.adjust(2, 2, 1)
    else: # free
        builder.button(text="10 Data", callback_data=LimitCB(val=10))
        builder.button(text="15 Data (Max Gratis)", callback_data=LimitCB(val=15))
        builder.button(text="🔒 25 Data (Pro)", callback_data=LimitCB(val=-99))
        builder.button(text="🔒 50 Data (Pro)", callback_data=LimitCB(val=-99))
        builder.button(text="🔒 Custom (Pro)", callback_data=LimitCB(val=-99))
        builder.adjust(2, 2, 1)

    builder.button(text="← Kembali", callback_data=ActionCB(action="back_to_input"))
    return builder.as_markup()

def get_format_keyboard(selected_formats: List[str], user_id: int = 0) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    tier = get_user_tier(user_id)

    for fmt in ALL_DISPLAY_FORMATS:
        if fmt == "CSV":
            icon = "✅" if fmt in selected_formats else "❌"
            builder.button(text=f"{icon} {fmt}", callback_data=FormatCB(name=fmt))
        elif fmt == "Excel (XLSX)":
            if tier in ("admin", "pro", "starter"):
                icon = "✅" if fmt in selected_formats else "❌"
                builder.button(text=f"{icon} {fmt}", callback_data=FormatCB(name=fmt))
            else:
                builder.button(text=f"🔒 {fmt} (Pro)", callback_data=FormatCB(name=f"locked_{fmt}"))
        else: # JSON, DOCX
            if tier in ("admin", "pro"):
                icon = "✅" if fmt in selected_formats else "❌"
                builder.button(text=f"{icon} {fmt}", callback_data=FormatCB(name=fmt))
            else:
                builder.button(text=f"🔒 {fmt} (Pro)", callback_data=FormatCB(name=f"locked_{fmt}"))
    
    builder.button(text="🚀 Mulai Scraping Sekarang", callback_data=ActionCB(action="start_scrape"), style="primary")
    builder.button(text="← Kembali", callback_data=ActionCB(action="back_to_limit"))
    builder.adjust(2, 2, 1, 1)
    return builder.as_markup()

def get_result_actions_keyboard(target_type: str) -> InlineKeyboardMarkup:
    builder = InlineKeyboardBuilder()
    builder.button(text="🔄 Scrape Ulang", callback_data=MenuCB(action=target_type), style="primary")
    builder.button(text="🏠 Menu Utama", callback_data=MenuCB(action="main"))
    builder.adjust(2)
    return builder.as_markup()

# ----------------------------------------------------
# Command Parsers & Core Scraper Logics
# ----------------------------------------------------
def parse_command(text: str):
    parts = text.split()
    if len(parts) < 3:
        return None, None, None
    
    try:
        limit = int(parts[1])
    except ValueError:
        return None, None, None

    formats_found = []
    query_start_idx = 2
    
    for i in range(2, len(parts)):
        token = parts[i].lower().strip(',')
        sub_tokens = [t.strip() for t in token.split(',') if t.strip()]
        
        all_valid = True
        for st in sub_tokens:
            if st not in AVAILABLE_FORMATS:
                all_valid = False
                break
                
        if all_valid and len(sub_tokens) > 0:
            for st in sub_tokens:
                formats_found.append(AVAILABLE_FORMATS[st])
            query_start_idx = i + 1
        else:
            break
            
    if len(formats_found) > 0 and query_start_idx < len(parts):
        formats = list(dict.fromkeys(formats_found))
        query = ' '.join(parts[query_start_idx:])
    else:
        formats = ["CSV"]
        query = ' '.join(parts[2:])
        
    return limit, formats, query

async def execute_places_scrape(message: types.Message, limit: int, formats: List[str], query: str, user_id: int):
    formats_str = ", ".join(formats) if formats else "CSV"
    initial_text = (
        f"🎯 <b>Target Tempat:</b> <code>{query}</code>\n"
        f"📊 <b>Limit:</b> <code>{limit}</code>\n"
        f"📄 <b>Format Output:</b> <code>{formats_str}</code>\n\n"
        f"⏳ <b>Progress:</b> <code>[░░░░░░░░░░] 0% (0/{limit})</code>\n"
        f"<i>🔍 Menunggu antrean...</i>"
    )
    status_msg = await message.answer(initial_text)
    try:
        async with scrape_semaphore:
            tracker = TelegramProgressTracker(status_msg, target_type="places", query=query, limit=limit, formats=formats)
            await tracker.update(0, limit, "Memulai pencarian di Google Maps...")
            
            scraper = GoogleMapsScraperAsync(
                query=query, 
                limit=limit, 
                headless=True, 
                output_dir=OUTPUT_DIR,
                progress_callback=tracker.update
            )
            
            start_time = time.time()
            await scraper.scrape()
            saved_files = scraper.save_data(formats=formats)
            end_time = time.time()
            duration_str = format_duration(end_time - start_time)
            
            if not saved_files:
                await message.answer(
                    f"⚠️ <b>Tidak ada data lokasi yang ditemukan</b>\n"
                    f"Keyword: <code>{query}</code>", 
                    reply_markup=get_result_actions_keyboard("places")
                )
                return
                
            for file_path in saved_files:
                document = FSInputFile(file_path)
                await message.answer_document(document=document)
                if not file_path.endswith('.csv'):
                    os.remove(file_path)
                    
            actual_count = len(scraper.results)
            msg_text = (
                f"✅ <b>Scraping selesai!</b>\n\n"
                f"📍 <b>Keyword:</b> <code>{query}</code>\n"
                f"📊 Berhasil mendapatkan <b>{actual_count}</b> data lokasi.\n"
                f"📄 <b>Format File:</b> <code>{formats_str}</code>\n"
                f"⏱️ Waktu Eksekusi: <b>{duration_str}</b>"
            )
            
            # Catat kuota harian & timestamp cooldown untuk user_id
            increment_usage(user_id, limit)

            await message.answer(msg_text, reply_markup=get_result_actions_keyboard("places"))
    except Exception as e:
        error_text = (
            f"❌ <b>Terjadi kesalahan saat ekstraksi data:</b>\n"
            f"<code>{e}</code>"
        )
        await message.answer(error_text, reply_markup=get_result_actions_keyboard("places"))
    finally:
        try:
            await status_msg.delete()
        except:
            pass

async def execute_reviews_scrape(message: types.Message, limit: int, formats: List[str], url: str, user_id: int):
    formats_str = ", ".join(formats) if formats else "CSV"
    initial_text = (
        f"🎯 <b>Target Ulasan:</b> <code>{url}</code>\n"
        f"📊 <b>Limit:</b> <code>{limit}</code>\n"
        f"📄 <b>Format Output:</b> <code>{formats_str}</code>\n\n"
        f"⏳ <b>Progress:</b> <code>[░░░░░░░░░░] 0% (0/{limit})</code>\n"
        f"<i>🔍 Menunggu antrean...</i>"
    )
    status_msg = await message.answer(initial_text)
    try:
        async with scrape_semaphore:
            tracker = TelegramProgressTracker(status_msg, target_type="reviews", query=url, limit=limit, formats=formats)
            await tracker.update(0, limit, "Memulai menarik ulasan...")
            
            scraper = GoogleMapsReviewScraperAsync(
                target_url=url, 
                limit=limit, 
                headless=True, 
                output_dir=OUTPUT_DIR,
                progress_callback=tracker.update
            )
            
            start_time = time.time()
            await scraper.scrape()
            saved_files = scraper.save_data(formats=formats)
            end_time = time.time()
            duration_str = format_duration(end_time - start_time)
            
            if not saved_files:
                await message.answer(
                    f"⚠️ <b>Tidak ada data ulasan yang berhasil ditarik</b>\n"
                    f"Target URL: <code>{url}</code>", 
                    reply_markup=get_result_actions_keyboard("reviews")
                )
                return
                
            for file_path in saved_files:
                document = FSInputFile(file_path)
                await message.answer_document(document=document)
                if not file_path.endswith('.csv'):
                    os.remove(file_path)
                    
            actual_count = len(scraper.results)
            place_name = getattr(scraper, 'place_name', '').strip()
            place_header = f"🏢 <b>{place_name}</b>\n" if place_name else ""
            
            msg_text = (
                f"✅ <b>Scraping Ulasan selesai!</b>\n\n"
                f"{place_header}"
                f"📊 Berhasil mendapatkan <b>{actual_count}</b> ulasan.\n"
                f"📄 <b>Format File:</b> <code>{formats_str}</code>\n"
                f"⏱️ Waktu Eksekusi: <b>{duration_str}</b>"
            )

            # Catat kuota harian & timestamp cooldown untuk user_id
            increment_usage(user_id, limit)

            await message.answer(msg_text, reply_markup=get_result_actions_keyboard("reviews"))
    except Exception as e:
        error_text = (
            f"❌ <b>Terjadi kesalahan saat ekstraksi ulasan:</b>\n"
            f"<code>{e}</code>"
        )
        await message.answer(error_text, reply_markup=get_result_actions_keyboard("reviews"))
    finally:
        try:
            await status_msg.delete()
        except:
            pass

# ----------------------------------------------------
# Command Handlers
# ----------------------------------------------------
@dp.message(Command("start", "menu"))
async def send_welcome(message: types.Message, state: FSMContext):
    await state.clear()
    await message.reply(WELCOME_TEXT, reply_markup=get_main_menu_keyboard(message.from_user.id))

@dp.message(Command("help"))
async def help_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    help_text = (
        "💡 <b>Panduan Penggunaan Scraper Bot</b>\n\n"
        "1. Klik <b>📍 Scrape Tempat</b> atau ketik <code>/places</code> untuk mencari tempat.\n"
        "2. Klik <b>⭐ Scrape Ulasan</b> atau ketik <code>/reviews</code> untuk mengambil ulasan dari URL Maps.\n"
        "3. Ketik <code>/profile</code> untuk mengecek kuota & cooldown akun Anda.\n"
        "4. Format perintah manual langsung:\n\n"
        "📍 <b>Scrape Lokasi (Places)</b>\n"
        "<code>/places &lt;limit&gt; [format] &lt;keyword&gt;</code>\n"
        "<i>Contoh:</i> <code>/places 10 csv,xlsx Hotel di Bandung</code>\n\n"
        "⭐ <b>Scrape Ulasan (Reviews)</b>\n"
        "<code>/reviews &lt;limit&gt; [format] &lt;url_maps&gt;</code>\n"
        "<i>Contoh:</i> <code>/reviews 20 csv https://maps.app.goo.gl/...</code>"
    )
    await message.reply(help_text, reply_markup=get_sub_menu_back_keyboard())

@dp.message(Command("profile", "myinfo"))
async def profile_cmd(message: types.Message, state: FSMContext):
    await state.clear()
    profile_text = get_profile_text(message.from_user.id)
    await message.reply(profile_text, reply_markup=get_sub_menu_back_keyboard())

@dp.message(Command("settier"))
async def set_tier_cmd(message: types.Message):
    if not is_admin(message.from_user.id):
        await message.reply("⛔ Perintah ini hanya dapat dijalankan oleh Admin.")
        return
        
    parts = message.text.split()
    if len(parts) < 3:
        await message.reply("⚠️ Format: <code>/settier &lt;user_id&gt; &lt;free|starter|pro&gt; [jumlah_hari]</code>\n<i>Contoh:</i> <code>/settier 123456789 starter 30</code>")
        return
        
    try:
        target_uid = int(parts[1])
        tier_name = parts[2].lower()
        days = int(parts[3]) if len(parts) >= 4 else 30
        
        if tier_name not in ("free", "starter", "pro"):
            await message.reply("❌ Tier tidak valid. Pilih antara: free, starter, atau pro.")
            return
            
        set_user_tier(target_uid, tier_name, days)
        await message.reply(f"✅ Berhasil mengubah tier pengguna <code>{target_uid}</code> menjadi <b>{tier_name.upper()}</b> selama <b>{days} hari</b>.")
    except Exception as e:
        await message.reply(f"❌ Terjadi kesalahan: {e}")

@dp.message(Command("places"))
async def places_cmd(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    limit, formats, query = parse_command(message.text)
    
    if not query:
        ok, err_msg = can_user_scrape(user_id, 10)
        if not ok:
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
            await message.reply(err_msg, reply_markup=back_kb.as_markup())
            return

        await state.clear()
        await state.set_state(ScrapeWizard.waiting_for_places_query)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
        prompt_msg = await message.answer(
            "📍 <b>Scrape Lokasi (Places)</b>\n\n"
            "Silakan ketik <b>keyword pencarian</b> yang ingin dicari di Google Maps:\n"
            "<i>Contoh: Hotel di Bandung, Apotek di Surabaya</i>",
            reply_markup=back_kb.as_markup()
        )
        await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="places")
        return

    # Perintah langsung (/places 10 csv Hotel)
    ok, err_msg = can_user_scrape(user_id, limit)
    if not ok:
        await message.reply(err_msg)
        return

    await execute_places_scrape(message, limit, formats, query, user_id=user_id)

@dp.message(Command("reviews"))
async def reviews_cmd(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    limit, formats, url = parse_command(message.text)
    
    if not url:
        ok, err_msg = can_user_scrape(user_id, 10)
        if not ok:
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
            await message.reply(err_msg, reply_markup=back_kb.as_markup())
            return

        await state.clear()
        await state.set_state(ScrapeWizard.waiting_for_reviews_url)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
        prompt_msg = await message.answer(
            "⭐ <b>Scrape Ulasan (Reviews)</b>\n\n"
            "Silakan kirimkan <b>URL Google Maps</b> tempat yang ingin diambil ulasannya:\n"
            "<i>Contoh: https://maps.app.goo.gl/...</i>",
            reply_markup=back_kb.as_markup()
        )
        await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="reviews")
        return

    if not url.startswith('http'):
        await message.answer("❌ URL tidak valid. Pastikan dimulai dengan http/https.")
        return

    ok, err_msg = can_user_scrape(user_id, limit)
    if not ok:
        await message.reply(err_msg)
        return

    await execute_reviews_scrape(message, limit, formats, url, user_id=user_id)

# ----------------------------------------------------
# FSM Message Handlers (Interactive Inputs)
# ----------------------------------------------------
@dp.message(ScrapeWizard.waiting_for_places_query)
async def process_places_query(message: types.Message, state: FSMContext):
    query = message.text.strip()
    if not query:
        await message.answer("❌ Keyword tidak boleh kosong. Silakan ketik keyword lokasi:")
        return

    data = await state.get_data()
    prompt_msg_id = data.get("prompt_msg_id")

    if prompt_msg_id:
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=prompt_msg_id)
        except:
            pass
    try:
        await message.delete()
    except:
        pass

    await state.update_data(target_type="places", query=query)
    await state.set_state(ScrapeWizard.selecting_limit)
    await message.answer(
        f"🔍 Keyword: <b>{query}</b>\n\n"
        f"Silakan pilih <b>jumlah data (limit)</b> yang ingin ditarik:",
        reply_markup=get_limit_keyboard(user_id=message.from_user.id)
    )

@dp.message(ScrapeWizard.waiting_for_reviews_url)
async def process_reviews_url(message: types.Message, state: FSMContext):
    url = message.text.strip()
    if not url.startswith('http'):
        await message.answer("❌ URL tidak valid. Pastikan diawali dengan http/https:")
        return

    data = await state.get_data()
    prompt_msg_id = data.get("prompt_msg_id")

    if prompt_msg_id:
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=prompt_msg_id)
        except:
            pass
    try:
        await message.delete()
    except:
        pass

    await state.update_data(target_type="reviews", query=url)
    await state.set_state(ScrapeWizard.selecting_limit)
    await message.answer(
        f"🔗 Target URL: <code>{url}</code>\n\n"
        f"Silakan pilih <b>jumlah ulasan (limit)</b> yang ingin ditarik:",
        reply_markup=get_limit_keyboard(user_id=message.from_user.id)
    )

@dp.message(ScrapeWizard.custom_limit_input)
async def process_custom_limit(message: types.Message, state: FSMContext):
    user_id = message.from_user.id
    try:
        limit = int(message.text.strip())
        if limit <= 0:
            raise ValueError()
    except ValueError:
        await message.answer("❌ Masukkan angka bulat positif yang valid (misal: 15, 50, 150):")
        return

    tier = get_user_tier(user_id)
    if tier == "free":
        limit = min(limit, 15)
    elif tier == "starter":
        limit = min(limit, 50)

    data = await state.get_data()
    custom_prompt_id = data.get("custom_prompt_id")

    if custom_prompt_id:
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=custom_prompt_id)
        except:
            pass
    try:
        await message.delete()
    except:
        pass

    if tier == "free":
        selected_formats = ["CSV"]
    elif tier == "starter":
        selected_formats = [f for f in data.get("formats", ["CSV"]) if f in ("CSV", "Excel (XLSX)")] or ["CSV"]
    else:
        selected_formats = data.get("formats", ["CSV"])

    await state.update_data(limit=limit, formats=selected_formats)
    await state.set_state(ScrapeWizard.selecting_formats)
    
    target_type = data.get("target_type")
    query = data.get("query")
    title_label = "Tempat" if target_type == "places" else "Ulasan"
    
    await message.answer(
        f"🎯 Target {title_label}: <b>{query}</b>\n"
        f"📊 Limit: <b>{limit}</b>\n\n"
        f"Silakan pilih <b>format file output</b> (klik untuk toggle ON/OFF):",
        reply_markup=get_format_keyboard(selected_formats, user_id=user_id)
    )

# ----------------------------------------------------
# Callback Query Handlers (Inline Button Clicks)
# ----------------------------------------------------
@dp.callback_query(MenuCB.filter())
async def handle_menu_cb(query: CallbackQuery, callback_data: MenuCB, state: FSMContext):
    user_id = query.from_user.id
    action = callback_data.action
    await query.answer()

    if action == "places":
        ok, err_msg = can_user_scrape(user_id, 10)
        if not ok:
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
            await query.message.edit_text(err_msg, reply_markup=back_kb.as_markup())
            return

        await state.clear()
        await state.set_state(ScrapeWizard.waiting_for_places_query)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
        prompt_text = (
            "📍 <b>Scrape Lokasi (Places)</b>\n\n"
            "Silakan ketik <b>keyword pencarian</b> yang ingin dicari di Google Maps:\n"
            "<i>Contoh: Hotel di Bandung, Restoran di Jakarta</i>"
        )
        try:
            prompt_msg = await query.message.edit_text(prompt_text, reply_markup=back_kb.as_markup())
        except Exception:
            prompt_msg = await query.message.answer(prompt_text, reply_markup=back_kb.as_markup())
        await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="places")

    elif action == "reviews":
        ok, err_msg = can_user_scrape(user_id, 10)
        if not ok:
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
            await query.message.edit_text(err_msg, reply_markup=back_kb.as_markup())
            return

        await state.clear()
        await state.set_state(ScrapeWizard.waiting_for_reviews_url)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
        prompt_text = (
            "⭐ <b>Scrape Ulasan (Reviews)</b>\n\n"
            "Silakan kirimkan <b>URL Google Maps</b> tempat yang ingin diambil ulasannya:\n"
            "<i>Contoh: https://maps.app.goo.gl/...</i>"
        )
        try:
            prompt_msg = await query.message.edit_text(prompt_text, reply_markup=back_kb.as_markup())
        except Exception:
            prompt_msg = await query.message.answer(prompt_text, reply_markup=back_kb.as_markup())
        await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="reviews")

    elif action == "profile":
        profile_text = get_profile_text(user_id)
        await query.message.edit_text(profile_text, reply_markup=get_sub_menu_back_keyboard())

    elif action == "admin_settier":
        if not is_admin(user_id):
            await query.answer("⛔ Akses Khusus Admin", show_alert=True)
            return

        await state.clear()
        await state.set_state(AdminTierWizard.waiting_for_user_id)
        
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=MenuCB(action="main"))
        
        prompt_text = (
            "👑 <b>Kelola Tier Pengguna (Admin Wizard)</b>\n\n"
            "Silakan ketik <b>Telegram User ID</b> pengguna yang ingin diubah tier-nya:\n"
            "<i>Contoh: 123456789</i>"
        )
        try:
            prompt_msg = await query.message.edit_text(prompt_text, reply_markup=back_kb.as_markup())
        except Exception:
            prompt_msg = await query.message.answer(prompt_text, reply_markup=back_kb.as_markup())
        await state.update_data(prompt_msg_id=prompt_msg.message_id)

    elif action == "formats":
        info_text = (
            "⚙️ <b>Format File Output Yang Didukung:</b>\n\n"
            "• <b>CSV</b>: Format standar ringan untuk spreadsheet / database (Semua Tier).\n"
            "• <b>Excel (XLSX)</b>: Format spreadsheet dengan styling tabel rapi (Tier Starter/Pro/Admin).\n"
            "• <b>JSON</b>: Format data terstruktur untuk developer / API (Tier Pro/Admin).\n"
            "• <b>Word (DOCX)</b>: Format dokumen laporan terformat (Tier Pro/Admin)."
        )
        await query.message.edit_text(info_text, reply_markup=get_sub_menu_back_keyboard())

    elif action == "help":
        help_text = (
            "💡 <b>Panduan Penggunaan Scraper Bot</b>\n\n"
            "1. Klik <b>📍 Scrape Tempat</b> atau ketik <code>/places</code> untuk mencari tempat.\n"
            "2. Klik <b>⭐ Scrape Ulasan</b> atau ketik <code>/reviews</code> untuk mengambil ulasan dari URL Maps.\n"
            "3. Klik <b>👤 Profil Akun</b> untuk mengecek status tier, kuota, & cooldown Anda.\n"
            "4. Format perintah manual langsung:\n\n"
            "📍 <b>Scrape Lokasi (Places)</b>\n"
            "<code>/places &lt;limit&gt; [format] &lt;keyword&gt;</code>\n"
            "<i>Contoh:</i> <code>/places 10 csv,xlsx Hotel di Bandung</code>\n\n"
            "⭐ <b>Scrape Ulasan (Reviews)</b>\n"
            "<code>/reviews &lt;limit&gt; [format] &lt;url_maps&gt;</code>\n"
            "<i>Contoh:</i> <code>/reviews 20 csv https://maps.app.goo.gl/...</code>"
        )
        await query.message.edit_text(help_text, reply_markup=get_sub_menu_back_keyboard())

    elif action == "main":
        await state.clear()
        await query.message.edit_text(WELCOME_TEXT, reply_markup=get_main_menu_keyboard(user_id))

# ----------------------------------------------------
# Admin Tier Interactive Wizard Handlers
# ----------------------------------------------------
@dp.message(AdminTierWizard.waiting_for_user_id)
async def process_admin_user_id(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        return
        
    txt = message.text.strip()
    try:
        target_uid = int(txt)
        if target_uid <= 0:
            raise ValueError()
    except ValueError:
        await message.answer("❌ User ID harus berupa angka bulat positif yang valid. Silakan coba lagi:")
        return

    data = await state.get_data()
    prompt_msg_id = data.get("prompt_msg_id")

    if prompt_msg_id:
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=prompt_msg_id)
        except:
            pass
    try:
        await message.delete()
    except:
        pass

    await state.update_data(target_uid=target_uid)
    await state.set_state(AdminTierWizard.selecting_tier)
    
    current_tier = get_user_tier(target_uid)
    
    tier_kb = InlineKeyboardBuilder()
    tier_kb.button(text="🆓 Free Tier", callback_data=AdminTierCB(action="set", name="free"))
    tier_kb.button(text="⭐ Starter Tier", callback_data=AdminTierCB(action="set", name="starter"))
    tier_kb.button(text="⚡ Pro Tier", callback_data=AdminTierCB(action="set", name="pro"))
    tier_kb.button(text="← Kembali", callback_data=AdminTierCB(action="back_uid", name=""))
    tier_kb.adjust(1, 2, 1)

    await message.answer(
        f"👑 <b>Kelola Tier Pengguna</b>\n\n"
        f"🆔 Target User ID: <code>{target_uid}</code>\n"
        f"🎖️ Status Tier Saat Ini: <b>{current_tier.upper()}</b>\n\n"
        f"Silakan pilih <b>Tier Baru</b> yang ingin diberikan:",
        reply_markup=tier_kb.as_markup()
    )

@dp.callback_query(AdminTierCB.filter())
async def handle_admin_tier_cb(query: CallbackQuery, callback_data: AdminTierCB, state: FSMContext):
    if not is_admin(query.from_user.id):
        await query.answer("⛔ Akses Khusus Admin", show_alert=True)
        return

    action = callback_data.action
    name = callback_data.name
    await query.answer()

    if action == "back_uid":
        await state.clear()
        await state.set_state(AdminTierWizard.waiting_for_user_id)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=MenuCB(action="main"))
        prompt_text = (
            "👑 <b>Kelola Tier Pengguna (Admin Wizard)</b>\n\n"
            "Silakan ketik <b>Telegram User ID</b> pengguna yang ingin diubah tier-nya:\n"
            "<i>Contoh: 123456789</i>"
        )
        prompt_msg = await query.message.edit_text(prompt_text, reply_markup=back_kb.as_markup())
        await state.update_data(prompt_msg_id=prompt_msg.message_id)

    elif action == "set":
        data = await state.get_data()
        target_uid = data.get("target_uid")

        if name == "free":
            set_user_tier(target_uid, "free", 0)
            await state.clear()
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="🏠 Menu Utama", callback_data=MenuCB(action="main"))
            await query.message.edit_text(
                f"✅ <b>Berhasil Mengubah Tier!</b>\n\n"
                f"🆔 User ID: <code>{target_uid}</code>\n"
                f"🎖️ Tier Baru: <b>FREE (Gratis)</b>\n"
                f"📅 Masa Aktif: Permanen",
                reply_markup=back_kb.as_markup()
            )
        else:
            await state.update_data(target_tier=name)
            await state.set_state(AdminTierWizard.selecting_duration)
            
            dur_kb = InlineKeyboardBuilder()
            dur_kb.button(text="7 Hari", callback_data=AdminDurationCB(days=7))
            dur_kb.button(text="30 Hari", callback_data=AdminDurationCB(days=30))
            dur_kb.button(text="90 Hari", callback_data=AdminDurationCB(days=90))
            dur_kb.button(text="365 Hari", callback_data=AdminDurationCB(days=365))
            dur_kb.button(text="✏️ Custom Hari", callback_data=AdminDurationCB(days=-1))
            dur_kb.button(text="← Kembali", callback_data=AdminTierCB(action="back_tier_select", name=""))
            dur_kb.adjust(2, 2, 1, 1)

            await query.message.edit_text(
                f"👑 <b>Pilih Masa Aktif (Durasi)</b>\n\n"
                f"🆔 User ID: <code>{target_uid}</code>\n"
                f"🎖️ Tier Baru: <b>{name.upper()}</b>\n\n"
                f"Silakan pilih durasi masa aktif langganan:",
                reply_markup=dur_kb.as_markup()
            )

    elif action == "back_tier_select":
        data = await state.get_data()
        target_uid = data.get("target_uid")
        current_tier = get_user_tier(target_uid)
        
        await state.set_state(AdminTierWizard.selecting_tier)
        tier_kb = InlineKeyboardBuilder()
        tier_kb.button(text="🆓 Free Tier", callback_data=AdminTierCB(action="set", name="free"))
        tier_kb.button(text="⭐ Starter Tier", callback_data=AdminTierCB(action="set", name="starter"))
        tier_kb.button(text="⚡ Pro Tier", callback_data=AdminTierCB(action="set", name="pro"))
        tier_kb.button(text="← Kembali", callback_data=AdminTierCB(action="back_uid", name=""))
        tier_kb.adjust(1, 2, 1)

        await query.message.edit_text(
            f"👑 <b>Kelola Tier Pengguna</b>\n\n"
            f"🆔 Target User ID: <code>{target_uid}</code>\n"
            f"🎖️ Status Tier Saat Ini: <b>{current_tier.upper()}</b>\n\n"
            f"Silakan pilih <b>Tier Baru</b> yang ingin diberikan:",
            reply_markup=tier_kb.as_markup()
        )

@dp.callback_query(AdminDurationCB.filter())
async def handle_admin_duration_cb(query: CallbackQuery, callback_data: AdminDurationCB, state: FSMContext):
    if not is_admin(query.from_user.id):
        await query.answer("⛔ Akses Khusus Admin", show_alert=True)
        return

    days = callback_data.days
    await query.answer()

    if days == -1:
        await state.set_state(AdminTierWizard.custom_duration_input)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=AdminTierCB(action="back_tier_select", name=""))
        prompt_msg = await query.message.answer(
            "✏️ Silakan ketik jumlah <b>hari masa aktif</b> (angka bulat, misal 14, 45, 180):",
            reply_markup=back_kb.as_markup()
        )
        await state.update_data(custom_dur_prompt_id=prompt_msg.message_id)
        try:
            await query.message.delete()
        except:
            pass
        return

    data = await state.get_data()
    target_uid = data.get("target_uid")
    target_tier = data.get("target_tier")

    set_user_tier(target_uid, target_tier, days)
    await state.clear()
    
    exp_date_str = (date.today() + timedelta(days=days)).isoformat()
    back_kb = InlineKeyboardBuilder()
    back_kb.button(text="🏠 Menu Utama", callback_data=MenuCB(action="main"))
    
    await query.message.edit_text(
        f"✅ <b>Berhasil Mengubah Tier Pengguna!</b>\n\n"
        f"🆔 User ID: <code>{target_uid}</code>\n"
        f"🎖️ Tier Baru: <b>{target_tier.upper()}</b>\n"
        f"📅 Durasi: <b>{days} Hari</b> (s.d. {exp_date_str})",
        reply_markup=back_kb.as_markup()
    )

@dp.message(AdminTierWizard.custom_duration_input)
async def process_custom_duration_input(message: types.Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        return

    try:
        days = int(message.text.strip())
        if days <= 0:
            raise ValueError()
    except ValueError:
        await message.answer("❌ Masukkan angka bulat positif jumlah hari yang valid (misal: 14, 45, 180):")
        return

    data = await state.get_data()
    custom_dur_prompt_id = data.get("custom_dur_prompt_id")

    if custom_dur_prompt_id:
        try:
            await bot.delete_message(chat_id=message.chat.id, message_id=custom_dur_prompt_id)
        except:
            pass
    try:
        await message.delete()
    except:
        pass

    target_uid = data.get("target_uid")
    target_tier = data.get("target_tier")

    set_user_tier(target_uid, target_tier, days)
    await state.clear()
    
    exp_date_str = (date.today() + timedelta(days=days)).isoformat()
    back_kb = InlineKeyboardBuilder()
    back_kb.button(text="🏠 Menu Utama", callback_data=MenuCB(action="main"))
    
    await message.answer(
        f"✅ <b>Berhasil Mengubah Tier Pengguna!</b>\n\n"
        f"🆔 User ID: <code>{target_uid}</code>\n"
        f"🎖️ Tier Baru: <b>{target_tier.upper()}</b>\n"
        f"📅 Durasi: <b>{days} Hari</b> (s.d. {exp_date_str})",
        reply_markup=back_kb.as_markup()
    )

@dp.callback_query(LimitCB.filter())
async def handle_limit_cb(query: CallbackQuery, callback_data: LimitCB, state: FSMContext):
    user_id = query.from_user.id
    val = callback_data.val

    if val == -99:
        tier = get_user_tier(user_id)
        if tier == "starter":
            await query.answer("🔒 Tier Starter maksimal 50 data per scraping. Upgrade ke Pro/Admin untuk limit hingga 500+ data.", show_alert=True)
        else:
            await query.answer("🔒 Tier Gratis maksimal 15 data per scraping. Upgrade ke Starter/Pro/Admin untuk limit lebih tinggi.", show_alert=True)
        return

    await query.answer()

    if val == -1:
        await state.set_state(ScrapeWizard.custom_limit_input)
        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_limit_from_custom"))
        prompt_msg = await query.message.answer(
            "✏️ Silakan ketik jumlah <b>limit custom</b> yang diinginkan (angka bulat):",
            reply_markup=back_kb.as_markup()
        )
        await state.update_data(custom_prompt_id=prompt_msg.message_id)
        try:
            await query.message.delete()
        except:
            pass
        return

    tier = get_user_tier(user_id)
    if tier == "free":
        val = min(val, 15)
    elif tier == "starter":
        val = min(val, 50)

    data = await state.get_data()
    if tier == "free":
        selected_formats = ["CSV"]
    elif tier == "starter":
        selected_formats = [f for f in data.get("formats", ["CSV"]) if f in ("CSV", "Excel (XLSX)")] or ["CSV"]
    else:
        selected_formats = data.get("formats", ["CSV"])

    await state.update_data(limit=val, formats=selected_formats)
    await state.set_state(ScrapeWizard.selecting_formats)

    target_type = data.get("target_type")
    q_str = data.get("query")
    title_label = "Tempat" if target_type == "places" else "Ulasan"

    await query.message.edit_text(
        f"🎯 Target {title_label}: <b>{q_str}</b>\n"
        f"📊 Limit: <b>{val}</b>\n\n"
        f"Silakan pilih <b>format file output</b> (klik untuk toggle ON/OFF):",
        reply_markup=get_format_keyboard(selected_formats, user_id=user_id)
    )

@dp.callback_query(FormatCB.filter())
async def handle_format_cb(query: CallbackQuery, callback_data: FormatCB, state: FSMContext):
    user_id = query.from_user.id
    fmt_name = callback_data.name

    if fmt_name.startswith("locked_"):
        await query.answer("🔒 Fitur format ini terkunci untuk tier Anda. Upgrade ke Starter/Pro untuk mengaktifkan.", show_alert=True)
        return

    tier = get_user_tier(user_id)
    if tier == "free" and fmt_name != "CSV":
        await query.answer("🔒 Tier Gratis hanya mendukung format CSV.", show_alert=True)
        return
    elif tier == "starter" and fmt_name not in ("CSV", "Excel (XLSX)"):
        await query.answer("🔒 Tier Starter mendukung format CSV & Excel (XLSX). Upgrade ke Pro untuk JSON & Word.", show_alert=True)
        return

    data = await state.get_data()
    selected_formats = data.get("formats", ["CSV"])

    if fmt_name in selected_formats:
        if len(selected_formats) > 1:
            selected_formats.remove(fmt_name)
            await query.answer(f"Format {fmt_name} dihapus.")
        else:
            await query.answer("⚠️ Minimal 1 format file harus dipilih!", show_alert=True)
            return
    else:
        selected_formats.append(fmt_name)
        await query.answer(f"Format {fmt_name} ditambahkan.")

    await state.update_data(formats=selected_formats)
    await query.message.edit_reply_markup(reply_markup=get_format_keyboard(selected_formats, user_id=user_id))

@dp.callback_query(ActionCB.filter())
async def handle_action_cb(query: CallbackQuery, callback_data: ActionCB, state: FSMContext):
    user_id = query.from_user.id
    action = callback_data.action
    await query.answer()

    if action in ("cancel", "back_to_main"):
        data = await state.get_data()
        prompt_msg_id = data.get("prompt_msg_id") or data.get("custom_prompt_id")
        if prompt_msg_id:
            try:
                await bot.delete_message(chat_id=query.message.chat.id, message_id=prompt_msg_id)
            except:
                pass
        await state.clear()
        try:
            await query.message.edit_text(WELCOME_TEXT, reply_markup=get_main_menu_keyboard(user_id))
        except Exception:
            await query.message.answer(WELCOME_TEXT, reply_markup=get_main_menu_keyboard(user_id))

    elif action == "back_to_input":
        data = await state.get_data()
        target_type = data.get("target_type", "places")
        await state.clear()

        try:
            await query.message.delete()
        except:
            pass

        back_kb = InlineKeyboardBuilder()
        back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))

        if target_type == "places":
            await state.set_state(ScrapeWizard.waiting_for_places_query)
            prompt_msg = await query.message.answer(
                "📍 <b>Scrape Lokasi (Places)</b>\n\n"
                "Silakan ketik <b>keyword pencarian</b> yang ingin dicari di Google Maps:\n"
                "<i>Contoh: Hotel di Bandung, Restoran di Jakarta</i>",
                reply_markup=back_kb.as_markup()
            )
            await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="places")
        else:
            await state.set_state(ScrapeWizard.waiting_for_reviews_url)
            prompt_msg = await query.message.answer(
                "⭐ <b>Scrape Ulasan (Reviews)</b>\n\n"
                "Silakan kirimkan <b>URL Google Maps</b> tempat yang ingin diambil ulasannya:\n"
                "<i>Contoh: https://maps.app.goo.gl/...</i>",
                reply_markup=back_kb.as_markup()
            )
            await state.update_data(prompt_msg_id=prompt_msg.message_id, target_type="reviews")

    elif action in ("back_to_limit", "back_to_limit_from_custom"):
        data = await state.get_data()
        target_type = data.get("target_type", "places")
        q_str = data.get("query", "")
        title_label = "Tempat" if target_type == "places" else "Ulasan"

        if action == "back_to_limit_from_custom":
            custom_prompt_id = data.get("custom_prompt_id")
            if custom_prompt_id:
                try:
                    await bot.delete_message(chat_id=query.message.chat.id, message_id=custom_prompt_id)
                except:
                    pass

        await state.set_state(ScrapeWizard.selecting_limit)
        try:
            await query.message.edit_text(
                f"🎯 Target {title_label}: <b>{q_str}</b>\n\n"
                f"Silakan pilih <b>jumlah data (limit)</b> yang ingin ditarik:",
                reply_markup=get_limit_keyboard(user_id=user_id)
            )
        except Exception:
            await query.message.answer(
                f"🎯 Target {title_label}: <b>{q_str}</b>\n\n"
                f"Silakan pilih <b>jumlah data (limit)</b> yang ingin ditarik:",
                reply_markup=get_limit_keyboard(user_id=user_id)
            )

    elif action == "start_scrape":
        data = await state.get_data()
        target_type = data.get("target_type")
        q_str = data.get("query")
        limit = data.get("limit", 10)
        formats = data.get("formats", ["CSV"])

        ok, err_msg = can_user_scrape(user_id, limit)
        if not ok:
            back_kb = InlineKeyboardBuilder()
            back_kb.button(text="← Kembali", callback_data=ActionCB(action="back_to_main"))
            await query.message.answer(err_msg, reply_markup=back_kb.as_markup())
            return

        await state.clear()
        try:
            await query.message.delete()
        except:
            pass

        if target_type == "places":
            await execute_places_scrape(query.message, limit, formats, q_str, user_id=user_id)
        elif target_type == "reviews":
            await execute_reviews_scrape(query.message, limit, formats, q_str, user_id=user_id)

# ----------------------------------------------------
# Telegram Bot Command Registration (Autocomplete & Scopes)
# ----------------------------------------------------
async def setup_bot_commands(bot: Bot):
    # 1. Menu Autocomplete Default untuk Seluruh Pengguna Publik (User Biasa)
    user_commands = [
        BotCommand(command="start", description="👋 Mulai Bot & Tampilkan Menu Utama"),
        BotCommand(command="places", description="📍 Scrape Data Tempat Google Maps"),
        BotCommand(command="reviews", description="⭐ Scrape Data Ulasan Google Maps"),
        BotCommand(command="profile", description="👤 Cek Status Tier, Kuota & Cooldown"),
        BotCommand(command="help", description="💡 Panduan Penggunaan Bot"),
    ]
    
    # Set menu untuk scope default dan semua private chats
    await bot.set_my_commands(commands=user_commands, scope=BotCommandScopeDefault())
    await bot.set_my_commands(commands=user_commands, scope=BotCommandScopeAllPrivateChats())

    # 2. Menu Autocomplete Khusus Admin (Owner & Whitelist)
    # Menampilkan perintah umum PLUS perintah khusus Admin (/settier)
    admin_commands = user_commands + [
        BotCommand(command="settier", description="👑 [Admin] Set Tier User (/settier id tier hari)"),
    ]
    
    for admin_id in ALLOWED_USERS:
        try:
            await bot.set_my_commands(commands=admin_commands, scope=BotCommandScopeChat(chat_id=admin_id))
        except Exception as e:
            logging.warning(f"Gagal memasang command scope khusus admin ID {admin_id}: {e}")

# ----------------------------------------------------
# Main Execution Entrypoint
# ----------------------------------------------------
async def main():
    global scrape_semaphore
    
    logging.info(f"Mempersiapkan Bot. Allowed Users: {ALLOWED_USERS}")
    scrape_semaphore = asyncio.Semaphore(2)
    
    await setup_bot_commands(bot)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
