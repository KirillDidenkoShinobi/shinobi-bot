import asyncio
import json
import os
import time
import random
import string
from datetime import datetime

from aiogram import Bot, Dispatcher, F, BaseMiddleware
from aiogram.filters import Command
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import (
    Message,
    ChatMemberUpdated,
    ChatJoinRequest,
    InlineKeyboardMarkup,
    InlineKeyboardButton,
    CallbackQuery,
    FSInputFile,
    ReplyKeyboardMarkup,
    KeyboardButton,
)

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

from schedule_system import (
    init_schedule_system,
    schedule_router,
    show_schedule_menu,
)


# =========================================================
# НАСТРОЙКИ
# =========================================================

BOT_TOKEN = "8935701609:AAHzwwEsgoLjnMqRS8JWFtcV-tu4MH0HXFE"

# ID группы Shinobi Team
CHAT_ID = -1004368701081

# ID администраторов, которым доступна админ-панель
OWNER_IDS = [1197818066]
ADMIN_IDS = OWNER_IDS.copy()

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "data.json")
EXCEL_FILE = os.path.join(BASE_DIR, "shinobi_links_log.xlsx")
ADMIN_STAFF_CHAT_ID = -1004462686385
ERROR_DELETE_AFTER = 15


# =========================================================
# БОТ
# =========================================================

# PythonAnywhere Free: исходящие подключения должны идти через их HTTP-прокси.
# Сначала берём адрес из переменных окружения PythonAnywhere; fallback оставлен
# на стандартный адрес их прокси.
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

# Управление доступом к командам в группах.
# Состояние хранится в data.json, поэтому не сбрасывается после перезапуска.
COMMAND_BUTTON_TEXTS = {
    "🔗 Создать ссылку", "📊 Моя статистика", "🔗 Мои ссылки",
    "👥 Приглашённые", "🏆 Топ", "📋 Мой график",
    "🏖 Запросить выходной", "🗓 График / заполнить",
    "⚙️ Админ-панель", "🗓 Управление графиком", "📥 Скачать Excel",
    "📊 Общая статистика", "📅 Excel с графиками", "🏖 Excel с выходными",
}
BLOCK_NOTICE_COOLDOWN = 4 * 60 * 60
_block_notice_last = {}

# Команды, которые обязательно подчиняются групповому запрету.
# В том числе графики: в запрещённой группе они не должны обходить middleware.
BLOCKABLE_GROUP_COMMANDS = {
    ".график", "график",
    ".мой график", "мой график",
}

def _looks_like_bot_command(text: str) -> bool:
    text = (text or "").strip()
    normalized = " ".join(text.casefold().split())
    return (
        normalized in BLOCKABLE_GROUP_COMMANDS
        or text.startswith("/")
        or text.startswith(".")
        or text in COMMAND_BUTTON_TEXTS
    )

def _group_is_installed(chat_id: int) -> bool:
    if "data" not in globals():
        return False
    return str(chat_id) in {str(x) for x in data.get("installed_groups", [])}

def _chat_commands_blocked(chat_id: int) -> bool:
    settings = data.get("group_command_access", {}) if "data" in globals() else {}
    return settings.get(str(chat_id), "allow") == "deny"

class GroupCommandAccessMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        if isinstance(event, Message) and event.chat.type in {"group", "supergroup"}:
            text = (event.text or "").strip()
            is_install = text.casefold() in {".установить", "/установить", "/install"}

            # До ручной установки бот не обслуживает команды в группе.
            # Команда .установить — единственное исключение.
            if not _group_is_installed(event.chat.id):
                if not is_install and _looks_like_bot_command(text):
                    return
            elif _chat_commands_blocked(event.chat.id) and _looks_like_bot_command(text) and not is_install:
                key = (event.chat.id, event.from_user.id if event.from_user else 0)
                now = time.time()
                if now - _block_notice_last.get(key, 0) >= BLOCK_NOTICE_COOLDOWN:
                    _block_notice_last[key] = now
                    code = "".join(random.choices(string.ascii_letters + string.digits, k=25))
                    try:
                        notice = await event.reply(f"error 404\n{code}")
                        async def _delete_error_message(chat_id: int, message_id: int):
                            await asyncio.sleep(ERROR_DELETE_AFTER)
                            try:
                                await bot.delete_message(chat_id, message_id)
                            except Exception:
                                pass
                        asyncio.create_task(_delete_error_message(notice.chat.id, notice.message_id))
                    except Exception:
                        pass
                return
        return await handler(event, middleware_data)

dp.message.outer_middleware(GroupCommandAccessMiddleware())


# Доступ к боту в ЛС разрешён только участникам группы Admin Staff.
# Проверка выполняется через Telegram при каждом действии в личном чате,
# поэтому после выхода/исключения из Admin Staff доступ прекращается.
async def _is_admin_staff_member(user_id: int) -> bool:
    try:
        member = await bot.get_chat_member(ADMIN_STAFF_CHAT_ID, user_id)
        status = getattr(member, "status", None)
        status_value = getattr(status, "value", status)
        return status_value in {"creator", "administrator", "member", "restricted"}
    except Exception:
        return False


class PrivateAdminStaffOnlyMiddleware(BaseMiddleware):
    async def __call__(self, handler, event, middleware_data):
        user = getattr(event, "from_user", None)
        if not user or user.is_bot:
            return await handler(event, middleware_data)

        is_private = False
        if isinstance(event, Message):
            is_private = event.chat.type == "private"
        elif isinstance(event, CallbackQuery):
            is_private = bool(event.message and event.message.chat.type == "private")

        if not is_private:
            return await handler(event, middleware_data)

        if await _is_admin_staff_member(user.id):
            return await handler(event, middleware_data)

        if isinstance(event, CallbackQuery):
            try:
                await event.answer("⛔ Доступ только для участников Admin Staff.", show_alert=True)
            except Exception:
                pass
        else:
            try:
                await event.answer(
                    "⛔ <b>Доступ запрещён</b>\n\n"
                    "Использовать Shinobi Team Bot в личных сообщениях могут только участники <b>Admin Staff</b>.",
                    parse_mode="HTML"
                )
            except Exception:
                pass
        return


dp.message.outer_middleware(PrivateAdminStaffOnlyMiddleware())
dp.callback_query.outer_middleware(PrivateAdminStaffOnlyMiddleware())

init_schedule_system(ADMIN_IDS)
dp.include_router(schedule_router)


# =========================================================
# JSON
# =========================================================

def default_data():
    return {
        "users": {},
        "links": {},
        "bot_admins": [],
        "known_groups": {},
        "installed_groups": [],
        "group_command_access": {},
        "pending_join_requests": {},
        "advertising_link": ""
    }


def load_data():

    if not os.path.exists(DATA_FILE):
        return default_data()

    try:

        with open(DATA_FILE, "r", encoding="utf-8") as file:
            loaded = json.load(file)

        loaded.setdefault("users", {})
        loaded.setdefault("links", {})
        loaded.setdefault("bot_admins", [])
        loaded.setdefault("known_groups", {})
        loaded.setdefault("installed_groups", [])
        loaded.setdefault("group_command_access", {})
        loaded.setdefault("pending_join_requests", {})
        loaded.setdefault("advertising_link", "")

        return loaded

    except (json.JSONDecodeError, OSError):

        return default_data()


data = load_data()

# Динамические администраторы сохраняются в data.json. Владелец всегда имеет доступ.
for _uid in data.get("bot_admins", []):
    try:
        _uid = int(_uid)
        if _uid not in ADMIN_IDS:
            ADMIN_IDS.append(_uid)
    except (TypeError, ValueError):
        pass
init_schedule_system(ADMIN_IDS)


def save_data():

    payload = json.dumps(
        data,
        ensure_ascii=False,
        indent=4
    )

    temp_file = os.path.join(
        BASE_DIR,
        f"data.{os.getpid()}.{int(time.time() * 1000)}.tmp"
    )

    try:

        with open(temp_file, "w", encoding="utf-8") as file:

            file.write(payload)
            file.flush()
            os.fsync(file.fileno())

        os.replace(temp_file, DATA_FILE)
        return

    except OSError as error:

        print(
            "Не удалось атомарно сохранить data.json:",
            error
        )

    finally:

        if os.path.exists(temp_file):

            try:
                os.remove(temp_file)
            except OSError:
                pass

    try:

        with open(DATA_FILE, "w", encoding="utf-8") as file:

            file.write(payload)
            file.flush()

    except OSError as error:

        print(
            "Не удалось сохранить data.json:",
            error
        )


# =========================================================
# ПОЛЬЗОВАТЕЛИ
# =========================================================

def get_user(user_id: int):

    uid = str(user_id)

    if uid not in data["users"]:

        data["users"][uid] = {
            "id": user_id,
            "name": "",
            "username": "",
            "links": [],
            "total_invited": 0,
            "invited_users": []
        }

    user = data["users"][uid]

    # Совместимость со старым data.json
    user.setdefault("id", user_id)
    user.setdefault("name", "")
    user.setdefault("username", "")
    user.setdefault("links", [])
    user.setdefault("total_invited", 0)
    user.setdefault("invited_users", [])

    return user


def update_user(tg_user):

    user = get_user(tg_user.id)

    user["id"] = tg_user.id
    user["name"] = tg_user.full_name
    user["username"] = tg_user.username or ""

    save_data()

    return user


def display_user(user):

    username = user.get("username", "")

    if username:
        return f"@{username}"

    return user.get("name", "Пользователь")


def is_admin(user_id: int):

    return user_id in ADMIN_IDS


# =========================================================
# КЛАВИАТУРА
# =========================================================

def quick_reply_keyboard(user_id: int):
    # Обычные пользователи видят все доступные им функции бота.
    rows = [
        [KeyboardButton(text="🔗 Создать ссылку"), KeyboardButton(text="📊 Моя статистика")],
        [KeyboardButton(text="🔗 Мои ссылки"), KeyboardButton(text="👥 Приглашённые")],
        [KeyboardButton(text="🏆 Топ"), KeyboardButton(text="📋 Мой график")],
        [KeyboardButton(text="🏖 Запросить выходной"), KeyboardButton(text="🗓 График / заполнить")],
    ]

    # У администраторов дополнительно отдельный блок административных кнопок.
    if is_admin(user_id):
        rows.extend([
            [KeyboardButton(text="⚙️ Админ-панель"), KeyboardButton(text="🗓 Управление графиком")],
            [KeyboardButton(text="📥 Скачать Excel"), KeyboardButton(text="📊 Общая статистика")],
            [KeyboardButton(text="📅 Excel с графиками"), KeyboardButton(text="🏖 Excel с выходными")],
        ])

    return ReplyKeyboardMarkup(
        keyboard=rows,
        resize_keyboard=True,
        is_persistent=True,
        input_field_placeholder="Выбери действие или напиши сообщение…",
    )

def main_keyboard(user_id: int):

    buttons = [
        [
            InlineKeyboardButton(
                text="🔗 Создать ссылку",
                callback_data="create_link"
            )
        ],
        [
            InlineKeyboardButton(
                text="📊 Моя статистика",
                callback_data="my_stats"
            ),
            InlineKeyboardButton(
                text="🔗 Мои ссылки",
                callback_data="my_links"
            )
        ],
        [
            InlineKeyboardButton(
                text="👥 Приглашённые",
                callback_data="my_refs"
            ),
            InlineKeyboardButton(
                text="🏆 Топ",
                callback_data="top"
            )
        ]
    ]

    buttons.append([
        InlineKeyboardButton(text="📋 Мой график", callback_data="sch_mine"),
        InlineKeyboardButton(text="🏖 Запросить выходной", callback_data="sch_dayoff")
    ])
    buttons.append([InlineKeyboardButton(text="🗓 График / заполнить", callback_data="sch_menu")])

    if is_admin(user_id):

        buttons.append([
            InlineKeyboardButton(
                text="⚙️ Админ-панель",
                callback_data="admin_panel"
            )
        ])

    return InlineKeyboardMarkup(
        inline_keyboard=buttons
    )


def admin_keyboard():

    return InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="📥 Скачать Excel",
                    callback_data="export_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📊 Общая статистика",
                    callback_data="admin_stats"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏆 Топ",
                    callback_data="top"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🗓 Управление графиком",
                    callback_data="admin_schedule"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📅 Excel с графиками",
                    callback_data="export_schedule_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏖 Excel с выходными",
                    callback_data="export_dayoff_excel"
                )
            ],
            [
                InlineKeyboardButton(
                    text="👥 Пользователи бота",
                    callback_data="admin_users_0"
                )
            ],
            [
                InlineKeyboardButton(
                    text="🏘 Группы и доступ к командам",
                    callback_data="admin_groups_0"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📩 Заявки на вступление",
                    callback_data="admin_join_requests"
                )
            ],
            [
                InlineKeyboardButton(
                    text="📣 Ссылка для рекламы",
                    callback_data="advertising_link"
                )
            ]
        ]
    )


# =========================================================
# /start
# =========================================================

@dp.message(Command("start"))
async def start_handler(message: Message, state: FSMContext):

    # Персональные ссылки создаём только через ЛС
    if message.chat.type != "private":
        return

    update_user(message.from_user)

    payload = ""
    if message.text:
        parts = message.text.split(maxsplit=1)
        if len(parts) > 1:
            payload = parts[1].strip()

    if payload == "schedule":
        await show_schedule_menu(message, state)
        return

    await message.answer(
        "🥷 <b>SHINOBI TEAM BOT</b>\n\n"
        f"👤 {message.from_user.full_name}\n"
        f"🆔 ID: <code>{message.from_user.id}</code>\n\n"
        "Ты успешно авторизован в системе.\n\n"
        "Используй автоматизированную систему Shinobi Team для управления графиком, персональными ссылками и рабочими инструментами 👇",
        reply_markup=main_keyboard(message.from_user.id),
        parse_mode="HTML"
    )
    await message.answer("⚡ Быстрые действия закреплены под строкой ввода.", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(F.text == "📋 Мой график")
async def quick_my_schedule(message: Message):
    # Переиспользуем текстовую команду, которую обрабатывает модуль графика
    from schedule_system import personal_week_schedule
    await personal_week_schedule(message)

@dp.message(F.text == "🏖 Запросить выходной")
async def quick_dayoff(message: Message):
    if message.chat.type != "private":
        info=await message.bot.get_me()
        await message.answer("🏖 Запрос выходного оформляется в ЛС с ботом.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Открыть бота",url=f"https://t.me/{info.username}?start=schedule")]]))
        return
    await message.answer("🏖 Открой меню графика и нажми «Запросить выходной».", reply_markup=main_keyboard(message.from_user.id))

@dp.message(Command("menu"))
async def quick_menu(message: Message):
    await message.answer("⚡ Быстрые действия:", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(Command("link"))
async def link_command(message: Message):

    # Если команду написали в группе
    if message.chat.type in {"group", "supergroup"}:

        if message.chat.id != CHAT_ID:
            return

        bot_info = await bot.get_me()

        bot_url = f"https://t.me/{bot_info.username}?start=link"

        keyboard = InlineKeyboardMarkup(
            inline_keyboard=[
                [
                    InlineKeyboardButton(
                        text="🔗 Получить свою ссылку",
                        url=bot_url
                    )
                ]
            ]
        )

        await message.reply(
            "🥷 <b>SHINOBI TEAM</b>\n\n"
            f"👤 {message.from_user.full_name}\n\n"
            "Чтобы получить персональную пригласительную ссылку, "
            "перейди в ЛС бота 👇",
            reply_markup=keyboard,
            parse_mode="HTML"
        )

        return

    # Если /link написали сразу в ЛС
    if message.chat.type == "private":

        await create_ref_link(
            message.from_user,
            message
        )

@dp.message(F.text == "..ссылка")
async def ref_link_chat(message: Message):

    # Команда работает только в группе
    if message.chat.type not in {"group", "supergroup"}:
        return

    # Только в нужной группе
    if message.chat.id != CHAT_ID:
        return

    bot_info = await bot.get_me()

    bot_url = f"https://t.me/{bot_info.username}?start=link"

    keyboard = InlineKeyboardMarkup(
        inline_keyboard=[
            [
                InlineKeyboardButton(
                    text="🔗 Получить ссылку",
                    url=bot_url
                )
            ]
        ]
    )

    await message.reply(
        f"👤 <b>{message.from_user.full_name}</b>\n\n"
        "Для получения персональной реферальной ссылки "
        "перейди в ЛС бота и нажми <b>START</b> 👇",
        reply_markup=keyboard,
        parse_mode="HTML"
    )

# =========================================================
# /menu
# =========================================================

@dp.message(Command("menu"))
async def menu_handler(message: Message):

    update_user(message.from_user)

    await message.answer(
        "🥷 <b>SHINOBI TEAM BOT</b>\n\n"
        "Выбери нужное действие:",
        reply_markup=main_keyboard(message.from_user.id),
        parse_mode="HTML"
    )


# =========================================================
# СОЗДАНИЕ РЕФЕРАЛЬНОЙ ССЫЛКИ
# =========================================================

async def _send_created_link(tg_user, target, creates_join_request: bool = False):
    user = update_user(tg_user)
    link_number = len(user["links"]) + 1

    try:
        invite = await bot.create_chat_invite_link(
            chat_id=CHAT_ID,
            name=f"REF_{tg_user.id}_{link_number}",
            creates_join_request=creates_join_request
        )
    except Exception as error:
        text = (
            "❌ <b>Не удалось создать ссылку.</b>\n\n"
            "Проверь, что бот является администратором группы и имеет право создавать "
            "пригласительные ссылки.\n\n"
            f"<code>{error}</code>"
        )
        if isinstance(target, Message):
            await target.answer(text, parse_mode="HTML")
        else:
            await target.message.answer(text, parse_mode="HTML")
        return

    link = invite.invite_link
    data["links"][link] = {
        "owner_id": tg_user.id,
        "created": datetime.now().isoformat(),
        "invited": 0,
        "users": [],
        "join_request": bool(creates_join_request)
    }
    user["links"].append(link)
    save_data()

    mode = "📩 По заявке" if creates_join_request else "⚡ Вход сразу"
    text = (
        "🔗 <b>НОВАЯ РЕФЕРАЛЬНАЯ ССЫЛКА</b>\n\n"
        f"{link}\n\n"
        f"👤 Создатель: {tg_user.full_name}\n"
        f"🆔 ID: <code>{tg_user.id}</code>\n"
        f"📎 Ссылка №{link_number}\n"
        f"🚪 Режим: <b>{mode}</b>\n"
        "👥 Приглашено: <b>0</b>\n\n"
        + ("Заявки на вход будут видны администраторам бота и в Admin Staff."
           if creates_join_request else "Пользователь сможет войти сразу без подтверждения.")
    )
    if isinstance(target, Message):
        await target.answer(text, parse_mode="HTML", disable_web_page_preview=True)
    else:
        await target.message.answer(text, parse_mode="HTML", disable_web_page_preview=True)


async def create_ref_link(tg_user, target):
    # Только администраторы выбирают тип ссылки. Для обычных пользователей
    # сохраняется прежнее поведение — ссылка с мгновенным входом.
    if is_admin(tg_user.id):
        kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="⚡ Без заявки", callback_data="create_link_direct")],
            [InlineKeyboardButton(text="📩 С заявкой на вход", callback_data="create_link_request")],
        ])
        text = "🔗 <b>ТИП ПРИГЛАСИТЕЛЬНОЙ ССЫЛКИ</b>\n\nВыбери режим входа в Shinobi Team:"
        if isinstance(target, Message):
            await target.answer(text, reply_markup=kb, parse_mode="HTML")
        else:
            await target.message.answer(text, reply_markup=kb, parse_mode="HTML")
        return
    await _send_created_link(tg_user, target, False)


@dp.callback_query(F.data.in_({"create_link_direct", "create_link_request"}))
async def create_link_mode_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    await callback.answer("Создаю ссылку...")
    await _send_created_link(callback.from_user, callback, callback.data == "create_link_request")


# =========================================================
# /link
# =========================================================

@dp.message(F.text == "📋 Мой график")
async def quick_my_schedule(message: Message):
    # Переиспользуем текстовую команду, которую обрабатывает модуль графика
    from schedule_system import personal_week_schedule
    await personal_week_schedule(message)

@dp.message(F.text == "🏖 Запросить выходной")
async def quick_dayoff(message: Message):
    if message.chat.type != "private":
        info=await message.bot.get_me()
        await message.answer("🏖 Запрос выходного оформляется в ЛС с ботом.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="Открыть бота",url=f"https://t.me/{info.username}?start=schedule")]]))
        return
    await message.answer("🏖 Открой меню графика и нажми «Запросить выходной».", reply_markup=main_keyboard(message.from_user.id))

@dp.message(Command("menu"))
async def quick_menu(message: Message):
    await message.answer("⚡ Быстрые действия:", reply_markup=quick_reply_keyboard(message.from_user.id))

@dp.message(Command("link"))
async def link_command(message: Message):

    await create_ref_link(
        message.from_user,
        message
    )


# =========================================================
# КНОПКА СОЗДАТЬ ССЫЛКУ
# =========================================================

@dp.callback_query(F.data == "create_link")
async def link_callback(callback: CallbackQuery):

    await callback.answer(
        "Создаю ссылку..."
    )

    await create_ref_link(
        callback.from_user,
        callback
    )


# =========================================================
# ЕДИНАЯ ССЫЛКА ДЛЯ РЕКЛАМЫ
# =========================================================

@dp.callback_query(F.data == "advertising_link")
async def advertising_link_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return

    await callback.answer("Получаю рекламную ссылку...")
    link = data.get("advertising_link", "")

    # Если ссылка уже есть в базе, повторно её не создаём.
    if link and link in data.get("links", {}):
        await callback.message.answer(
            "📣 <b>ССЫЛКА ДЛЯ РЕКЛАМЫ</b>\n\n"
            f"{link}\n\n"
            "Это единая ссылка для рекламных постов. Новую при каждом нажатии бот не создаёт.",
            parse_mode="HTML", disable_web_page_preview=True
        )
        return

    try:
        invite = await bot.create_chat_invite_link(
            chat_id=CHAT_ID,
            name="SHINOBI_ADVERTISING",
            creates_join_request=True
        )
    except Exception as error:
        await callback.message.answer(
            "❌ <b>Не удалось создать рекламную ссылку.</b>\n\n"
            "Проверь права бота на создание пригласительных ссылок.\n\n"
            f"<code>{error}</code>", parse_mode="HTML"
        )
        return

    link = invite.invite_link
    data["advertising_link"] = link
    data.setdefault("links", {})[link] = {
        "owner_id": None,
        "type": "advertising",
        "label": "Ссылка для рекламы",
        "created": datetime.now().isoformat(),
        "invited": 0,
        "users": [],
        "join_request": True
    }
    save_data()

    await callback.message.answer(
        "✅ <b>СОЗДАНА ССЫЛКА ДЛЯ РЕКЛАМЫ</b>\n\n"
        f"{link}\n\n"
        "В источнике приглашения она будет отображаться как <b>«Ссылка для рекламы»</b>.",
        parse_mode="HTML", disable_web_page_preview=True
    )


def invite_source_text(invite_url: str) -> str:
    if not invite_url:
        return "Ссылка не относится к боту"
    info = data.get("links", {}).get(invite_url)
    if not info:
        return "Ссылка не относится к боту"
    if info.get("type") == "advertising":
        return "Ссылка для рекламы"
    owner_id = info.get("owner_id")
    owner = get_user(owner_id) if owner_id else None
    return owner.get("name", "Ссылка не относится к боту") if owner else "Ссылка не относится к боту"


# =========================================================
# ЗАЯВКИ НА ВСТУПЛЕНИЕ
# =========================================================

def _join_request_key(chat_id: int, user_id: int) -> str:
    return f"{chat_id}:{user_id}"


def _join_request_keyboard(chat_id: int, user_id: int):
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text="✅ Принять", callback_data=f"jr_accept:{chat_id}:{user_id}"),
        InlineKeyboardButton(text="❌ Отклонить", callback_data=f"jr_decline:{chat_id}:{user_id}")
    ]])


def _join_request_text(req: dict) -> str:
    username = f"@{req['username']}" if req.get("username") else "—"
    source = req.get("link_source") or invite_source_text(req.get("invite_link", ""))
    return (
        "📩 <b>НОВАЯ ЗАЯВКА НА ВСТУПЛЕНИЕ</b>\n\n"
        f"👤 {req.get('name', 'Без имени')}\n"
        f"🔹 {username}\n"
        f"🆔 <code>{req.get('user_id')}</code>\n"
        f"🏘 Группа: <b>{req.get('chat_title', 'Shinobi Team')}</b>\n"
        f"🔗 Источник: <b>{source}</b>\n\n"
        "Решение синхронизируется для всех администраторов."
    )


@dp.chat_join_request()
async def join_request_handler(event: ChatJoinRequest):
    if event.chat.id != CHAT_ID:
        return
    invite_url = event.invite_link.invite_link if event.invite_link else ""
    link_info = data.get("links", {}).get(invite_url, {})
    key = _join_request_key(event.chat.id, event.from_user.id)
    req = {
        "chat_id": event.chat.id,
        "chat_title": event.chat.title or "Shinobi Team",
        "user_id": event.from_user.id,
        "name": event.from_user.full_name,
        "username": event.from_user.username or "",
        "invite_link": invite_url,
        "link_owner_id": link_info.get("owner_id"),
        "link_source": invite_source_text(invite_url),
        "created": datetime.now().isoformat(),
        "status": "pending"
    }
    data.setdefault("pending_join_requests", {})[key] = req
    save_data()
    text = _join_request_text(req)
    kb = _join_request_keyboard(event.chat.id, event.from_user.id)

    # Уведомление каждому администратору в ЛС.
    for admin_id in set(ADMIN_IDS):
        try:
            await bot.send_message(admin_id, text, reply_markup=kb, parse_mode="HTML")
        except Exception:
            pass
    # И единое уведомление в группе Admin Staff.
    try:
        await bot.send_message(ADMIN_STAFF_CHAT_ID, text, reply_markup=kb, parse_mode="HTML")
    except Exception:
        pass


@dp.callback_query(F.data.startswith("jr_accept:") | F.data.startswith("jr_decline:"))
async def join_request_decision(callback: CallbackQuery):
    # Заявки на вступление могут обрабатывать:
    # 1) владелец/администраторы бота;
    # 2) любой обычный участник группы Admin Staff.
    # Посторонние пользователи доступа к кнопкам не получают.
    can_decide = is_admin(callback.from_user.id) or await _is_admin_staff_member(callback.from_user.id)
    if not can_decide:
        await callback.answer("⛔ Принимать заявки могут только участники Admin Staff и администрация.", show_alert=True)
        return
    action, chat_id_s, user_id_s = callback.data.split(":", 2)
    chat_id, user_id = int(chat_id_s), int(user_id_s)
    key = _join_request_key(chat_id, user_id)
    req = data.setdefault("pending_join_requests", {}).get(key)
    if not req or req.get("status") != "pending":
        await callback.answer("Эта заявка уже обработана.", show_alert=True)
        try:
            await callback.message.edit_reply_markup(reply_markup=None)
        except Exception:
            pass
        return
    try:
        if action == "jr_accept":
            await bot.approve_chat_join_request(chat_id, user_id)
            req["status"] = "approved"
            result = "✅ ПРИНЯТА"
        else:
            await bot.decline_chat_join_request(chat_id, user_id)
            req["status"] = "declined"
            result = "❌ ОТКЛОНЕНА"
    except Exception as error:
        await callback.answer(f"Ошибка Telegram: {error}", show_alert=True)
        return
    req["decided_by"] = callback.from_user.id
    req["decided_at"] = datetime.now().isoformat()
    save_data()
    await callback.answer("Готово")
    try:
        await callback.message.edit_text(
            _join_request_text(req) + f"\n\n<b>{result}</b>\nРешение: {callback.from_user.full_name}",
            parse_mode="HTML", reply_markup=None
        )
    except Exception:
        pass


@dp.callback_query(F.data == "admin_join_requests")
async def admin_join_requests_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    pending = [r for r in data.get("pending_join_requests", {}).values() if r.get("status") == "pending"]
    if not pending:
        await callback.message.edit_text(
            "📩 <b>ЗАЯВКИ НА ВСТУПЛЕНИЕ</b>\n\nСейчас активных заявок нет.",
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")]]),
            parse_mode="HTML"
        )
        await callback.answer()
        return
    rows=[]
    for req in pending[-30:][::-1]:
        rows.append([InlineKeyboardButton(
            text=f"📩 {req.get('name','Без имени')}"[:55],
            callback_data=f"admin_jr:{req['chat_id']}:{req['user_id']}"
        )])
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    await callback.message.edit_text(
        f"📩 <b>ЗАЯВКИ НА ВСТУПЛЕНИЕ</b>\n\nОжидают решения: <b>{len(pending)}</b>",
        reply_markup=InlineKeyboardMarkup(inline_keyboard=rows), parse_mode="HTML"
    )
    await callback.answer()


@dp.callback_query(F.data.startswith("admin_jr:"))
async def admin_join_request_open(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("Нет доступа.", show_alert=True)
        return
    _, chat_id_s, user_id_s = callback.data.split(":", 2)
    req = data.get("pending_join_requests", {}).get(_join_request_key(int(chat_id_s), int(user_id_s)))
    if not req or req.get("status") != "pending":
        await callback.answer("Заявка уже обработана.", show_alert=True)
        return
    await callback.message.edit_text(_join_request_text(req), reply_markup=_join_request_keyboard(int(chat_id_s), int(user_id_s)), parse_mode="HTML")
    await callback.answer()


# =========================================================
# ОТСЛЕЖИВАНИЕ ВСТУПЛЕНИЙ
# =========================================================

@dp.chat_member()
async def member_join_handler(event: ChatMemberUpdated):

    # Только нужная группа
    if event.chat.id != CHAT_ID:
        return

    old_status = event.old_chat_member.status
    new_status = event.new_chat_member.status

    joined = (
        old_status in {"left", "kicked"}
        and new_status in {
            "member",
            "administrator",
            "restricted"
        }
    )

    if not joined:
        return

    # Если Telegram не передал ссылку
    if not event.invite_link:
        return

    used_link = event.invite_link.invite_link

    joined_user = event.new_chat_member.user

    # Ботов не считаем
    if joined_user.is_bot:
        return

    # Чужая/ручная ссылка не относится к реферальной системе бота.
    if used_link not in data.get("links", {}):
        return

    link_info = data["links"][used_link]

    # Рекламная ссылка общая и не принадлежит конкретному пользователю.
    if link_info.get("type") == "advertising":
        link_info.setdefault("invited", 0)
        link_info.setdefault("users", [])
        if any(x.get("id") == joined_user.id for x in link_info["users"]):
            return
        link_info["invited"] += 1
        link_info["users"].append({
            "id": joined_user.id,
            "name": joined_user.full_name,
            "username": joined_user.username or "",
            "joined": datetime.now().isoformat(),
            "source": "Ссылка для рекламы"
        })
        save_data()
        return

    owner_id = str(link_info.get("owner_id"))
    if owner_id not in data["users"]:
        return
    owner = data["users"][owner_id]
    owner.setdefault("invited_users", [])
    owner.setdefault("total_invited", 0)

    link_info.setdefault(
        "invited",
        0
    )

    link_info.setdefault(
        "users",
        []
    )

    joined_id = joined_user.id

    # =====================================================
    # АНТИНАКРУТКА
    # =====================================================
    #
    # Если человек уже был засчитан этому владельцу,
    # повторно его не считаем.
    #
    # =====================================================

    if joined_id in owner["invited_users"]:
        return

    # Записываем приглашённого
    owner["invited_users"].append(
        joined_id
    )

    owner["total_invited"] += 1

    link_info["invited"] += 1

    link_info["users"].append({
        "id": joined_user.id,
        "name": joined_user.full_name,
        "username": joined_user.username or "",
        "joined": datetime.now().isoformat()
    })

    save_data()

    # =====================================================
    # УВЕДОМЛЕНИЕ ВЛАДЕЛЬЦУ ССЫЛКИ
    # =====================================================

    try:

        username_text = ""

        if joined_user.username:

            username_text = (
                f"\n🔹 @{joined_user.username}"
            )

        await bot.send_message(
            int(owner_id),
            "🎉 <b>НОВЫЙ ПРИГЛАШЁННЫЙ!</b>\n\n"
            f"👤 {joined_user.full_name}"
            f"{username_text}\n"
            f"🆔 <code>{joined_user.id}</code>\n\n"
            f"📊 Всего приглашено: "
            f"<b>{owner['total_invited']}</b>",
            parse_mode="HTML"
        )

    except Exception:
        pass


# =========================================================
# МОЯ СТАТИСТИКА
# =========================================================

async def send_stats(user_id: int, target):

    user = get_user(user_id)

    links = user.get(
        "links",
        []
    )

    text = (
        "📊 <b>ТВОЯ СТАТИСТИКА</b>\n\n"
        f"🆔 ID: <code>{user_id}</code>\n"
        f"👥 Всего приглашено: "
        f"<b>{user.get('total_invited', 0)}</b>\n"
        f"🔗 Создано ссылок: "
        f"<b>{len(links)}</b>\n"
    )

    if links:

        text += (
            "\n<b>Статистика ссылок:</b>\n\n"
        )

        for number, link in enumerate(
            links,
            start=1
        ):

            info = data["links"].get(
                link
            )

            if not info:
                continue

            text += (
                f"🔗 Ссылка #{number}\n"
                f"👥 Приглашено: "
                f"<b>{info.get('invited', 0)}</b>\n"
                f"{link}\n\n"
            )

            if len(text) > 3500:

                if isinstance(target, Message):

                    await target.answer(
                        text,
                        parse_mode="HTML",
                        disable_web_page_preview=True
                    )

                else:

                    await target.message.answer(
                        text,
                        parse_mode="HTML",
                        disable_web_page_preview=True
                    )

                text = ""

    if text:

        if isinstance(target, Message):

            await target.answer(
                text,
                parse_mode="HTML",
                disable_web_page_preview=True
            )

        else:

            await target.message.answer(
                text,
                parse_mode="HTML",
                disable_web_page_preview=True
            )


# =========================================================
# /stats
# =========================================================

@dp.message(Command("stats"))
async def stats_command(message: Message):

    await send_stats(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА СТАТИСТИКА
# =========================================================

@dp.callback_query(F.data == "my_stats")
async def stats_callback(callback: CallbackQuery):

    await callback.answer()

    await send_stats(
        callback.from_user.id,
        callback
    )


# =========================================================
# МОИ ССЫЛКИ
# =========================================================

async def send_links(user_id: int, target):

    user = get_user(user_id)

    links = user.get(
        "links",
        []
    )

    if not links:

        text = (
            "❌ У тебя пока нет реферальных ссылок.\n\n"
            "Нажми «🔗 Создать ссылку»."
        )

    else:

        text = (
            "🔗 <b>ТВОИ ССЫЛКИ</b>\n\n"
        )

        for number, link in enumerate(
            links,
            start=1
        ):

            info = data["links"].get(
                link,
                {}
            )

            text += (
                f"<b>#{number}</b>\n"
                f"{link}\n"
                f"👥 Приглашено: "
                f"<b>{info.get('invited', 0)}</b>\n\n"
            )

    if isinstance(target, Message):

        await target.answer(
            text,
            parse_mode="HTML",
            disable_web_page_preview=True
        )

    else:

        await target.message.answer(
            text,
            parse_mode="HTML",
            disable_web_page_preview=True
        )


# =========================================================
# /links
# =========================================================

@dp.message(Command("links"))
async def links_command(message: Message):

    await send_links(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА МОИ ССЫЛКИ
# =========================================================

@dp.callback_query(F.data == "my_links")
async def links_callback(callback: CallbackQuery):

    await callback.answer()

    await send_links(
        callback.from_user.id,
        callback
    )


# =========================================================
# ПРИГЛАШЁННЫЕ
# =========================================================

async def send_refs(user_id: int, target):

    user = get_user(user_id)

    refs = []

    for link in user.get(
        "links",
        []
    ):

        link_info = data["links"].get(
            link
        )

        if not link_info:
            continue

        for ref in link_info.get(
            "users",
            []
        ):

            refs.append(ref)

    if not refs:

        text = (
            "👥 Ты пока никого не пригласил."
        )

        if isinstance(target, Message):

            await target.answer(
                text
            )

        else:

            await target.message.answer(
                text
            )

        return

    text = (
        "👥 <b>ТВОИ ПРИГЛАШЁННЫЕ</b>\n\n"
    )

    for number, ref in enumerate(
        refs,
        start=1
    ):

        username = ref.get(
            "username",
            ""
        )

        if username:

            name = f"@{username}"

        else:

            name = ref.get(
                "name",
                "Пользователь"
            )

        line = (
            f"{number}. {name}\n"
            f"🆔 <code>{ref.get('id')}</code>\n\n"
        )

        if len(text) + len(line) > 3500:

            if isinstance(target, Message):

                await target.answer(
                    text,
                    parse_mode="HTML"
                )

            else:

                await target.message.answer(
                    text,
                    parse_mode="HTML"
                )

            text = ""

        text += line

    if text:

        if isinstance(target, Message):

            await target.answer(
                text,
                parse_mode="HTML"
            )

        else:

            await target.message.answer(
                text,
                parse_mode="HTML"
            )


# =========================================================
# /refs
# =========================================================

@dp.message(Command("refs"))
async def refs_command(message: Message):

    await send_refs(
        message.from_user.id,
        message
    )


# =========================================================
# КНОПКА ПРИГЛАШЁННЫЕ
# =========================================================

@dp.callback_query(F.data == "my_refs")
async def refs_callback(callback: CallbackQuery):

    await callback.answer()

    await send_refs(
        callback.from_user.id,
        callback
    )


# =========================================================
# ТОП
# =========================================================

async def send_top(target):

    users = list(
        data["users"].items()
    )

    users.sort(
        key=lambda item: item[1].get(
            "total_invited",
            0
        ),
        reverse=True
    )

    text = (
        "🏆 <b>ТОП SHINOBI TEAM</b>\n\n"
    )

    medals = [
        "🥇",
        "🥈",
        "🥉"
    ]

    position = 0

    for uid, user in users:

        count = user.get(
            "total_invited",
            0
        )

        if count <= 0:
            continue

        position += 1

        if position <= 3:

            icon = medals[
                position - 1
            ]

        else:

            icon = f"{position}."

        text += (
            f"{icon} "
            f"{display_user(user)} — "
            f"<b>{count}</b>\n"
        )

        if position >= 20:
            break

    if position == 0:

        text += (
            "Пока приглашений нет."
        )

    if isinstance(target, Message):

        await target.answer(
            text,
            parse_mode="HTML"
        )

    else:

        await target.message.answer(
            text,
            parse_mode="HTML"
        )


# =========================================================
# /top
# =========================================================

@dp.message(Command("top"))
async def top_command(message: Message):

    await send_top(message)


# =========================================================
# КНОПКА ТОП
# =========================================================

@dp.callback_query(F.data == "top")
async def top_callback(callback: CallbackQuery):

    await callback.answer()

    await send_top(callback)


# =========================================================
# СОЗДАНИЕ EXCEL
# =========================================================

def create_excel():

    workbook = Workbook()

    # =====================================================
    # ЛИСТ №1 — РЕФЕРАЛЬНЫЕ ССЫЛКИ
    # =====================================================

    sheet = workbook.active

    sheet.title = (
        "Реферальные ссылки"
    )

    headers = [
        "№",
        "Ссылка",
        "ID создателя",
        "Username",
        "Имя",
        "Приглашено",
        "Дата создания"
    ]

    sheet.append(headers)

    for cell in sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    number = 1

    for link, info in data["links"].items():

        owner_id = str(
            info.get(
                "owner_id",
                ""
            )
        )

        owner = data["users"].get(
            owner_id,
            {}
        )

        username = owner.get(
            "username",
            ""
        )

        if username:

            username = (
                f"@{username}"
            )

        created = info.get(
            "created",
            ""
        )

        try:

            created = (
                datetime
                .fromisoformat(created)
                .strftime(
                    "%d.%m.%Y %H:%M:%S"
                )
            )

        except Exception:
            pass

        sheet.append([
            number,
            link,
            owner_id,
            username,
            owner.get(
                "name",
                ""
            ),
            info.get(
                "invited",
                0
            ),
            created
        ])

        number += 1

    sheet.freeze_panes = "A2"

    sheet.auto_filter.ref = (
        sheet.dimensions
    )

    sheet.column_dimensions[
        "A"
    ].width = 8

    sheet.column_dimensions[
        "B"
    ].width = 45

    sheet.column_dimensions[
        "C"
    ].width = 20

    sheet.column_dimensions[
        "D"
    ].width = 25

    sheet.column_dimensions[
        "E"
    ].width = 30

    sheet.column_dimensions[
        "F"
    ].width = 15

    sheet.column_dimensions[
        "G"
    ].width = 23


    # =====================================================
    # ЛИСТ №2 — ПРИГЛАШЁННЫЕ
    # =====================================================

    refs_sheet = workbook.create_sheet(
        "Приглашённые"
    )

    ref_headers = [
        "№",
        "Ссылка",
        "ID создателя",
        "Username создателя",
        "ID приглашённого",
        "Username приглашённого",
        "Имя приглашённого",
        "Дата вступления"
    ]

    refs_sheet.append(
        ref_headers
    )

    for cell in refs_sheet[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ref_number = 1

    for link, info in data["links"].items():

        owner_id = str(
            info.get(
                "owner_id",
                ""
            )
        )

        owner = data["users"].get(
            owner_id,
            {}
        )

        owner_username = owner.get(
            "username",
            ""
        )

        if owner_username:

            owner_username = (
                f"@{owner_username}"
            )

        for ref in info.get(
            "users",
            []
        ):

            ref_username = ref.get(
                "username",
                ""
            )

            if ref_username:

                ref_username = (
                    f"@{ref_username}"
                )

            joined = ref.get(
                "joined",
                ""
            )

            try:

                joined = (
                    datetime
                    .fromisoformat(joined)
                    .strftime(
                        "%d.%m.%Y %H:%M:%S"
                    )
                )

            except Exception:
                pass

            refs_sheet.append([
                ref_number,
                link,
                owner_id,
                owner_username,
                ref.get(
                    "id",
                    ""
                ),
                ref_username,
                ref.get(
                    "name",
                    ""
                ),
                joined
            ])

            ref_number += 1

    refs_sheet.freeze_panes = (
        "A2"
    )

    refs_sheet.auto_filter.ref = (
        refs_sheet.dimensions
    )

    refs_sheet.column_dimensions[
        "A"
    ].width = 8

    refs_sheet.column_dimensions[
        "B"
    ].width = 45

    refs_sheet.column_dimensions[
        "C"
    ].width = 22

    refs_sheet.column_dimensions[
        "D"
    ].width = 25

    refs_sheet.column_dimensions[
        "E"
    ].width = 22

    refs_sheet.column_dimensions[
        "F"
    ].width = 27

    refs_sheet.column_dimensions[
        "G"
    ].width = 30

    refs_sheet.column_dimensions[
        "H"
    ].width = 23

    # =====================================================
    # ОБЩЕЕ ВЫРАВНИВАНИЕ
    # =====================================================

    for current_sheet in [
        sheet,
        refs_sheet
    ]:

        for row in current_sheet.iter_rows():

            for cell in row:

                cell.alignment = Alignment(
                    vertical="center"
                )

    workbook.save(
        EXCEL_FILE
    )

    return EXCEL_FILE


# =========================================================
# ПОСТОЯННЫЕ НИЖНИЕ КНОПКИ
# =========================================================

@dp.message(F.text == "🔗 Создать ссылку")
async def reply_create_link(message: Message):
    await link_command(message)

@dp.message(F.text == "📊 Моя статистика")
async def reply_my_stats(message: Message):
    await stats_command(message)

@dp.message(F.text == "🔗 Мои ссылки")
async def reply_my_links(message: Message):
    await links_command(message)

@dp.message(F.text == "👥 Приглашённые")
async def reply_my_refs(message: Message):
    await refs_command(message)

@dp.message(F.text == "🏆 Топ")
async def reply_top(message: Message):
    await top_command(message)

@dp.message(F.text == "🗓 График / заполнить")
async def reply_schedule_menu(message: Message, state: FSMContext):
    if message.chat.type != "private":
        await message.answer("🗓 Заполнение графика доступно в личных сообщениях с ботом.")
        return
    await show_schedule_menu(message, state)

@dp.message(F.text == "⚙️ Админ-панель")
async def reply_admin_panel(message: Message):
    await admin_command(message)

async def _admin_shortcut(message: Message, callback_data: str, title: str):
    if not is_admin(message.from_user.id):
        await message.answer("❌ У тебя нет доступа к этой функции.")
        return
    await message.answer(
        title,
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[
            InlineKeyboardButton(text="Открыть", callback_data=callback_data)
        ]])
    )

@dp.message(F.text == "🗓 Управление графиком")
async def reply_admin_schedule(message: Message):
    await _admin_shortcut(message, "admin_schedule", "🗓 Управление графиком")

@dp.message(F.text == "📥 Скачать Excel")
async def reply_export_excel(message: Message):
    await _admin_shortcut(message, "export_excel", "📥 Excel со ссылками")

@dp.message(F.text == "📊 Общая статистика")
async def reply_admin_stats(message: Message):
    await _admin_shortcut(message, "admin_stats", "📊 Общая статистика")

@dp.message(F.text == "📅 Excel с графиками")
async def reply_schedule_excel(message: Message):
    await _admin_shortcut(message, "export_schedule_excel", "📅 Excel с графиками")

@dp.message(F.text == "🏖 Excel с выходными")
async def reply_dayoff_excel(message: Message):
    await _admin_shortcut(message, "export_dayoff_excel", "🏖 Excel с запросами выходных")

# =========================================================
# /admin
# =========================================================

@dp.message(Command("admin"))
async def admin_command(message: Message):

    if not is_admin(
        message.from_user.id
    ):

        await message.answer(
            "❌ У тебя нет доступа "
            "к админ-панели."
        )

        return

    await message.answer(
        "⚙️ <b>АДМИН-ПАНЕЛЬ</b>\n\n"
        "🥷 Shinobi Team · Automation Bot\n\n"
        "Выбери действие:",
        reply_markup=admin_keyboard(),
        parse_mode="HTML"
    )


# =========================================================
# КНОПКА АДМИН-ПАНЕЛЬ
# =========================================================

@dp.callback_query(
    F.data == "admin_panel"
)
async def admin_panel_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer()

    await callback.message.answer(
        "⚙️ <b>АДМИН-ПАНЕЛЬ</b>\n\n"
        "🥷 Shinobi Team · Automation Bot\n\n"
        "Выбери действие:",
        reply_markup=admin_keyboard(),
        parse_mode="HTML"
    )



# =========================================================
# ГРУППЫ БОТА / ДОСТУП К КОМАНДАМ
# =========================================================

def admin_groups_keyboard(page: int = 0):
    groups = []
    installed = {str(x) for x in data.get("installed_groups", [])}
    for cid, g in data.get("known_groups", {}).items():
        if str(cid) in installed and g.get("bot_status") in {"administrator", "creator"}:
            groups.append((cid, g))
    groups.sort(key=lambda item: (item[1].get("title") or item[0]).casefold())
    per_page = 8
    pages = max(1, (len(groups) + per_page - 1) // per_page)
    page = max(0, min(page, pages - 1))
    rows = []
    for cid, g in groups[page * per_page:(page + 1) * per_page]:
        denied = data.get("group_command_access", {}).get(str(cid), "allow") == "deny"
        badge = "🔴" if denied else "🟢"
        title = g.get("title") or f"Группа {cid}"
        rows.append([InlineKeyboardButton(text=f"{badge} {title}"[:60], callback_data=f"admin_group_{cid}_{page}")])
    if not groups:
        rows.append([InlineKeyboardButton(text="Группы пока не обнаружены", callback_data="noop")])
    nav=[]
    if page > 0: nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"admin_groups_{page-1}"))
    nav.append(InlineKeyboardButton(text=f"{page+1}/{pages}", callback_data="noop"))
    if page + 1 < pages: nav.append(InlineKeyboardButton(text="➡️", callback_data=f"admin_groups_{page+1}"))
    rows.append(nav)
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


def admin_group_actions_keyboard(chat_id: int, page: int = 0):
    denied = data.get("group_command_access", {}).get(str(chat_id), "allow") == "deny"
    action_text = "✅ Разрешить команды" if denied else "⛔ Запретить команды"
    action = "allow" if denied else "deny"
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=action_text, callback_data=f"admin_groupaccess_{action}_{chat_id}_{page}")],
        [InlineKeyboardButton(text="⬅️ К группам", callback_data=f"admin_groups_{page}")]
    ])


@dp.message(F.text.casefold().in_({".установить", "/установить", "/install"}))
async def install_group_command(message: Message):
    if message.chat.type not in {"group", "supergroup"}:
        await message.answer("Эта команда используется непосредственно в группе.")
        return

    # Устанавливать Shinobi Team Bot может только администратор/владелец группы.
    try:
        member = await bot.get_chat_member(message.chat.id, message.from_user.id)
        if member.status not in {"administrator", "creator"}:
            await message.reply("❌ Установить бота может только администратор группы.")
            return
        me = await bot.get_me()
        bot_member = await bot.get_chat_member(message.chat.id, me.id)
        if bot_member.status not in {"administrator", "creator"}:
            await message.reply("❌ Сначала выдайте боту права администратора, затем повторите <code>.установить</code>.", parse_mode="HTML")
            return
    except Exception:
        await message.reply("❌ Не удалось проверить права администратора. Проверьте права бота и повторите команду.")
        return

    cid = str(message.chat.id)
    installed = [str(x) for x in data.setdefault("installed_groups", [])]
    already = cid in installed
    if not already:
        data["installed_groups"].append(cid)

    data.setdefault("known_groups", {})[cid] = {
        "id": message.chat.id,
        "title": message.chat.title or f"Группа {message.chat.id}",
        "bot_status": bot_member.status,
        "installed_by": message.from_user.id,
        "installed_at": datetime.now().isoformat(),
    }
    data.setdefault("group_command_access", {}).setdefault(cid, "allow")
    save_data()

    if already:
        await message.reply("✅ <b>Shinobi Team Bot уже установлен в этой группе.</b>\n\nКоманды работают согласно настройкам доступа владельца.", parse_mode="HTML")
    else:
        await message.reply(
            "🥷 <b>SHINOBI TEAM BOT · УСТАНОВКА</b>\n\n"
            "✅ Группа зарегистрирована в системе.\n"
            f"🆔 <code>{message.chat.id}</code>\n"
            "🟢 Команды бота разрешены.\n\n"
            "Теперь группа отображается у владельца в разделе «🏘 Группы и доступ к командам».",
            parse_mode="HTML"
        )


@dp.callback_query(F.data.startswith("admin_groups_"))
async def admin_groups_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    page = int(callback.data.rsplit("_", 1)[1])
    await callback.answer()
    installed = {str(x) for x in data.get("installed_groups", [])}
    count = sum(1 for cid, g in data.get("known_groups", {}).items() if str(cid) in installed and g.get("bot_status") in {"administrator", "creator"})
    await callback.message.answer(
        f"🏘 <b>ГРУППЫ SHINOBI TEAM</b>\n\nГрупп, где бот администратор: <b>{count}</b>\n"
        "🟢 команды разрешены · 🔴 команды запрещены\n\nВыбери группу:",
        reply_markup=admin_groups_keyboard(page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_group_") & ~F.data.startswith("admin_groups_") & ~F.data.startswith("admin_groupaccess_"))
async def admin_group_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    rest = callback.data[len("admin_group_"):]
    chat_s, page_s = rest.rsplit("_", 1)
    chat_id, page = int(chat_s), int(page_s)
    g = data.get("known_groups", {}).get(str(chat_id))
    if not g:
        await callback.answer("Группа не найдена", show_alert=True); return
    denied = data.get("group_command_access", {}).get(str(chat_id), "allow") == "deny"
    await callback.answer()
    await callback.message.answer(
        f"🏘 <b>{g.get('title') or 'Группа'}</b>\n🆔 <code>{chat_id}</code>\n"
        f"🤖 Статус бота: <b>{g.get('bot_status', 'administrator')}</b>\n"
        f"⌨️ Команды: <b>{'ЗАПРЕЩЕНЫ' if denied else 'РАЗРЕШЕНЫ'}</b>\n\n"
        "При запрете любые команды бота в этой группе не выполняются. Ответ error 404 показывается одному пользователю не чаще одного раза в 4 часа.",
        reply_markup=admin_group_actions_keyboard(chat_id, page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_groupaccess_"))
async def admin_group_access_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    rest = callback.data[len("admin_groupaccess_"):]
    action, rest = rest.split("_", 1)
    chat_s, page_s = rest.rsplit("_", 1)
    chat_id, page = int(chat_s), int(page_s)
    data.setdefault("group_command_access", {})[str(chat_id)] = "deny" if action == "deny" else "allow"
    save_data()
    await callback.answer("Команды запрещены" if action == "deny" else "Команды разрешены")
    await callback.message.answer(
        "🔴 Команды бота в этой группе теперь запрещены." if action == "deny" else "🟢 Команды бота в этой группе снова работают в штатном режиме.",
        reply_markup=admin_group_actions_keyboard(chat_id, page))


@dp.my_chat_member()
async def remember_bot_group_status(event: ChatMemberUpdated):
    if event.chat.type not in {"group", "supergroup"}:
        return
    status = event.new_chat_member.status
    # Сам факт добавления бота в группу больше не считается установкой.
    # Обновляем статус только для групп, где уже выполняли .установить.
    if _group_is_installed(event.chat.id):
        data.setdefault("known_groups", {}).setdefault(str(event.chat.id), {
            "id": event.chat.id,
            "title": event.chat.title or f"Группа {event.chat.id}",
        })
        data["known_groups"][str(event.chat.id)]["title"] = event.chat.title or f"Группа {event.chat.id}"
        data["known_groups"][str(event.chat.id)]["bot_status"] = status
        save_data()


# =========================================================
# ПОЛЬЗОВАТЕЛИ БОТА / ПРАВА АДМИНИСТРАТОРА
# =========================================================

def admin_users_keyboard(page: int = 0):
    users = list(data.get("users", {}).values())
    users.sort(key=lambda u: (u.get("name") or u.get("username") or str(u.get("id"))).casefold())
    per_page = 8
    pages = max(1, (len(users) + per_page - 1) // per_page)
    page = max(0, min(page, pages - 1))
    rows = []
    for u in users[page * per_page:(page + 1) * per_page]:
        uid = int(u.get("id"))
        badge = "👑" if uid in OWNER_IDS else ("🛡" if is_admin(uid) else "👤")
        label = u.get("name") or ("@" + u.get("username", "") if u.get("username") else str(uid))
        rows.append([InlineKeyboardButton(text=f"{badge} {label}"[:55], callback_data=f"admin_user_{uid}_{page}")])
    nav = []
    if page > 0:
        nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"admin_users_{page-1}"))
    nav.append(InlineKeyboardButton(text=f"{page+1}/{pages}", callback_data="noop"))
    if page + 1 < pages:
        nav.append(InlineKeyboardButton(text="➡️", callback_data=f"admin_users_{page+1}"))
    if nav: rows.append(nav)
    rows.append([InlineKeyboardButton(text="⬅️ Админ-панель", callback_data="admin_panel")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


class InviteCountEdit(StatesGroup):
    waiting_value = State()


def admin_user_actions_keyboard(uid: int, page: int = 0):
    rows = [
        [InlineKeyboardButton(text="➕ Добавить", callback_data=f"admin_inv_add_{uid}_{page}"),
         InlineKeyboardButton(text="➖ Отнять", callback_data=f"admin_inv_sub_{uid}_{page}")],
        [InlineKeyboardButton(text="✏️ Установить число", callback_data=f"admin_inv_set_{uid}_{page}")]
    ]
    if uid not in OWNER_IDS:
        if is_admin(uid):
            rows.append([InlineKeyboardButton(text="🛡 Снять права админа", callback_data=f"admin_demote_{uid}_{page}")])
        else:
            rows.append([InlineKeyboardButton(text="🛡 Сделать админом бота", callback_data=f"admin_promote_{uid}_{page}")])
        rows.append([InlineKeyboardButton(text="🗑 Удалить пользователя", callback_data=f"admin_userdel_{uid}_{page}")])
    rows.append([InlineKeyboardButton(text="⬅️ К пользователям", callback_data=f"admin_users_{page}")])
    return InlineKeyboardMarkup(inline_keyboard=rows)


@dp.callback_query(F.data.startswith("admin_users_"))
async def admin_users_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    page = int(callback.data.rsplit("_", 1)[1])
    await callback.answer()
    await callback.message.answer(
        f"👥 <b>ПОЛЬЗОВАТЕЛИ БОТА</b>\n\nЗарегистрировано: <b>{len(data.get('users', {}))}</b>\n👑 — владелец · 🛡 — админ · 👤 — пользователь\n\nВыбери пользователя:",
        reply_markup=admin_users_keyboard(page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_user_"))
async def admin_user_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True); return
    parts = callback.data.split("_")
    uid, page = int(parts[2]), int(parts[3])
    u = data.get("users", {}).get(str(uid))
    if not u:
        await callback.answer("Пользователь не найден", show_alert=True); return
    role = "Владелец" if uid in OWNER_IDS else ("Администратор бота" if is_admin(uid) else "Пользователь")
    username = f"@{u.get('username')}" if u.get('username') else "—"
    await callback.answer()
    await callback.message.answer(
        f"👤 <b>{u.get('name') or 'Без имени'}</b>\n🆔 <code>{uid}</code>\n🔗 {username}\n🔐 Роль: <b>{role}</b>\n👥 Приглашено: <b>{u.get('total_invited', 0)}</b>",
        reply_markup=admin_user_actions_keyboard(uid, page), parse_mode="HTML")


@dp.callback_query(F.data.startswith("admin_inv_"))
async def admin_invite_edit_start(callback: CallbackQuery, state: FSMContext):
    if not is_admin(callback.from_user.id):
        await callback.answer("❌ Нет доступа", show_alert=True)
        return
    parts = callback.data.split("_")
    action, uid, page = parts[2], int(parts[3]), int(parts[4])
    if str(uid) not in data.get("users", {}):
        await callback.answer("Пользователь не найден", show_alert=True)
        return
    await state.set_state(InviteCountEdit.waiting_value)
    await state.update_data(invite_action=action, invite_uid=uid, invite_page=page)
    labels = {"add": "добавить", "sub": "отнять", "set": "установить"}
    await callback.answer()
    await callback.message.answer(
        f"👥 Сейчас приглашено: <b>{data['users'][str(uid)].get('total_invited', 0)}</b>\n\n"
        f"Введи целое число, которое нужно <b>{labels[action]}</b>.\nНапример: <code>5</code>",
        parse_mode="HTML"
    )


@dp.message(InviteCountEdit.waiting_value)
async def admin_invite_edit_value(message: Message, state: FSMContext):
    if not is_admin(message.from_user.id):
        await state.clear()
        return
    raw = (message.text or "").strip()
    if not raw.isdigit():
        await message.answer("❌ Введи целое неотрицательное число, например <code>5</code>.", parse_mode="HTML")
        return
    value = int(raw)
    ctx = await state.get_data()
    action, uid, page = ctx.get("invite_action"), int(ctx.get("invite_uid")), int(ctx.get("invite_page", 0))
    user = data.get("users", {}).get(str(uid))
    if not user:
        await state.clear(); await message.answer("❌ Пользователь больше не найден."); return
    old = max(0, int(user.get("total_invited", 0) or 0))
    if action == "add": new = old + value
    elif action == "sub": new = max(0, old - value)
    elif action == "set": new = value
    else:
        await state.clear(); await message.answer("❌ Неизвестное действие."); return
    user["total_invited"] = new
    save_data()
    await state.clear()
    await message.answer(
        f"✅ <b>Количество приглашённых изменено</b>\n\nБыло: <b>{old}</b>\nСтало: <b>{new}</b>",
        parse_mode="HTML", reply_markup=admin_user_actions_keyboard(uid, page)
    )


@dp.callback_query(F.data.startswith("admin_promote_"))
async def admin_promote_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid not in ADMIN_IDS: ADMIN_IDS.append(uid)
    admins = {int(x) for x in data.get("bot_admins", []) if str(x).isdigit()}; admins.add(uid)
    data["bot_admins"] = sorted(admins); save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Администратор добавлен")
    await callback.message.answer("✅ Пользователю выданы права администратора бота. Ему доступна админ-панель.", reply_markup=admin_user_actions_keyboard(uid, page))


@dp.callback_query(F.data.startswith("admin_demote_"))
async def admin_demote_callback(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца снять нельзя", show_alert=True); return
    if uid in ADMIN_IDS: ADMIN_IDS.remove(uid)
    data["bot_admins"] = [int(x) for x in data.get("bot_admins", []) if int(x) != uid]
    save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Права сняты")
    await callback.message.answer("✅ Права администратора сняты.", reply_markup=admin_user_actions_keyboard(uid, page))


@dp.callback_query(F.data.startswith("admin_userdel_"))
async def admin_user_delete_confirm(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца удалить нельзя", show_alert=True); return
    kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="✅ Да, удалить", callback_data=f"admin_userdelok_{uid}_{page}"), InlineKeyboardButton(text="❌ Отмена", callback_data=f"admin_user_{uid}_{page}")]])
    await callback.answer(); await callback.message.answer("⚠️ Удалить пользователя из базы бота и его персональные ссылки?", reply_markup=kb)


@dp.callback_query(F.data.startswith("admin_userdelok_"))
async def admin_user_delete(callback: CallbackQuery):
    if not is_admin(callback.from_user.id): await callback.answer("❌ Нет доступа", show_alert=True); return
    _, _, uid_s, page_s = callback.data.split("_"); uid, page = int(uid_s), int(page_s)
    if uid in OWNER_IDS: await callback.answer("Владельца удалить нельзя", show_alert=True); return
    user = data.get("users", {}).pop(str(uid), None)
    # Удаляем созданные им ссылки из общей базы.
    for key in list(data.get("links", {})):
        rec = data["links"][key]
        creator = rec.get("creator_id", rec.get("owner_id", rec.get("user_id"))) if isinstance(rec, dict) else None
        if str(creator) == str(uid): data["links"].pop(key, None)
    if uid in ADMIN_IDS: ADMIN_IDS.remove(uid)
    data["bot_admins"] = [int(x) for x in data.get("bot_admins", []) if int(x) != uid]
    save_data(); init_schedule_system(ADMIN_IDS)
    await callback.answer("Пользователь удалён")
    await callback.message.answer("🗑 Пользователь удалён из базы бота. При следующем /start он зарегистрируется заново.", reply_markup=admin_users_keyboard(page))


@dp.callback_query(F.data == "noop")
async def noop_callback(callback: CallbackQuery):
    await callback.answer()

# =========================================================
# ОБЩАЯ СТАТИСТИКА
# =========================================================

@dp.callback_query(
    F.data == "admin_stats"
)
async def admin_stats_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer()

    total_users = len(
        data["users"]
    )

    total_links = len(
        data["links"]
    )

    total_invited = sum(
        user.get(
            "total_invited",
            0
        )
        for user in data[
            "users"
        ].values()
    )

    await callback.message.answer(
        "📊 <b>ОБЩАЯ СТАТИСТИКА</b>\n\n"
        f"👤 Пользователей бота: "
        f"<b>{total_users}</b>\n"
        f"🔗 Создано ссылок: "
        f"<b>{total_links}</b>\n"
        f"👥 Приглашено: "
        f"<b>{total_invited}</b>",
        parse_mode="HTML"
    )


# =========================================================
# EXCEL
# =========================================================

@dp.callback_query(
    F.data == "export_excel"
)
async def export_excel_callback(
    callback: CallbackQuery
):

    if not is_admin(
        callback.from_user.id
    ):

        await callback.answer(
            "❌ Нет доступа",
            show_alert=True
        )

        return

    await callback.answer(
        "📊 Создаю Excel..."
    )

    try:

        filename = create_excel()

        document = FSInputFile(
            filename
        )

        total_invited = sum(
            user.get(
                "total_invited",
                0
            )
            for user in data[
                "users"
            ].values()
        )

        await callback.message.answer_document(
            document=document,
            caption=(
                "📊 <b>SHINOBI TEAM</b>\n"
                "<b>Реферальный отчёт</b>\n\n"
                f"👤 Пользователей: "
                f"<b>{len(data['users'])}</b>\n"
                f"🔗 Ссылок: "
                f"<b>{len(data['links'])}</b>\n"
                f"👥 Приглашено: "
                f"<b>{total_invited}</b>\n\n"
                "📄 Лист 1 — ссылки\n"
                "📄 Лист 2 — приглашённые"
            ),
            parse_mode="HTML"
        )

    except Exception as error:

        await callback.message.answer(
            "❌ Ошибка создания Excel:\n\n"
            f"<code>{error}</code>",
            parse_mode="HTML"
        )

    finally:

        if os.path.exists(
            EXCEL_FILE
        ):

            try:

                os.remove(
                    EXCEL_FILE
                )

            except OSError:
                pass


# =========================================================
# ЗАПУСК
# =========================================================

async def main():
    print("Удаляю старый webhook...")

    # Удаляем webhook перед запуском polling
    await bot.delete_webhook(drop_pending_updates=True)

    print("Webhook удалён.")
    print("=" * 50)
    print(" SHINOBI TEAM REFERRAL BOT")
    print(" БОТ ЗАПУЩЕН")
    print("=" * 50)

    await dp.start_polling(
        bot,
        allowed_updates=[
            "message",
            "callback_query",
            "chat_member",
            "my_chat_member",
            "chat_join_request"
        ]
    )


if __name__ == "__main__":
    asyncio.run(main())
