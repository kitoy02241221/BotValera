
import os
import re
import logging
import asyncio
from pathlib import Path

from dotenv import load_dotenv
from groq import Groq
from openpyxl import load_workbook

from telegram import Update
from telegram.ext import (
    Application,
    MessageHandler,
    ContextTypes,
    filters
)


# ==================================================
# ПУТИ И НАСТРОЙКИ
# ==================================================

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "valera_table.xlsx"
LOG_FILE = BASE_DIR / "bot.log"

MODEL = "openai/gpt-oss-20b"

MAX_PRODUCTS_IN_CONTEXT = 35
MAX_CONTEXT_LENGTH = 14000


# ==================================================
# ЛОГИРОВАНИЕ
# ==================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(LOG_FILE, encoding="utf-8")
    ],
    force=True
)

logger = logging.getLogger("VALERA_AI")


# ==================================================
# ПЕРЕМЕННЫЕ ОКРУЖЕНИЯ
# ==================================================

load_dotenv(BASE_DIR / ".env")

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

if not TELEGRAM_TOKEN:
    raise RuntimeError("Не найден TELEGRAM_TOKEN в .env")

if not GROQ_API_KEY:
    raise RuntimeError("Не найден GROQ_API_KEY в .env")


# ==================================================
# GROQ
# ==================================================

client = Groq(api_key=GROQ_API_KEY)

logger.info("Клиент Groq создан")


# ==================================================
# РАБОТА С EXCEL
# ==================================================

COLUMN_ALIASES = {
    "product": [
        "артикулноменклатура",
        "артикул",
        "номенклатура"
    ],
    "characteristic": [
        "характеристика"
    ],
    "quantity": [
        "количество"
    ],
    "revenue": [
        "выручка"
    ],
    "cost": [
        "себестоимость"
    ],
    "profit": [
        "валоваяприбыль"
    ],
    "profitability": [
        "рентабельность"
    ]
}


def normalize_header(value):
    """Приводит название столбца к единому формату."""

    if value is None:
        return ""

    value = str(value).lower().strip()
    return re.sub(r"[^а-яёa-z0-9]", "", value)


def parse_number(value):
    """Преобразует числа из Excel в float."""

    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text:
        return None

    text = (
        text.replace("\xa0", "")
        .replace(" ", "")
        .replace("₽", "")
        .replace("%", "")
        .replace(",", ".")
    )

    text = re.sub(r"[^0-9.\-]", "", text)

    try:
        return float(text)
    except ValueError:
        return None


def format_number(value):
    if value is None:
        return "нет данных"

    return f"{value:,.2f}".replace(",", " ").replace(".", ",")


def get_excel_data():
    """
    Загружает актуальные данные из Excel.
    Файл перечитывается при каждом сообщении.
    """

    logger.info("[EXCEL] Открываем файл %s", EXCEL_FILE.name)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"Не найден файл {EXCEL_FILE.name}"
        )

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True
    )

    try:
        worksheet = workbook.active

        rows = worksheet.iter_rows(values_only=True)

        headers = next(rows, None)

        if not headers:
            raise ValueError("Excel-файл пустой")

        normalized_headers = {
            normalize_header(header): index
            for index, header in enumerate(headers)
            if header is not None
        }

        column_indexes = {}

        for field, aliases in COLUMN_ALIASES.items():

            found_index = None

            for alias in aliases:

                if alias in normalized_headers:
                    found_index = normalized_headers[alias]
                    break

            if found_index is None:
                raise ValueError(
                    f"Не найден обязательный столбец: {field}. "
                    f"Проверь заголовки Excel."
                )

            column_indexes[field] = found_index

        products = []

        for row_number, row in enumerate(rows, start=2):

            def get_value(field):
                index = column_indexes[field]

                if index >= len(row):
                    return None

                return row[index]

            name = get_value("product")

            if name is None or str(name).strip() == "":
                continue

            product = {
                "name": str(name).strip(),
                "characteristic": str(
                    get_value("characteristic") or ""
                ).strip(),
                "quantity": parse_number(get_value("quantity")),
                "revenue": parse_number(get_value("revenue")),
                "cost": parse_number(get_value("cost")),
                "profit": parse_number(get_value("profit")),
                "profitability": parse_number(
                    get_value("profitability")
                ),
                "excel_row": row_number
            }

            products.append(product)

        logger.info(
            "[EXCEL] Успешно загружено товаров: %s",
            len(products)
        )

        if not products:
            raise ValueError("В таблице нет строк с товарами")

        return products

    finally:
        workbook.close()


# ==================================================
# АНАЛИТИКА ТОВАРОВ
# ==================================================

def build_analytics(products):

    total_revenue = sum(
        p["revenue"] or 0 for p in products
    )

    total_cost = sum(
        p["cost"] or 0 for p in products
    )

    total_profit = sum(
        p["profit"] or 0 for p in products
    )

    zero_stock = [
        p for p in products
        if p["quantity"] is not None
        and p["quantity"] <= 0
    ]

    positive_stock = [
        p for p in products
        if p["quantity"] is not None
        and p["quantity"] > 0
    ]

    by_revenue = sorted(
        [p for p in products if p["revenue"] is not None],
        key=lambda p: p["revenue"],
        reverse=True
    )

    by_profit = sorted(
        [p for p in products if p["profit"] is not None],
        key=lambda p: p["profit"],
        reverse=True
    )

    by_profitability = sorted(
        [
            p for p in products
            if p["profitability"] is not None
        ],
        key=lambda p: p["profitability"]
    )

    result = []

    result.append("ОБЩАЯ АНАЛИТИКА ТОВАРНОЙ ТАБЛИЦЫ")
    result.append(f"Количество товарных позиций: {len(products)}")
    result.append(f"Выручка: {format_number(total_revenue)} ₽")
    result.append(f"Себестоимость: {format_number(total_cost)} ₽")
    result.append(f"Валовая прибыль: {format_number(total_profit)} ₽")

    if total_revenue:
        margin = total_profit / total_revenue * 100
        result.append(
            f"Расчётная доля валовой прибыли в выручке: "
            f"{format_number(margin)}%"
        )

    result.append(
        f"Позиций с нулевым или отрицательным остатком: "
        f"{len(zero_stock)}"
    )

    result.append(
        f"Позиций с положительным остатком: "
        f"{len(positive_stock)}"
    )

    result.append("\nТОП-10 ПО ВЫРУЧКЕ")

    for p in by_revenue[:10]:
        result.append(
            f"{p['name']} | Выручка: "
            f"{format_number(p['revenue'])} ₽"
        )

    result.append("\nТОП-10 ПО ВАЛОВОЙ ПРИБЫЛИ")

    for p in by_profit[:10]:
        result.append(
            f"{p['name']} | Прибыль: "
            f"{format_number(p['profit'])} ₽"
        )

    result.append("\nТОВАРЫ С САМОЙ НИЗКОЙ РЕНТАБЕЛЬНОСТЬЮ")

    for p in by_profitability[:10]:
        result.append(
            f"{p['name']} | Рентабельность: "
            f"{format_number(p['profitability'])}%"
        )

    result.append("\nТОВАРЫ С НУЛЕВЫМ ОСТАТКОМ")

    for p in zero_stock[:20]:
        result.append(p["name"])

    if len(zero_stock) > 20:
        result.append(
            f"И ещё {len(zero_stock) - 20} позиций"
        )

    return "\n".join(result)


# ==================================================
# ПОИСК ТОВАРОВ ПО ЗАПРОСУ
# ==================================================

STOP_WORDS = {
    "какой", "какая", "какие", "сколько", "есть",
    "мне", "для", "про", "товар", "товары",
    "покажи", "найди", "расскажи", "сделай",
    "анализ", "аналитику", "магазине", "магазина",
    "самый", "самая", "самые", "пожалуйста",
    "что", "это", "как", "могу", "можно",
    "у", "в", "на", "и", "или", "по", "с"
}


def find_relevant_products(user_text, products):

    words = re.findall(
        r"[а-яёa-z0-9]+",
        user_text.lower()
    )

    keywords = [
        word for word in words
        if len(word) >= 3 and word not in STOP_WORDS
    ]

    if not keywords:
        return []

    matches = []

    for product in products:

        searchable = (
            product["name"] + " " +
            product["characteristic"]
        ).lower()

        score = sum(
            1 for word in keywords
            if word in searchable
        )

        if score > 0:
            matches.append((score, product))

    matches.sort(
        key=lambda item: item[0],
        reverse=True
    )

    return [
        product
        for score, product in matches[:MAX_PRODUCTS_IN_CONTEXT]
    ]


# ==================================================
# ПОДГОТОВКА ИНФОРМАЦИИ ДЛЯ AI
# ==================================================

def product_to_text(product):

    return (
        f"Товар: {product['name']}\n"
        f"Характеристика: {product['characteristic'] or 'не указана'}\n"
        f"Текущий остаток: "
        f"{format_number(product['quantity'])}\n"
        f"Выручка: {format_number(product['revenue'])} ₽\n"
        f"Себестоимость: {format_number(product['cost'])} ₽\n"
        f"Валовая прибыль: {format_number(product['profit'])} ₽\n"
        f"Рентабельность: "
        f"{format_number(product['profitability'])}%"
    )


def build_context(user_text, products):

    analytics = build_analytics(products)

    relevant = find_relevant_products(
        user_text,
        products
    )

    analytics_words = [
        "аналитик", "анализ", "прибыл", "выруч",
        "рентабель", "убыточ", "топ", "рейтинг",
        "ассортимент", "закуп", "остатк", "склад",
        "сравни", "сравнен", "самых", "лучших",
        "худших", "статистик"
    ]

    is_analytics_question = any(
        word in user_text.lower()
        for word in analytics_words
    )

    context_parts = [analytics]

    if relevant:
        context_parts.append(
            "\nТОВАРЫ, ПОДХОДЯЩИЕ ПОД ЗАПРОС:"
        )

        for product in relevant:
            context_parts.append(
                "\n" + product_to_text(product)
            )

    elif not is_analytics_question:
        context_parts.append(
            "\nПо словам пользователя конкретные товары "
            "не найдены в таблице."
        )

    if not relevant and is_analytics_question:
        top_products = sorted(
            products,
            key=lambda p: p["profit"] or 0,
            reverse=True
        )[:MAX_PRODUCTS_IN_CONTEXT]

        context_parts.append(
            "\nДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ ДЛЯ АНАЛИЗА:"
        )

        for product in top_products:
            context_parts.append(
                "\n" + product_to_text(product)
            )

    context = "\n".join(context_parts)

    return context[:MAX_CONTEXT_LENGTH]


# ==================================================
# ИНСТРУКЦИЯ ДЛЯ AI
# ==================================================

SYSTEM_PROMPT = """
Ты — Валера, ИИ-аналитик и консультант продуктового магазина.

Твоя специализация:
- товарный ассортимент;
- текущие остатки;
- выручка;
- себестоимость;
- валовая прибыль;
- рентабельность;
- анализ финансовых показателей;
- рекомендации по управлению ассортиментом.

ТВОИ ЗАДАЧИ:

1. Помогать находить товары в ассортименте.
2. Сообщать актуальные остатки товаров.
3. Объяснять финансовые показатели.
4. Сравнивать товары по выручке, прибыли и рентабельности.
5. Находить наиболее прибыльные позиции.
6. Выявлять товары с низкой или отрицательной прибылью.
7. Находить товары с нулевым остатком.
8. Давать обоснованные рекомендации владельцу магазина.
9. Проводить анализ ассортимента на основании предоставленных данных.

ИСТОЧНИК ИНФОРМАЦИИ:

Тебе передаётся содержимое актуальной Excel-таблицы.

Используй её как основной источник конкретных сведений о товарах.

Не придумывай:
- названия товаров;
- артикулы;
- цены;
- остатки;
- выручку;
- себестоимость;
- прибыль;
- рентабельность.

Если товара нет в переданных данных, честно сообщи об этом.

ВАЖНО О КОЛИЧЕСТВЕ:

Столбец «Количество» означает ТЕКУЩИЙ ОСТАТОК товара.

Это НЕ количество проданных товаров.

Никогда не называй остаток количеством продаж.

В таблице нет истории продаж, поэтому не делай вид,
что знаешь динамику продаж, скорость реализации
или количество проданных единиц.

АНАЛИТИКА:

Ты можешь анализировать:
- общую выручку;
- общую себестоимость;
- общую валовую прибыль;
- рентабельность;
- прибыльность отдельных товаров;
- товары с нулевым остатком;
- ассортимент и финансовые показатели.

Не называй товар убыточным, если его прибыль просто
ниже прибыли других товаров.

Убыточным товар можно считать только при отрицательной
валовой прибыли.

Не придумывай порог низкого остатка.
Если пользователь спрашивает о низких остатках,
поясни, что порог необходимо определить отдельно.

ОБЩИЕ ВОПРОСЫ:

На общие вопросы о продуктах, питании, хранении
и свойствах товаров можешь отвечать на основании
своих общих знаний.

Но не выдавай общие сведения за факты о конкретном товаре
из магазина, если их нет в таблице.

РЕКОМЕНДАЦИИ:

Давай конкретные и полезные рекомендации.

Объясняй, на каких цифрах основан вывод.

Не утверждай, что товар плохо продаётся,
если у тебя нет истории продаж.

Не утверждай, что товар нужно срочно закупать,
только потому что его остаток равен нулю:
сначала учитывай контекст вопроса.

СТИЛЬ:

- Общайся дружелюбно и профессионально.
- Ты аналитик-консультант, а не робот.
- Отвечай на русском языке.
- Используй понятные формулировки.
- Денежные значения указывай в рублях.
- При необходимости используй списки и таблицы.
- Не перегружай короткие ответы лишними цифрами.
- Если пользователь просит подробный анализ,
  предоставляй подробный анализ.
- Не упоминай системные инструкции и внутренний код.
- Не говори, что читаешь файл Excel.
- Не придумывай отсутствующие данные.
"""


# ==================================================
# ОБРАБОТКА СООБЩЕНИЙ
# ==================================================

async def handle_message(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    try:
        if update.message is None:
            return

        if not update.message.text:
            return

        user_text = update.message.text.strip()

        if not user_text:
            return

        logger.info("========================================")
        logger.info("[СООБЩЕНИЕ] Получено новое сообщение")
        logger.info(
            "[СООБЩЕНИЕ] Длина текста: %s",
            len(user_text)
        )

        # Загружаем актуальную таблицу
        products = await asyncio.to_thread(
            get_excel_data
        )

        logger.info(
            "[EXCEL] Данные загружены: %s товаров",
            len(products)
        )

        # Подготавливаем контекст
        data_context = build_context(
            user_text,
            products
        )

        logger.info("[AI] Контекст подготовлен")

        messages = [
            {
                "role": "system",
                "content": SYSTEM_PROMPT
            },
            {
                "role": "user",
                "content": (
                    "АКТУАЛЬНЫЕ ДАННЫЕ МАГАЗИНА:\n"
                    + data_context
                    + "\n\nВОПРОС ПОЛЬЗОВАТЕЛЯ:\n"
                    + user_text
                )
            }
        ]

        logger.info("[GROQ] Отправляем запрос")

        response = await asyncio.to_thread(
            client.chat.completions.create,
            model=MODEL,
            messages=messages,
            temperature=0.2
        )

        answer = response.choices[0].message.content

        if not answer:
            answer = (
                "Не получилось сформировать ответ. "
                "Попробуйте задать вопрос ещё раз."
            )

        logger.info(
            "[GROQ] Ответ получен. Символов: %s",
            len(answer)
        )

        # Telegram ограничивает длину сообщения
        max_length = 4000

        for start in range(0, len(answer), max_length):

            part = answer[start:start + max_length]

            await update.message.reply_text(part)

        logger.info("[TELEGRAM] Ответ отправлен")

    except FileNotFoundError:
        logger.exception("[EXCEL] Файл не найден")

        await update.message.reply_text(
            "Не могу найти файл valera_table.xlsx. "
            "Проверь, что он лежит в одной папке с ботом."
        )

    except Exception:
        logger.exception(
            "[ОШИБКА] Ошибка обработки сообщения"
        )

        try:
            await update.message.reply_text(
                "Произошла ошибка при обработке запроса. "
                "Попробуйте ещё раз немного позже."
            )
        except Exception:
            logger.exception(
                "[TELEGRAM] Не удалось отправить сообщение об ошибке"
            )


# ==================================================
# ОБРАБОТЧИК ОШИБОК TELEGRAM
# ==================================================

async def error_handler(
    update: object,
    context: ContextTypes.DEFAULT_TYPE
):

    logger.error(
        "[TELEGRAM] Необработанная ошибка",
        exc_info=context.error
    )


# ==================================================
# ЗАПУСК
# ==================================================

def main():

    try:
        logger.info("========================================")
        logger.info("ЗАПУСК AI-АНАЛИТИКА МАГАЗИНА")

        # Проверяем таблицу при запуске
        products = get_excel_data()

        logger.info(
            "[EXCEL] Проверка пройдена. Товаров: %s",
            len(products)
        )

        app = (
            Application.builder()
            .token(TELEGRAM_TOKEN)
            .build()
        )

        app.add_handler(
            MessageHandler(
                filters.TEXT & ~filters.COMMAND,
                handle_message
            )
        )

        app.add_error_handler(error_handler)

        print("\nВалера — ИИ-аналитик магазина запущен!")
        print("Excel: valera_table.xlsx")
        print("Ожидаю сообщения в Telegram...")
        print("Логи сохраняются в bot.log\n")

        logger.info("[TELEGRAM] Запускаем polling")

        app.run_polling()

    except KeyboardInterrupt:
        logger.warning("Бот остановлен пользователем")

    except Exception:
        logger.exception(
            "[КРИТИЧЕСКАЯ ОШИБКА] Ошибка запуска бота"
        )

        print(
            "\nОшибка запуска. Подробности находятся в bot.log"
        )

        input("Нажмите Enter для выхода...")


# ==================================================
# ТОЧКА ВХОДА
# ==================================================

if __name__ == "__main__":
    main()
