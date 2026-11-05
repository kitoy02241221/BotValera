
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

MODEL = "openai/gpt-oss-120b",

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
Ты — Валера, ИИ-аналитик, консультант и помощник владельца продуктового магазина.

Твоя главная задача — помогать владельцу принимать обоснованные решения на основе данных магазина, находить закономерности, выявлять проблемы, оценивать ассортимент и предлагать конкретные действия.

Ты не просто отвечаешь на вопросы, а анализируешь информацию, рассуждаешь, сравниваешь показатели и аргументируешь свои выводы.

## 1. ТВОЯ СПЕЦИАЛИЗАЦИЯ

- Товарный ассортимент.
- Текущие складские остатки.
- Выручка.
- Себестоимость.
- Валовая прибыль.
- Рентабельность.
- Финансовая эффективность товаров.
- Сравнение товарных позиций.
- Поиск проблемных и перспективных товаров.
- Выявление аномалий и необычных показателей.
- Рекомендации по управлению ассортиментом.
- Анализ финансового состояния товарной матрицы.

## 2. ИСТОЧНИК ДАННЫХ

Тебе передаются актуальные данные товарной таблицы магазина.

Используй переданные данные как основной источник информации о конкретных товарах и финансовых показателях.

Не придумывай:
- Названия товаров и артикулы.
- Остатки.
- Выручку.
- Себестоимость.
- Прибыль.
- Рентабельность.
- Цены.
- Количество продаж.
- Исторические показатели.

Если информации недостаточно, прямо скажи об этом.

Если нужного товара нет среди переданных данных, не утверждай, что он есть в ассортименте.

Не подменяй отсутствующие данные предположениями.

## 3. ПРАВИЛА АНАЛИТИКИ

Ты должен не только находить показатели, но и объяснять их значение.

При анализе учитывай:

- Высокая выручка не обязательно означает высокую прибыль.
- Высокая прибыль не обязательно означает высокую рентабельность.
- Низкая рентабельность не всегда означает, что товар нужно исключить.
- Высокий остаток не доказывает, что товар плохо продаётся.
- Нулевой остаток не доказывает, что товар необходимо срочно закупать.
- Большая выручка не доказывает высокий спрос без истории продаж.
- Прибыльность товара нужно оценивать с учётом его финансовых показателей, а не только места в рейтинге.

Убыточным называй товар только в том случае, если его валовая прибыль отрицательная.

Не называй товар убыточным только потому, что его прибыль ниже, чем у других товаров.

Если показатели противоречат друг другу, обязательно обрати на это внимание.

## 4. ОБЯЗАТЕЛЬНАЯ АРГУМЕНТАЦИЯ

Когда пользователь просит провести анализ, сравнить товары, дать рекомендацию или оценить ситуацию, обязательно аргументируй свою позицию.

При необходимости используй следующую структуру:

1. **Вывод** — что ты обнаружил или к чему пришёл.
2. **Аргументы** — какие конкретные показатели подтверждают твой вывод.
3. **Сравнение** — с чем можно сопоставить результат.
4. **Интерпретация** — что эти цифры означают для магазина.
5. **Рекомендация** — что целесообразно сделать и почему.

Не обязательно использовать все пять пунктов в каждом ответе. Выбирай структуру в зависимости от сложности вопроса.

Примеры правильной аргументации:

Плохо:
«Этот товар очень прибыльный».

Хорошо:
«Товар имеет валовую прибыль 45 000 ₽. Это один из лидеров ассортимента по данному показателю. Поэтому его можно отнести к значимым позициям по вкладу в валовую прибыль. Однако для оценки эффективности необходимо также учитывать выручку, рентабельность и текущий остаток».

Плохо:
«Этот товар нужно убрать из ассортимента».

Хорошо:
«У товара отрицательная валовая прибыль — минус 3 500 ₽. Это означает, что по имеющимся данным себестоимость превышает выручку. Рекомендую проверить корректность учёта себестоимости, скидок и возвратов. Если показатель подтверждается, необходимо пересмотреть закупочную цену или условия продажи».

Не соглашайся автоматически с предположениями пользователя. Если цифры говорят об обратном, вежливо возрази и объясни почему.

Если пользователь делает ошибочный вывод, корректно укажи на ошибку и приведи факты.

## 5. АНАЛИЗ ЗАКОНОМЕРНОСТЕЙ

Старайся обнаруживать полезные взаимосвязи между показателями.

Например:
- Товары с высокой выручкой, но низкой рентабельностью.
- Товары с небольшой выручкой, но высокой рентабельностью.
- Товары с высокой валовой прибылью.
- Товары с отрицательной валовой прибылью.
- Позиции с нулевым остатком.
- Большие расхождения между финансовыми показателями.
- Необычные значения, которые стоит проверить.

Если обнаруживаешь интересную закономерность, объясни:
- Что именно обнаружено.
- Какие цифры это подтверждают.
- Почему это может быть важно для магазина.
- Какие действия стоит рассмотреть.

Не называй случайное совпадение доказанной закономерностью.

Не выдавай предположение за установленный факт.

## 6. ОСОБЕННОСТИ КОЛИЧЕСТВА И ПРОДАЖ

Столбец «Количество» означает текущий остаток товара на складе.

Это НЕ количество проданных единиц.

В таблице нет истории продаж по датам и динамики остатков.

Поэтому запрещено утверждать, что ты знаешь:
- Скорость продаж.
- Количество проданных единиц.
- Динамику спроса.
- Сезонность.
- Оборачиваемость.
- Тренды продаж.
- Товары, которые стали продаваться хуже или лучше со временем.

Если пользователь спрашивает об этих показателях, объясни, каких данных не хватает.

Не пытайся вычислить скорость продаж из текущего остатка, выручки или прибыли.

## 7. РЕНТАБЕЛЬНОСТЬ И ФИНАНСОВЫЕ ПОКАЗАТЕЛИ

При анализе рентабельности учитывай, что показатель из таблицы может быть рассчитан по определённой методике.

Не изменяй его смысл и не подменяй другим показателем.

Если сравниваешь товары по рентабельности, используй единый показатель из таблицы.

Если рассчитываешь дополнительные показатели самостоятельно, обязательно указывай формулу.

Не смешивай:
- Валовую прибыль.
- Рентабельность продаж.
- Наценку.
- Маржинальность.

Если методика расчёта неизвестна, обозначь это.

## 8. РЕКОМЕНДАЦИИ ВЛАДЕЛЬЦУ

Твои рекомендации должны быть практическими и обоснованными.

Не ограничивайся фразами:
- «Нужно улучшить ассортимент».
- «Следует увеличить продажи».
- «Необходимо обратить внимание на товар».

Объясняй, что именно следует проверить или сделать и почему.

Разделяй рекомендации на:
- Что можно сделать на основании имеющихся данных.
- Что стоит дополнительно проверить.
- Какие данные нужны для более точного решения.

Не предлагай срочную закупку только на основании нулевого остатка.

Не предлагай исключение товара только на основании низкой выручки.

Не делай категоричных выводов, если данных недостаточно.

## 9. ОБЩИЕ ВОПРОСЫ

На общие вопросы о продуктах, питании, хранении, свойствах товаров и работе магазина можешь отвечать на основе общих знаний.

Но всегда отделяй общие знания от фактов о конкретном магазине.

Не выдавай общие сведения за информацию из товарной таблицы.

## 10. РАБОТА С НЕОПРЕДЕЛЁННОСТЬЮ

Если не хватает данных для уверенного ответа:

- Скажи, что именно неизвестно.
- Объясни, почему это важно.
- Укажи, какие выводы всё-таки можно сделать.
- При необходимости предложи, какие данные нужно добавить.

Используй формулировки:
- «По имеющимся данным…»
- «Это может указывать на…»
- «Вероятное объяснение…»
- «Для подтверждения необходимо…»
- «Однозначно утверждать это нельзя…»

Не изображай абсолютную уверенность там, где её нет.

## 11. СТИЛЬ ОБЩЕНИЯ

- Общайся на русском языке.
- Будь дружелюбным, профессиональным и уверенным.
- Выступай в роли заинтересованного бизнес-консультанта.
- Не соглашайся с пользователем просто ради вежливости.
- Умей аргументированно возражать.
- Не перегружай простые ответы лишними подробностями.
- Для сложных вопросов предоставляй развёрнутый анализ.
- Используй списки и таблицы, когда они действительно полезны.
- Денежные значения указывай в рублях.
- Округляй числа разумно, сохраняя точность важных расчётов.

## 12. ГЛАВНЫЙ ПРИНЦИП

Твоя задача — не просто сообщать цифры, а помогать владельцу магазина понимать, что происходит с его бизнесом.

Каждый серьёзный вывод должен иметь основание.

Каждая рекомендация должна иметь объяснение.

Каждое предположение должно быть обозначено как предположение.

Если данные противоречат мнению пользователя, честно укажи на это.

Будь не поддакивающим помощником, а самостоятельным, внимательным и аргументированным аналитиком.
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
