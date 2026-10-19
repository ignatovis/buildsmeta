#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Маппинг файла "Тест для маппинга.xlsx":
  - Лист: Согласованная Смета
  - A = № п/п, B = № раздела, C = Кодификатор (записываем сюда), D = Наименование

Алгоритм (v4, правила v2.0):
  1. keyword rules с SMR-контекстом родительского раздела (column B)
  2. Exact match по примерам
  3. Fuzzy match (порог 85) с фильтром по SMR-домену
  4. Наследование SMR-кода родительского раздела из иерархии column B

  Изменения v4 (2026-05-05):
  - Раздел 19 убран: реставрация → 6.8/6.9/7.4/8.5/8.3.2-8.3.4
  - Раздел 2: БНС(несущие) → 2.1, БСС(ограждение)/шпунт → 2.4
  - Добавлен 8.5 (реставрация внутренняя)
  - Добавлены разделы 15, 17, 20
  - Наружное освещение: 11.11 vs 12.99 по контексту предка
"""
import sys, re, os, argparse, openpyxl
from fuzzywuzzy import fuzz, process

sys.stdout.reconfigure(encoding="utf-8")

_parser = argparse.ArgumentParser(add_help=False)
_parser.add_argument("-i", "--input",  default="Тест для маппинга.xlsx")
_parser.add_argument("-o", "--output", default="Тест для маппинга_READY.xlsx")
_args, _ = _parser.parse_known_args()

INPUT_FILE   = _args.input
OUTPUT_FILE  = _args.output
EXAMPLES_DIR = "Пример смет"

CODE_RE = re.compile(r"^(\d+\.(?:\d+\.)*)")


def extract_code(text):
    if not text:
        return ""
    m = CODE_RE.match(str(text).strip())
    if m:
        return m.group(1)
    raw = str(text).strip()
    if re.match(r"^\d+(\.\d+)*\.?$", raw):
        if not raw.endswith("."):
            raw += "."
        return raw
    return ""


def norm(text):
    return str(text).strip().lower() if text is not None else ""


def contains(name_n, *keywords):
    return any(kw in name_n for kw in keywords)


def word_in(name_n, word):
    """Проверяет наличие слова/аббревиатуры как отдельного токена (не как подстроку)."""
    return bool(re.search(r'(?<![а-яёa-z0-9])' + re.escape(word) + r'(?![а-яёa-z0-9])', name_n))


# ── Фильтр нулевых строк по столбцу стоимости ──────────────────────────────

COST_HEADERS = ["всего", "итого", "стоимость", "сумма", "общая стоимость"]


def norm_header(text):
    """Нормализует заголовок: строчные буквы, переносы строк → пробел."""
    if text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text)).strip().lower()


def find_cost_col(ws):
    """Ищет 0-based индекс первого столбца стоимости по заголовку в первых 3 строках.
    Возвращает None если подходящий заголовок не найден."""
    for row in ws.iter_rows(min_row=1, max_row=3, values_only=True):
        for j, cell in enumerate(row):
            h = norm_header(cell)
            if h and any(kw in h for kw in COST_HEADERS):
                return j
    return None


def is_zero_cost(val):
    """True если стоимость строки = пусто / прочерк / 0."""
    if val is None:
        return True
    s = str(val).strip()
    if not s or s in ("-", "–", "—"):
        return True
    try:
        return float(s.replace(",", ".").replace(" ", "").replace("\xa0", "")) == 0.0
    except (ValueError, TypeError):
        return False


def smr_top(smr_code):
    """Возвращает первый уровень SMR кода: '12.1.' → '12', '2.6.' → '2'"""
    if not smr_code:
        return ""
    return smr_code.rstrip(".").split(".")[0]


def get_parent_b(b_val):
    """Возвращает родительский номер раздела: '11.2.5' → '11.2', '11' → None"""
    if not b_val:
        return None
    parts = str(b_val).split(".")
    if len(parts) <= 1:
        return None
    return ".".join(parts[:-1])


def find_ancestor_smr(b_val, b_smr_map):
    """Ищет SMR-код ближайшего предка в иерархии column B."""
    cur = get_parent_b(b_val)
    while cur is not None:
        if b_smr_map.get(cur):
            return b_smr_map[cur]
        cur = get_parent_b(cur)
    return ""


def smr_compatible(code, smr_ctx_code):
    """True если два SMR-кода совместимы в иерархии."""
    if not code or not smr_ctx_code:
        return True
    top_code = smr_top(code)
    top_ctx  = smr_top(smr_ctx_code)
    if top_code != top_ctx:
        return False
    # Для раздела 1 проверяем подраздел строго: 1.2. ≠ 1.3.
    if top_code == "1":
        parts_ctx  = smr_ctx_code.rstrip(".").split(".")
        parts_code = code.rstrip(".").split(".")
        if len(parts_ctx) >= 2 and len(parts_code) >= 2:
            # "99" — прочие, совместимы с любым подразделом
            if parts_code[1] == "99":
                return True
            if parts_ctx[1] != parts_code[1]:
                return False
    return True


def smr_correct(name, smr_ctx_code):
    """
    Переопределяет код с учётом SMR-контекста родительского раздела.
    Вызывается когда keyword дал результат несовместимый с контекстом.
    """
    n = norm(name)
    top = smr_top(smr_ctx_code)

    if top == "12":  # Благоустройство
        if contains(n, "земляные", "засыпка", "грунт", "выемка", "насыпь", "вывоз"):
            return "12.1."
        if contains(n, "асфальт", "тротуар", "покрытие", "плитка", "брусчатка",
                    "бортовой камень", "твёрд", "тверд", "дорог", "проезд",
                    "резиновая крошка", "epdm", "террасн"):
            return "12.2."
        if contains(n, "газон", "озеленен", "деревь", "кустарник", "растени",
                    "решетка", "газонная", "плодородн"):
            return "12.3."
        if contains(n, "маф", "скамья", "урна", "фонтан", "велопарковк", "площадка",
                    "качел", "горк"):
            return "12.4."
        return smr_ctx_code   # наследуем контекст

    if top == "2":  # Котлован
        if contains(n, "земляные", "разработк", "вывоз", "засыпка", "грунт", "выемка"):
            return "2.6."
        return smr_ctx_code

    return smr_ctx_code  # для прочих контекстов просто наследуем


def map_code(name_orig, smr_ctx_code=""):
    """
    Возвращает SMR-кодификатор для строки.
    smr_ctx_code — SMR-код родительского раздела.
    """
    if not name_orig:
        return ""
    n = norm(name_orig)
    ctx = smr_top(smr_ctx_code)

    # ── РЕСТАВРАЦИЯ: высший приоритет (раздел 19 убран, коды по типу работ) ──
    _is_restav = contains(n, "реставрац", "воссоздани", "вычинка",
                           "инъектировани", "инъецировани",
                           "расчистка историческ")
    if _is_restav:
        if contains(n, "кровля", "кровли"):
            if contains(n, "демонтаж"):
                return "7.4.1."
            if contains(n, "воссоздани"):
                return "7.4.2."
            if contains(n, "элемент"):
                return "7.4.3."
            if contains(n, "фронтон"):
                return "7.4.6."
            return "7.4."
        if contains(n, "окно", "витраж"):
            if contains(n, "демонтаж"):
                return "6.9.1."
            if contains(n, "воссоздани", "монтаж"):
                return "6.9.3."
            return "6.9."
        if contains(n, "двер"):
            if contains(n, "воссоздани", "монтаж"):
                return "8.3.3."
            return "8.3.2."
        if contains(n, "ворот"):
            return "8.3.4."
        if contains(n, "фасад", "кладк", "лепнин", "кирпич", "штукатурн"):
            if contains(n, "демонтаж"):
                return "6.8.1."
            if contains(n, "кладка", "кладк"):
                return "6.8.3."
            if contains(n, "штукатурн"):
                return "6.8.4."
            if contains(n, "декоратив", "лепнин"):
                return "6.8.5."
            if contains(n, "крылец", "крыльц"):
                return "6.8.6."
            if contains(n, "навес"):
                return "6.8.7."
            if contains(n, "ограждени", "перил"):
                return "6.8.8."
            return "6.8."
        if contains(n, "интерьер", "внутренн", "потолок", "пол", "стена", "свод",
                    "паркет", "лестниц"):
            if contains(n, "демонтаж"):
                return "8.5.1."
            if contains(n, "усилени"):
                return "8.5.2."
            if contains(n, "декоратив"):
                return "8.5.5."
            if contains(n, "потолок"):
                return "8.5.6."
            if contains(n, "пол"):
                return "8.5.7."
            if contains(n, "стена"):
                return "8.5.8."
            if contains(n, "лестниц"):
                return "8.5.10."
            return "8.5."
        return "6.8.99."

    # ── РАЗДЕЛ 16: РД / 18: Наружные сети ────────────────────────────────────
    if contains(n, "рабочая документация", "авторский надзор",
                "рабочее проектировани", "архитектурные решения",
                "конструктивные и объемно", "схема планировочн"):
        return "16."
    if contains(n, "наружные сети"):
        return "18."

    # ── РАЗДЕЛ 12: Благоустройство (до раздела 11!) ───────────────────────────
    # Проверяем здесь, чтобы "благоустройство и наружное освещение" не попало в 11.11.
    if contains(n, "благоустройств"):
        return "12."
    if ctx == "12":  # Если мы в контексте благоустройства — применяем правила сразу
        if contains(n, "тротуар", "брусчатка", "бортовой камень", "асфальт",
                    "твёрдое покрытие", "твердое покрытие"):
            return "12.2."
        # газонная решетка как твёрдое покрытие (с гравием, щебнем) → 12.2.
        if "газонная решетка" in n and contains(n, "гравие", "щебен", "заполнени"):
            return "12.2."
        if contains(n, "озеленени", "деревья", "газон"):
            return "12.3."
        if contains(n, "маф", "детские площадки", "велопарковка", "фонтан"):
            return "12.4."
        if contains(n, "земляные работы", "земляные", "вывоз грунта", "вывоз и утилизаци",
                    "обратная засыпка", "разработка грунта", "выемка грунта",
                    "насыпь", "уплотненный грунт"):
            return "12.1."

    # ── РАЗДЕЛ 1: Подготовительные работы ────────────────────────────────────
    # Сводный заголовок раздела → общий код "1."
    if n.startswith("подготовительные работы"):
        return "1."
    if contains(n, "банковская гарантия", "обеспечение финансов"):
        return "1.1."
    if contains(n, "страховани"):
        return "1.1."
    if contains(n, "мобилизац", "демобилизац"):
        return "1.2."
    if contains(n, "бытовой городок", "строительный городок",
                "временные офис", "временный офис", "ограждение строит",
                "ограждение площадк", "ограждение стройплощадк",
                "ограждение участк", "пункт мойки", "временные дороги",
                "временных внутриплощадочн", "обустройство помещений",
                "штаб строительства", "временная подсветка",
                "перенос крун", "перенос ктпн"):
        return "1.2."
    if contains(n, "содержание площадки", "содержание временных",
                "ежемесячная уборка", "уборка снега",
                "меры предосторожности в зимн", "вывоз строительного мусора",
                "ручные инструменты", "механизац", "кран башенн",
                "коммунальн", "гарантийн", "охрана площадки",
                "ежемесячная охрана", "генподрядн", "пропускная система",
                "биометрический скуд", "обслуживание системы"):
        return "1.3."
    if contains(n, "итр ", "управление проектом", "вспомогательный персонал"):
        return "1.3."
    if contains(n, "геодезич"):
        return "1.3."
    if contains(n, "организация работ по проведению", "плановых проверок",
                "согласование и совместный приказ", "открытие совместного приказ",
                "техничес надзор"):
        return "1.3."
    if contains(n, "мокап"):
        return "1.4."
    if contains(n, "организация строительств", "этапность"):
        return "1.5."
    if contains(n, "ппр", "технологических карт", "лабораторн испытани"):
        return "1.3."
    if contains(n, "восстановление покрытия", "защита резиновым",
                "защита металлическими", "содержание, ремонт и защита"):
        return "1.99."

    # ── РАЗДЕЛ 2: Котлован ────────────────────────────────────────────────────
    # БНС (несущие буровые сваи) → 2.1; БСС / шпунт / стена в грунте → 2.4
    if contains(n, "буросекущ", "бсс", "шпунт", "стена в грунте", "свг",
                "ограждение котлован", "ограждающих конструкц котлован"):
        return "2.4."
    if contains(n, "буронабивн", "бнс", "свайное основани", "устройство свай",
                "форшахта") and not contains(n, "ограждени"):
        return "2.1."
    if contains(n, "срез оголовк", "срубка оголовк"):
        return "2.2."
    if contains(n, "распорная система", "извлечение шпунта"):
        return "2.3."
    if contains(n, "строительное водопонижени", "верхностный водоотлив",
                "строительный водоотли"):
        return "2.5."
    if contains(n, "земляные работы", "разработка грунта", "вывоз грунта",
                "перемещение грунта", "обратная засыпка", "устройство котлован",
                "доработка грунта", "вывоз и утилизаци"):
        if ctx == "12":
            return "12.1."
        return "2.6."
    if contains(n, "котлован"):
        return "2."
    if contains(n, "демонтаж перекрытия в канале", "демонтаж бсс",
                "демонтаж и переустройство подземных",
                "демонтаж надземных сооружений тп", "демонтаж конструкций"):
        return "1.99."
    if contains(n, "демонтаж", "ручная разборка", "ручной демонтаж"):
        return "1.2."
    if contains(n, "геологические изыскания"):
        return "1.99."
    if contains(n, "ручная разборка здания", "снос"):
        return "1.99."

    # ── РАЗДЕЛ 3: Гидроизоляция ───────────────────────────────────────────────
    if contains(n, "виброизол", "виброзащитн"):
        return "3.3."
    if contains(n, "подготовка основания", "уплотнение основания",
                "щебеночное основание"):
        return "3.1."
    if contains(n, "гидроизол"):
        if contains(n, "кровл"):
            return "7."
        return "3.2."
    if contains(n, "дренаж"):
        return "3.99."

    # ── РАЗДЕЛ 4: Конструктивные решения ──────────────────────────────────────
    if contains(n, "металлоконструкц"):
        return "4.3."
    if contains(n, "усиление фундаментов", "усиление здания",
                "усиление конструкций", "усиление вертикальн",
                "усиление сводов", "устройство металлоконструкц в подземн",
                "устройство ростверка", "устройство закладных"):
        return "4.99."
    if contains(n, "самонесущие ограждающ", "самонесущих конструкц"):
        return "4.99."
    if contains(n, "ж/б конструкции", "монолитн конструкц", "монолитные работы",
                "возведение несущих", "несущих конструкций"):
        if contains(n, "подземн"):
            if contains(n, "горизонтальн", "перекрыти", "плит"):
                return "4.1.2."
            if contains(n, "вертикальн", "стен", "колонн", "пилон"):
                return "4.1.3."
            return "4.1."
        if contains(n, "надземн", "типовой этаж"):
            if contains(n, "горизонтальн", "перекрыти"):
                return "4.2.3."
            if contains(n, "вертикальн", "стен", "колонн"):
                return "4.2.4."
            return "4.2."
        return "4."
    if contains(n, "фундаментная плита", "фундаментной плит"):
        return "4.1.1."
    if contains(n, "устройство траншеи", "выемка грунта",
                "установка стоек переопирани", "устройство фундаментной"):
        if ctx == "12":
            return "12.1."
        return "4.99."

    # ── РАЗДЕЛ 5: Общестроительные работы ─────────────────────────────────────
    if contains(n, "светопрозрачные перегородки"):
        return "5.3."
    if contains(n, "перегородки"):
        if contains(n, "квартир", "mr base"):
            return "14.1."
        if contains(n, "паркинг", "подземн"):
            return "5.1."
        return "5.2."

    # ── РАЗДЕЛ 6: Фасадные работы ─────────────────────────────────────────────
    if contains(n, "баннерная конструкция", "удерживающих конструкц",
                "ипрс", "фасадной сетки", "самонесущая баннерн"):
        return "6.99."
    if contains(n, "разборка/восстановление исторических фасадов",
                "разборка исторических фасадов"):
        return "19.1."
    if contains(n, "фасад"):
        if contains(n, "освещени", "осветительн", "подсветк"):
            return "11.12."
        if contains(n, "витраж", "стоечно-ригельная"):
            if contains(n, "стилобат", "1й", "1-й"):
                return "6.1."
            return "6.2."
        if contains(n, "навесн", "вентфасад", "нвф"):
            if contains(n, "стилобат", "1й", "цоколь"):
                return "6.3."
            return "6.4."
        if contains(n, "модульн"):
            return "6.5."
        if contains(n, "обслуживания фасадов", "соф"):
            return "6.6."
        if contains(n, "штукатурн"):
            return "6.7."
        if contains(n, "демонтаж фасадн"):
            return "6.8."
        if contains(n, "откос", "отлив", "парапет"):
            return "6.10."
        return "6."
    if contains(n, "витраж", "стоечно-ригельная"):
        if contains(n, "стилобат", "1й", "1-й"):
            return "6.1."
        return "6.2."
    if contains(n, "устройство фасадов"):
        return "6."

    # ── РАЗДЕЛ 7: Кровля ──────────────────────────────────────────────────────
    if contains(n, "кровля", "кровли", "временная кровля"):
        if contains(n, "паркинг", "подземн", "стилобат"):
            return "7.1."
        if contains(n, "парапет", "лестниц", "ограждени"):
            return "7.3."
        if contains(n, "воронк", "люк"):
            return "7.99."
        return "7.2."

    # ── РАЗДЕЛ 8: Отделочные работы МОП и паркинга ───────────────────────────
    if contains(n, "двери", "ворота", "противопожарная штора"):
        return "8.3."
    if contains(n, "навигация и обеспечение доступа"):
        return "13.99."
    if contains(n, "отделк"):
        if contains(n, "квартир", "mr base"):
            return "14."
        if contains(n, "паркинг", "подземн", "рамп"):
            return "8.1.1."
        if contains(n, "лобби", "гранд"):
            return "8.2.1."
        if contains(n, "воинской части"):
            return "8.2."
        return "8.2."

    # ── РАЗДЕЛ 9: Лифты ───────────────────────────────────────────────────────
    if contains(n, "лифт", "подъёмник", "подъемник"):
        if contains(n, "отделка кабин"):
            return "9.3."
        if contains(n, "управление лифтами", "вызывная"):
            return "9.99."
        return "9.2."

    # ── РАЗДЕЛ 10: ВИС механические ──────────────────────────────────────────
    if contains(n, "водоснабжени", "хвс", "гвс", "горячее водоснабж",
                "холодное водоснабж", "хоз-питьевого водоснабж"):
        return "10.1."
    if contains(n, "газовое пожаротушени", "порошковое пожаротушени"):
        return "10.9."
    if contains(n, "водяное пожаротушени", "спринклерн",
                "противопожарный водопровод", "аупт",
                "автоматических установок пожаротушени"):
        return "10.2."
    if contains(n, "канализац", "водоотведени", "жироуловит", "дренажных насосов"):
        return "10.3."
    if contains(n, "отоплени", "теплоснабж", "радиатор", "конвектор"):
        return "10.4."
    if contains(n, "противодымная вентиляц", "противодымной вентиляц",
                "дымоудалени", "подпор воздуха", "противодымн вентиляц"):
        return "10.7."
    if contains(n, "тепловой пункт", " итп", "итп "):
        return "10.8."
    if contains(n, "вентиляц"):
        if contains(n, "приточно-вытяжная") and contains(n, "квартир"):
            return "14.9."
        return "10.5."
    if contains(n, "кондиционировани"):
        if contains(n, "квартир"):
            return "14.10."
        return "10.6."

    # ── РАЗДЕЛ 11: ВИС электрические и слаботочные ───────────────────────────
    # ВАЖНО: молниезащита до блока электрооборудования, т.к. "молниезащита" содержит "щит"
    if contains(n, "молниезащит", "заземлени", "системы уравнивания потенциалов"):
        return "11.2."
    if contains(n, "система электроснабжения и электроосвещения",
                "внутреннее электрооборудование",
                "электроснабжени", "электроосвещени",
                "электромонтажн", "переустройство системы электроснабж"):
        if contains(n, "кабелен", "лотки"):
            return "11.1.1."
        if contains(n, "щитовое", "вру", "грщ", "щит", "распределительные силовые"):
            return "11.1.3."
        if contains(n, "шинопровод"):
            return "11.1.5."
        if contains(n, "зарядные устройства", "дгу", "ибп", "резервные источники"):
            return "11.1.8."
        return "11.1."
    if contains(n, "сети связи"):
        return "11.3."
    if contains(n, "радиофикац", "го и чс", "го чс"):
        return "11.3.1."
    if contains(n, "структурированная кабельная система систем связи", "скс сс"):
        return "11.3.2."
    if contains(n, "стс мгн", "сигнализации для санузлов мгн"):
        return "11.3.3."
    if contains(n, "сусс", "усиления сотовой"):
        return "11.3.4."
    if contains(n, "wi-fi", "wifi", "беспроводная локальная вычилит"):
        return "11.3.5."
    if contains(n, "музыкальная трансляция"):
        return "11.3."
    if contains(n, "пожарная сигнализац", "апс", "соуэ", "аспс", "аспа",
                "система пожарной сигнализации и автоматики"):
        return "11.4."
    if contains(n, "система оповещения и управления эвакуацией"):
        return "11.4."
    if contains(n, "противопожарные системы"):
        return "11.4."
    # СКУД / охрана — ВАЖНО: проверяем строгие вхождения, чтобы не поймать "энергоресурсов" как "сов"
    if contains(n, "скуд", "контроль и управление доступом"):
        return "11.5.3."
    if contains(n, "охранного видеонаблюдени", "система охранного"):
        return "11.5.2."
    # "сов" только как отдельное слово/аббревиатура (не как часть "ресурсов")
    if word_in(n, "сов"):
        return "11.5.2."
    if contains(n, "сотс", "охранно-тревожной"):
        return "11.5.4."
    if contains(n, "скс сб", "систем безопасности"):
        return "11.5.1."
    if contains(n, "домофон", "сдс"):
        return "11.5."
    if contains(n, "система контроля и управления доступом",
                "охранная сигнализац", "система охранного видеонаблюд"):
        return "11.5."
    if contains(n, "оздс", "дератизац"):
        return "11.10."
    # АСКУЭ / учёт — ПЕРЕД наружным освещением
    if contains(n, "аскуэ", "учёта электроэнергии", "учета электроэнергии",
                "учёта воды", "учета воды", "учёта тепла", "учета тепла",
                "автоматизированная система учета", "аскувт", "асувт", "асуэ",
                "аскув", "аскут", "учета энергоресурс", "учёта энергоресурс"):
        return "11.7."
    if contains(n, "автоматизац", "диспетчеризац", "диспетчерского контроля", "асуд"):
        if contains(n, "контроль загазованност"):
            return "11.6.7."
        if contains(n, "лифт"):
            return "11.6.3."
        if contains(n, "противодымн"):
            return "11.6.2."
        if contains(n, "канализац", "дренаж"):
            return "11.6.4."
        if contains(n, "водоснабж"):
            return "11.6.6."
        if contains(n, "холодоснабж"):
            return "11.6.8."
        if contains(n, "пожаротушени"):
            return "11.6.5."
        if contains(n, "общеобменн вентиляц", "вентиляции"):
            return "11.6.1."
        return "11.6."
    if contains(n, "наружное освещени", "опоры освещени"):
        # Если контекст B-предка — благоустройство → 12.99, иначе → 11.11
        if smr_ctx_code.startswith("12"):
            return "12.99."
        return "11.11."
    if contains(n, "фасадное освещени"):
        return "11.12."
    if contains(n, "есцз", "ецсз", "цифровой среды здания", "цифровая среда здания"):
        return "11.8."
    if contains(n, "умная квартира", "умный офис"):
        return "11.9."

    # ── РАЗДЕЛ 12: Благоустройство (продолжение) ──────────────────────────────
    if contains(n, "тротуар", "брусчатка", "бортовой камень", "асфальт"):
        return "12.2."
    if contains(n, "озеленени", "деревья", "газон"):
        return "12.3."
    if contains(n, "маф", "детские площадки", "велопарковка", "фонтан"):
        return "12.4."

    # ── РАЗДЕЛ 13: ТХ ────────────────────────────────────────────────────────
    if contains(n, "технологические решения"):
        if contains(n, "паркинг", "автостоянки"):
            return "13.1."
        if contains(n, "апс", "автоматизированная парковочная"):
            return "13.2."
        if contains(n, "мусоро"):
            return "13.3."
        return "13.4."
    if contains(n, "мусороудалени", "мусоропровод"):
        return "13.3."
    if contains(n, "система «анти", "противодействи террористич"):
        return "13.99."
    if contains(n, "навигация"):
        return "13.99."

    # ── РАЗДЕЛ 14: MR Base ───────────────────────────────────────────────────
    if contains(n, "mr base", "mrbase"):
        return "14."
    if contains(n, "white box", "whitebox"):
        return "14."
    if contains(n, "умная"):
        return "11.9."

    # ── РАЗДЕЛ 15: MR Ready ──────────────────────────────────────────────────
    if contains(n, "mr ready", "mrready"):
        return "15."
    if contains(n, "чистовая отделка квартир", "финишная отделка квартир"):
        return "15."
    if contains(n, "меблировка", "меблировани"):
        return "15.2."
    if contains(n, "сантехника квартир", "сантехнические приборы квартир"):
        return "15.3."
    if contains(n, "кухонное оборудование квартир", "встроенная техника"):
        return "15.4."

    # ── РАЗДЕЛ 17: Офис продаж ───────────────────────────────────────────────
    if contains(n, "офис продаж", "офиса продаж"):
        return "17."
    if contains(n, "шоурум", "showroom"):
        return "17."

    # ── РАЗДЕЛ 20: MR Shell & Core ───────────────────────────────────────────
    if contains(n, "mr shell", "shell & core", "shell&core", "shell and core"):
        return "20."

    # ── Общие заголовки разделов ──────────────────────────────────────────────
    if contains(n, "механические инженерные системы"):
        return "10."
    if contains(n, "электрические и слаботочные инженерные системы",
                "электрические инженерные системы"):
        return "11."
    # "внутренние инженерные системы" и "инженерные системы" без уточнения —
    # смешанные заголовки (10+11), код не назначаем, чтобы не блокировать ветки 11.x

    return ""


def load_examples(folder):
    mapping = {}
    for fname in sorted(os.listdir(folder)):
        if not fname.endswith(".xlsx"):
            continue
        fpath = os.path.join(folder, fname)
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True)
        except Exception:
            continue
        ws = wb.active
        header_row_num = 1
        code_cols, name_cols = [], []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True)):
            has_code = any(row[j] and "кодификатор" in norm(row[j]) for j in range(len(row)))
            if has_code:
                header_row_num = i + 1
                for j, cell in enumerate(row):
                    if cell and "кодификатор" in norm(cell):
                        code_cols.append(j)
                    if cell and "наименование" in norm(cell):
                        name_cols.append(j)
                break
        for row in ws.iter_rows(min_row=header_row_num + 1, values_only=True):
            if not row:
                continue
            pairs = []
            if len(row) >= 7:
                if row[1] and row[2]: pairs.append((str(row[1]), str(row[2])))
                if row[3] and row[4]: pairs.append((str(row[3]), str(row[4])))
                if row[5] and row[6]: pairs.append((str(row[5]), str(row[6])))
            for ci, ni in zip(code_cols, name_cols):
                if ci < len(row) and ni < len(row) and row[ci] and row[ni]:
                    pairs.append((str(row[ci]), str(row[ni])))
            for raw_code, raw_name in pairs:
                code = extract_code(raw_code)
                if code and raw_name:
                    key = norm(raw_name)
                    if key and key not in mapping:
                        mapping[key] = code
        wb.close()
    print(f"  Примеры: {len(mapping)} пар")
    return mapping


GENERIC_LOCATION = [
    "корпус ", "стилобат", "секция ", "лот №", "башня ",
    "очередь ", "блок ", "подъезд", "тип ", "тип к",
]


def is_location_header(name):
    if not name:
        return False
    n = norm(name)
    if any(kw in n for kw in GENERIC_LOCATION) and len(n) < 40:
        return True
    return False


def main():
    print("=" * 60)
    print(f"Маппинг: {INPUT_FILE}  →  {OUTPUT_FILE}")
    print("=" * 60)

    print("\n[1] Загружаем примеры смет...")
    emap = load_examples(EXAMPLES_DIR)
    ekeys = list(emap.keys())

    print("\n[2] Читаем файл и строим иерархию разделов...")
    wb = openpyxl.load_workbook(INPUT_FILE)
    ws = wb.active

    # Ищем столбец стоимости
    cost_col_idx = find_cost_col(ws)
    if cost_col_idx is not None:
        print(f"  Столбец стоимости: {openpyxl.utils.get_column_letter(cost_col_idx + 1)}")
    else:
        print("  Столбец стоимости: не найден (фильтр по нулям отключён)")

    # Собираем все строки
    rows_data = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=False), start=2):
        if len(row) < 4:
            continue
        cost_val = row[cost_col_idx].value if (cost_col_idx is not None and cost_col_idx < len(row)) else None
        rows_data.append({
            "excel_row": i,
            "section_b": str(row[1].value).strip() if row[1].value is not None else None,
            "cell_c": row[2],
            "code_c": str(row[2].value).strip() if row[2].value is not None else None,
            "name_d": row[3].value if row[3].value is not None else None,
            "cost_val": cost_val,
        })

    # ── ФАЗА 1: Назначаем SMR-коды строкам-разделам (с B value) ─────────────
    # b_smr[b_val] = SMR-код, назначенный этому разделу
    b_smr = {}

    # Сначала обрабатываем из примеров (уже заполненные файлы)
    for rd in rows_data:
        if rd["code_c"] and rd["section_b"]:
            b_smr[rd["section_b"]] = rd["code_c"]

    # Для незаполненных B-строк: keyword/exact/fuzzy
    for rd in rows_data:
        bv = rd["section_b"]
        if not bv or bv in b_smr:
            continue
        name = rd["name_d"]
        if not name:
            continue
        # Получаем SMR-контекст от предка
        ancestor_smr = find_ancestor_smr(bv, b_smr)
        code = map_code(str(name), ancestor_smr)
        if not code:
            n = norm(name)
            if n in emap:
                code = emap[n]
        if not code and norm(name) and len(norm(name)) >= 8:
            match = process.extractOne(
                norm(name), ekeys,
                scorer=fuzz.token_sort_ratio,
                score_cutoff=85)
            if match:
                code = emap[match[0]]
        # Проверяем совместимость с контекстом предка — исправляем до сохранения в b_smr
        if code and ancestor_smr and not smr_compatible(code, ancestor_smr):
            corrected = smr_correct(str(name), ancestor_smr)
            if corrected:
                code = corrected
        if code:
            b_smr[bv] = code

    # ── ФАЗА 2: Назначаем SMR-коды всем строкам (с учётом иерархии B) ────────
    print("\n[3] Проставляем коды...")
    filled = 0
    already = 0
    inherited = 0
    skipped_zero = 0
    manual = []

    current_b = None   # последний встреченный B-раздел

    for rd in rows_data:
        if rd["section_b"]:
            current_b = rd["section_b"]

        cell_c = rd["cell_c"]
        name_d = rd["name_d"]

        if rd["code_c"]:
            already += 1
            continue
        if not name_d:
            continue

        # ── ФИЛЬТР НУЛЕВОЙ СТОИМОСТИ ─────────────────────────────────────────
        # Пропускаем строки с нулевой/пустой стоимостью — код не нужен
        # Исключение: строки с непустым B (заголовки разделов) всегда обрабатываем
        if cost_col_idx is not None and not rd["section_b"]:
            if is_zero_cost(rd["cost_val"]):
                skipped_zero += 1
                continue

        name_n = norm(name_d)

        # SMR-контекст: собственный код раздела или предка
        smr_ctx_code = ""
        if current_b:
            smr_ctx_code = b_smr.get(current_b, "")
            if not smr_ctx_code:
                smr_ctx_code = find_ancestor_smr(current_b, b_smr)

        assigned = ""

        # ── ПРИОРИТЕТ 1: Keyword rules (с SMR-контекстом) ───────────────────
        assigned = map_code(str(name_d), smr_ctx_code)

        # ── ПРИОРИТЕТ 2: Если keyword дал результат — проверяем совместимость
        if assigned and smr_ctx_code and not smr_compatible(assigned, smr_ctx_code):
            # Несовместим с контекстом — пробуем исправить
            corrected = smr_correct(str(name_d), smr_ctx_code)
            if corrected:
                assigned = corrected

        # ── ПРИОРИТЕТ 3: Exact match по примерам ────────────────────────────
        if not assigned and name_n in emap:
            cand = emap[name_n]
            # Принимаем только если совместим с контекстом
            if not smr_ctx_code or smr_compatible(cand, smr_ctx_code):
                assigned = cand
            else:
                corrected = smr_correct(str(name_d), smr_ctx_code)
                assigned = corrected if corrected else ""

        # ── ПРИОРИТЕТ 4: Fuzzy match (контекстно-фильтрованный) ─────────────
        if not assigned and name_n and len(name_n) >= 8:
            ctx_top = smr_top(smr_ctx_code)
            if ctx_top:
                filtered = [k for k in ekeys
                            if smr_top(emap[k]) == ctx_top]
            else:
                filtered = ekeys

            if filtered:
                match = process.extractOne(
                    name_n, filtered,
                    scorer=fuzz.token_sort_ratio,
                    score_cutoff=85)
                if match:
                    cand = emap[match[0]]
                    # Проверяем совместимость и с контекстом
                    if smr_ctx_code and not smr_compatible(cand, smr_ctx_code):
                        corrected = smr_correct(str(name_d), smr_ctx_code)
                        assigned = corrected if corrected else ""
                    else:
                        assigned = cand

            if not assigned:
                match2 = process.extractOne(
                    name_n, ekeys,
                    scorer=fuzz.token_sort_ratio,
                    score_cutoff=90)
                if match2:
                    cand = emap[match2[0]]
                    if not smr_ctx_code or smr_compatible(cand, smr_ctx_code):
                        assigned = cand

        # ── ПРИОРИТЕТ 5: Наследование от родительского SMR-раздела ──────────
        if not assigned and smr_ctx_code:
            assigned = smr_ctx_code
            inherited += 1

        if assigned:
            cell_c.value = assigned
            # Обновляем b_smr для текущего раздела если ещё не установлен
            if current_b and not b_smr.get(current_b):
                b_smr[current_b] = assigned
            filled += 1
        else:
            manual.append(
                f"  строка {rd['excel_row']:4d} | B={str(rd['section_b'])[:8]:8s} | D={str(name_d)[:70]}")

    print(f"\n[4] Результаты:")
    print(f"    Уже были с кодом:         {already}")
    print(f"    Keyword/exact/fuzzy:      {filled - inherited}")
    print(f"    По наследованию раздела:  {inherited}")
    print(f"    Итого проставлено:        {filled}")
    print(f"    Пропущено (нулевая $):    {skipped_zero}")
    print(f"    Осталось пустых:          {len(manual)}")

    wb.save(OUTPUT_FILE)
    print(f"\nСохранено: {OUTPUT_FILE}")

    if manual:
        print(f"\nТРЕБУЮТ РУЧНОЙ ПРОВЕРКИ ({len(manual)} строк):")
        for line in manual[:60]:
            print(line)
        if len(manual) > 60:
            print(f"  ... и ещё {len(manual) - 60} строк")

    print("\nГотово!")


if __name__ == "__main__":
    main()
