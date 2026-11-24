"""
poster.py — постер «Расписание» на неделю для публикации в TG-канал.

Фон (assets/poster/template.png) — бланк из макета с вычищенной сеткой:
на нём уже есть «РАСПИСАНИЕ», дни недели, логотип, адрес и пустая плашка
месяца. Всё, что меняется от недели к неделе, рисуется здесь:
месяц, неделя месяца, числа, линии, кружки, время и подписи.

Координаты и размеры сняты с макета 1280×1280 (см. CLAUDE.md, раздел
«Постер расписания»). Шрифт — Atyp Display.
"""

from __future__ import annotations

import io
import os
from dataclasses import dataclass
from datetime import date
from functools import lru_cache
from typing import Sequence

from PIL import Image, ImageDraw, ImageFont

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "poster")

MONTHS_RU = [
    "январь", "февраль", "март", "апрель", "май", "июнь",
    "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
]
MONTHS_RU_GEN = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря",
]

ACTIVITY_LABEL = "плавание"
REMOTE_LABEL   = "удалённо"

# ── Геометрия макета (пиксели, холст 1280×1280) ─────────────────
# Центры кружков дней Пн..Вс: столбец = число, строка = день недели.
DAY_CX = [444.5, 539.5, 640.5, 739.5, 834.5, 937.5, 1030.5]
DAY_CY = [442.5, 544.5, 642.5, 739.5, 838.5, 937.5, 1039.5]
# Радиусы кружков тренировок: растут к выходным (Вт/Чт/Сб — как в макете).
# Пиксельные позиции линий (вертикаль от числа, горизонталь от дня) — как в макете.
LINE_X = [445, 539, 641, 740, 835, 939, 1031]
LINE_Y = [443, 544, 643, 739, 838, 939, 1039]
TRAINING_R = [12, 15, 22, 30, 40, 50, 55]
EMPTY_R    = 10                       # пустой кружок дня без тренировки
LINES_START = 400                     # откуда идут линии (x для горизонталей, y для вертикалей)

DATE_LEFT     = [411, 511, 612, 712, 812, 912, 1012]   # левый край чисел
DATE_SLOT_W   = 52                                     # ширина двузначного числа
DATE_BASELINE = 359

PILL_CX, PILL_BASELINE = 749, 253     # плашка месяца (центр букв как в макете)
PILL_STEP_MAX   = 80                  # шаг между буквами месяца
PILL_SPAN_MAX   = 540                 # максимальный разброс центров букв

WEEK_CIRCLE_C = (340.0, 339.0)        # центр цифры недели в кружке

TIME_LEFT        = 418                # левый край времени
TIME_ABOVE_LINE  = 9                  # базовая линия времени над горизонталью
LABEL_TOP        = 401                # верх вертикальной подписи
LABEL_GAP        = 10                 # отступ подписи от вертикальной линии

# ── Цвета ───────────────────────────────────────────────────────
BLACK       = (0, 0, 0)
BLUE        = (0, 59, 255)
RED         = (255, 0, 0)
DOT_GREY    = (126, 126, 126)
DIAG_GREY   = (150, 150, 150)
EMPTY_FILL  = (236, 236, 236)

# Начало и конец пунктирной диагонали (от кружка недели вниз-вправо)
DIAG_FROM = (366.0, 361.5)
DIAG_TO   = (1080.5, 1090.0)

SS = 4  # суперсэмплинг для сглаживания кружков и диагонали


@dataclass(frozen=True)
class PosterDay:
    """Один день недели на постере."""
    date: date
    time: str = ""        # "20:00"; пусто — время не задано
    remote: bool = False  # удалённая тренировка

    @property
    def has_training(self) -> bool:
        return bool(self.time) or self.remote


# ═══════════════════════════════════════════════════════════════
# ДАННЫЕ
# ═══════════════════════════════════════════════════════════════

def week_month_and_number(days: Sequence[PosterDay]) -> tuple[int, int]:
    """Месяц недели и её номер в месяце — по четвергу.

    Неделя, перешедшая через границу месяцев, относится к тому месяцу,
    где большинство её дней (там же четверг). Номер — какой по счёту
    четверг месяца: 18–24 мая 2026 → (5, 3).
    """
    thursday = days[3].date
    return thursday.month, (thursday.day - 1) // 7 + 1


def date_range_label(days: Sequence[PosterDay]) -> str:
    """«18–24 мая» / «29 июня – 5 июля»."""
    first, last = days[0].date, days[-1].date
    if first.month == last.month:
        return f"{first.day}–{last.day} {MONTHS_RU_GEN[last.month - 1]}"
    return (f"{first.day} {MONTHS_RU_GEN[first.month - 1]} – "
            f"{last.day} {MONTHS_RU_GEN[last.month - 1]}")


# ═══════════════════════════════════════════════════════════════
# ОТРИСОВКА
# ═══════════════════════════════════════════════════════════════

@lru_cache(maxsize=None)
def _font(weight: str, size: int) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(os.path.join(ASSETS_DIR, f"AtypDisplay-{weight}.ttf"), size)


@lru_cache(maxsize=1)
def _template() -> Image.Image:
    return Image.open(os.path.join(ASSETS_DIR, "template.png")).convert("RGB")


def _text_mask(text: str, font: ImageFont.FreeTypeFont, tracking: float = 0.0
               ) -> tuple[Image.Image, int, int]:
    """Рисует строку (с трекингом) в маску.

    Возвращает (маска, обрезанная по чернилам по горизонтали;
    смещение базовой линии от верха маски; ascent шрифта).
    По вертикали маска охватывает всю высоту шрифта, чтобы подписи
    с выносными элементами и без них стояли на одной базовой линии.
    """
    ascent, descent = font.getmetrics()
    pad = 4
    width = int(font.getlength(text) + abs(tracking) * len(text)) + pad * 2 + 8
    height = ascent + descent + pad * 2
    mask = Image.new("L", (width, height), 0)
    d = ImageDraw.Draw(mask)
    baseline = pad + ascent
    x = float(pad)
    for ch in text:
        d.text((x, baseline), ch, font=font, fill=255, anchor="ls")
        x += font.getlength(ch) + tracking
    bbox = mask.getbbox()
    if not bbox:
        return mask, baseline, ascent
    left, right = bbox[0], bbox[2]
    return mask.crop((left, 0, right, height)), baseline, ascent


def _paste_text(img: Image.Image, text: str, font, *, left: float, baseline: float,
                color=BLACK, tracking: float = 0.0) -> int:
    """Ставит текст так, что левый край чернил = left, базовая линия = baseline.
    Возвращает ширину текста по чернилам."""
    mask, base_off, _ = _text_mask(text, font, tracking)
    img.paste(color, (round(left), round(baseline - base_off)), mask)
    return mask.width


def _paste_text_centered(img, text, font, *, cx: float, baseline: float,
                         color=BLACK, tracking: float = 0.0) -> None:
    mask, base_off, _ = _text_mask(text, font, tracking)
    img.paste(color, (round(cx - mask.width / 2), round(baseline - base_off)), mask)


def _paste_vertical_label(img, text, font, *, baseline_x: float, top: float,
                          side: str = "left", color=BLACK, tracking: float = 0.0) -> None:
    """Подпись, повёрнутая на 90° против часовой (читается снизу вверх).

    Верх букв смотрит влево. side="left": базовая линия на baseline_x
    (текст левее линии); side="right": левый край букв на baseline_x.
    """
    mask, base_off, _ = _text_mask(text, font, tracking)
    rotated = mask.rotate(90, expand=True)
    # после поворота: x маски = вертикальная координата исходной (верх → лево)
    if side == "left":
        x = baseline_x - base_off
    else:
        ink = mask.getbbox()
        x = baseline_x - (ink[1] if ink else 0)
    img.paste(color, (round(x), round(top)), rotated)


def _dotted_hline(d: ImageDraw.ImageDraw, x0: int, x1: int, y: int, on: int, off: int, color) -> None:
    x = x0
    while x < x1:
        d.line([(x, y), (min(x + on - 1, x1 - 1), y)], fill=color, width=1)
        x += on + off


def _dotted_vline(d: ImageDraw.ImageDraw, x: int, y0: int, y1: int, on: int, off: int, color) -> None:
    y = y0
    while y < y1:
        d.line([(x, y), (x, min(y + on - 1, y1 - 1))], fill=color, width=1)
        y += on + off


def _circle(img: Image.Image, cx: float, cy: float, r: float, fill, outline=None) -> None:
    """Сглаженный кружок: рисуем в SS раз крупнее и уменьшаем."""
    x0, y0 = int(cx - r) - 2, int(cy - r) - 2
    size = int(2 * r) + 6
    big = Image.new("L", (size * SS, size * SS), 0)
    d = ImageDraw.Draw(big)
    bx, by, br = (cx - x0) * SS, (cy - y0) * SS, r * SS
    d.ellipse([bx - br, by - br, bx + br, by + br], fill=255)
    mask = big.resize((size, size), Image.LANCZOS)
    if outline is None:
        img.paste(fill, (x0, y0), mask)
        return
    img.paste(outline, (x0, y0), mask)
    inner = Image.new("L", (size * SS, size * SS), 0)
    ImageDraw.Draw(inner).ellipse([bx - br + SS, by - br + SS, bx + br - SS, by + br - SS], fill=255)
    img.paste(fill, (x0, y0), inner.resize((size, size), Image.LANCZOS))


def _diagonal(img: Image.Image) -> None:
    """Тонкая точечная диагональ через все дни."""
    (fx, fy), (tx, ty) = DIAG_FROM, DIAG_TO
    x0, y0 = int(fx) - 2, int(fy) - 2
    w, h = int(tx - x0) + 3, int(ty - y0) + 3
    big = Image.new("L", (w * SS, h * SS), 0)
    d = ImageDraw.Draw(big)
    dx, dy = tx - fx, ty - fy
    length = (dx * dx + dy * dy) ** 0.5
    ux, uy = dx / length, dy / length
    step, dot = 4.0, 1.0
    t = 0.0
    while t < length:
        ax, ay = (fx + ux * t - x0) * SS, (fy + uy * t - y0) * SS
        bx, by = (fx + ux * (t + dot) - x0) * SS, (fy + uy * (t + dot) - y0) * SS
        d.line([(ax, ay), (bx, by)], fill=255, width=SS)
        t += step
    img.paste(DIAG_GREY, (x0, y0), big.resize((w, h), Image.LANCZOS))


def render_schedule_poster(days: Sequence[PosterDay]) -> bytes:
    """Рисует постер на неделю (7 дней Пн..Вс) и возвращает PNG."""
    if len(days) != 7:
        raise ValueError(f"Нужно 7 дней недели, получено {len(days)}")

    img = _template().copy()
    d = ImageDraw.Draw(img)

    # 1. Диагональ (под всем остальным)
    _diagonal(img)

    # 2. Линии от числа и от дня недели к кружку
    for i, day in enumerate(days):
        cx, cy = DAY_CX[i], DAY_CY[i]
        px, py = LINE_X[i], LINE_Y[i]
        r = TRAINING_R[i] if day.has_training else EMPTY_R
        if day.has_training:
            _dotted_hline(d, LINES_START, int(cx - r), py, 5, 5, BLACK)
            _dotted_vline(d, px, LINES_START, int(cy - r), 5, 5, BLACK)
        else:
            _dotted_hline(d, LINES_START, int(cx - r), py, 2, 2, DOT_GREY)
            _dotted_vline(d, px, LINES_START, int(cy - r), 2, 2, DOT_GREY)

    # 3. Кружки
    for i, day in enumerate(days):
        cx, cy = DAY_CX[i], DAY_CY[i]
        if day.has_training:
            _circle(img, cx, cy, TRAINING_R[i], RED if i >= 5 else BLUE)
        else:
            _circle(img, cx, cy, EMPTY_R, EMPTY_FILL, outline=BLACK)

    # 4. Числа
    f_date = _font("Light", 50)
    for i, day in enumerate(days):
        text = str(day.date.day)
        mask, base_off, _ = _text_mask(text, f_date, -2)
        left = DATE_LEFT[i] if len(text) == 2 else DATE_LEFT[i] + (DATE_SLOT_W - mask.width) / 2
        img.paste(BLACK, (round(left), round(DATE_BASELINE + 1 - base_off)), mask)

    # 5. Месяц в плашке и неделя месяца в кружке
    month, week_no = week_month_and_number(days)
    letters = list(MONTHS_RU[month - 1])
    f_month = _font("Medium", 50)
    step = min(PILL_STEP_MAX, PILL_SPAN_MAX / max(1, len(letters) - 1))
    first_cx = PILL_CX - step * (len(letters) - 1) / 2
    for k, ch in enumerate(letters):
        _paste_text_centered(img, ch, f_month, cx=first_cx + step * k, baseline=PILL_BASELINE)

    f_week = _font("Semibold", 40)
    mask, base_off, _ = _text_mask(str(week_no), f_week)
    ink = mask.getbbox()
    ink_h = ink[3] - ink[1]
    img.paste(BLACK, (round(WEEK_CIRCLE_C[0] - mask.width / 2),
                      round(WEEK_CIRCLE_C[1] - ink_h / 2 - ink[1])), mask)

    # 6. Время слева над линией и вертикальная подпись над кружком
    f_small = _font("Light", 23)
    for i, day in enumerate(days):
        if not day.has_training:
            continue
        cx, cy, r = DAY_CX[i], DAY_CY[i], TRAINING_R[i]

        # Подпись: над кружком слева от вертикали; если вертикаль короче
        # подписи (понедельник) — справа от кружка.
        label_len = _text_mask(ACTIVITY_LABEL, f_small, -2)[0].width
        label_fits_above = LABEL_TOP + label_len <= cy - r - 6
        label_right_x = cx + r + 6
        if label_fits_above:
            _paste_vertical_label(img, ACTIVITY_LABEL, f_small, baseline_x=LINE_X[i] - LABEL_GAP,
                                  top=LABEL_TOP, side="left", tracking=-2)
        else:
            _paste_vertical_label(img, ACTIVITY_LABEL, f_small, baseline_x=label_right_x,
                                  top=LABEL_TOP, side="right", tracking=-2)

        # Время: над горизонталью слева. Если не влезает до кружка
        # (понедельник) — под горизонталью, левее подписи.
        time_text = day.time or REMOTE_LABEL
        mask, base_off, _ = _text_mask(time_text, f_small, -2)
        ink = mask.getbbox() or (0, 0, mask.width, mask.height)
        if TIME_LEFT + mask.width <= cx - r - 5:
            left, baseline = TIME_LEFT, cy - TIME_ABOVE_LINE
        else:
            right_limit = label_right_x - 4 if not label_fits_above else cx + r
            left = max(383, min(TIME_LEFT, right_limit - mask.width))
            top = cy + r + 4                                 # ниже кружка
            baseline = top - ink[1] + base_off
        img.paste(BLACK, (round(left), round(baseline - base_off)), mask)

    buf = io.BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


if __name__ == "__main__":
    # Демо: неделя из макета (18–24 мая 2026: Вт и Чт 20:00, Сб 14:00)
    import sys
    demo = [
        PosterDay(date(2026, 5, 18)),
        PosterDay(date(2026, 5, 19), "20:00"),
        PosterDay(date(2026, 5, 20)),
        PosterDay(date(2026, 5, 21), "20:00"),
        PosterDay(date(2026, 5, 22)),
        PosterDay(date(2026, 5, 23), "14:00"),
        PosterDay(date(2026, 5, 24)),
    ]
    out = sys.argv[1] if len(sys.argv) > 1 else "poster_demo.png"
    with open(out, "wb") as fh:
        fh.write(render_schedule_poster(demo))
    print(f"Сохранено: {out}")
