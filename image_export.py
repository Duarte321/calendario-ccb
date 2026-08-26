import calendar
import textwrap
from io import BytesIO

from PIL import Image, ImageDraw, ImageFont

MESES = {
    1: "JANEIRO", 2: "FEVEREIRO", 3: "MARÇO", 4: "ABRIL",
    5: "MAIO", 6: "JUNHO", 7: "JULHO", 8: "AGOSTO",
    9: "SETEMBRO", 10: "OUTUBRO", 11: "NOVEMBRO", 12: "DEZEMBRO",
}
DIAS = ["DOM", "SEG", "TER", "QUA", "QUI", "SEX", "SÁB"]

NAVY = (6, 29, 51)
NAVY_2 = (11, 45, 77)
GOLD = (224, 167, 47)
GOLD_LIGHT = (255, 241, 198)
WHITE = (255, 255, 255)
INK = (23, 36, 54)
MUTED = (94, 106, 122)
GRID = (211, 217, 225)
EMPTY = (241, 243, 246)
NOTICE_BG = (255, 248, 229)
NOTICE_TEXT = (112, 76, 8)


def _font(size, bold=False):
    names = [
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for name in names:
        try:
            return ImageFont.truetype(name, size=size)
        except Exception:
            pass
    return ImageFont.load_default()


def _fit_lines(draw, text, font, max_width, max_lines):
    words = str(text or "").replace("\n", " ").split()
    if not words:
        return []
    lines = []
    current = ""
    for word in words:
        candidate = (current + " " + word).strip()
        width = draw.textbbox((0, 0), candidate, font=font)[2]
        if width <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word
            if len(lines) >= max_lines:
                break
    if current and len(lines) < max_lines:
        lines.append(current)
    if len(lines) == max_lines and len(words) > 1:
        original = " ".join(words)
        rendered = " ".join(lines)
        if len(rendered) < len(original):
            last = lines[-1]
            while last and draw.textbbox((0, 0), last + "...", font=font)[2] > max_width:
                last = last[:-1]
            lines[-1] = last.rstrip() + "..."
    return lines


def _center_text(draw, box, text, font, fill):
    x1, y1, x2, y2 = box
    bb = draw.textbbox((0, 0), text, font=font)
    w = bb[2] - bb[0]
    h = bb[3] - bb[1]
    draw.text((x1 + (x2 - x1 - w) / 2, y1 + (y2 - y1 - h) / 2 - 2), text, font=font, fill=fill)


def gerar_imagem_mes(ano, mes, agenda, aviso=""):
    """Gera PNG 1080x1350. agenda deve conter tuplas (date, evento)."""
    width, height = 1080, 1350
    img = Image.new("RGB", (width, height), (248, 249, 251))
    draw = ImageDraw.Draw(img)

    title_font = _font(54, True)
    subtitle_font = _font(24, True)
    month_font = _font(38, True)
    dow_font = _font(22, True)
    day_font = _font(28, True)
    event_font = _font(17, True)
    time_font = _font(16, True)
    footer_font = _font(18, False)
    notice_title_font = _font(20, True)
    notice_font = _font(18, False)

    # Cabeçalho premium
    draw.rounded_rectangle((35, 28, 1045, 190), radius=28, fill=NAVY)
    draw.text((75, 57), "AGENDA MUSICAL", font=title_font, fill=WHITE)
    draw.text((78, 126), "Região de Jaciara - MT", font=subtitle_font, fill=(242, 196, 94))
    draw.text((795, 70), "ENSaiOS LOCAIS".upper(), font=_font(17, True), fill=(219, 228, 238))
    draw.text((795, 103), "Calendário oficial", font=_font(17, False), fill=(242, 196, 94))

    # Mês
    draw.rounded_rectangle((35, 215, 1045, 300), radius=20, fill=WHITE, outline=GRID, width=2)
    _center_text(draw, (35, 215, 1045, 300), f"{MESES[int(mes)]} {int(ano)}", month_font, NAVY_2)

    left, right = 35, 1045
    grid_top = 325
    header_h = 54
    col_w = (right - left) / 7

    for i, nome in enumerate(DIAS):
        x1 = int(left + i * col_w)
        x2 = int(left + (i + 1) * col_w)
        draw.rectangle((x1, grid_top, x2, grid_top + header_h), fill=NAVY_2, outline=(41, 77, 105), width=1)
        _center_text(draw, (x1, grid_top, x2, grid_top + header_h), nome, dow_font, WHITE)

    calendar.setfirstweekday(calendar.SUNDAY)
    semanas = calendar.monthcalendar(int(ano), int(mes))
    notice_height = 145 if aviso else 72
    footer_y = 1298
    grid_bottom = footer_y - notice_height - 22
    row_h = (grid_bottom - (grid_top + header_h)) / len(semanas)

    eventos_por_dia = {}
    for dt, evt in agenda:
        if dt.year == int(ano) and dt.month == int(mes):
            eventos_por_dia.setdefault(dt.day, []).append(evt)

    for r, semana in enumerate(semanas):
        y1 = int(grid_top + header_h + r * row_h)
        y2 = int(grid_top + header_h + (r + 1) * row_h)
        for c, dia in enumerate(semana):
            x1 = int(left + c * col_w)
            x2 = int(left + (c + 1) * col_w)
            has_event = dia and dia in eventos_por_dia
            bg = EMPTY if dia == 0 else GOLD_LIGHT if has_event else WHITE
            draw.rectangle((x1, y1, x2, y2), fill=bg, outline=GRID, width=2)
            if dia == 0:
                continue

            draw.text((x1 + 12, y1 + 9), str(dia), font=day_font, fill=INK)

            if has_event:
                cy = y1 + 48
                max_w = x2 - x1 - 22
                for evt in eventos_por_dia[dia][:2]:
                    titulo = str(evt.get("titulo", "ENSAIO LOCAL")).title()
                    local = str(evt.get("local", "")).title()
                    hora = str(evt.get("hora", ""))
                    for line in _fit_lines(draw, titulo, event_font, max_w, 2):
                        draw.text((x1 + 11, cy), line, font=event_font, fill=INK)
                        cy += 21
                    for line in _fit_lines(draw, local, event_font, max_w, 2):
                        draw.text((x1 + 11, cy), line, font=event_font, fill=INK)
                        cy += 21
                    draw.text((x1 + 11, cy + 2), hora, font=time_font, fill=NOTICE_TEXT)
                    cy += 29
                    if cy > y2 - 18:
                        break

    notice_y = int(grid_top + header_h + len(semanas) * row_h + 18)
    if aviso:
        draw.rounded_rectangle((35, notice_y, 1045, notice_y + 122), radius=18, fill=NOTICE_BG, outline=(238, 213, 148), width=2)
        draw.text((58, notice_y + 18), "AVISO IMPORTANTE", font=notice_title_font, fill=NOTICE_TEXT)
        lines = _fit_lines(draw, aviso, notice_font, 930, 3)
        yy = notice_y + 52
        for line in lines:
            draw.text((58, yy), line, font=notice_font, fill=NOTICE_TEXT)
            yy += 25
    else:
        draw.rounded_rectangle((35, notice_y, 1045, notice_y + 52), radius=16, fill=WHITE, outline=GRID, width=2)
        _center_text(draw, (35, notice_y, 1045, notice_y + 52), "Dias em dourado indicam ensaio local", footer_font, MUTED)

    # Rodapé
    draw.line((55, footer_y - 12, 1025, footer_y - 12), fill=(220, 224, 230), width=2)
    _center_text(draw, (35, footer_y, 1045, 1338), "Agenda Musical • Região de Jaciara - MT", footer_font, NAVY_2)

    out = BytesIO()
    img.save(out, format="PNG", optimize=True)
    out.seek(0)
    return out
