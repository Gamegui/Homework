#!/usr/bin/env python3
"""Чистка групповой презентации ПИ-262: та же тёмная тема, без мусора, с фактами."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BG = RGBColor(0x0C, 0x16, 0x26)
CARD = RGBColor(0x16, 0x25, 0x3A)
GOLD = RGBColor(0xD3, 0xAA, 0x4C)
WHITE = RGBColor(0xF5, 0xF6, 0xF8)
GRAY = RGBColor(0xAE, 0xB8, 0xC5)
FONT = "Aptos"

prs = Presentation()
prs.slide_width = Inches(13.33)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]
TOTAL = 13

def bg(slide, color=BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def box(slide, l, t, w, h, fill_color=None, line_color=None):
    from pptx.enum.shapes import MSO_SHAPE
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if fill_color else MSO_SHAPE.RECTANGLE,
                               Inches(l), Inches(t), Inches(w), Inches(h))
    if line_color:
        s.line.color.rgb = line_color
        s.line.width = Pt(1.5)
    else:
        s.line.fill.background()
    if fill_color:
        s.fill.solid()
        s.fill.fore_color.rgb = fill_color
    return s

def txt(shape, text, size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    shape.text_frame.clear()
    shape.text_frame.word_wrap = True
    p = shape.text_frame.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = FONT
    p.alignment = align
    return shape.text_frame

def bullets(shape, items, size=16, space=Pt(6)):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, (head, body) in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = space
        if head:
            r = p.add_run()
            r.text = head
            r.font.size = Pt(size)
            r.font.bold = True
            r.font.color.rgb = GOLD
            r.font.name = FONT
        if body:
            r = p.add_run()
            r.text = body
            r.font.size = Pt(size)
            r.font.color.rgb = WHITE
            r.font.name = FONT
    return tf

def footer(slide, n):
    t = slide.shapes.add_textbox(Inches(0.6), Inches(6.95), Inches(11), Inches(0.35))
    txt(t, "Основы российской государственности • 2026", size=10, color=GRAY)
    t2 = slide.shapes.add_textbox(Inches(12.1), Inches(6.95), Inches(0.7), Inches(0.35))
    txt(t2, f"{n}", size=10, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)

def header(slide, kicker, title, n):
    t = slide.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(11), Inches(0.45))
    txt(t, kicker, size=13, bold=True, color=GOLD)
    t = slide.shapes.add_textbox(Inches(0.7), Inches(0.85), Inches(11), Inches(0.8))
    txt(t, title, size=30, bold=True, color=WHITE)
    t = slide.shapes.add_textbox(Inches(12.1), Inches(0.5), Inches(0.7), Inches(0.6))
    txt(t, f"{n:02d}", size=28, bold=True, color=GOLD, align=PP_ALIGN.RIGHT)
    box(slide, 0.7, 1.75, 11.9, 0.03, GOLD)

def note(slide, text):
    slide.notes_slide.placeholders[1].text = text

def card(slide, rows):
    """Правая карточка: Дата / Место / Командующий / Противник."""
    c = box(slide, 9.0, 2.1, 3.6, 4.35, CARD)
    tf = c.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.15)
    first = True
    for head, body in rows:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(10)
        r = p.add_run()
        r.text = head + "\n"
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = GOLD
        r.font.name = FONT
        r = p.add_run()
        r.text = body
        r.font.size = Pt(14)
        r.font.color.rgb = WHITE
        r.font.name = FONT

def itog(slide, text):
    c = box(slide, 0.7, 5.6, 7.9, 1.0, CARD, GOLD)
    tf = c.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "ИТОГ  "
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = GOLD
    r.font.name = FONT
    r = p.add_run()
    r.text = text
    r.font.size = Pt(14)
    r.font.color.rgb = WHITE
    r.font.name = FONT

# ================= 1. ТИТУЛ =================
s = prs.slides.add_slide(BLANK)
bg(s)
t = s.shapes.add_textbox(Inches(0.7), Inches(0.7), Inches(11.9), Inches(0.5))
txt(t, "ОСНОВЫ РОССИЙСКОЙ ГОСУДАРСТВЕННОСТИ", size=15, bold=True, color=GOLD)
box(s, 0.7, 1.25, 1.6, 0.06, GOLD)
t = s.shapes.add_textbox(Inches(0.7), Inches(1.6), Inches(11.9), Inches(1.6))
txt(t, "Ключевые военные победы в истории России", size=40, bold=True, color=WHITE)
t = s.shapes.add_textbox(Inches(0.7), Inches(3.2), Inches(11.9), Inches(0.6))
txt(t, "От русских княжеств до победы в Великой Отечественной войне", size=18, color=GRAY)
c = box(s, 0.7, 4.2, 5.5, 2.2, CARD)
tf = c.text_frame
tf.clear()
tf.word_wrap = True
tf.margin_left = Inches(0.3)
tf.margin_top = Inches(0.2)
p = tf.paragraphs[0]
r = p.add_run()
r.text = "ВЫПОЛНИЛИ\n"
r.font.size = Pt(12)
r.font.bold = True
r.font.color.rgb = GOLD
r.font.name = FONT
r = p.add_run()
r.text = "Антонов Савелий • Романов Даниил\nДанилов Матвей • Ващенко Роман\nГруппа ПИ-262"
r.font.size = Pt(16)
r.font.color.rgb = WHITE
r.font.name = FONT
t = s.shapes.add_textbox(Inches(7.6), Inches(4.9), Inches(4.9), Inches(1.2))
txt(t, "Доклад • 5–7 минут", size=16, color=GRAY, align=PP_ALIGN.RIGHT)
footer(s, 1)
note(s, "Титул: тема, состав группы. 15 секунд.")

# ================= 2. ВВЕДЕНИЕ =================
s = prs.slides.add_slide(BLANK)
bg(s)
header(s, "ВВЕДЕНИЕ", "О каких победах идёт речь", 2)
t = s.shapes.add_textbox(Inches(0.7), Inches(2.1), Inches(11.9), Inches(0.7))
txt(t, "Не только современная РФ, а военные победы разных периодов российской истории.", size=17, color=WHITE)
periods = [("1242–1380", "Русские княжества"), ("1612–1812", "Русское государство и империя"), ("1942–1945", "СССР в Великой Отечественной")]
for i, (y, d) in enumerate(periods):
    c = box(s, 0.7 + i * 4.1, 3.1, 3.8, 1.7, CARD, GOLD)
    tf = c.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = Inches(0.25)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = y + "\n"
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = GOLD
    r.font.name = FONT
    r = p.add_run()
    r.text = d
    r.font.size = Pt(15)
    r.font.color.rgb = WHITE
    r.font.name = FONT
t = s.shapes.add_textbox(Inches(0.7), Inches(5.3), Inches(11.9), Inches(1.0))
txt(t, "Критерий: не просто выигранный бой, а влияние победы на ход войны и развитие государства.",
    size=15, bold=True, color=GRAY)
footer(s, 2)
note(s, "Введение 30 сек: охватываем все периоды, критерий — влияние на историю.")

# ================= СРАЖЕНИЯ =================
battles = [
    dict(n=3, kicker="1242 • ЧУДСКОЕ ОЗЕРО", title="Ледовое побоище",
         card=[("ДАТА", "5 апреля 1242"), ("МЕСТО", "Чудское озеро"), ("ПОЛКОВОДЕЦ", "Александр Невский"), ("ПРОТИВНИК", "Ливонский орден")],
         items=[("Силы: ", "несколько тысяч с каждой стороны (точные данные не сохранились)."),
                ("Ход боя: ", "слабый центр заманил рыцарскую «свинью», фланги окружили её на льду."),
                ("Потери врага: ", "сотни рыцарей погибли, часть войска взята в плен."),
                ("После боя: ", "Орден запросил мир, отказался от претензий на Псков и Новгород.")],
         itog="крестоносцы остановлены, северо-запад отстоян. День славы — 18 апреля.",
         note="Тактика: слабый центр, сильные фланги, окружение. Цифры — оценки, точных нет."),
    dict(n=4, kicker="1380 • КУЛИКОВО ПОЛЕ", title="Куликовская битва",
         card=[("ДАТА", "8 сентября 1380"), ("МЕСТО", "Куликово поле, у Дона"), ("ПОЛКОВОДЕЦ", "Дмитрий Донской"), ("ПРОТИВНИК", "войско хана Мамая")],
         items=[("Силы: ", "десятки тысяч с обеих сторон — оценки историков расходятся."),
                ("Ход боя: ", "тяжёлая сеча в центре; решающий удар засадного полка Боброка."),
                ("Цена: ", "потери русских огромны — полегло множество воевод и воинов."),
                ("После боя: ", "Мамай бежал; иго продлится ещё 100 лет, но страх сломан.")],
         itog="первая победа объединённого войска над Ордой. День славы — 21 сентября.",
         note="Главное: засадный полк и то, что полки разных княжеств бились как единое войско."),
    dict(n=5, kicker="1612 • МОСКВА", title="Освобождение Москвы",
         card=[("ДАТЫ", "август — октябрь 1612"), ("МЕСТО", "Москва, Кремль"), ("РУКОВОДИТЕЛИ", "Минин и Пожарский"), ("ПРОТИВНИК", "польский гарнизон, гетман Ходкевич")],
         items=[("22–24 августа: ", "ополчение отбило все попытки Ходкевича прорваться в Кремль."),
                ("Осада: ", "гарнизон остался без еды и помощи."),
                ("26 октября (4 ноября): ", "капитуляция польского гарнизона Кремля."),
                ("После: ", "в 1613 году избран Михаил Романов — конец Смуты.")],
         itog="народ спас государство, когда его почти не было. 4 ноября — День народного единства.",
         note="Два этапа: разбили деблокаду, потом дожали гарнизон голодом. Победа ополчения, не царя."),
    dict(n=6, kicker="1709 • ПОЛТАВА", title="Полтавская битва",
         card=[("ДАТА", "8 июля (27 июня) 1709"), ("МЕСТО", "под Полтавой"), ("ПОЛКОВОДЕЦ", "Пётр I"), ("ПРОТИВНИК", "Карл XII, Швеция")],
         items=[("Перед боем: ", "гарнизон Полтавы выдержал 20 штурмов и дал Петру время."),
                ("Силы: ", "≈42 тыс. и 100+ орудий против ≈17–20 тыс. шведов с 4 орудиями."),
                ("Ход боя: ", "редуты измотали шведов; 2 часа решающего боя — и армия Карла бежит."),
                ("Потери: ", "шведы — 9+ тыс. убитых, ~19 тыс. пленных; русские — 1345 убитых, 3290 раненых.")],
         itog="перелом Северной войны, выход к Балтике. День славы — 10 июля.",
         note="Сильные факты: 20 штурмов, редуты, артиллерия, разгром за 2 часа, Переволочна."),
    dict(n=7, kicker="1812 • БОРОДИНО — МОСКВА", title="Отечественная война 1812 года",
         card=[("ДАТЫ", "июнь — декабрь 1812"), ("ГЛАВНОЕ СРАЖЕНИЕ", "Бородино, 7 сентября"), ("ПОЛКОВОДЕЦ", "М. И. Кутузов"), ("ПРОТИВНИК", "Наполеон, «Великая армия»")],
         items=[("Силы при Бородине: ", "≈120–150 тыс. против ≈135 тыс.; потери — десятки тысяч с обеих сторон."),
                ("Ход боя: ", "флеши и батарея Раевского по 8 раз переходили из рук в руки."),
                ("Замысел Кутузова: ", "русские отошли, но сохранили армию; Москва стала для врага ловушкой."),
                ("Финал: ", "пожары, голод, морозы, партизаны — к декабрю армия Наполеона уничтожена.")],
         itog="Россия выстояла; в 1814 русские вошли в Париж. День Бородина — 8 сентября.",
         note="Честно: Бородино — не отдельная победа, а часть кампании. Стратегически — победа."),
    dict(n=8, kicker="1942–1943 • СТАЛИНГРАД", title="Сталинградская битва",
         card=[("ДАТЫ", "17 июля 1942 — 2 февраля 1943"), ("МЕСТО", "Сталинград, Волга"), ("ОПЕРАЦИЯ", "«Уран», с 19 ноября"), ("ПРОТИВНИК", "6-я армия Паулюса")],
         items=[("Оборона: ", "июль–ноябрь: бои за каждый дом, город удержан на полосе у Волги."),
                ("Перелом: ", "удары по флангам — окружение ~300-тысячной группировки."),
                ("Кольцо: ", "срыв деблокады Манштейна, сжатие котла."),
                ("2 февраля 1943: ", "капитуляция Паулюса. Потери сторон — более миллиона человек.")],
         itog="коренной перелом войны. День славы — 2 февраля.",
         note="Три фазы: выстоять — окружить — уничтожить. Даты «Урана» и капитуляции — точно."),
    dict(n=9, kicker="1943 • КУРСК — ОРЁЛ — ХАРЬКОВ", title="Курская битва",
         card=[("ДАТЫ", "5 июля — 23 августа 1943"), ("МЕСТО", "Курский выступ"), ("КУЛЬМИНАЦИЯ", "Прохоровка, 12 июля"), ("ЗАМЫСЕЛ ВРАГА", "операция «Цитадель»")],
         items=[("Оборона: ", "немецкие клинья за неделю увязли в глубокой обороне."),
                ("Прохоровка: ", "крупнейшее встречное танковое сражение истории."),
                ("Наступление: ", "освобождены Орёл, Белгород, Харьков."),
                ("5 августа: ", "первый салют в Москве — в честь Орла и Белгорода.")],
         itog="крах последнего наступления Германии на востоке. День славы — 23 августа.",
         note="Логика: вымотать в обороне — разбить в наступлении. Прохоровка и первый салют."),
    dict(n=10, kicker="1945 • БЕРЛИН", title="Взятие Берлина. Победа",
         card=[("ДАТЫ", "16 апреля — 9 мая 1945"), ("ФРОНТЫ", "Жуков, Конев, Рокоссовский"), ("СИМВОЛ", "Знамя над Рейхстагом"), ("ЦЕНА", "≈27 млн жизней")],
         items=[("Прорыв: ", "форсирование Одера и Нейсе, бои за Зееловские высоты."),
                ("Штурм: ", "окружение Берлина, уличные бои."),
                ("30 апреля: ", "Знамя Победы над Рейхстагом (Егоров и Кантария)."),
                ("Ночь на 9 мая: ", "подписан Акт о безоговорочной капитуляции Германии.")],
         itog="сокрушение нацизма. 9 Мая — главный праздник воинской славы.",
         note="Финал линии: назвать фронты, Рейхстаг, цену Победы. Пауза перед 27 млн."),
]

for b in battles:
    s = prs.slides.add_slide(BLANK)
    bg(s)
    header(s, b["kicker"], b["title"], b["n"])
    t = s.shapes.add_textbox(Inches(0.7), Inches(2.1), Inches(7.9), Inches(3.3))
    bullets(t, b["items"], size=15)
    card(s, b["card"])
    itog(s, b["itog"])
    footer(s, b["n"])
    note(s, b["note"] + " Время — 40–50 секунд на слайд.")

# ================= 11. СОВРЕМЕННОСТЬ =================
s = prs.slides.add_slide(BLANK)
bg(s)
header(s, "ПОСЛЕ 1991 ГОДА", "Современный период — честная оговорка", 11)
t = s.shapes.add_textbox(Inches(0.7), Inches(2.1), Inches(11.9), Inches(2.6))
bullets(t, [
    ("Факт: ", "после распада СССР Россия участвовала в военных конфликтах и операциях: Чечня (1990–2000-е), 2008 год, Сирия (с 2015)."),
    ("Проблема: ", "официального перечня «ключевых побед» по этому периоду — как в законе о днях славы — нет."),
    ("Решение: ", "в доклад взяты исторические победы с научным консенсусом; современность не смешиваем с историческим списком."),
], size=17, space=Pt(10))
footer(s, 11)
note(s, "30 сек: показать, что группа осознанно ограничилась историческими победами. Сильный ход.")

# ================= 12. ВЫВОД =================
s = prs.slides.add_slide(BLANK)
bg(s)
header(s, "ВЫВОД", "Почему эти победы важны", 12)
t = s.shapes.add_textbox(Inches(0.7), Inches(2.1), Inches(11.9), Inches(0.7))
txt(t, "Они меняли не исход боя, а положение государства.", size=17, color=WHITE)
blocks = [("ЗАЩИТА", "территории\nи независимости"), ("ПЕРЕЛОМ", "хода крупных\nвойн"), ("ГОСУДАРСТВЕННОСТЬ", "развитие\nи консолидация")]
for i, (h, d) in enumerate(blocks):
    c = box(s, 0.7 + i * 4.1, 3.1, 3.8, 1.9, CARD, GOLD)
    tf = c.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = h + "\n"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = GOLD
    r.font.name = FONT
    r = p.add_run()
    r.text = d
    r.font.size = Pt(15)
    r.font.color.rgb = WHITE
    r.font.name = FONT
t = s.shapes.add_textbox(Inches(0.7), Inches(5.4), Inches(11.9), Inches(0.8))
txt(t, "Военная победа важна историческими последствиями, а не только выигранным боем.",
    size=15, bold=True, color=GRAY)
footer(s, 12)
note(s, "Вывод 30 сек: три слова — защита, перелом, государственность.")

# ================= 13. ИСТОЧНИКИ =================
s = prs.slides.add_slide(BLANK)
bg(s)
header(s, "ИСТОЧНИКИ", "Основные материалы", 13)
src = ["Федеральный закон № 32-ФЗ «О днях воинской славы и памятных датах России».",
       "Большая российская энциклопедия — статьи по основным сражениям.",
       "Президентская библиотека им. Б. Н. Ельцина.",
       "Официальные исторические и архивные материалы."]
for i, text in enumerate(src):
    y = 2.2 + i * 0.85
    c = box(s, 0.7, y, 0.55, 0.55, GOLD)
    txt(c, str(i + 1), size=16, bold=True, color=BG, align=PP_ALIGN.CENTER)
    c.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    c.text_frame.margin_left = 0
    c.text_frame.margin_right = 0
    t = s.shapes.add_textbox(Inches(1.5), Inches(y), Inches(10.9), Inches(0.6))
    txt(t, text, size=15, color=WHITE)
t = s.shapes.add_textbox(Inches(0.7), Inches(5.8), Inches(11.9), Inches(0.6))
txt(t, "Спасибо за внимание!", size=22, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
footer(s, 13)
note(s, "Финал: источники + благодарность. Если спросят про цифры — по старым векам это оценки.")

prs.save("/home/user/Homework/klyuchevye-voennye-pobedy-rossii/prezentatsiya-gruppa-PI-262.pptx")
print("saved")
