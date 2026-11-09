from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
try:
    from pptx.dml.color import RgbColor
except ImportError:
    from pptx.dml.color import RGBColor as RgbColor

# Создаем презентацию
prs = Presentation()

# Настраиваем размеры слайдов (16:9)
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Цветовая схема
BG_COLOR = RgbColor(255, 248, 240)  # Мягкий кремовый цвет
TEXT_COLOR = RgbColor(51, 51, 51)   # Темно-серый текст
ACCENT_COLOR = RgbColor(70, 130, 180)  # Стальной синий для заголовков
FORMULA_COLOR = RgbColor(139, 0, 0)    # Темно-красный для формул

def add_title_slide(prs, title, subtitle):
    """Добавляет титульный слайд"""
    slide_layout = prs.slide_layouts[6]  # Пустой слайд
    slide = prs.slides.add_slide(slide_layout)

    # Фон
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11.333), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = title
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = ACCENT_COLOR
    title_para.alignment = PP_ALIGN.CENTER

    # Подзаголовок
    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(3.8), Inches(11.333), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.text = subtitle
    subtitle_para.font.size = Pt(28)
    subtitle_para.font.color.rgb = TEXT_COLOR
    subtitle_para.alignment = PP_ALIGN.CENTER

    return slide

def add_content_slide(prs, title, content_lines, is_formula=False):
    """Добавляет слайд с содержанием"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # Фон
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.333), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = title
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = ACCENT_COLOR

    # Содержание
    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(12.333), Inches(5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True

    for i, line in enumerate(content_lines):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        para.text = line
        para.font.size = Pt(24) if is_formula else Pt(22)
        para.font.color.rgb = FORMULA_COLOR if is_formula else TEXT_COLOR
        para.space_after = Pt(10) if is_formula else Pt(8)

    return slide

def add_task_slide(prs, task_number, title, content):
    """Добавляет слайд с условием задачи"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # Фон
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # Заголовок с номером
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.333), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = f"Задача {task_number}. {title}"
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = ACCENT_COLOR

    # Содержание
    content_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(12.333), Inches(5))
    text_frame = content_box.text_frame
    text_frame.word_wrap = True

    for i, line in enumerate(content):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        para.text = line
        para.font.size = Pt(22)
        para.font.color.rgb = TEXT_COLOR
        para.space_after = Pt(8)

    return slide

def add_solution_slide(prs, task_number, content_lines):
    """Добавляет слайд с решением задачи"""
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)

    # Фон
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = BG_COLOR

    # Заголовок
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12.333), Inches(1))
    title_frame = title_box.text_frame
    title_para = title_frame.paragraphs[0]
    title_para.text = f"Задача {task_number}. Решение"
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = ACCENT_COLOR

    # Решение
    solution_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.6), Inches(12.333), Inches(5.2))
    text_frame = solution_box.text_frame
    text_frame.word_wrap = True

    for i, line in enumerate(content_lines):
        if i == 0:
            para = text_frame.paragraphs[0]
        else:
            para = text_frame.add_paragraph()

        # Определяем, это формула или обычный текст
        if "Формула:" in line or "=" in line and len(line) < 50:
            para.font.color.rgb = FORMULA_COLOR
            para.font.size = Pt(20)
        else:
            para.font.color.rgb = TEXT_COLOR
            para.font.size = Pt(18)

        para.text = line
        para.space_after = Pt(6)

    return slide

# === СОЗДАНИЕ СЛАЙДОВ ===

# Слайд 1: Титульный
add_title_slide(
    prs,
    "Решение задач по теме",
    "\"Сила тяжести и закон всемирного тяготения\"\\\n9 класс, урок 24"
)

# Слайд 2: Цели урока
add_content_slide(prs, "Цели урока", [
    "• Систематизировать знания по теме",
    "  «Сила тяжести и закон всемирного тяготения»",
    "• Отработать навыки применения формул",
    "• Научиться решать задачи разного уровня сложности"
])

# Слайд 3: Вопросы для актуализации знаний (часть 1)
add_content_slide(prs, "Актуализация знаний", [
    "1. Какая сила действует на любое тело у поверхности Земли?",
    "2. Как направлена сила тяжести?",
    "3. По какой формуле вычисляется сила тяжести?",
    "4. Чему равно ускорение свободного падения?"
])

# Слайд 4: Вопросы для актуализации знаний (часть 2)
add_content_slide(prs, "Актуализация знаний (продолжение)", [
    "5. Чему численно равна гравитационная постоянная?",
    "6. В чём заключается закон всемирного тяготения?",
    "7. От чего зависит сила тяжести?",
    "8. Какая закономерность связывает g с высотой?"
])

# Слайд 5: Основные формулы
add_content_slide(prs, "Основные формулы", [
    "Сила тяжести:",
    "  Fтяж = m · g",
    "",
    "Закон всемирного тяготения:",
    "  F = G · (m₁ · m₂) / r²",
    "",
    "Зависимость силы тяжести от высоты:",
    "  Fтяж(h) = G · (M · m) / (R + h)²"
], is_formula=True)

# Слайд 6: Фундаментальные константы
add_content_slide(prs, "Фундаментальные константы", [
    "Гравитационная постоянная:",
    "  G = 6,67 · 10⁻¹¹ Н·м²/кг²",
    "",
    "Масса Земли:",
    "  M = 5,97 · 10²⁴ кг",
    "",
    "Радиус Земли:",
    "  R = 6,371 · 10⁶ м",
    "",
    "Ускорение свободного падения:",
    "  g = 9,8 м/с²"
], is_formula=True)

# Слайд 7: Задача 1
add_task_slide(prs, 1, "Базовая задача", [
    "Условие:",
    "",
    "Тело массой 5 кг находится на поверхности Земли.",
    "Определите силу тяжести, действующую на это тело."
])

# Слайд 8: Решение задачи 1
add_solution_slide(prs, 1, [
    "Дано:",
    "  m = 5 кг",
    "  g = 9,8 м/с²",
    "",
    "Найти: Fтяж — ?",
    "",
    "Решение:",
    "Формула: Fтяж = m · g",
    "",
    "Fтяж = 5 кг · 9,8 м/с² = 49 Н",
    "",
    "Ответ: Сила тяжести равна 49 Н."
])

# Слайд 9: Задача 2
add_task_slide(prs, 2, "Базовая задача", [
    "Условие:",
    "",
    "Сила тяжести, действующая на тело массой 500 кг,",
    "равна 4900 Н. Определите ускорение свободного падения.",
    "Почему полученное значение говорит о высоте орбиты?"
])

# Слайд 10: Решение задачи 2
add_solution_slide(prs, 2, [
    "Дано:",
    "  m = 500 кг",
    "  Fтяж = 4900 Н",
    "",
    "Найти: g — ?",
    "",
    "Решение:",
    "Формула: Fтяж = m · g",
    "",
    "g = Fтяж / m = 4900 Н / 500 кг = 9,8 м/с²",
    "",
    "Так как g на поверхности Земли равно 9,8 м/с²,",
    "то тело находится у поверхности Земли или",
    "на очень небольшой высоте.",
    "",
    "Ответ: g = 9,8 м/с², тело на поверхности Земли."
])

# Слайд 11: Задача 3
add_task_slide(prs, 3, "Средняя задача", [
    "Условие:",
    "",
    "Определите силу притяжения между двумя",
    "космическими кораблями массой по 5000 кг каждый,",
    "находящимися на расстоянии 100 м друг от друга."
])

# Слайд 12: Решение задачи 3
add_solution_slide(prs, 3, [
    "Дано:",
    "  m₁ = m₂ = 5000 кг",
    "  r = 100 м",
    "  G = 6,67 · 10⁻¹¹ Н·м²/кг²",
    "",
    "Найти: F — ?",
    "",
    "Решение:",
    "Формула: F = G · (m₁ · m₂) / r²",
    "",
    "F = 6,67 · 10⁻¹¹ · (5000 · 5000) / 100²",
    "F = 6,67 · 10⁻¹¹ · 2500 = 1,67 · 10⁻⁷ Н",
    "",
    "Ответ: F ≈ 1,67 · 10⁻⁷ Н (очень небольшая сила)"
])

# Слайд 13: Задача 4
add_task_slide(prs, 4, "Средняя задача", [
    "Условие:",
    "",
    "На какой высоте над поверхностью Земли сила тяжести",
    "будет в 4 раза меньше, чем на её поверхности?"
])

# Слайд 14: Решение задачи 4
add_solution_slide(prs, 4, [
    "Дано:",
    "  Fтяж(h) / Fтяж(поверхность) = 1/4",
    "  R = 6,371 · 10⁶ м",
    "",
    "Найти: h — ?",
    "",
    "Решение:",
    "Fтяж(h)/Fтяж(поверхность) = R²/(R + h)²",
    "",
    "R²/(R + h)² = 1/4",
    "R/(R + h) = 1/2",
    "2R = R + h",
    "h = R",
    "",
    "h = 6,371 · 10⁶ м = 6371 км",
    "",
    "Ответ: h = 6371 км"
])

# Слайд 15: Задача 5
add_task_slide(prs, 5, "Повышенная задача", [
    "Условие:",
    "",
    "Спутник движется по круговой орбите вокруг Земли",
    "на высоте 1000 км. Определите скорость спутника",
    "и период его обращения."
])

# Слайд 16: Решение задачи 5
add_solution_slide(prs, 5, [
    "Дано:",
    "  h = 10⁶ м",
    "  M = 5,97 · 10²⁴ кг",
    "  R = 6,371 · 10⁶ м",
    "",
    "Найти: v — ?, T — ?",
    "",
    "Решение:",
    "R + h = 7,371 · 10⁶ м",
    "",
    "v = √(G · M / (R + h)) = √(6,67·10⁻¹¹ · 5,97·10²⁴ / 7,371·10⁶)",
    "v ≈ 7350 м/с = 7,35 км/с",
    "",
    "T = 2π(R + h)/v = 2π · 7,371·10⁶ / 7350 ≈ 1992 с",
    "T ≈ 33 минуты",
    "",
    "Ответ: v = 7,35 км/с, T ≈ 33 мин"
])

# Слайд 17: Задача 6
add_task_slide(prs, 6, "Повышенная задача", [
    "Условие:",
    "",
    "Два тела массами m₁ = 2 кг и m₂ = 5 кг находятся",
    "на расстоянии 0,1 м друг от друга. Определите,",
    "с каким ускорением будет сближаться каждое тело."
])

# Слайд 18: Решение задачи 6
add_solution_slide(prs, 6, [
    "Дано:",
    "  m₁ = 2 кг, m₂ = 5 кг",
    "  r = 0,1 м",
    "  G = 6,67 · 10⁻¹¹ Н·м²/кг²",
    "",
    "Найти: a₁ — ?, a₂ — ?",
    "",
    "Решение:",
    "Сила: F = G·(m₁·m₂)/r² = 6,67·10⁻¹¹ · (2·5)/0,01",
    "F = 6,67·10⁻⁸ Н",
    "",
    "a₁ = F/m₁ = 6,67·10⁻⁸/2 = 3,34·10⁻⁸ м/с²",
    "a₂ = F/m₂ = 6,67·10⁻⁸/5 = 1,33·10⁻⁸ м/с²",
    "",
    "Ответ: a₁ = 3,34·10⁻⁸ м/с², a₂ = 1,33·10⁻⁸ м/с²"
])

# Слайд 19: Итоги урока
add_content_slide(prs, "Итоги урока", [
    "• Мы повторили основные формулы по теме",
    "«Сила тяжести и закон всемирного тяготения»",
    "",
    "• Решили задачи разного уровня сложности:",
    "  — базовые (применение формул)",
    "  — средние (анализ и вычисления)",
    "  — повышенные (комплексные задачи)",
    "",
    "• Закрепили навыки перевода единиц и",
    "  работы со степенями"
])

# Слайд 20: Типичные ошибки
add_content_slide(prs, "Типичные ошибки при решении задач", [
    "1. Подмена понятий веса и силы тяжести",
    "",
    "2. Забывание направления силы тяжести",
    "   (всегда вертикально вниз к центру Земли)",
    "",
    "3. Путаница в единицах измерения",
    "   (необходимо переводить все в систему СИ)",
    "",
    "4. Ошибки в работе со степенями",
    "   (особенно при возведении в квадрат)"
])

# Слайд 21: Домашнее задание
add_content_slide(prs, "Домашнее задание", [
    "1. Решить задачи:",
    "   а) Спутник вращается вокруг Земли по круговой",
    "      орбите на высоте 500 км. Определите первую",
    "      космическую скорость спутника.",
    "   б) На какой высоте над поверхностью Земли сила",
    "      тяжести в 2 раза меньше, чем на поверхности?",
    "",
    "2. Повторить формулы по теме",
    "   «Сила тяжести и закон всемирного тяготения»",
    "",
    "3. Подготовиться к контрольной работе"
])

# Слайд 22: Спасибо за внимание!
add_title_slide(prs, "Спасибо за внимание!", "До встречи на следующем уроке!")

# Сохранение презентации
output_path = r"C:\f1ex\school\physics\lessons\9\9_24_Решение_задач_Сила_тяжести_и_закон_всемирного_тяготения\урок_24_сила_тяжести.pptx"
prs.save(output_path)

print(f"Презентация успешно сохранена: {output_path}")
print(f"Всего слайдов: {len(prs.slides)}")
