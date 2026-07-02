"""
Анализ рисков регионов по аналогии с методологией ОСФ (скилл osf-risk-analysis).
Источник данных: Рейтинг регионов Итоги 2025 (xlsx)
Максимум баллов: 190 (блок 1 max 120 + блок 2 max 100 + блок 3 max 50+30*)
Порог «высокий рейтинг»: 130 баллов (≈68% от 190, аналог 80% у ОСФ)
"""

import openpyxl
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import os, warnings
warnings.filterwarnings('ignore')

# ── Пути ──────────────────────────────────────────────────────────────────────
XLSX_PATH = "/tmp/Reiting-regionov-Itogi-2025.xlsx"
OUT_DIR   = "/home/ubuntu/region_risk"
os.makedirs(OUT_DIR, exist_ok=True)

# ── Порог рейтинга (аналог 80 баллов у ОСФ, но шкала 0-190) ─────────────────
THRESHOLD = 130   # ~68% от 190 — «достаточный уровень антидопинговой работы»
MAX_SCORE = 190

# ── Цвета (точно по скиллу) ───────────────────────────────────────────────────
INK  = "#0F2D52"; SUB = "#7C8DA6"; GRID = "#EEF2F7"
FONT_FAMILY = "Inter, Segoe UI, Roboto, Arial, sans-serif"
ZONE_COLORS = {
    "🔴 КРИТИЧЕСКАЯ": "#DC2626",
    "🟠 УМЕРЕННАЯ":   "#F59E0B",
    "🟢 НИЗКАЯ":      "#10B981",
}

# ── 1. Загрузка данных ────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
ws = wb['Рейтинг регионов 2025']

# Заголовки (строка 3)
headers_row = [c.value for c in ws[3]]

# Данные регионов (строки 4..94)
rows = []
for row in ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True):
    fo   = row[0]   # Федеральный округ
    name = row[1]   # Субъект РФ
    if not name or not fo:
        continue
    name = str(name).strip().replace('\xa0', '').replace('  ', ' ')
    fo   = str(fo).strip()
    # Пропускаем строки-заголовки
    if name.lower() in ('субъект рф', 'субъект') or fo.lower() == 'фо':
        continue

    # Блок 1: колонки C..L (индексы 2..11), итог M (12)
    b1_vals = [row[i] if isinstance(row[i], (int, float)) else 0 for i in range(2, 12)]
    b1_total = row[12] if isinstance(row[12], (int, float)) else sum(b1_vals)

    # Блок 2: колонки N..S (13..18), итог T (19)
    b2_vals = [row[i] if isinstance(row[i], (int, float)) else 0 for i in range(13, 19)]
    b2_total = row[19] if isinstance(row[19], (int, float)) else sum(b2_vals)

    # Блок 3: колонки W..Z (22..25), итог AA (26)
    b3_vals = [row[i] if isinstance(row[i], (int, float)) else 0 for i in range(22, 26)]
    b3_total = row[26] if isinstance(row[26], (int, float)) else sum(b3_vals)

    total = row[27] if isinstance(row[27], (int, float)) else (b1_total + b2_total + b3_total)
    place = row[28] if isinstance(row[28], (int, float)) else None

    rows.append({
        'ФО': fo,
        'Регион': name,
        'Блок 1 (Взаимодействие с РУСАДА)': b1_total if b1_total else 0,
        'Блок 2 (Образование)': b2_total if b2_total else 0,
        'Блок 3 (Профилактика здравоохранения)': b3_total if b3_total else 0,
        'Итого баллов': total if total else 0,
        'Место': place,
        # Детали блока 1
        'B1: Ответственные/Сайт': b1_vals[0],
        'B1: Повышение квалификации': b1_vals[1],
        'B1: Соглашение с РУСАДА': b1_vals[2],
        'B1: План-график 2026': b1_vals[3],
        'B1: Отчёт за 2025': b1_vals[4],
        'B1: Опрос': b1_vals[5],
        'B1: Участие в мероприятиях': b1_vals[6],
        'B1: Коммуникация': b1_vals[7],
        'B1: Мониторинг (штраф)': b1_vals[8],
        'B1: Организация мероприятий *': b1_vals[9],
        # Детали блока 2
        'B2: Квалификация лекторов': b2_vals[0],
        'B2: Обновление программ СШ': b2_vals[1],
        'B2: НМС для СШ': b2_vals[2],
        'B2: Соглашение о персданных': b2_vals[3],
        'B2: Уровень АД образования': b2_vals[4],
        'B2: Стенды': b2_vals[5],
        # Детали блока 3
        'B3: Ответственный Минздрава': b3_vals[0],
        'B3: Программа для медперсонала': b3_vals[1],
        'B3: Межведомственное взаимодействие': b3_vals[2],
        'B3: Инновационные системы *': b3_vals[3],
    })

df = pd.DataFrame(rows)
df['Итого баллов'] = pd.to_numeric(df['Итого баллов'], errors='coerce').fillna(0).astype(int)
df['Место'] = pd.to_numeric(df['Место'], errors='coerce')

# ── 2. Расчёт зоны риска ──────────────────────────────────────────────────────
# Логика адаптирована из скилла ОСФ:
# 🔴 КРИТИЧЕСКАЯ — итого < 50% от максимума (< 95 баллов)
# 🟠 УМЕРЕННАЯ   — 50%–68% (95–129 баллов)
# 🟢 НИЗКАЯ      — ≥ 68% (≥ 130 баллов)

def get_zone(score):
    if score < 95:
        return "🔴 КРИТИЧЕСКАЯ"
    elif score < THRESHOLD:
        return "🟠 УМЕРЕННАЯ"
    else:
        return "🟢 НИЗКАЯ"

df['Зона риска'] = df['Итого баллов'].apply(get_zone)

# ── 3. Квадранты (аналог матрицы ОСФ) ────────────────────────────────────────
# Ось X: Итого баллов (рейтинг)
# Ось Y: Зона риска (критическая/умеренная/низкая)
# Квадранты:
#   Рейтинг ≥ 130 + Зона НИЗКАЯ    → «Поддерживать, мониторить»
#   Рейтинг ≥ 130 + Зона УМЕРЕННАЯ → «Усилить образование; внесоревновательный контроль»
#   Рейтинг < 130 + Зона УМЕРЕННАЯ → «Приоритизировать невыполненные критерии»
#   Рейтинг < 130 + Зона КРИТИЧЕСКАЯ → «Срочное вмешательство РУСАДА»

def get_quadrant(row):
    score = row['Итого баллов']
    zone  = row['Зона риска']
    if score >= THRESHOLD and zone == "🟢 НИЗКАЯ":
        return "✅ Высокий рейтинг + Низкий риск"
    elif score >= THRESHOLD and zone == "🟠 УМЕРЕННАЯ":
        return "⚠️ Высокий рейтинг + Умеренный риск"
    elif score < THRESHOLD and zone == "🟠 УМЕРЕННАЯ":
        return "🔶 Низкий рейтинг + Умеренный риск"
    elif score < THRESHOLD and zone == "🔴 КРИТИЧЕСКАЯ":
        return "🚨 Низкий рейтинг + Критический риск"
    else:
        return "🔶 Низкий рейтинг + Умеренный риск"

df['Квадрант'] = df.apply(get_quadrant, axis=1)

# ── 4. Рекомендации ───────────────────────────────────────────────────────────
RECS = {
    "✅ Высокий рейтинг + Низкий риск":
        "Поддерживать достигнутый уровень; регулярный мониторинг; обмен лучшими практиками с другими регионами",
    "⚠️ Высокий рейтинг + Умеренный риск":
        "Усилить образовательные программы; провести беседы с ответственными специалистами; усилить внесоревновательный контроль",
    "🔶 Низкий рейтинг + Умеренный риск":
        "Приоритизировать невыполненные критерии рейтинга; профилактика; подтянуть рейтинг до порогового значения",
    "🚨 Низкий рейтинг + Критический риск":
        "Срочное вмешательство РУСАДА: выездные проверки, усиленный контроль, индивидуальный план исправления",
}
df['Рекомендация'] = df['Квадрант'].map(RECS)

# ── 5. Невыполненные критерии (блок 1) ───────────────────────────────────────
CRITERIA_B1 = {
    'B1: Ответственные/Сайт': 'Ответственные/Сайт (15)',
    'B1: Повышение квалификации': 'Повышение квалификации (15)',
    'B1: Соглашение с РУСАДА': 'Соглашение с РУСАДА (10)',
    'B1: План-график 2026': 'План-график 2026 (15)',
    'B1: Отчёт за 2025': 'Отчёт за 2025 (15)',
    'B1: Опрос': 'Опрос (15)',
    'B1: Участие в мероприятиях': 'Участие в мероприятиях (10)',
    'B1: Коммуникация': 'Коммуникация (5)',
}
CRITERIA_B2 = {
    'B2: Квалификация лекторов': 'Квалификация лекторов (15)',
    'B2: Обновление программ СШ': 'Обновление программ СШ (35)',
    'B2: НМС для СШ': 'НМС для СШ (10)',
    'B2: Соглашение о персданных': 'Соглашение о персданных (10)',
    'B2: Уровень АД образования': 'Уровень АД образования (20)',
    'B2: Стенды': 'Стенды (10)',
}
CRITERIA_B3 = {
    'B3: Ответственный Минздрава': 'Ответственный Минздрава (10)',
    'B3: Программа для медперсонала': 'Программа для медперсонала (25)',
    'B3: Межведомственное взаимодействие': 'Межведомственное взаимодействие (15)',
}

ALL_CRITERIA = {**CRITERIA_B1, **CRITERIA_B2, **CRITERIA_B3}

def get_not_done(row):
    not_done = []
    for col, label in ALL_CRITERIA.items():
        val = row.get(col, 0)
        if val is None or val == 0:
            not_done.append(label)
    return "; ".join(not_done) if not_done else "—"

def get_done(row):
    done = []
    for col, label in ALL_CRITERIA.items():
        val = row.get(col, 0)
        if val and val > 0:
            done.append(label)
    return "; ".join(done) if done else "—"

df['Выполненные критерии'] = df.apply(get_done, axis=1)
df['Невыполненные критерии'] = df.apply(get_not_done, axis=1)

# ── 6. Процент выполнения ─────────────────────────────────────────────────────
df['% выполнения'] = (df['Итого баллов'] / MAX_SCORE * 100).round(1)

# ── 7. Сортировка ─────────────────────────────────────────────────────────────
zone_order = {"🔴 КРИТИЧЕСКАЯ": 0, "🟠 УМЕРЕННАЯ": 1, "🟢 НИЗКАЯ": 2}
df['_zone_order'] = df['Зона риска'].map(zone_order)
df = df.sort_values(['_zone_order', 'Итого баллов']).reset_index(drop=True)
df.drop(columns=['_zone_order'], inplace=True)
df.index = range(1, len(df)+1)

print(f"Всего регионов: {len(df)}")
print(df[['Регион', 'Итого баллов', 'Зона риска', 'Квадрант']].head(10))

# ── 8. Excel-отчёт ───────────────────────────────────────────────────────────
OUT_XLSX = os.path.join(OUT_DIR, "region_risk_final.xlsx")

wb_out = openpyxl.Workbook()

# --- Лист «Матрица регионов» ---
ws_main = wb_out.active
ws_main.title = "Матрица регионов"

# Цвета заливки
FILL_RED    = PatternFill("solid", fgColor="FEE2E2")
FILL_ORANGE = PatternFill("solid", fgColor="FEF3C7")
FILL_GREEN  = PatternFill("solid", fgColor="D1FAE5")
FILL_HEADER = PatternFill("solid", fgColor="0F2D52")

FONT_HEADER = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
FONT_BOLD   = Font(bold=True, name="Calibri", size=9)
FONT_NORM   = Font(name="Calibri", size=9)

BORDER_THIN = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC'),
)

COLS = [
    '№', 'ФО', 'Регион', 'Итого баллов', '% выполнения', 'Место',
    'Зона риска', 'Квадрант',
    'Блок 1 (Взаимодействие с РУСАДА)',
    'Блок 2 (Образование)',
    'Блок 3 (Профилактика здравоохранения)',
    'Выполненные критерии', 'Невыполненные критерии', 'Рекомендация'
]

# Заголовок
ws_main.append(['АНАЛИЗ РИСКОВ РЕГИОНОВ — РЕЙТИНГ РУСАДА 2025'])
ws_main.merge_cells(f'A1:{get_column_letter(len(COLS))}1')
ws_main['A1'].font = Font(bold=True, size=13, color="0F2D52", name="Calibri")
ws_main['A1'].alignment = Alignment(horizontal='center')
ws_main.row_dimensions[1].height = 22

ws_main.append([f'Максимум баллов: {MAX_SCORE} | Порог «высокий рейтинг»: {THRESHOLD} баллов | Регионов: {len(df)}'])
ws_main.merge_cells(f'A2:{get_column_letter(len(COLS))}2')
ws_main['A2'].font = Font(italic=True, size=9, color="7C8DA6", name="Calibri")
ws_main['A2'].alignment = Alignment(horizontal='center')
ws_main.row_dimensions[2].height = 16

ws_main.append([])  # пустая строка

# Заголовки столбцов (строка 4)
ws_main.append(COLS)
for col_idx, col_name in enumerate(COLS, start=1):
    cell = ws_main.cell(row=4, column=col_idx)
    cell.fill = FILL_HEADER
    cell.font = FONT_HEADER
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = BORDER_THIN
ws_main.row_dimensions[4].height = 36

# Данные
for i, row in df.iterrows():
    data_row = [
        i,
        row['ФО'],
        row['Регион'],
        row['Итого баллов'],
        f"{row['% выполнения']}%",
        row['Место'],
        row['Зона риска'],
        row['Квадрант'],
        row['Блок 1 (Взаимодействие с РУСАДА)'],
        row['Блок 2 (Образование)'],
        row['Блок 3 (Профилактика здравоохранения)'],
        row['Выполненные критерии'],
        row['Невыполненные критерии'],
        row['Рекомендация'],
    ]
    ws_main.append(data_row)
    row_num = ws_main.max_row

    # Цвет строки по зоне
    zone = row['Зона риска']
    if zone == "🔴 КРИТИЧЕСКАЯ":
        fill = FILL_RED
    elif zone == "🟠 УМЕРЕННАЯ":
        fill = FILL_ORANGE
    else:
        fill = FILL_GREEN

    for col_idx in range(1, len(COLS)+1):
        cell = ws_main.cell(row=row_num, column=col_idx)
        cell.fill = fill
        cell.font = FONT_NORM
        cell.border = BORDER_THIN
        cell.alignment = Alignment(vertical='top', wrap_text=(col_idx >= 12))

    ws_main.row_dimensions[row_num].height = 40 if row['Невыполненные критерии'] != '—' else 18

# Ширины столбцов
col_widths = [5, 7, 28, 12, 12, 8, 18, 38, 14, 14, 14, 55, 55, 60]
for i, w in enumerate(col_widths, start=1):
    ws_main.column_dimensions[get_column_letter(i)].width = w

# Заморозить строку 4
ws_main.freeze_panes = 'A5'

# --- Лист «Сводка по ФО» ---
ws_fo = wb_out.create_sheet("Сводка по ФО")
fo_summary = df.groupby('ФО').agg(
    Регионов=('Регион', 'count'),
    Критических=('Зона риска', lambda x: (x == '🔴 КРИТИЧЕСКАЯ').sum()),
    Умеренных=('Зона риска', lambda x: (x == '🟠 УМЕРЕННАЯ').sum()),
    Низких=('Зона риска', lambda x: (x == '🟢 НИЗКАЯ').sum()),
    Средний_балл=('Итого баллов', 'mean'),
    Мин_балл=('Итого баллов', 'min'),
    Макс_балл=('Итого баллов', 'max'),
).reset_index()
fo_summary['Средний_балл'] = fo_summary['Средний_балл'].round(1)
fo_summary = fo_summary.sort_values('Критических', ascending=False)

ws_fo.append(['СВОДКА ПО ФЕДЕРАЛЬНЫМ ОКРУГАМ'])
ws_fo.merge_cells('A1:H1')
ws_fo['A1'].font = Font(bold=True, size=12, color="0F2D52", name="Calibri")
ws_fo['A1'].alignment = Alignment(horizontal='center')
ws_fo.append([])

fo_headers = ['ФО', 'Регионов', '🔴 Критических', '🟠 Умеренных', '🟢 Низких', 'Средний балл', 'Мин балл', 'Макс балл']
ws_fo.append(fo_headers)
for col_idx in range(1, len(fo_headers)+1):
    cell = ws_fo.cell(row=3, column=col_idx)
    cell.fill = FILL_HEADER
    cell.font = FONT_HEADER
    cell.alignment = Alignment(horizontal='center')
    cell.border = BORDER_THIN

for _, row in fo_summary.iterrows():
    ws_fo.append([
        row['ФО'], row['Регионов'], row['Критических'],
        row['Умеренных'], row['Низких'],
        row['Средний_балл'], row['Мин_балл'], row['Макс_балл']
    ])
    r = ws_fo.max_row
    for c in range(1, 9):
        cell = ws_fo.cell(row=r, column=c)
        cell.border = BORDER_THIN
        cell.font = FONT_NORM
        cell.alignment = Alignment(horizontal='center')
    if row['Критических'] > 0:
        ws_fo.cell(row=r, column=3).fill = FILL_RED
    if row['Умеренных'] > 0:
        ws_fo.cell(row=r, column=4).fill = FILL_ORANGE
    if row['Низких'] > 0:
        ws_fo.cell(row=r, column=5).fill = FILL_GREEN

for i, w in enumerate([10, 10, 16, 16, 12, 14, 12, 12], start=1):
    ws_fo.column_dimensions[get_column_letter(i)].width = w

wb_out.save(OUT_XLSX)
print(f"\nExcel сохранён: {OUT_XLSX}")

# ── 9. Plotly HTML-дашборд ───────────────────────────────────────────────────
OUT_HTML = os.path.join(OUT_DIR, "region_risk_dashboard.html")

# Цвета точек
color_map = {
    "🔴 КРИТИЧЕСКАЯ": "#DC2626",
    "🟠 УМЕРЕННАЯ":   "#F59E0B",
    "🟢 НИЗКАЯ":      "#10B981",
}

fig = make_subplots(
    rows=2, cols=2,
    subplot_titles=[
        "Матрица рисков регионов (Рейтинг × Зона)",
        "Топ-20 регионов с наибольшим риском",
        "Распределение регионов по квадрантам",
        "Средний балл по федеральным округам",
    ],
    specs=[[{"type": "scatter"}, {"type": "bar"}],
           [{"type": "pie"},    {"type": "bar"}]],
    vertical_spacing=0.15,
    horizontal_spacing=0.1,
)

# --- График 1: Scatter матрица ---
import random
random.seed(42)
zone_y = {"🔴 КРИТИЧЕСКАЯ": 1, "🟠 УМЕРЕННАЯ": 2, "🟢 НИЗКАЯ": 3}

for zone, color in color_map.items():
    sub = df[df['Зона риска'] == zone]
    jitter_y = [zone_y[zone] + random.uniform(-0.25, 0.25) for _ in range(len(sub))]
    fig.add_trace(go.Scatter(
        x=sub['Итого баллов'],
        y=jitter_y,
        mode='markers+text',
        name=zone,
        marker=dict(color=color, size=10, opacity=0.82, line=dict(width=0.5, color='white')),
        text=sub['Регион'].apply(lambda x: x[:12] + '…' if len(x) > 12 else x),
        textposition='top center',
        textfont=dict(size=7),
        hovertemplate=(
            "<b>%{customdata[0]}</b><br>"
            "ФО: %{customdata[1]}<br>"
            "Баллы: %{x}<br>"
            "Место: %{customdata[2]}<br>"
            "Зона: " + zone + "<extra></extra>"
        ),
        customdata=list(zip(sub['Регион'], sub['ФО'], sub['Место'].fillna('—'))),
    ), row=1, col=1)

fig.add_vline(x=THRESHOLD, line_dash="dash", line_color=INK, line_width=1.5,
              annotation_text=f"Порог {THRESHOLD}", annotation_position="top right",
              row=1, col=1)

fig.update_yaxes(
    tickvals=[1, 2, 3],
    ticktext=["🔴 КРИТИЧЕСКАЯ", "🟠 УМЕРЕННАЯ", "🟢 НИЗКАЯ"],
    row=1, col=1
)
fig.update_xaxes(title_text="Итого баллов (макс. 190)", row=1, col=1)

# --- График 2: Топ-20 рисковых регионов (горизонтальный бар) ---
top20 = df[df['Зона риска'].isin(['🔴 КРИТИЧЕСКАЯ', '🟠 УМЕРЕННАЯ'])].sort_values('Итого баллов').tail(20)
bar_colors = [color_map[z] for z in top20['Зона риска']]

fig.add_trace(go.Bar(
    x=top20['Итого баллов'],
    y=top20['Регион'].apply(lambda x: x[:25] + '…' if len(x) > 25 else x),
    orientation='h',
    marker_color=bar_colors,
    marker_cornerradius=5,
    text=top20['Итого баллов'],
    textposition='outside',
    cliponaxis=False,
    hovertemplate="<b>%{y}</b><br>Баллы: %{x}<extra></extra>",
    showlegend=False,
), row=1, col=2)
fig.update_xaxes(title_text="Итого баллов", row=1, col=2)

# --- График 3: Pie по квадрантам ---
quad_counts = df['Квадрант'].value_counts()
quad_colors_map = {
    "✅ Высокий рейтинг + Низкий риск":   "#10B981",
    "⚠️ Высокий рейтинг + Умеренный риск": "#F59E0B",
    "🔶 Низкий рейтинг + Умеренный риск":  "#FB923C",
    "🚨 Низкий рейтинг + Критический риск": "#DC2626",
}
pie_colors = [quad_colors_map.get(q, "#94A3B8") for q in quad_counts.index]

fig.add_trace(go.Pie(
    labels=quad_counts.index,
    values=quad_counts.values,
    marker_colors=pie_colors,
    textinfo='label+percent+value',
    textfont=dict(size=9),
    hole=0.35,
    showlegend=False,
), row=2, col=1)

# --- График 4: Средний балл по ФО ---
fo_avg = df.groupby('ФО')['Итого баллов'].mean().sort_values()
fo_bar_colors = [
    "#DC2626" if v < 95 else ("#F59E0B" if v < THRESHOLD else "#10B981")
    for v in fo_avg.values
]

fig.add_trace(go.Bar(
    x=fo_avg.values,
    y=fo_avg.index,
    orientation='h',
    marker_color=fo_bar_colors,
    marker_cornerradius=5,
    text=[f"{v:.0f}" for v in fo_avg.values],
    textposition='outside',
    cliponaxis=False,
    showlegend=False,
    hovertemplate="<b>%{y}</b><br>Средний балл: %{x:.1f}<extra></extra>",
), row=2, col=2)
# Вертикальная линия порога на графике 4 через shapes
fig.add_shape(type='line', x0=THRESHOLD, x1=THRESHOLD, y0=-0.5, y1=len(fo_avg)-0.5,
              line=dict(dash='dash', color=INK, width=1.5), row=2, col=2)
fig.update_xaxes(title_text="Средний балл", row=2, col=2)

# --- Общий стиль ---
fig.update_layout(
    title=dict(
        text="<b>Анализ рисков регионов | Рейтинг РУСАДА 2025</b>",
        font=dict(size=18, color=INK, family=FONT_FAMILY),
        x=0.5,
    ),
    height=900,
    paper_bgcolor="white",
    plot_bgcolor=GRID,
    font=dict(family=FONT_FAMILY, color=INK),
    legend=dict(
        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
        font=dict(size=10),
    ),
    margin=dict(t=100, b=60, l=60, r=60),
)

fig.write_html(OUT_HTML)
print(f"Дашборд сохранён: {OUT_HTML}")

# ── 10. Итоговая статистика ───────────────────────────────────────────────────
print("\n=== ИТОГОВАЯ СТАТИСТИКА ===")
print(df['Зона риска'].value_counts())
print(f"\nПорог рейтинга: {THRESHOLD} баллов")
print(f"Всего регионов: {len(df)}")
