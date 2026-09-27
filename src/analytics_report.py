import os
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.utils import get_column_letter


# ============================================================
# 1. ЛОГИРОВАНИЕ
# ============================================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("report_generation.log", encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


# ============================================================
# 2. СТИЛИ EXCEL
# ============================================================
header_font = Font(name='Calibri', size=12, bold=True, color="000080")
data_font = Font(name='Calibri', size=11)

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
header_border = Border(
    top=Side(style='thick', color="000080"),
    left=Side(style='thin'),
    right=Side(style='thin'),
    bottom=Side(style='thin')
)

align_left = Alignment(horizontal='left', vertical='center', indent=1)
align_right = Alignment(horizontal='right', vertical='center')
align_center = Alignment(horizontal='center', vertical='center')

header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")


def style_sheet(ws, df, start_row=0, start_col=0):
    max_row = start_row + len(df) + 1
    max_col = start_col + len(df.columns)

    for row in ws.iter_rows(min_row=start_row + 2, max_row=max_row,
                            min_col=start_col + 1, max_col=max_col):
        for cell in row:
            cell.font = data_font
            cell.border = thin_border
            if cell.column == start_col + 1:
                cell.alignment = align_left
            else:
                cell.alignment = align_right

    for col_num in range(start_col + 1, max_col + 1):
        cell = ws.cell(row=start_row + 1, column=col_num)
        cell.font = header_font
        cell.border = header_border
        cell.fill = header_fill
        cell.alignment = align_center

    for column in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column[0].column)
        for cell in column:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = min(max_length + 4, 30)

    ws.row_dimensions[start_row + 1].height = 22


# ============================================================
# 3. ПУТИ
# ============================================================
base_dir = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(base_dir, '..', 'data', 'sales_raw.csv')
output_dir = os.path.join(base_dir, '..', 'output')
charts_dir = os.path.join(output_dir, 'charts')

os.makedirs(output_dir, exist_ok=True)
os.makedirs(charts_dir, exist_ok=True)

date_str = datetime.now().strftime('%Y-%m-%d')
report_file = os.path.join(output_dir, f"report_{date_str}.xlsx")


def clean_charts_folder(folder_path):
    for file in Path(folder_path).glob("*.png"):
        try:
            file.unlink()
            logger.info(f"Удалён старый график: {file.name}")
        except Exception as e:
            logger.warning(f"Не удалось удалить {file.name}: {e}")


# ============================================================
# 4. ЗАПУСК И ВАЛИДАЦИЯ
# ============================================================
logger.info("Начало генерации отчёта")

clean_charts_folder(charts_dir)

if os.path.exists(report_file):
    try:
        with open(report_file, 'a'):
            pass
    except PermissionError:
        error_msg = (
            f"Файл {report_file} заблокирован (скорее всего, открыт в Excel). "
            "Закройте файл и запустите скрипт снова."
        )
        logger.error(error_msg)
        raise SystemExit(1)
    else:
        logger.info("Файл свободен, продолжаем...")

try:
    df = pd.read_csv(data_path)
    logger.info(f"Данные загружены: {len(df)} строк из {data_path}")
except FileNotFoundError:
    logger.error(f"Не найден файл с данными: {data_path}")
    raise SystemExit(1)

df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')
df['Месяц'] = df['Дата'].dt.to_period('M')

# Улучшенная валидация: показываем, в каких колонках и сколько пропусков
null_counts = df.isnull().sum()
if null_counts.any():
    problem_cols = {col: count for col, count in null_counts.items() if count > 0}
    problem_message = ", ".join([f"{col}: {count}" for col, count in problem_cols.items()])
    logger.warning(f"Обнаружены пропуски в данных: {problem_message}")
else:
    logger.info("Данные чистые: пропусков нет.")

# Дополнительная проверка: отрицательные суммы
if (df['Сумма'] < 0).any():
    neg_count = (df['Сумма'] < 0).sum()
    logger.warning(f"Обнаружено {neg_count} отрицательных значений в колонке 'Сумма'. Проверьте данные.")


# ============================================================
# 5. АГРЕГАЦИЯ
# ============================================================
region_agg = df.groupby(['Регион']).agg(
    Выручка=('Сумма', 'sum'),
    Средний_чек=('Сумма', lambda x: x.sum() / x.count()),
    Продажи=('Количество', 'sum')
).reset_index()
region_agg = region_agg.sort_values('Выручка', ascending=False)

category_agg = df.groupby(['Категория']).agg(
    Выручка=('Сумма', 'sum'),
    Продажи=('Количество', 'sum')
).reset_index().sort_values('Выручка', ascending=False)

monthly_trend = df.groupby('Месяц')['Сумма'].sum().reset_index()
monthly_trend.columns = ['Месяц', 'Выручка']

top_regions = region_agg.head(3)
top_categories = category_agg.head(3)

logger.info("Агрегация завершена")


# ============================================================
# 6. ГРАФИКИ
# ============================================================
plt.figure(figsize=(10, 6))
plt.bar(region_agg['Регион'], region_agg['Выручка'], color='#4c72b0')
plt.title('Выручка по регионам')
plt.xlabel('Регион')
plt.ylabel('Выручка, ₽')
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(charts_dir, 'revenue_by_region.png'))
plt.close()

plt.figure(figsize=(10, 6))
plt.plot(
    monthly_trend['Месяц'].astype(str),
    monthly_trend['Выручка'],
    marker='o',
    linewidth=2,
    markersize=6
)
plt.title('Динамика выручки по месяцам')
plt.xlabel('Месяц')
plt.ylabel('Выручка, ₽')
plt.grid(True, linestyle='--', alpha=0.3)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(os.path.join(charts_dir, 'monthly_trend.png'))
plt.close()

plt.figure(figsize=(8, 8))
plt.pie(
    category_agg['Выручка'],
    labels=category_agg['Категория'],
    autopct='%1.1f%%',
    startangle=90
)
plt.title('Доля выручки по категориям')
plt.tight_layout()
plt.savefig(os.path.join(charts_dir, 'category_share.png'))
plt.close()

logger.info(f"Графики сохранены в {charts_dir}")


# ============================================================
# 7. СОХРАНЕНИЕ В EXCEL
# ============================================================
try:
    with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
        region_agg.to_excel(writer, sheet_name='Регионы', index=False)
        style_sheet(writer.sheets['Регионы'], region_agg)

        category_agg.to_excel(writer, sheet_name='Категории', index=False)
        style_sheet(writer.sheets['Категории'], category_agg)

        monthly_trend.to_excel(writer, sheet_name='Динамика', index=False)
        style_sheet(writer.sheets['Динамика'], monthly_trend)

        top_regions.to_excel(writer, sheet_name='Топы', index=False)
        style_sheet(writer.sheets['Топы'], top_regions)

        ws_top = writer.sheets['Топы']
        ws_top.cell(row=1, column=1).value = (
            f"Топ-3 региона (по выручке) на {datetime.now().strftime('%d.%m.%Y')}"
        )
        ws_top.cell(row=1, column=1).font = Font(name='Calibri', size=14, bold=True)

    logger.info(f"[OK] Отчёт успешно создан: {report_file}")

except Exception as e:
    logger.exception(f"Критическая ошибка при записи Excel: {e}")
    raise


print('[OK] Отчёт сформирован')
print(f'Файл: {report_file}')
print(f'Графики сохранены в: {charts_dir}')
