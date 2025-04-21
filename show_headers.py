import streamlit as st
import pandas as pd
import openpyxl
import io
from utils import load_excel_file
import os

st.set_page_config(
    page_title="Заголовки шаблонов маркетплейсов",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Заголовки шаблонов маркетплейсов")
st.markdown("""
### Справочная информация для сопоставления колонок

Эта страница показывает заголовки колонок из шаблонов различных маркетплейсов для удобства создания маппингов.
""")

# Поиск доступных файлов шаблонов
template_files = []
assets_dir = "attached_assets"
if os.path.exists(assets_dir):
    for file in os.listdir(assets_dir):
        if file.endswith(".xlsx"):
            template_files.append(os.path.join(assets_dir, file))

col1, col2 = st.columns(2)

with col1:
    st.subheader("🟣 Шаблон Wildberries")
    try:
        # Загружаем тестовый файл шаблона Wildberries
        wb_file = None
        
        # Ищем шаблон Wildberries среди доступных файлов
        for file_path in template_files:
            if "тачки" in file_path.lower() or "wildberries" in file_path.lower():
                wb_file = file_path
                break
        
        if wb_file is None and len(template_files) > 0:
            wb_file = template_files[0]  # Берем первый файл, если не нашли подходящего
        
        if wb_file:
            with open(wb_file, "rb") as f:
                workbook, sheets = load_excel_file(f)
                
                # Ищем лист "Товары"
                target_sheet = None
                for sheet_name in sheets:
                    if sheet_name.lower() == "товары":
                        target_sheet = sheet_name
                        break
                
                if target_sheet is None and len(sheets) > 0:
                    target_sheet = sheets[0]
                
                if target_sheet:
                    st.success(f"Файл: {os.path.basename(wb_file)}, Лист: {target_sheet}")
                    
                    # Получаем заголовки (обычно в 3-й строке для Wildberries)
                    sheet = workbook[target_sheet]
                    header_row = 3  # Типичная строка для заголовков Wildberries
                    
                    headers = []
                    for cell in sheet[header_row]:
                        if cell.value is not None and str(cell.value).strip() != "":
                            headers.append(str(cell.value))
                    
                    # Отображаем заголовки в два столбика
                    if headers:
                        st.markdown("#### Заголовки колонок:")
                        
                        # Разделяем заголовки на две колонки
                        half_length = len(headers) // 2 + len(headers) % 2  # Первая колонка может быть на 1 больше
                        first_half = headers[:half_length]
                        second_half = headers[half_length:]
                        
                        # Создаем колонки
                        col_a, col_b = st.columns(2)
                        
                        # Отображаем первую половину
                        with col_a:
                            for i, header in enumerate(first_half, 1):
                                st.markdown(f"{i}. **{header}**")
                                
                        # Отображаем вторую половину
                        with col_b:
                            for i, header in enumerate(second_half, half_length + 1):
                                st.markdown(f"{i}. **{header}**")
                    else:
                        st.warning("Заголовки не найдены. Попробуйте изменить номер строки заголовков.")
                else:
                    st.error("Подходящий лист не найден в файле.")
        else:
            st.warning("Файл шаблона Wildberries не найден в директории assets.")
            
    except Exception as e:
        st.error(f"Ошибка при чтении шаблона Wildberries: {str(e)}")

with col2:
    st.subheader("🔶 Шаблон Ozon")
    try:
        # Загружаем тестовый файл шаблона Ozon
        ozon_file = None
        
        # Ищем шаблон Ozon среди доступных файлов
        for file_path in template_files:
            if "атё" in file_path.lower() or "ozon" in file_path.lower():
                ozon_file = file_path
                break
        
        if ozon_file is None and len(template_files) > 1:
            ozon_file = template_files[1]  # Берем второй файл, если не нашли подходящего
        
        if ozon_file:
            with open(ozon_file, "rb") as f:
                workbook, sheets = load_excel_file(f)
                
                # Ищем лист "Шаблон"
                target_sheet = None
                for sheet_name in sheets:
                    if sheet_name.lower() == "шаблон":
                        target_sheet = sheet_name
                        break
                
                if target_sheet is None and len(sheets) > 0:
                    target_sheet = sheets[0]
                
                if target_sheet:
                    st.success(f"Файл: {os.path.basename(ozon_file)}, Лист: {target_sheet}")
                    
                    # Получаем заголовки (обычно во 2-й строке для Ozon)
                    sheet = workbook[target_sheet]
                    header_row = 2  # Типичная строка для заголовков Ozon
                    
                    headers = []
                    for cell in sheet[header_row]:
                        if cell.value is not None and str(cell.value).strip() != "":
                            headers.append(str(cell.value))
                    
                    # Отображаем заголовки в два столбика
                    if headers:
                        st.markdown("#### Заголовки колонок:")
                        
                        # Разделяем заголовки на две колонки
                        half_length = len(headers) // 2 + len(headers) % 2  # Первая колонка может быть на 1 больше
                        first_half = headers[:half_length]
                        second_half = headers[half_length:]
                        
                        # Создаем колонки
                        col_a, col_b = st.columns(2)
                        
                        # Отображаем первую половину
                        with col_a:
                            for i, header in enumerate(first_half, 1):
                                st.markdown(f"{i}. **{header}**")
                                
                        # Отображаем вторую половину
                        with col_b:
                            for i, header in enumerate(second_half, half_length + 1):
                                st.markdown(f"{i}. **{header}**")
                    else:
                        st.warning("Заголовки не найдены. Попробуйте изменить номер строки заголовков.")
                else:
                    st.error("Подходящий лист не найден в файле.")
        else:
            st.warning("Файл шаблона Ozon не найден в директории assets.")
            
    except Exception as e:
        st.error(f"Ошибка при чтении шаблона Ozon: {str(e)}")

# Добавляем таблицу для сопоставления
st.divider()
st.subheader("🔄 Таблица для сопоставления")
st.caption("Используйте эту таблицу как справочный материал при создании маппингов")

# Создадим 5 пар столбцов для отображения соответствий колонок
mapping_data = []

# Соответствие Wildberries → Ozon (ключевые колонки)
mapping_data.append({
    "Wildberries": "Наименование", 
    "→": "→",
    "Ozon": "Название",
    "| Wildberries": "Артикул продавца",
    "→ ": "→",
    "Ozon ": "Артикул"
})

mapping_data.append({
    "Wildberries": "Цена, руб.*", 
    "→": "→",
    "Ozon": "Розничная цена",
    "| Wildberries": "Вес в упаковке, г*",
    "→ ": "→",
    "Ozon ": "Вес товара, г"
})

mapping_data.append({
    "Wildberries": "Артикул производителя", 
    "→": "→",
    "Ozon": "Штрихкод",
    "| Wildberries": "Ставка НДС (10%, 20%)",
    "→ ": "→",
    "Ozon ": "НДС, %"
})

mapping_data.append({
    "Wildberries": "Глубина упаковки, мм*", 
    "→": "→",
    "Ozon": "Длина упаковки, мм",
    "| Wildberries": "Ширина упаковки, мм*",
    "→ ": "→",
    "Ozon ": "Ширина упаковки, мм"
})

mapping_data.append({
    "Wildberries": "Высота упаковки, мм*", 
    "→": "→",
    "Ozon": "Высота упаковки, мм",
    "| Wildberries": "Бренд*",
    "→ ": "→",
    "Ozon ": "Торговая марка"
})

mapping_data.append({
    "Wildberries": "Описание", 
    "→": "→",
    "Ozon": "Описание",
    "| Wildberries": "Гарантийный срок",
    "→ ": "→",
    "Ozon ": "Гарантийный срок"
})

# Отображаем таблицу
if mapping_data:
    mapping_df = pd.DataFrame(mapping_data)
    st.dataframe(mapping_df, use_container_width=True, hide_index=True)

# Добавляем кнопку для скачивания шаблона маппинга
st.subheader("📥 Скачать шаблон для маппинга")
st.markdown("""
В этом шаблоне вы можете:
1. Отредактировать соответствия заголовков
2. Добавить новые соответствия
3. Загрузить готовый шаблон в основное приложение для автоматического маппинга
""")

download_col1, download_col2 = st.columns([1, 1])

with download_col1:
    # Создаем Excel файл с шаблоном маппинга
    def create_mapping_template():
        # Создаем DataFrame для маппинга
        wb_headers = []
        oz_headers = []
        
        # Пытаемся собрать заголовки из предыдущего анализа
        try:
            for item in mapping_data:
                if "Wildberries" in item and item["Wildberries"] and item["Wildberries"] not in wb_headers:
                    wb_headers.append(item["Wildberries"])
                if "| Wildberries" in item and item["| Wildberries"] and item["| Wildberries"] not in wb_headers:
                    wb_headers.append(item["| Wildberries"])
                if "Ozon" in item and item["Ozon"] and item["Ozon"] not in oz_headers:
                    oz_headers.append(item["Ozon"])
                if "Ozon " in item and item["Ozon "] and item["Ozon "] not in oz_headers:
                    oz_headers.append(item["Ozon "])
        except:
            # В случае ошибки добавляем базовые заголовки
            wb_headers = ["Наименование", "Артикул продавца", "Цена, руб.*", "Вес в упаковке, г*", 
                        "Артикул производителя", "Ставка НДС (10%, 20%)", "Глубина упаковки, мм*", 
                        "Ширина упаковки, мм*", "Высота упаковки, мм*", "Бренд*", "Описание", 
                        "Гарантийный срок"]
            oz_headers = ["Название", "Артикул", "Розничная цена", "Вес товара, г", "Штрихкод", 
                        "НДС, %", "Длина упаковки, мм", "Ширина упаковки, мм", "Высота упаковки, мм", 
                        "Торговая марка", "Описание", "Гарантийный срок"]
        
        # Создаем DataFrame в два столбца (Wildberries и Ozon бок о бок)
        # Сначала определяем максимальную длину списков
        max_length = max(len(wb_headers), len(oz_headers))
        
        # Создаем четыре столбца данных: Wildberries1, Ozon1, Wildberries2, Ozon2
        wb_col1 = []
        oz_col1 = []
        wb_col2 = []
        oz_col2 = []
        
        # Делим заголовки на две части и распределяем по столбцам
        half_length = (max_length + 1) // 2
        
        for i in range(half_length):
            if i < len(wb_headers):
                wb_col1.append(wb_headers[i])
                oz_val = oz_headers[i] if i < len(oz_headers) else ""
                oz_col1.append(oz_val)
            else:
                wb_col1.append("")
                oz_col1.append("")
        
        for i in range(half_length, max_length):
            if i < len(wb_headers):
                wb_col2.append(wb_headers[i])
                oz_val = oz_headers[i] if i < len(oz_headers) else ""
                oz_col2.append(oz_val)
            else:
                wb_col2.append("")
                oz_col2.append("")
        
        # Выравниваем длины столбцов
        while len(wb_col2) < len(wb_col1):
            wb_col2.append("")
            oz_col2.append("")
        
        # Создаем DataFrame с заголовками в два столбца (Wildberries1, Ozon1, Wildberries2, Ozon2)
        mapping_template = pd.DataFrame({
            "Wildberries (1)": wb_col1,
            "Ozon (1)": oz_col1,
            "Wildberries (2)": wb_col2,
            "Ozon (2)": oz_col2
        })
        
        # Создаем буфер для сохранения Excel файла
        output = io.BytesIO()
        
        # Создаем Excel-писателя
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Записываем DataFrame в Excel
            mapping_template.to_excel(writer, sheet_name='Маппинг', index=False)
            
            # Получаем рабочую книгу и лист
            workbook = writer.book
            worksheet = writer.sheets['Маппинг']
            
            # Устанавливаем ширину колонок
            worksheet.column_dimensions['A'].width = 40
            worksheet.column_dimensions['B'].width = 40
            
            # Добавляем инструкцию на отдельном листе
            instruction_sheet = workbook.create_sheet(title='Инструкция')
            
            # Добавляем текст инструкции
            instruction_sheet['A1'] = 'Инструкция по использованию шаблона маппинга'
            instruction_sheet['A3'] = '1. В колонке "Wildberries" перечислены заголовки колонок из шаблона Wildberries'
            instruction_sheet['A4'] = '2. В колонке "Ozon" перечислены соответствующие заголовки колонок из шаблона Ozon'
            instruction_sheet['A5'] = '3. Отредактируйте соответствия по необходимости'
            instruction_sheet['A6'] = '4. Вы можете добавить новые строки для дополнительных соответствий'
            instruction_sheet['A7'] = '5. Загрузите отредактированный файл в основное приложение для применения вашего маппинга'
            instruction_sheet['A9'] = 'Важно: сохраняйте структуру и названия листов для корректной работы импорта'
            
            # Устанавливаем ширину колонок в инструкции
            instruction_sheet.column_dimensions['A'].width = 120
        
        # Сохраняем Excel файл в буфер
        workbook.save(output)
        output.seek(0)
        
        return output

    # Кнопка для скачивания шаблона
    template_buffer = create_mapping_template()
    st.download_button(
        label="📝 Скачать шаблон маппинга (Excel)",
        data=template_buffer,
        file_name="template_mapping.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Скачайте шаблон для создания своего маппинга"
    )

with download_col2:
    st.info("""
    **Как использовать шаблон:**
    
    1. Скачайте Excel-файл с шаблоном маппинга
    2. Откройте его в Excel или другом редакторе
    3. Отредактируйте соответствия между колонками
    4. Сохраните файл
    5. Загрузите его в основное приложение
    6. Выберите опцию "Загрузить маппинг из файла"
    """)

# Добавляем кнопку для возврата на главную страницу
if st.button("⬅️ Вернуться на главную страницу"):
    st.switch_page("app.py")