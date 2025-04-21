import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import io
import tempfile
import os
import re
from fuzzywuzzy import fuzz
from utils import (
    load_excel_file, 
    save_excel_file, 
    map_columns_automatically, 
    transfer_data_between_tables,
    preview_data
)

# Настройка страницы
st.set_page_config(
    page_title="Маппинг таблиц маркетплейсов",
    page_icon="📊",
    layout="wide"
)

# Инициализация состояний сессии
if 'source_file' not in st.session_state:
    st.session_state.source_file = None
if 'target_file' not in st.session_state:
    st.session_state.target_file = None
if 'source_data' not in st.session_state:
    st.session_state.source_data = None
if 'target_data' not in st.session_state:
    st.session_state.target_data = None
if 'source_columns' not in st.session_state:
    st.session_state.source_columns = None
if 'target_columns' not in st.session_state:
    st.session_state.target_columns = None
if 'column_mapping' not in st.session_state:
    st.session_state.column_mapping = {}
if 'mapping_complete' not in st.session_state:
    st.session_state.mapping_complete = False
if 'transfer_complete' not in st.session_state:
    st.session_state.transfer_complete = False
if 'auto_mapped' not in st.session_state:
    st.session_state.auto_mapped = False
if 'source_sheet_name' not in st.session_state:
    st.session_state.source_sheet_name = None
if 'target_sheet_name' not in st.session_state:
    st.session_state.target_sheet_name = None
if 'source_sheets' not in st.session_state:
    st.session_state.source_sheets = []
if 'target_sheets' not in st.session_state:
    st.session_state.target_sheets = []
if 'source_workbook' not in st.session_state:
    st.session_state.source_workbook = None
if 'target_workbook' not in st.session_state:
    st.session_state.target_workbook = None
if 'preview_result' not in st.session_state:
    st.session_state.preview_result = None
if 'source_header_row' not in st.session_state:
    st.session_state.source_header_row = 1
if 'target_header_row' not in st.session_state:
    st.session_state.target_header_row = 1

# Заголовок и описание
st.title("🔄 Маппинг таблиц маркетплейсов")
st.markdown("""
### Инструмент для переноса данных между шаблонами таблиц маркетплейсов

Этот инструмент позволяет:
- Загружать таблицы Excel из Ozon, Wildberries и других источников
- Автоматически сопоставлять колонки между таблицами
- Вручную настраивать сопоставление колонок
- Переносить данные с сохранением исходной структуры файлов
- Экспортировать обновленные данные в исходном формате
""")

st.divider()

# Основное содержимое: два столбца
col1, col2 = st.columns(2)

with col1:
    st.subheader("📤 Исходная таблица (Откуда)")
    source_file = st.file_uploader("Загрузите исходную таблицу (xlsx)", type=['xlsx'], key="source_uploader")
    
    if source_file is not None and source_file != st.session_state.source_file:
        st.session_state.source_file = source_file
        try:
            source_workbook, source_sheets = load_excel_file(source_file)
            st.session_state.source_workbook = source_workbook
            st.session_state.source_sheets = source_sheets
            
            if len(source_sheets) > 0:
                default_sheet = 0
                st.session_state.source_sheet_name = source_sheets[default_sheet]
            else:
                st.error("В исходном файле не найдено листов!")
                st.session_state.source_data = None
        except Exception as e:
            st.error(f"Ошибка при загрузке исходного файла: {str(e)}")
            st.session_state.source_data = None
    
    if st.session_state.source_workbook is not None and st.session_state.source_sheets:
        col1a, col1b = st.columns([3, 1])
        with col1a:
            selected_source_sheet = st.selectbox(
                "Выберите лист в исходной таблице", 
                st.session_state.source_sheets,
                index=st.session_state.source_sheets.index(st.session_state.source_sheet_name) if st.session_state.source_sheet_name in st.session_state.source_sheets else 0
            )
        
        with col1b:
            # Добавляем выбор строки с заголовками
            source_header_row = st.number_input(
                "Строка заголовков",
                min_value=1,
                max_value=20,
                value=st.session_state.source_header_row,
                step=1,
                key="source_header_input"
            )
        
        if selected_source_sheet != st.session_state.source_sheet_name or source_header_row != st.session_state.source_header_row:
            st.session_state.source_sheet_name = selected_source_sheet
            st.session_state.source_header_row = source_header_row
            st.session_state.auto_mapped = False
            st.session_state.mapping_complete = False
            st.session_state.transfer_complete = False
        
        try:
            sheet = st.session_state.source_workbook[st.session_state.source_sheet_name]
            # Определяем заголовки колонок (используем выбранную пользователем строку)
            header_row = st.session_state.source_header_row
            headers = []
            column_indices = []
            
            # Собираем заголовки и их индексы
            for i, cell in enumerate(sheet[header_row]):
                if cell.value is not None and str(cell.value).strip() != "":
                    headers.append(str(cell.value))
                    column_indices.append(i)
                
            # Проверяем на дубликаты и исправляем
            unique_headers = {}
            for i, header in enumerate(headers):
                if header in unique_headers:
                    # Если заголовок уже существует, добавляем суффикс
                    counter = 1
                    new_header = f"{header}_{counter}"
                    while new_header in unique_headers:
                        counter += 1
                        new_header = f"{header}_{counter}"
                    headers[i] = new_header
                unique_headers[headers[i]] = True
                
            # Читаем данные начиная со следующей строки после заголовка
            data = []
            for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                if any(cell is not None for cell in row):
                    # Берем только данные из столбцов с заголовками
                    row_data = [row[idx] for idx in column_indices]
                    data.append(row_data)
            
            # Создаем DataFrame только с непустыми заголовками
            df = pd.DataFrame(data, columns=headers)
            
            st.session_state.source_data = df
            st.session_state.source_columns = headers
            
            # Показываем предпросмотр исходной таблицы
            st.write("Предпросмотр исходной таблицы:")
            st.dataframe(df.head(5))
            
        except Exception as e:
            st.error(f"Ошибка при обработке исходного файла: {str(e)}")
            st.session_state.source_data = None

with col2:
    st.subheader("📥 Целевая таблица (Куда)")
    target_file = st.file_uploader("Загрузите целевую таблицу (xlsx)", type=['xlsx'], key="target_uploader")
    
    if target_file is not None and target_file != st.session_state.target_file:
        st.session_state.target_file = target_file
        try:
            target_workbook, target_sheets = load_excel_file(target_file)
            st.session_state.target_workbook = target_workbook
            st.session_state.target_sheets = target_sheets
            
            if len(target_sheets) > 0:
                default_sheet = 0
                st.session_state.target_sheet_name = target_sheets[default_sheet]
            else:
                st.error("В целевом файле не найдено листов!")
                st.session_state.target_data = None
        except Exception as e:
            st.error(f"Ошибка при загрузке целевого файла: {str(e)}")
            st.session_state.target_data = None
    
    if st.session_state.target_workbook is not None and st.session_state.target_sheets:
        col2a, col2b = st.columns([3, 1])
        with col2a:
            selected_target_sheet = st.selectbox(
                "Выберите лист в целевой таблице", 
                st.session_state.target_sheets,
                index=st.session_state.target_sheets.index(st.session_state.target_sheet_name) if st.session_state.target_sheet_name in st.session_state.target_sheets else 0
            )
        
        with col2b:
            # Добавляем выбор строки с заголовками
            target_header_row = st.number_input(
                "Строка заголовков",
                min_value=1,
                max_value=20,
                value=st.session_state.target_header_row,
                step=1,
                key="target_header_input"
            )
        
        if selected_target_sheet != st.session_state.target_sheet_name or target_header_row != st.session_state.target_header_row:
            st.session_state.target_sheet_name = selected_target_sheet
            st.session_state.target_header_row = target_header_row
            st.session_state.auto_mapped = False
            st.session_state.mapping_complete = False
            st.session_state.transfer_complete = False
        
        try:
            sheet = st.session_state.target_workbook[st.session_state.target_sheet_name]
            # Определяем заголовки колонок (используем выбранную пользователем строку)
            header_row = st.session_state.target_header_row
            headers = []
            column_indices = []
            
            # Собираем заголовки и их индексы
            for i, cell in enumerate(sheet[header_row]):
                if cell.value is not None and str(cell.value).strip() != "":
                    headers.append(str(cell.value))
                    column_indices.append(i)
                
            # Проверяем на дубликаты и исправляем
            unique_headers = {}
            for i, header in enumerate(headers):
                if header in unique_headers:
                    # Если заголовок уже существует, добавляем суффикс
                    counter = 1
                    new_header = f"{header}_{counter}"
                    while new_header in unique_headers:
                        counter += 1
                        new_header = f"{header}_{counter}"
                    headers[i] = new_header
                unique_headers[headers[i]] = True
                
            # Читаем данные начиная со следующей строки после заголовка
            data = []
            for row in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                if any(cell is not None for cell in row):
                    # Берем только данные из столбцов с заголовками
                    row_data = [row[idx] for idx in column_indices]
                    data.append(row_data)
            
            # Создаем DataFrame только с непустыми заголовками
            df = pd.DataFrame(data, columns=headers)
            
            st.session_state.target_data = df
            st.session_state.target_columns = headers
            
            # Показываем предпросмотр целевой таблицы
            st.write("Предпросмотр целевой таблицы:")
            st.dataframe(df.head(5))
            
        except Exception as e:
            st.error(f"Ошибка при обработке целевого файла: {str(e)}")
            st.session_state.target_data = None

st.divider()

# Раздел автоматического и ручного маппинга
if st.session_state.source_data is not None and st.session_state.target_data is not None:
    st.header("🔄 Сопоставление колонок")
    
    # Кнопка для автоматического маппинга
    if not st.session_state.auto_mapped:
        if st.button("📊 Автоматический маппинг колонок"):
            with st.spinner("Выполняется автоматическое сопоставление колонок..."):
                st.session_state.column_mapping = map_columns_automatically(
                    st.session_state.source_columns,
                    st.session_state.target_columns
                )
                st.session_state.auto_mapped = True
                st.rerun()
    
    # Отображение и редактирование маппинга
    if st.session_state.auto_mapped:
        st.subheader("Настройка сопоставления колонок")
        
        # Получаем список всех колонок целевой таблицы для выбора
        all_target_columns = ["Не переносить"] + list(st.session_state.target_columns)
        
        # Разделяем колонки на автоматически сопоставленные и несопоставленные
        auto_mapped_columns = []
        unmapped_columns = []
        
        for src_col in st.session_state.source_columns:
            if src_col in st.session_state.column_mapping:
                auto_mapped_columns.append(src_col)
            else:
                unmapped_columns.append(src_col)
        
        # Создаем форму для сопоставления
        with st.form(key="mapping_form"):
            st.markdown("### ✅ Автоматически сопоставленные колонки:")
            st.caption("Слева - исходные заголовки, справа - выбор целевых заголовков.")
            
            # Создаем маппинги для автоматически сопоставленных колонок
            for src_col in auto_mapped_columns:
                # Определяем текущее значение маппинга
                current_mapping = st.session_state.column_mapping.get(src_col, None)
                
                # Определяем индекс текущего выбора
                selected_idx = 0
                if current_mapping is not None:
                    try:
                        selected_idx = all_target_columns.index(current_mapping)
                    except ValueError:
                        # Если значение не найдено, добавляем его в список
                        all_target_columns.append(current_mapping)
                        selected_idx = len(all_target_columns) - 1
                
                # Компактное отображение колонок
                cols = st.columns([3, 3])
                with cols[0]:
                    st.markdown(f"**{src_col}**", help="Исходная колонка")
                with cols[1]:
                    mapping = st.selectbox(
                        "Целевая колонка", 
                        options=all_target_columns,
                        index=selected_idx,
                        key=f"auto_map_{src_col}",
                        label_visibility="collapsed"
                    )
                    
                    # Обновляем маппинг
                    if mapping != "Не переносить":
                        st.session_state.column_mapping[src_col] = mapping
                    elif src_col in st.session_state.column_mapping:
                        del st.session_state.column_mapping[src_col]
            
            # Добавляем отступ между группами
            st.divider()
            
            # Если есть несопоставленные колонки
            if unmapped_columns:
                st.markdown("### ❌ Несопоставленные колонки:")
                st.caption("Для этих колонок нужно вручную выбрать целевые заголовки из выпадающего списка.")
                
                # Создаем маппинги для несопоставленных колонок
                for src_col in unmapped_columns:
                    # Определяем текущее значение маппинга (должно быть None)
                    current_mapping = st.session_state.column_mapping.get(src_col, None)
                    
                    # Определяем индекс текущего выбора
                    selected_idx = 0
                    if current_mapping is not None:
                        try:
                            selected_idx = all_target_columns.index(current_mapping)
                        except ValueError:
                            all_target_columns.append(current_mapping)
                            selected_idx = len(all_target_columns) - 1
                    
                    # Компактное отображение колонок
                    cols = st.columns([3, 3])
                    with cols[0]:
                        st.markdown(f"**{src_col}**", help="Исходная колонка")
                    with cols[1]:
                        mapping = st.selectbox(
                            "Целевая колонка",
                            options=all_target_columns,
                            index=selected_idx,
                            key=f"unmap_{src_col}",
                            label_visibility="collapsed"
                        )
                        
                        # Обновляем маппинг
                        if mapping != "Не переносить":
                            st.session_state.column_mapping[src_col] = mapping
                        elif src_col in st.session_state.column_mapping:
                            del st.session_state.column_mapping[src_col]
            
            # Показываем сводку по сопоставлению
            total_src = len(st.session_state.source_columns)
            total_mapped = len(st.session_state.column_mapping)
            st.info(f"Сопоставлено {total_mapped} из {total_src} исходных колонок")
            
            # Кнопка для завершения маппинга
            submitted = st.form_submit_button("✅ Подтвердить сопоставление")
            if submitted:
                st.session_state.mapping_complete = True
                st.success("Сопоставление колонок выполнено успешно!")
                st.rerun()
    
    # Отображение результатов маппинга
    if st.session_state.mapping_complete:
        st.subheader("Результаты сопоставления")
        
        # Создаем таблицу сопоставлений
        mapping_data = []
        for src_col, tgt_col in st.session_state.column_mapping.items():
            mapping_data.append({"Исходная колонка": src_col, "Целевая колонка": tgt_col})
        
        if mapping_data:
            mapping_df = pd.DataFrame(mapping_data)
            st.dataframe(mapping_df, use_container_width=True)
        else:
            st.warning("Нет сопоставленных колонок!")
        
        # Кнопка для выполнения переноса данных
        if not st.session_state.transfer_complete:
            if st.button("📤 Выполнить перенос данных"):
                with st.spinner("Выполняется перенос данных..."):
                    try:
                        # Предварительный просмотр результата
                        preview_df = preview_data(
                            st.session_state.source_data, 
                            st.session_state.target_data,
                            st.session_state.column_mapping
                        )
                        st.session_state.preview_result = preview_df
                        st.session_state.transfer_complete = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"Ошибка при переносе данных: {str(e)}")
                        st.session_state.transfer_complete = False
        
        # Отображение результатов переноса и кнопка скачивания
        if st.session_state.transfer_complete and st.session_state.preview_result is not None:
            st.subheader("Предпросмотр результата")
            st.dataframe(st.session_state.preview_result.head(10), use_container_width=True)
            
            # Кнопка для скачивания результата
            if st.button("💾 Скачать обновленный файл"):
                with st.spinner("Подготовка файла для скачивания..."):
                    try:
                        # Подготовка обновленного файла
                        output = io.BytesIO()
                        
                        # Перенос данных в целевой файл с сохранением форматирования
                        result_workbook = transfer_data_between_tables(
                            st.session_state.source_data,
                            st.session_state.target_workbook,
                            st.session_state.target_sheet_name,
                            st.session_state.column_mapping,
                            st.session_state.target_header_row
                        )
                        
                        # Сохранение результата в BytesIO буфер
                        result_workbook.save(output)
                        output.seek(0)
                        
                        # Определение имени выходного файла
                        original_filename = st.session_state.target_file.name
                        filename_parts = os.path.splitext(original_filename)
                        output_filename = f"{filename_parts[0]}_обновленный{filename_parts[1]}"
                        
                        # Скачивание файла
                        st.download_button(
                            label="📥 Скачать результат",
                            data=output,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.success(f"Файл '{output_filename}' готов к скачиванию!")
                        
                    except Exception as e:
                        st.error(f"Ошибка при подготовке файла: {str(e)}")
        
        # Кнопка для сброса маппинга и начала заново
        if st.button("🔄 Сбросить и начать заново"):
            st.session_state.column_mapping = {}
            st.session_state.mapping_complete = False
            st.session_state.transfer_complete = False
            st.session_state.auto_mapped = False
            st.session_state.preview_result = None
            st.rerun()

# Инструкции и пояснения
with st.expander("ℹ️ Инструкция по использованию"):
    st.markdown("""
    ### Как пользоваться инструментом:
    
    1. **Загрузка файлов**:
       - В левой колонке загрузите исходный файл, из которого будут взяты данные
       - В правой колонке загрузите целевой файл, в который будут переноситься данные
       - При необходимости выберите нужные листы в каждом файле
       - Укажите номер строки, содержащей заголовки для каждой таблицы
    
    2. **Сопоставление колонок**:
       - Нажмите кнопку "Автоматический маппинг колонок" для автоматического сопоставления
       - Система автоматически найдет соответствия между колонками на основе их названий
       - Проверьте результаты сопоставления в таблице
       - Нажмите "Подтвердить сопоставление" для завершения настройки
    
    3. **Перенос данных**:
       - После подтверждения сопоставления нажмите "Выполнить перенос данных"
       - Проверьте предпросмотр результата
       - Нажмите "Скачать обновленный файл" для получения файла с перенесенными данными
    
    4. **Начать заново**:
       - Для сброса текущего сопоставления и начала нового нажмите "Сбросить и начать заново"
    
    ### Особенности работы:
    - Инструмент сохраняет исходное форматирование и структуру целевого файла
    - Поддерживаются файлы Excel (.xlsx) от Ozon, Wildberries и других источников
    - Автоматический маппинг основан на сходстве заголовков колонок (например, "Артикул" и "Код товара")
    - Приложение игнорирует пустые столбцы и автоматически обрабатывает дублирующиеся заголовки
    """)

# Добавляем информацию в нижней части основного интерфейса
st.divider()
st.caption("""
**О приложении**: Инструмент для маппинга и переноса данных между таблицами маркетплейсов. 
Поддерживает работу с Ozon, Wildberries и другими Excel-шаблонами. **Версия:** 1.0
""")
