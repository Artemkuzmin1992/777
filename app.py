import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import io
from datetime import datetime
import time
import zipfile
import base64

# Создаем базовый путь для работы с относительными путями
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Создаем необходимые папки
for folder in ["data", "static", "backups"]:
    os.makedirs(os.path.join(BASE_DIR, folder), exist_ok=True)

# Настройка страницы
st.set_page_config(
    page_title="ProductTableManager",
    page_icon="📊",
    layout="wide"
)

# Блок для всего приложения в try-except
try:
    # Импорт модулей
    try:
        from marketplace_detection import detect_marketplace
        from utils import convert_table_format
    except ImportError as e:
        st.error(f"Ошибка импорта модулей: {str(e)}")
        st.stop()
    
    # Заголовок и описание
    st.title("ProductTableManager")
    st.subheader("Инструмент для работы с таблицами товаров с разных маркетплейсов")
    
    # Функция для загрузки и отображения логотипа
    def load_logo(marketplace):
        logo_paths = {
            "ozon": "attached_assets/ozon.png",
            "wildberries": "attached_assets/wildberries.png",
            "lemanpro": "attached_assets/lemanpro.png",
            "yandex_market": "attached_assets/yandex-market.png",
            "vse_instrumenty": "attached_assets/vse-instrumenty.png",
            "sber_mega_market": "attached_assets/sbermegamarket.png"
        }
        
        try:
            if marketplace in logo_paths:
                full_path = os.path.join(BASE_DIR, logo_paths[marketplace])
                if os.path.exists(full_path):
                    return st.image(full_path, width=100)
                else:
                    st.warning(f"Логотип не найден по пути: {full_path}")
        except Exception as e:
            st.warning(f"Не удалось загрузить логотип {marketplace}: {str(e)}")
    
    # Функция для создания загружаемого файла
    def create_download_link(df, filename):
        try:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            b64 = base64.b64encode(excel_data).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Скачать файл</a>'
            return href
        except Exception as e:
            st.error(f"Ошибка создания ссылки для скачивания: {str(e)}")
            return None
    
    # Создаем вкладки с использованием radio вместо st.tabs
    marketplace_options = ["Ozon", "Wildberries", "ЛеманПро", "Яндекс.Маркет", "Все инструменты", "СберМегаМаркет"]
    selected_marketplace = st.radio("Выберите маркетплейс:", marketplace_options)
    
    # Отображение контента для выбранного маркетплейса
    st.subheader(f"Работа с таблицами {selected_marketplace}")
    
    # Логотип для выбранного маркетплейса
    marketplace_key_map = {
        "Ozon": "ozon",
        "Wildberries": "wildberries",
        "ЛеманПро": "lemanpro",
        "Яндекс.Маркет": "yandex_market",
        "Все инструменты": "vse_instrumenty",
        "СберМегаМаркет": "sber_mega_market"
    }
    
    selected_key = marketplace_key_map.get(selected_marketplace)
    if selected_key:
        load_logo(selected_key)
    
    # Загрузка файла
    uploaded_file = st.file_uploader("Загрузите Excel файл с товарами", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Сохраняем файл
            original_filename = uploaded_file.name
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            save_path = os.path.join(BASE_DIR, "data", f"original_{timestamp}_{original_filename}")
            
            with open(save_path, "wb") as f:
                f.write(uploaded_file.getbuffer())
            
            # Читаем Excel
            df = pd.read_excel(uploaded_file)
            
            # Если DataFrame пустой или имеет проблемы
            if df.empty:
                st.warning("Загруженный файл не содержит данных.")
                st.stop()
            
            # Отображаем первые 5 строк
            st.subheader("Предварительный просмотр данных")
            st.dataframe(df.head())
            
            # Определяем маркетплейс
            detected_marketplace = detect_marketplace(df)
            if detected_marketplace:
                st.success(f"Обнаружен формат маркетплейса: {detected_marketplace}")
            else:
                st.warning("Не удалось определить формат маркетплейса автоматически.")
                detected_marketplace = "Unknown"
            
            # Конвертируем данные
            st.subheader("Конвертация формата таблицы")
            target_formats = [m for m in marketplace_options if m != detected_marketplace]
            
            target_marketplace = st.selectbox(
                "Выберите целевой формат для конвертации:", 
                target_formats
            )
            
            if st.button("Конвертировать"):
                with st.spinner("Конвертация..."):
                    try:
                        # Используем функцию конвертации из utils.py
                        converted_df = convert_table_format(df, detected_marketplace, target_marketplace)
                        
                        # Предварительный просмотр конвертированных данных
                        st.subheader("Предварительный просмотр конвертированных данных")
                        st.dataframe(converted_df.head())
                        
                        # Создаем ссылку для скачивания
                        converted_filename = f"converted_{detected_marketplace}_to_{target_marketplace}_{timestamp}.xlsx"
                        download_link = create_download_link(converted_df, converted_filename)
                        
                        if download_link:
                            st.markdown(download_link, unsafe_allow_html=True)
                            
                            # Сохраняем конвертированный файл
                            converted_path = os.path.join(BASE_DIR, "data", converted_filename)
                            converted_df.to_excel(converted_path, index=False)
                            st.success(f"Файл успешно конвертирован и сохранен!")
                    except Exception as e:
                        st.error(f"Ошибка конвертации: {str(e)}")
        
        except Exception as e:
            st.error(f"Ошибка обработки файла: {str(e)}")
    
    # Информация о приложении
    with st.expander("О приложении"):
        st.write("""
        ### ProductTableManager
        
        Инструмент для работы с таблицами товаров с разных маркетплейсов.
        
        Функции:
        - Определение формата таблицы товаров
        - Конвертация между форматами различных маркетплейсов
        - Предварительный просмотр данных
        - Сохранение и экспорт результатов
        
        Поддерживаемые маркетплейсы:
        - Ozon
        - Wildberries
        - ЛеманПро
        - Яндекс.Маркет
        - Все инструменты
        - СберМегаМаркет
        
        Версия: 1.0
        """)

except Exception as e:
    st.error("Произошла непредвиденная ошибка!")
    st.error(f"Тип ошибки: {type(e).__name__}")
    st.error(f"Сообщение об ошибке: {str(e)}")
    st.write(f"Python версия: {sys.version}")
    st.write(f"Текущая директория: {os.getcwd()}")
    st.write(f"Список файлов в директории: {os.listdir()}")
    st.exception(e)
