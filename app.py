import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import io
from datetime import datetime
import base64

# Определяем базовый путь безопасным способом
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()

# Настройка страницы
st.set_page_config(
    page_title="ProductTableManager",
    page_icon="📊",
    layout="wide"
)

# Блок для всего приложения в try-except
try:
    # Пытаемся создать необходимые папки
    try:
        for folder in ["data", "static", "backups"]:
            folder_path = os.path.join(BASE_DIR, folder)
            os.makedirs(folder_path, exist_ok=True)
    except Exception as e:
        st.warning(f"Информация: {str(e)}")
        # Продолжаем работу без создания папок

    # Импорт модулей
    try:
        sys.path.append(BASE_DIR)
        from marketplace_detection import detect_marketplace
        from utils import convert_table_format
        st.sidebar.success("Модули успешно импортированы")
    except Exception as e:
        st.sidebar.error(f"Ошибка импорта модулей: {str(e)}")
        # Можно продолжить с минимальной функциональностью
    
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
                logo_path = os.path.join(BASE_DIR, logo_paths[marketplace])
                if os.path.exists(logo_path):
                    return st.image(logo_path, width=100)
                else:
                    st.info(f"Логотип не найден по пути: {logo_path}")
        except Exception as e:
            st.info(f"Не удалось загрузить логотип {marketplace}")
    
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
            # Чтение Excel-файла
            df = pd.read_excel(uploaded_file)
            
            # Если DataFrame пустой или имеет проблемы
            if df.empty:
                st.warning("Загруженный файл не содержит данных.")
                st.stop()
            
            # Отображаем первые 5 строк
            st.subheader("Предварительный просмотр данных")
            st.dataframe(df.head())
            
            # Определяем маркетплейс
            try:
                detected_marketplace = detect_marketplace(df)
                if detected_marketplace:
                    st.success(f"Обнаружен формат маркетплейса: {detected_marketplace}")
                else:
                    st.warning("Не удалось определить формат маркетплейса автоматически.")
                    detected_marketplace = "Unknown"
            except Exception as e:
                st.warning(f"Ошибка при определении маркетплейса: {str(e)}")
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
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        converted_filename = f"converted_{detected_marketplace}_to_{target_marketplace}_{timestamp}.xlsx"
                        download_link = create_download_link(converted_df, converted_filename)
                        
                        if download_link:
                            st.markdown(download_link, unsafe_allow_html=True)
                            st.success(f"Файл успешно конвертирован!")
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

    # Отладочная информация в сайдбаре
    with st.sidebar:
        with st.expander("Отладочная информация"):
            st.write(f"Python версия: {sys.version}")
            st.write(f"Текущая директория: {BASE_DIR}")
            
            try:
                st.write(f"Файлы в директории: {[f for f in os.listdir(BASE_DIR) if os.path.isfile(os.path.join(BASE_DIR, f))]}")
            except Exception as e:
                st.write(f"Не удалось получить список файлов: {str(e)}")

except Exception as e:
    st.error("Произошла непредвиденная ошибка!")
    st.error(f"Тип ошибки: {type(e).__name__}")
    st.error(f"Сообщение об ошибке: {str(e)}")
    
    # Отладочная информация
    st.write(f"Python версия: {sys.version}")
    st.write(f"Текущая директория: {os.getcwd()}")
    try:
        st.write(f"Список файлов в директории: {os.listdir()}")
    except:
        st.write("Не удалось получить список файлов")
    
    st.exception(e)
