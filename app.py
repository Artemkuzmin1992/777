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

# Пытаемся создать необходимые папки
try:
    for folder in ["data", "static", "backups"]:
        folder_path = os.path.join(BASE_DIR, folder)
        os.makedirs(folder_path, exist_ok=True)
except Exception as e:
    pass  # Молча продолжаем работу без создания папок

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
        sys.path.append(BASE_DIR)
        from marketplace_detection import detect_marketplace
        from utils import convert_table_format
    except Exception as e:
        st.sidebar.error(f"Ошибка импорта модулей: {str(e)}")
    
    # Заголовок и описание
    st.title("ProductTableManager")
    st.write("Инструмент для работы с таблицами товаров с разных маркетплейсов")
    
    # Функция для загрузки и отображения логотипа
    def load_logo(marketplace):
        logo_paths = {
            "Ozon": "attached_assets/ozon.png",
            "Wildberries": "attached_assets/wildberries.png",
            "ЛеманПро": "attached_assets/lemanpro.png",
            "Яндекс.Маркет": "attached_assets/yandex-market.png",
            "Все инструменты": "attached_assets/vse-instrumenty.png",
            "СберМегаМаркет": "attached_assets/sbermegamarket.png"
        }
        
        try:
            if marketplace in logo_paths:
                logo_path = os.path.join(BASE_DIR, logo_paths[marketplace])
                if os.path.exists(logo_path):
                    return st.image(logo_path, width=100)
                else:
                    pass  # Молча игнорируем отсутствие логотипа
        except Exception as e:
            pass  # Молча игнорируем ошибки
    
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
    
    # Создаем вкладки с использованием st.tabs
    marketplace_options = ["Ozon", "Wildberries", "ЛеманПро", "Яндекс.Маркет", "Все инструменты", "СберМегаМаркет"]
    
    try:
        # Пробуем использовать оригинальные tabs
        marketplace_tabs = st.tabs(marketplace_options)
    except Exception as e:
        # Если не сработало, используем radio как запасной вариант
        st.error(f"Не удалось создать tabs, используем radio: {str(e)}")
        selected_tab_index = st.radio("Выберите маркетплейс:", range(len(marketplace_options)), format_func=lambda x: marketplace_options[x])
        # Создаем фейковые tabs для совместимости
        marketplace_tabs = [st.container() for _ in marketplace_options]
        # Показываем только выбранный таб
        for i, tab in enumerate(marketplace_tabs):
            if i != selected_tab_index:
                tab.empty()
    
    # Обработка каждого таба
    for i, tab in enumerate(marketplace_tabs):
        with tab:
            marketplace = marketplace_options[i]
            
            # Отображаем логотип
            load_logo(marketplace)
            
            # Загрузка файла
            uploaded_file = st.file_uploader(f"Загрузите Excel файл с товарами {marketplace}", type=["xlsx", "xls"], key=f"upload_{marketplace}")
            
            if uploaded_file is not None:
                try:
                    # Читаем Excel-файл
                    df = pd.read_excel(uploaded_file)
                    
                    if df.empty:
                        st.warning("Загруженный файл не содержит данных")
                        st.stop()
                    
                    # Отображаем первые строки
                    st.subheader("Предварительный просмотр данных")
                    st.dataframe(df.head())
                    
                    # Определяем маркетплейс автоматически
                    try:
                        detected_marketplace = detect_marketplace(df)
                        if detected_marketplace:
                            st.success(f"Обнаружен формат маркетплейса: {detected_marketplace}")
                        else:
                            st.warning("Не удалось определить формат маркетплейса автоматически")
                            detected_marketplace = marketplace  # Используем текущий таб как формат
                    except Exception as e:
                        st.warning(f"Ошибка при определении маркетплейса: {str(e)}")
                        detected_marketplace = marketplace  # Используем текущий таб как формат
                    
                    # Конвертация
                    st.subheader("Конвертация формата таблицы")
                    target_formats = [m for m in marketplace_options if m != detected_marketplace]
                    
                    target_marketplace = st.selectbox(
                        "Выберите целевой формат для конвертации:", 
                        target_formats,
                        key=f"target_{marketplace}"
                    )
                    
                    if st.button("Конвертировать", key=f"convert_{marketplace}"):
                        with st.spinner("Конвертация..."):
                            try:
                                # Конвертируем таблицу
                                converted_df = convert_table_format(df, detected_marketplace, target_marketplace)
                                
                                # Отображаем результат
                                st.subheader("Предварительный просмотр конвертированных данных")
                                st.dataframe(converted_df.head())
                                
                                # Создаем ссылку для скачивания
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                converted_filename = f"converted_{detected_marketplace}_to_{target_marketplace}_{timestamp}.xlsx"
                                download_link = create_download_link(converted_df, converted_filename)
                                
                                if download_link:
                                    st.markdown(download_link, unsafe_allow_html=True)
                                    st.success(f"Файл успешно конвертирован!")
                                    
                                    # Попытка сохранить файл локально (может не работать на хостинге)
                                    try:
                                        converted_path = os.path.join(BASE_DIR, "data", converted_filename)
                                        converted_df.to_excel(converted_path, index=False)
                                    except:
                                        pass  # Игнорируем ошибки сохранения
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
    st.exception(e)
