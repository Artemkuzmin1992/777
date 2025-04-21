import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime
import base64
import sys

# Безопасное определение BASE_DIR
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

# Импорт модулей
sys.path.append(BASE_DIR)
try:
    # Это импорт модулей из оригинального проекта
    from marketplace_detection import detect_marketplace
    from utils import convert_table_format
except Exception as e:
    st.sidebar.error(f"Ошибка импорта модулей: {str(e)}")

# Функция для создания ссылки скачивания
def create_download_link(df, filename):
    """Создает ссылку для скачивания DataFrame как Excel-файла"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Скачать файл</a>'
    return href

# Заголовок приложения
st.title("ProductTableManager")
st.write("Инструмент для работы с таблицами товаров с разных маркетплейсов")

# Создаем табы для разных маркетплейсов
marketplaces = ["Ozon", "Wildberries", "ЛеманПро", "Яндекс.Маркет", "Все инструменты", "СберМегаМаркет"]
tabs = st.tabs(marketplaces)

# Функция для отображения логотипа
def show_logo(marketplace):
    try:
        logo_path = {
            "Ozon": os.path.join(BASE_DIR, "attached_assets", "ozon.png"),
            "Wildberries": os.path.join(BASE_DIR, "attached_assets", "wildberries.png"),
            "ЛеманПро": os.path.join(BASE_DIR, "attached_assets", "lemanpro.png"),
            "Яндекс.Маркет": os.path.join(BASE_DIR, "attached_assets", "yandex-market.png"),
            "Все инструменты": os.path.join(BASE_DIR, "attached_assets", "vse-instrumenty.png"),
            "СберМегаМаркет": os.path.join(BASE_DIR, "attached_assets", "sbermegamarket.png")
        }
        if marketplace in logo_path and os.path.exists(logo_path[marketplace]):
            st.image(logo_path[marketplace], width=100)
    except Exception as e:
        pass  # Игнорируем ошибки с логотипами

# Обработка для каждого таба
for i, tab in enumerate(tabs):
    with tab:
        marketplace = marketplaces[i]
        
        # Показываем логотип
        show_logo(marketplace)
        
        # Загрузка файла
        uploaded_file = st.file_uploader(f"Загрузите таблицу товаров {marketplace}", type=["xlsx", "xls"], key=f"upload_{marketplace}")
        
        if uploaded_file is not None:
            try:
                # Чтение Excel-файла
                df = pd.read_excel(uploaded_file)
                
                if df.empty:
                    st.warning("Загруженный файл не содержит данных")
                    st.stop()
                
                # Отображаем первые строки
                st.subheader("Предварительный просмотр данных")
                st.dataframe(df.head())
                
                # Определяем маркетплейс
                try:
                    detected_marketplace = detect_marketplace(df)
                    if detected_marketplace:
                        st.success(f"Обнаружен формат маркетплейса: {detected_marketplace}")
                    else:
                        st.warning("Не удалось определить формат маркетплейса")
                        detected_marketplace = marketplace
                except Exception as e:
                    st.warning(f"Ошибка при определении маркетплейса: {str(e)}")
                    detected_marketplace = marketplace
                
                # Конвертация
                st.subheader("Конвертация формата таблицы")
                target_formats = [m for m in marketplaces if m != detected_marketplace]
                
                target_marketplace = st.selectbox(
                    "Выберите целевой формат для конвертации:", 
                    target_formats,
                    key=f"target_{marketplace}"
                )
                
                if st.button("Конвертировать", key=f"convert_{marketplace}"):
                    with st.spinner("Выполняется конвертация..."):
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
                                st.success(f"Таблица успешно конвертирована!")
                        except Exception as e:
                            st.error(f"Ошибка конвертации: {str(e)}")
            except Exception as e:
                st.error(f"Ошибка при обработке файла: {str(e)}")

# Информация о приложении
with st.expander("О приложении"):
    st.write("""
    ### ProductTableManager
    
    Инструмент для работы с таблицами товаров с разных маркетплейсов.
    
    **Функциональность**:
    - Определение формата таблицы товаров
    - Конвертация между форматами различных маркетплейсов
    - Предварительный просмотр данных
    - Экспорт конвертированных таблиц
    
    **Поддерживаемые маркетплейсы**:
    - Ozon
    - Wildberries
    - ЛеманПро
    - Яндекс.Маркет
    - Все инструменты
    - СберМегаМаркет
    
    **Версия**: 1.0
    """)
