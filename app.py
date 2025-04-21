import streamlit as st
import pandas as pd
import numpy as np
import os
import sys
import io
from datetime import datetime
import base64

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

# CSS для улучшения внешнего вида
st.markdown("""
<style>
    .main {
        padding-top: 1rem;
    }
    .block-container {
        padding-top: 1rem;
    }
    h1 {
        font-size: 2.5rem;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 5px;
        padding-left: 10px;
        padding-right: 10px;
        font-size: 14px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #4e88e5 !important;
        color: white !important;
    }
    .upload-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        padding: 20px;
        margin-top: 20px;
        border: 2px dashed #4e88e5;
        border-radius: 10px;
        background-color: #f8f9fa;
    }
    .success-box {
        background-color: #d1e7dd;
        color: #0a3622;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
    }
    .warning-box {
        background-color: #fff3cd;
        color: #856404;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
    }
    .button-primary {
        background-color: #4e88e5;
        color: white;
        padding: 10px 20px;
        border-radius: 5px;
        border: none;
        cursor: pointer;
        font-weight: bold;
    }
    .logo-container {
        display: flex;
        justify-content: center;
        margin-bottom: 20px;
    }
</style>
""", unsafe_allow_html=True)

# Безопасный импорт модулей
try:
    sys.path.append(BASE_DIR)
    from marketplace_detection import detect_marketplace
    from utils import convert_table_format
except Exception as e:
    st.sidebar.error(f"Ошибка импорта модулей: {str(e)}")

# Глобальные переменные
MARKETPLACE_NAMES = ["Ozon", "Wildberries", "ЛеманПро", "Яндекс.Маркет", "Все инструменты", "СберМегаМаркет"]
MARKETPLACE_KEYS = ["ozon", "wildberries", "lemanpro", "yandex_market", "vse_instrumenty", "sber_mega_market"]

# Функция для загрузки и отображения логотипа
def load_logo(marketplace):
    """Загружает и отображает логотип маркетплейса"""
    logo_files = {
        "Ozon": "ozon.png",
        "Wildberries": "wildberries.png",
        "ЛеманПро": "lemanpro.png",
        "Яндекс.Маркет": "yandex-market.png",
        "Все инструменты": "vse-instrumenty.png",
        "СберМегаМаркет": "sbermegamarket.png"
    }
    
    try:
        if marketplace in logo_files:
            logo_path = os.path.join(BASE_DIR, "attached_assets", logo_files[marketplace])
            if os.path.exists(logo_path):
                st.markdown('<div class="logo-container">', unsafe_allow_html=True)
                st.image(logo_path, width=150)
                st.markdown('</div>', unsafe_allow_html=True)
            else:
                pass
    except Exception as e:
        pass

# Функция для создания ссылки скачивания
def create_download_link(df, filename):
    """Создает ссылку для скачивания DataFrame как Excel-файла"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    download_str = f"""
    <div style="text-align: center; margin-top: 20px;">
        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
           download="{filename}" 
           style="background-color: #4CAF50; color: white; padding: 12px 20px; 
                  text-decoration: none; border-radius: 5px; font-weight: bold;">
            Скачать конвертированный файл
        </a>
    </div>
    """
    return download_str

# Заголовок приложения
st.title("ProductTableManager")
st.write("Инструмент для работы с таблицами товаров разных маркетплейсов")
st.markdown("---")

try:
    # Создаем табы
    tabs = st.tabs(MARKETPLACE_NAMES)
    
    # Обработка каждого таба
    for i, tab in enumerate(tabs):
        with tab:
            marketplace = MARKETPLACE_NAMES[i]
            
            # Отображаем логотип
            load_logo(marketplace)
            
            # Контейнер для загрузки файла
            st.markdown('<div class="upload-container">', unsafe_allow_html=True)
            st.subheader(f"Загрузите таблицу {marketplace}")
            uploaded_file = st.file_uploader("", type=["xlsx", "xls"], key=f"upload_{marketplace}")
            st.markdown('</div>', unsafe_allow_html=True)
            
            if uploaded_file is not None:
                try:
                    # Чтение файла
                    df = pd.read_excel(uploaded_file)
                    
                    # Базовая валидация
                    if df.empty:
                        st.warning("Загруженный файл не содержит данных")
                        st.stop()
                    
                    # Отображение предпросмотра
                    st.subheader("Предварительный просмотр данных")
                    st.dataframe(df.head())
                    
                    # Определение маркетплейса
                    try:
                        detected_marketplace = detect_marketplace(df)
                        if detected_marketplace:
                            st.markdown(f'<div class="success-box">Обнаружен формат маркетплейса: {detected_marketplace}</div>', unsafe_allow_html=True)
                        else:
                            st.markdown(f'<div class="warning-box">Не удалось определить формат маркетплейса. Используется: {marketplace}</div>', unsafe_allow_html=True)
                            detected_marketplace = marketplace
                    except Exception as e:
                        st.markdown(f'<div class="warning-box">Ошибка при определении маркетплейса: {str(e)}</div>', unsafe_allow_html=True)
                        detected_marketplace = marketplace
                    
                    # Конвертация формата
                    st.markdown("---")
                    st.subheader("Конвертация формата таблицы")
                    
                    # Выбор целевого формата
                    col1, col2 = st.columns([3, 1])
                    with col1:
                        target_formats = [m for m in MARKETPLACE_NAMES if m != detected_marketplace]
                        target_marketplace = st.selectbox(
                            "Выберите целевой формат для конвертации:", 
                            target_formats,
                            key=f"target_{marketplace}"
                        )
                    
                    # Кнопка конвертации
                    with col2:
                        convert_clicked = st.button("Конвертировать", key=f"convert_{marketplace}", 
                                                   help="Нажмите для конвертации таблицы в выбранный формат")
                    
                    if convert_clicked:
                        with st.spinner("Выполняется конвертация..."):
                            try:
                                # Конвертация
                                converted_df = convert_table_format(df, detected_marketplace, target_marketplace)
                                
                                # Успешная конвертация
                                st.markdown('<div class="success-box" style="text-align: center; padding: 15px;">'
                                          '<h3 style="margin: 0;">Таблица успешно конвертирована!</h3>'
                                          '</div>', unsafe_allow_html=True)
                                
                                # Предпросмотр конвертированных данных
                                st.subheader("Предварительный просмотр конвертированных данных")
                                st.dataframe(converted_df.head())
                                
                                # Ссылка для скачивания
                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                converted_filename = f"converted_{detected_marketplace}_to_{target_marketplace}_{timestamp}.xlsx"
                                download_link = create_download_link(converted_df, converted_filename)
                                
                                if download_link:
                                    st.markdown(download_link, unsafe_allow_html=True)
                                    
                                    # Опциональное сохранение файла
                                    try:
                                        os.makedirs(os.path.join(BASE_DIR, "data"), exist_ok=True)
                                        converted_path = os.path.join(BASE_DIR, "data", converted_filename)
                                        converted_df.to_excel(converted_path, index=False)
                                    except:
                                        pass
                            except Exception as e:
                                st.error(f"Ошибка конвертации: {str(e)}")
                except Exception as e:
                    st.error(f"Ошибка при обработке файла: {str(e)}")
    
    # Информация о приложении
    st.markdown("---")
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

except Exception as e:
    # В случае ошибки с табами, используем альтернативный интерфейс
    st.error("Не удалось создать интерфейс с вкладками. Используется альтернативный интерфейс.")
    st.error(f"Причина: {str(e)}")
    
    # Выбор маркетплейса через радио-кнопки
    marketplace = st.radio("Выберите маркетплейс:", MARKETPLACE_NAMES)
    
    # Отображаем логотип
    load_logo(marketplace)
    
    # Загрузка файла
    st.subheader(f"Загрузите таблицу {marketplace}")
    uploaded_file = st.file_uploader("", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Процесс обработки такой же, как и в основном интерфейсе...
        try:
            # Чтение файла
            df = pd.read_excel(uploaded_file)
            
            # Отображение предпросмотра
            st.subheader("Предварительный просмотр данных")
            st.dataframe(df.head())
            
            # Определение маркетплейса
            detected_marketplace = detect_marketplace(df)
            if detected_marketplace:
                st.success(f"Обнаружен формат маркетплейса: {detected_marketplace}")
            else:
                st.warning("Не удалось определить формат маркетплейса")
                detected_marketplace = marketplace
            
            # Конвертация формата
            st.subheader("Конвертация формата таблицы")
            
            # Выбор целевого формата
            target_formats = [m for m in MARKETPLACE_NAMES if m != detected_marketplace]
            target_marketplace = st.selectbox("Выберите целевой формат для конвертации:", target_formats)
            
            # Кнопка конвертации
            if st.button("Конвертировать"):
                with st.spinner("Выполняется конвертация..."):
                    try:
                        # Конвертация
                        converted_df = convert_table_format(df, detected_marketplace, target_marketplace)
                        
                        # Предпросмотр конвертированных данных
                        st.subheader("Предварительный просмотр конвертированных данных")
                        st.dataframe(converted_df.head())
                        
                        # Ссылка для скачивания
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        converted_filename = f"converted_{detected_marketplace}_to_{target_marketplace}_{timestamp}.xlsx"
                        download_link = create_download_link(converted_df, converted_filename)
                        
                        if download_link:
                            st.markdown(download_link, unsafe_allow_html=True)
                            st.success(f"Файл успешно конвертирован!")
                    except Exception as e:
                        st.error(f"Ошибка конвертации: {str(e)}")
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {str(e)}")
