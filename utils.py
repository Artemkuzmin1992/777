import pandas as pd
import numpy as np

def convert_table_format(df, source_marketplace, target_marketplace):
    """
    Конвертирует таблицу из формата одного маркетплейса в другой с сопоставлением колонок.
    
    Args:
        df (pd.DataFrame): Исходная таблица
        source_marketplace (str): Исходный маркетплейс
        target_marketplace (str): Целевой маркетплейс
    
    Returns:
        pd.DataFrame: Конвертированная таблица
    """
    # Создаем копию исходной таблицы для работы
    df_source = df.copy()
    
    # Словари соответствия колонок для каждого маркетплейса
    marketplace_column_maps = {
        "Ozon": {
            "ID товара": "product_id",
            "Артикул": "sku",
            "Название": "title",
            "Цена": "price",
            "Остаток": "stock",
            "Бренд": "brand",
            "Категория": "category",
            "Описание": "description",
            "Изображение": "image_url",
            "Штрихкод": "barcode"
        },
        "Wildberries": {
            "Номенклатура": "product_id",
            "Артикул поставщика": "sku",
            "Предмет": "title",
            "Цена СП": "price",
            "Остаток": "stock",
            "Бренд": "brand",
            "Категория": "category",
            "Описание": "description",
            "Медиафайлы": "image_url",
            "Баркод": "barcode"
        },
        "ЛеманПро": {
            "ID товара": "product_id",
            "Артикул": "sku",
            "Наименование": "title",
            "Цена": "price",
            "Количество": "stock",
            "Бренд": "brand",
            "Категория": "category",
            "Описание товара": "description",
            "Фото": "image_url",
            "Штрихкод": "barcode"
        },
        "Яндекс.Маркет": {
            "marketSku": "product_id",
            "vendorCode": "sku",
            "title": "title",
            "price": "price",
            "stock": "stock",
            "vendor": "brand",
            "categoryName": "category",
            "description": "description",
            "imageUrl": "image_url",
            "barcode": "barcode"
        },
        "Все инструменты": {
            "Код товара": "product_id",
            "Артикул": "sku",
            "Наименование": "title",
            "Цена": "price",
            "Наличие": "stock",
            "Производитель": "brand",
            "Категория": "category",
            "Описание": "description",
            "Изображение": "image_url",
            "Штрихкод": "barcode"
        },
        "СберМегаМаркет": {
            "ID": "product_id",
            "Артикул": "sku",
            "Наименование": "title",
            "Цена продажи": "price",
            "Остаток": "stock",
            "Бренд": "brand",
            "Категория": "category",
            "Описание": "description",
            "Ссылка на изображение": "image_url",
            "Штрихкод": "barcode"
        }
    }
    
    # Дополнительные колонки, которые могут отличаться в разных маркетплейсах
    additional_columns = {
        "Ozon": {
            "Вес упаковки, г": "weight",
            "Ширина упаковки, мм": "package_width",
            "Высота упаковки, мм": "package_height",
            "Длина упаковки, мм": "package_length",
            "Ссылка на товар": "product_url"
        },
        "Wildberries": {
            "Вес": "weight",
            "Ширина": "package_width",
            "Высота": "package_height",
            "Длина": "package_length",
            "Ссылка": "product_url",
            "Размер": "size"
        },
        "ЛеманПро": {
            "Вес, г": "weight",
            "Ширина, мм": "package_width",
            "Высота, мм": "package_height",
            "Длина, мм": "package_length",
            "Ссылка на товар": "product_url"
        },
        "Яндекс.Маркет": {
            "weight": "weight",
            "width": "package_width",
            "height": "package_height",
            "length": "package_length",
            "url": "product_url"
        },
        "Все инструменты": {
            "Вес (кг)": "weight",
            "Ширина (см)": "package_width",
            "Высота (см)": "package_height",
            "Длина (см)": "package_length",
            "Ссылка на карточку": "product_url"
        },
        "СберМегаМаркет": {
            "Вес": "weight",
            "Ширина": "package_width",
            "Высота": "package_height",
            "Длина": "package_length",
            "Ссылка": "product_url"
        }
    }
    
    # Объединяем основные и дополнительные колонки
    for marketplace in marketplace_column_maps:
        marketplace_column_maps[marketplace].update(additional_columns.get(marketplace, {}))
    
    # Получаем маппинги для исходного и целевого маркетплейсов
    source_map = marketplace_column_maps.get(source_marketplace, {})
    target_map = marketplace_column_maps.get(target_marketplace, {})
    
    # Если не удалось найти маппинги, возвращаем исходную таблицу с информацией
    if not source_map or not target_map:
        df_source["conversion_info"] = f"Не удалось найти маппинг для {source_marketplace} или {target_marketplace}"
        return df_source
    
    # Создаем словарь для обратного маппинга целевого маркетплейса
    target_reverse_map = {v: k for k, v in target_map.items()}
    
    # Создаем промежуточный DataFrame с унифицированными колонками
    df_unified = pd.DataFrame()
    
    # Маппим колонки из исходного формата в унифицированный
    for src_col, unified_col in source_map.items():
        if src_col in df_source.columns:
            df_unified[unified_col] = df_source[src_col]
    
    # Создаем результирующий DataFrame с колонками целевого маркетплейса
    df_target = pd.DataFrame()
    
    # Маппим из унифицированного формата в целевой
    for unified_col, target_col in target_reverse_map.items():
        if unified_col in df_unified.columns:
            df_target[target_col] = df_unified[unified_col]
        else:
            # Если унифицированная колонка отсутствует, заполняем значением по умолчанию
            df_target[target_col] = ""
    
    # Добавляем информационные колонки
    df_target["Исходный формат"] = source_marketplace
    df_target["Целевой формат"] = target_marketplace
    df_target["Дата конвертации"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Проверяем соответствие структуры, чтобы избежать ошибок
    if df_target.empty and not df_source.empty:
        # Если что-то пошло не так, возвращаем исходную таблицу с информацией
        df_source["conversion_info"] = f"Ошибка при конвертации из {source_marketplace} в {target_marketplace}"
        return df_source
    
    return df_target

def get_marketplace_columns(marketplace):
    """
    Возвращает список ожидаемых колонок для указанного маркетплейса.
    
    Args:
        marketplace (str): Название маркетплейса
        
    Returns:
        list: Список названий колонок
    """
    marketplace_columns = {
        "Ozon": [
            "ID товара", "Артикул", "Название", "Цена", "Остаток", "Бренд", 
            "Категория", "Описание", "Изображение", "Штрихкод", "Вес упаковки, г", 
            "Ширина упаковки, мм", "Высота упаковки, мм", "Длина упаковки, мм", "Ссылка на товар"
        ],
        "Wildberries": [
            "Номенклатура", "Артикул поставщика", "Предмет", "Цена СП", "Остаток", 
            "Бренд", "Категория", "Описание", "Медиафайлы", "Баркод", "Вес", 
            "Ширина", "Высота", "Длина", "Ссылка", "Размер"
        ],
        "ЛеманПро": [
            "ID товара", "Артикул", "Наименование", "Цена", "Количество", "Бренд", 
            "Категория", "Описание товара", "Фото", "Штрихкод", "Вес, г", 
            "Ширина, мм", "Высота, мм", "Длина, мм", "Ссылка на товар"
        ],
        "Яндекс.Маркет": [
            "marketSku", "vendorCode", "title", "price", "stock", "vendor", 
            "categoryName", "description", "imageUrl", "barcode", "weight", 
            "width", "height", "length", "url"
        ],
        "Все инструменты": [
            "Код товара", "Артикул", "Наименование", "Цена", "Наличие", "Производитель", 
            "Категория", "Описание", "Изображение", "Штрихкод", "Вес (кг)", 
            "Ширина (см)", "Высота (см)", "Длина (см)", "Ссылка на карточку"
        ],
        "СберМегаМаркет": [
            "ID", "Артикул", "Наименование", "Цена продажи", "Остаток", "Бренд", 
            "Категория", "Описание", "Ссылка на изображение", "Штрихкод", "Вес", 
            "Ширина", "Высота", "Длина", "Ссылка"
        ]
    }
    
    return marketplace_columns.get(marketplace, [])
