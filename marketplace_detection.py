import pandas as pd
from fuzzywuzzy import process

def detect_marketplace(df):
    """
    Определяет маркетплейс на основе заголовков таблицы.
    
    Args:
        df (pd.DataFrame): Таблица с данными товаров
    
    Returns:
        str: Название маркетплейса или None, если не удалось определить
    """
    try:
        # Получаем заголовки колонок
        headers = list(df.columns)
        
        # Характерные заголовки для каждого маркетплейса
        marketplace_headers = {
            "Ozon": ["ID", "Артикул", "Название", "Цена", "Остаток", "Ссылка на товар"],
            "Wildberries": ["Номенклатура", "Артикул поставщика", "Баркод", "Цена СП", "Предмет"],
            "ЛеманПро": ["ID", "Артикул", "Наименование", "Цена", "Количество"],
            "Яндекс.Маркет": ["marketSku", "title", "categoryName", "price", "vendorCode"],
            "Все инструменты": ["Код товара", "Наименование", "Цена", "Наличие", "Категория"],
            "СберМегаМаркет": ["ID", "Артикул", "Наименование", "Цена продажи", "Остаток"]
        }
        
        # Определяем схожесть заголовков с каждым маркетплейсом
        best_match = None
        highest_score = 0
        
        for marketplace, expected_headers in marketplace_headers.items():
            total_score = 0
            match_count = 0
            
            for header in expected_headers:
                # Находим лучшее совпадение и его оценку
                best_header_match, score = process.extractOne(header, headers)
                if score > 70:  # Порог схожести
                    total_score += score
                    match_count += 1
            
            # Считаем среднюю оценку, если найдено хотя бы одно совпадение
            if match_count > 0:
                average_score = total_score / match_count
                # Требуем минимальное количество совпадений (например, 2)
                if match_count >= 2 and average_score > highest_score:
                    highest_score = average_score
                    best_match = marketplace
        
        return best_match
    
    except Exception as e:
        print(f"Ошибка при определении маркетплейса: {str(e)}")
        return None
