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
        # Список заголовков
        headers = list(df.columns)
        
        # Характерные заголовки для каждого маркетплейса
        marketplace_headers = {
            "Ozon": ["Артикул", "Штрихкод", "Название", "Цена", "Остаток"],
            "Wildberries": ["Номенклатура", "Артикул поставщика", "Баркод", "Цена СП"],
            "ЛеманПро": ["ID", "Артикул", "Наименование", "Цена", "Количество"],
            "Яндекс.Маркет": ["marketSku", "title", "categoryName", "price", "stock"],
            "Все инструменты": ["Код товара", "Наименование", "Цена", "Наличие"],
            "СберМегаМаркет": ["Артикул", "Наименование", "Категория", "Цена продажи"]
        }
        
        # Определяем схожесть заголовков с каждым маркетплейсом
        best_match = None
        highest_score = 0
        
        for marketplace, expected_headers in marketplace_headers.items():
            total_score = 0
            
            for header in expected_headers:
                # Находим лучшее совпадение для каждого ожидаемого заголовка
                best_header_match, score = process.extractOne(header, headers)
                total_score += score
                
            average_score = total_score / len(expected_headers)
            
            if average_score > highest_score and average_score > 70:  # Порог схожести
                highest_score = average_score
                best_match = marketplace
        
        return best_match
    
    except Exception as e:
        print(f"Ошибка при определении маркетплейса: {str(e)}")
        return None
