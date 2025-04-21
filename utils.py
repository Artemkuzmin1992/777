import pandas as pd
import numpy as np

def convert_table_format(df, source_marketplace, target_marketplace):
    """
    Конвертирует таблицу из формата одного маркетплейса в другой.
    
    Args:
        df (pd.DataFrame): Исходная таблица
        source_marketplace (str): Исходный маркетплейс
        target_marketplace (str): Целевой маркетплейс
    
    Returns:
        pd.DataFrame: Конвертированная таблица
    """
    # Создаем копию исходной таблицы
    result_df = df.copy()
    
    # В реальном приложении здесь была бы логика преобразования
    # В этой упрощенной версии просто возвращаем исходную таблицу
    # с добавленным столбцом, указывающим на целевой формат
    
    result_df["converted_to"] = target_marketplace
    
    return result_df
