# Добавьте этот код между частью, где пользователь выбирает целевой маркетплейс и кнопкой конвертации

# Показываем сопоставление колонок
if 'utils' in sys.modules:
    try:
        with st.expander("Посмотреть схему сопоставления колонок"):
            # Получаем колонки для исходного и целевого маркетплейсов
            try:
                source_columns = utils.get_marketplace_columns(detected_marketplace)
                target_columns = utils.get_marketplace_columns(target_marketplace)
                
                # Создаем таблицу сопоставления
                mapping_data = []
                
                # Получаем маппинги для колонок
                source_map = getattr(utils, 'marketplace_column_maps', {}).get(detected_marketplace, {})
                target_map = getattr(utils, 'marketplace_column_maps', {}).get(target_marketplace, {})
                
                # Создаем обратный маппинг для целевого маркетплейса
                target_reverse_map = {v: k for k, v in target_map.items()}
                
                # Заполняем таблицу сопоставления
                for src_col, unified_col in source_map.items():
                    if unified_col in target_reverse_map:
                        target_col = target_reverse_map[unified_col]
                        mapping_data.append({"Исходная колонка": src_col, "Целевая колонка": target_col})
                
                # Отображаем таблицу сопоставления
                if mapping_data:
                    mapping_df = pd.DataFrame(mapping_data)
                    st.write("Таблица сопоставления колонок:")
                    st.dataframe(mapping_df)
                else:
                    st.write("Не удалось создать таблицу сопоставления.")
            except Exception as e:
                st.write(f"Ошибка при создании схемы сопоставления: {str(e)}")
    except Exception as e:
        st.write("Не удалось отобразить схему сопоставления колонок.")
