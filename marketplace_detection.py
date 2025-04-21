"""
Модуль для распознавания шаблонов маркетплейсов (Ozon, Wildberries, ЛеманПро, Яндекс.Маркет)
на основе заголовков и содержимого файлов Excel.
"""

def detect_marketplace_by_row_headers(normalized_columns, row_num):
    """
    Определяет тип маркетплейса на основе заголовков в указанной строке.
    
    Args:
        normalized_columns: Список нормализованных (lowercase) заголовков колонок
        row_num: Номер строки (начиная с 1), которая содержит заголовки
        
    Returns:
        tuple: (marketplace, confidence, additional_info)
            - marketplace: Строка с типом маркетплейса ('ozon', 'wildberries', 'lemanpro', 'yandex', 'other')
            - confidence: Число от 0 до 100, указывающее уровень уверенности в определении
            - additional_info: Словарь с дополнительной информацией (для диагностики)
    """
    # Эталонные заголовки для каждого маркетплейса
    marketplace_headers = {
        # Строка 4 - ЛеманПро
        'lemanpro_row4': [
            'guid', 'код тн вэд', 'наименование товара мерчанта', 
            'бренд товара', 'модель товара'
        ],
        # Строка 4 - Яндекс.Маркет
        'yandex_row4': [
            'ваш sku *', 'качество карточки', 'рекомендации по заполнению', 
            'название товара *', 'ссылка на изображение *'
        ],
        # Строка 2 - Ozon
        'ozon_row2': [
            'артикул*', 'название товара*', 'ссылка на главное фото*', 
            'цена, руб.*', 'бренд*'
        ],
        # Строка 2 - Яндекс.Маркет (альтернативный вариант)
        'yandex_row2': [
            'ваш sku *', 'качество карточки', 'рекомендации по заполнению', 
            'название товара *', 'ссылка на изображение *'
        ],
        # Строка 3 - Wildberries
        'wildberries_row3': [
            'артикул продавца', 'артикул wb', 'наименование', 
            'бренд', 'фото'
        ],
        # Строка 2 - Все инструменты
        'vseinstrumenty_row2': [
            'guid*', 'бренд', 'наименование', 'артикул', 'код тн вэд'
        ]
    }
    
    # Проверка каждого типа шаблона на основе номера строки
    results = {}
    
    # Для строки 4
    if row_num == 4:
        # Проверяем первые 5 колонок на совпадение с эталонными заголовками
        first_5_columns = normalized_columns[:min(5, len(normalized_columns))]
        
        # Проверка на ЛеманПро
        lemanpro_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['lemanpro_row4']) 
                                   if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Проверка на Яндекс.Маркет
        yandex_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['yandex_row4']) 
                                 if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Дополнительные проверки
        has_guid = any('guid' in col for col in normalized_columns)
        has_yandex_sku = any('ваш sku' in col for col in normalized_columns)
        has_quality = any('качество карточки' in col for col in normalized_columns)
        
        # Общие совпадения (просто наличие характерных заголовков)
        lemanpro_matches = sum(1 for header in marketplace_headers['lemanpro_row4'] if any(header.lower() in col for col in normalized_columns))
        yandex_matches = sum(1 for header in marketplace_headers['yandex_row4'] if any(header.lower() in col for col in normalized_columns))
        
        # Результаты для строки 4
        results['lemanpro'] = {
            'exact_matches': lemanpro_exact_matches,
            'total_matches': lemanpro_matches,
            'has_guid': has_guid,
            'confidence': calculate_confidence('lemanpro', lemanpro_exact_matches, has_guid, row_num)
        }
        
        results['yandex'] = {
            'exact_matches': yandex_exact_matches,
            'total_matches': yandex_matches,
            'has_sku': has_yandex_sku,
            'has_quality': has_quality,
            'confidence': calculate_confidence('yandex', yandex_exact_matches, has_yandex_sku or has_quality, row_num)
        }
    
    # Для строки 2 (Ozon, Яндекс или Все инструменты)
    elif row_num == 2:
        # Проверяем первые 5 колонок на совпадение с эталонными заголовками
        first_5_columns = normalized_columns[:min(5, len(normalized_columns))]
        
        # Проверка на Ozon
        ozon_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['ozon_row2']) 
                               if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Проверка на Яндекс.Маркет
        yandex_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['yandex_row2']) 
                                 if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Проверка на Все инструменты
        vseinstrumenty_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['vseinstrumenty_row2']) 
                                        if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Дополнительные проверки
        has_asterisk = any('*' in col for col in normalized_columns)
        has_ozon_price = any('цена, руб.*' in col for col in normalized_columns)
        has_ozon_article = any('артикул*' in col for col in normalized_columns)
        has_yandex_sku = any('ваш sku' in col for col in normalized_columns)
        has_quality = any('качество карточки' in col for col in normalized_columns)
        has_vi_guid = any('guid*' in col.lower() for col in normalized_columns)
        has_data_sheet = False  # Будет установлено true, если лист называется "Данные"
        
        # Общие совпадения (просто наличие характерных заголовков)
        ozon_matches = sum(1 for header in marketplace_headers['ozon_row2'] if any(header.lower() in col for col in normalized_columns))
        yandex_matches = sum(1 for header in marketplace_headers['yandex_row2'] if any(header.lower() in col for col in normalized_columns))
        vseinstrumenty_matches = sum(1 for header in marketplace_headers['vseinstrumenty_row2'] if any(header.lower() in col for col in normalized_columns))
        
        # Результаты для строки 2
        results['ozon'] = {
            'exact_matches': ozon_exact_matches,
            'total_matches': ozon_matches,
            'has_asterisk': has_asterisk,
            'has_ozon_price': has_ozon_price,
            'has_ozon_article': has_ozon_article,
            'confidence': calculate_confidence('ozon', ozon_exact_matches, has_ozon_price or has_ozon_article, row_num)
        }
        
        results['yandex'] = {
            'exact_matches': yandex_exact_matches,
            'total_matches': yandex_matches,
            'has_sku': has_yandex_sku,
            'has_quality': has_quality,
            'confidence': calculate_confidence('yandex', yandex_exact_matches, has_yandex_sku or has_quality, row_num)
        }
        
        results['vseinstrumenty'] = {
            'exact_matches': vseinstrumenty_exact_matches,
            'total_matches': vseinstrumenty_matches,
            'has_guid': has_vi_guid,
            'is_data_sheet': has_data_sheet,
            'confidence': calculate_confidence('vseinstrumenty', vseinstrumenty_exact_matches, has_vi_guid, row_num)
        }
    
    # Для строки 3 (Wildberries)
    elif row_num == 3:
        # Проверяем первые 5 колонок на совпадение с эталонными заголовками
        first_5_columns = normalized_columns[:min(5, len(normalized_columns))]
        
        # Проверка на Wildberries
        wb_exact_matches = sum(1 for i, header in enumerate(marketplace_headers['wildberries_row3']) 
                             if i < len(first_5_columns) and header.lower() in first_5_columns[i])
        
        # Дополнительные проверки
        has_wb_article = any('артикул wb' in col for col in normalized_columns)
        has_seller_article = any('артикул продавца' in col for col in normalized_columns)
        
        # Общие совпадения (просто наличие характерных заголовков)
        wb_matches = sum(1 for header in marketplace_headers['wildberries_row3'] if any(header.lower() in col for col in normalized_columns))
        
        # Результаты для строки 3
        results['wildberries'] = {
            'exact_matches': wb_exact_matches,
            'total_matches': wb_matches,
            'has_wb_article': has_wb_article,
            'has_seller_article': has_seller_article,
            'confidence': calculate_confidence('wildberries', wb_exact_matches, has_wb_article or has_seller_article, row_num)
        }
    
    # Определяем лучший результат
    best_marketplace = None
    best_confidence = 0
    
    for marketplace, data in results.items():
        if data['confidence'] > best_confidence:
            best_confidence = data['confidence']
            best_marketplace = marketplace
    
    # Если ничего не нашли или слишком низкая уверенность
    if best_marketplace is None or best_confidence < 50:
        return ('other', 0, results)
    
    return (best_marketplace, best_confidence, results)

def calculate_confidence(marketplace, exact_matches, has_key_field, row_num):
    """
    Рассчитывает уровень уверенности в определении маркетплейса.
    
    Args:
        marketplace: Название маркетплейса
        exact_matches: Количество точных совпадений заголовков
        has_key_field: Есть ли ключевое поле для этого маркетплейса
        row_num: Номер строки
        
    Returns:
        float: Уровень уверенности от 0 до 100
    """
    # Базовая оценка на основе точных совпадений
    base_confidence = 0
    
    # Проверка соответствия "маркетплейс + строка"
    row_match = False
    if (marketplace == 'ozon' and row_num == 2) or \
       (marketplace == 'wildberries' and row_num == 3) or \
       (marketplace == 'lemanpro' and row_num == 4) or \
       (marketplace == 'vseinstrumenty' and row_num == 2) or \
       (marketplace == 'yandex' and (row_num == 2 or row_num == 4)):
        row_match = True
        base_confidence += 20  # Бонус за совпадение типичной строки для этого маркетплейса
    
    # Оценка на основе точных совпадений (до 60 баллов)
    if exact_matches == 5:
        base_confidence += 60
    elif exact_matches == 4:
        base_confidence += 50
    elif exact_matches == 3:
        base_confidence += 40
    elif exact_matches == 2:
        base_confidence += 30
    elif exact_matches == 1:
        base_confidence += 20
    
    # Дополнительные баллы за наличие ключевого поля
    if has_key_field:
        base_confidence += 20
    
    # Максимум 100 баллов
    return min(100, base_confidence)

# Для тестирования
if __name__ == "__main__":
    print("Тестирование модуля распознавания маркетплейсов")
    
    # Тест 1: Шаблон Яндекс.Маркет
    test_yandex_columns = ['ваш sku *', 'качество карточки', 'рекомендации по заполнению', 
                           'название товара *', 'ссылка на изображение *']
    marketplace, confidence, details = detect_marketplace_by_row_headers(test_yandex_columns, 4)
    print(f"Тест 1 (Яндекс.Маркет): {marketplace}, уверенность: {confidence}%, детали: {details}")
    
    # Тест 2: Шаблон ЛеманПро
    test_lemanpro_columns = ['guid', 'код тн вэд', 'наименование товара мерчанта', 
                            'бренд товара', 'модель товара']
    marketplace, confidence, details = detect_marketplace_by_row_headers(test_lemanpro_columns, 4)
    print(f"Тест 2 (ЛеманПро): {marketplace}, уверенность: {confidence}%, детали: {details}")
    
    # Тест 3: Шаблон Все инструменты
    test_vseinstrumenty_columns = ['guid*', 'бренд', 'наименование', 'артикул', 'код тн вэд']
    marketplace, confidence, details = detect_marketplace_by_row_headers(test_vseinstrumenty_columns, 2)
    print(f"Тест 3 (Все инструменты): {marketplace}, уверенность: {confidence}%, детали: {details}")