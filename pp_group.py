import pandas as pd
import re

def standardize_profession(profession):
    if isinstance(profession, str):
        match = re.search(r'[А-Яа-яЁё]', profession)
        if match:
            first_russian_letter_index = match.start()
            profession = profession[first_russian_letter_index:]
            profession = profession[0].upper() + profession[1:]
    return profession

def is_noun(word):
    # Простая эвристика для определения существительных по окончаниям, включая знак "-"
    match = re.search(r'(тель|ник|тор|лог|ист|ер|арь|чик|щик|ец|от|чий|рка|га|сть|ца|до|на|ель|ок|ук|ра|вт|тр|ом|тик|льт|нт|ёр|рик|ур|ля|дел|ож|яр|чие|ат|ир|лт|ач|ка|мик|рт|мен|пед|лаз)(-|$)', word)
    if match:
        return True
    return False

def get_profession_base(profession):
    if isinstance(profession, str):
        words = profession.split()
        base_words = []
        for word in words:
            # Проверка на окончания перед дефисом
            parts = word.split('-')
            if is_noun(parts[0]):
                word = parts[0]  # Оставляем только часть до дефиса
            else:
                # Проверяем наличие дефиса
                if '-' in word:
                    # Проверяем, заканчивается ли часть после дефиса на одно из окончаний
                    if len(parts) > 1 and is_noun(parts[1]):
                        word = '-'.join(parts[:2])  # Оставляем обе части
                    else:
                        word = parts[0]  # Оставляем только первую часть
                else:
                    # Удаляем часть слова после любого другого знака препинания
                    word = re.split(r'[,:;!?]', word)[0]
            base_words.append(word)
            if is_noun(word):
                break
        base_profession = ' '.join(base_words)
        return base_profession
    return profession

# Загрузка таблицы "Потребность персонала"
file_path = 'Потребность персонала.xlsx'  # Укажите путь к вашему файлу Excel
sheet_name = 'data'  # Укажите название листа в файле
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Применение функции стандартизации к столбцу "Профессия"
df['Профессия'] = df['Профессия'].apply(standardize_profession)

# Получение базовой профессии для группировки
df['Профессия'] = df['Профессия'].apply(get_profession_base)

# Преобразование нечисловых значений в столбцах с числовыми данными
numeric_columns = [
    'Количество штатных единиц (2019)', 'Количество штатных единиц (2020)', 'Количество штатных единиц (2021)',
    'Потребность (2022)', 'Потребность (2023)', 'Потребность (2024)', 'Потребность (2025)',
    'Потребность (2026)', 'Потребность (2027)', 'Потребность (2028)', 'Потребность (2029)',
    'Потребность (2030)', 'Потребность (2031)'
]

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Группировка данных по "Базовая профессия" с суммированием числовых значений
grouped_df = df.groupby('Профессия', as_index=False)[numeric_columns].sum()

# Сохранение результата в новый файл
output_file_path = 'ПП_группировка.xlsx'
grouped_df.to_excel(output_file_path, index=False)

print(f"Стандартизированные и сгруппированные данные сохранены в файл {output_file_path}")