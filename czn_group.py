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

def extract_year(date):
    if isinstance(date, str) and '-' in date:
        return int(date.split('-')[0])
    elif isinstance(date, pd.Timestamp):
        return date.year
    return date

# Загрузка данных
file_path = 'ЦЗН.xlsx'  # Укажите путь к вашему файлу Excel
sheet_name = 'data'  # Укажите название листа в файле
df = pd.read_excel(file_path, sheet_name=sheet_name)

# Применение функции стандартизации к столбцам "Должность" и "Специальность"
df['Должность'] = df['Должность'].apply(standardize_profession)
df['Специальность'] = df['Специальность'].apply(standardize_profession)

# Получение базовой должности и специальности для группировки
df['Базовая должность'] = df['Должность'].apply(get_profession_base)
df['Базовая специальность'] = df['Специальность'].apply(get_profession_base)

# Преобразование столбца "Дата увольнение" в год
df['Год увольнения'] = df['Дата увольнения'].apply(extract_year)

# Группировка по базовой должности
grouped_dolzhnost_df = df.groupby('Базовая должность').agg({
    'Пол': lambda x: '; '.join(x),
    'Возраст на момент обращения': lambda x: '; '.join(map(str, x)),
    'Основание незанятости': lambda x: '; '.join(x),
    'Основание увольнения': lambda x: '; '.join(map(str, x)),
    'Средний заработок': lambda x: '; '.join(map(str, x)),
    'Год увольнения': lambda x: '; '.join(map(str, x)),
    'Стаж на последней работе': lambda x: '; '.join(map(str, x)),
    'Образование': lambda x: '; '.join(x)
}).reset_index()

# Сохранение результата для должностей в новый файл
output_file_path_dolzhnosti = 'ЦЗН_сгруппированные_должности.xlsx'
grouped_dolzhnost_df.to_excel(output_file_path_dolzhnosti, index=False)

# Группировка по базовой специальности
grouped_specialnost_df = df.groupby('Базовая специальность').agg({
    'Пол': lambda x: '; '.join(x),
    'Возраст на момент обращения': lambda x: '; '.join(map(str, x)),
    'Основание незанятости': lambda x: '; '.join(x),
    'Основание увольнения': lambda x: '; '.join(map(str, x)),
    'Средний заработок': lambda x: '; '.join(map(str, x)),
    'Год увольнения': lambda x: '; '.join(map(str, x)),
    'Стаж на последней работе': lambda x: '; '.join(map(str, x)),
    'Образование': lambda x: '; '.join(x)
}).reset_index()

# Сохранение результата для специальностей в новый файл
output_file_path_specialnosti = 'ЦЗН_сгруппированные_специальности.xlsx'
grouped_specialnost_df.to_excel(output_file_path_specialnosti, index=False)

print(f"Сгруппированные данные сохранены в файлы: {output_file_path_dolzhnosti} и {output_file_path_specialnosti}")