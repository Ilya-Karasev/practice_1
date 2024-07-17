import pandas as pd
import numpy as np

# Загрузка данных из Excel файла
file_path = 'Группированные_данные_по_отраслям.xlsx'
professions_df = pd.read_excel(file_path, sheet_name='Профессии')

# Загрузка данных из файла ПП_группировка
pp_professions_file_path = 'ПП_группировка.xlsx'
pp_professions_df = pd.read_excel(pp_professions_file_path)

# Функция для преобразования строковых чисел в числовой формат
def convert_to_numeric(series):
    def safe_convert(x):
        if isinstance(x, str):
            try:
                return pd.to_numeric(x.replace(';', ',').split(',')).mean()
            except ValueError:
                return np.nan
        return x
    return series.apply(safe_convert)



# Агрегация данных по профессиям
professions_agg = professions_df.groupby('Отрасль').agg({
    'Количество штатных единиц (2019)': 'sum',
    'Количество штатных единиц (2020)': 'sum',
    'Количество штатных единиц (2021)': 'sum',
    'Потребность (2022)': 'mean',
    'Потребность (2023)': 'mean',
    'Потребность (2024)': 'mean',
    'Потребность (2025)': 'mean',
    'Потребность (2026)': 'mean',
    'Потребность (2027)': 'mean',
    'Потребность (2028)': 'mean',
    'Потребность (2029)': 'mean',
    'Потребность (2030)': 'mean',
    'Потребность (2031)': 'mean',
}).reset_index()

# Агрегация данных по профессиям из файла ПП_профессии_стандарт_группировка
pp_professions_agg = pp_professions_df.groupby('Профессия').agg({
    'Количество штатных единиц (2019)': 'sum',
    'Количество штатных единиц (2020)': 'sum',
    'Количество штатных единиц (2021)': 'sum',
    'Потребность (2022)': 'mean',
    'Потребность (2023)': 'mean',
    'Потребность (2024)': 'mean',
    'Потребность (2025)': 'mean',
    'Потребность (2026)': 'mean',
    'Потребность (2027)': 'mean',
    'Потребность (2028)': 'mean',
    'Потребность (2029)': 'mean',
    'Потребность (2030)': 'mean',
    'Потребность (2031)': 'mean',
}).reset_index()

# Определение востребованных и невостребованных профессий
def determine_demand(prof_agg):
    demand_info = []
    for index, row in prof_agg.iterrows():
        отрасль = row['Отрасль']
        потребность = row.filter(like='Потребность').mean()

        # Логика для определения востребованности на основе других параметров
        if потребность > threshold:  # threshold - пороговое значение для определения востребованности
            demand_info.append((отрасль, 'Востребовано', потребность))
        else:
            demand_info.append((отрасль, 'Не востребовано', потребность))

    demand_df = pd.DataFrame(demand_info, columns=['Отрасль', 'Статус востребованности', 'Средняя потребность'])
    return demand_df

# Определение востребованных и невостребованных профессий для файла ПП_профессии_стандарт_группировка
def determine_profession_demand(pp_prof_agg):
    demand_info = []
    for index, row in pp_prof_agg.iterrows():
        профессия = row['Профессия']
        потребность = row.filter(like='Потребность').mean()

        # Логика для определения востребованности на основе других параметров
        if потребность > threshold2:  # threshold - пороговое значение для определения востребованности
            demand_info.append((профессия, 'Востребовано', потребность))
        else:
            demand_info.append((профессия, 'Не востребовано', потребность))

    demand_df = pd.DataFrame(demand_info, columns=['Профессия', 'Статус востребованности', 'Средняя потребность'])
    return demand_df

threshold = 100  # Пример порогового значения для определения востребованности
threshold2 = 1
demand_df = determine_demand(professions_agg)
profession_demand_df = determine_profession_demand(pp_professions_agg)

# Сохранение результатов в новый Excel файл
with pd.ExcelWriter('ПП_анализ.xlsx') as writer:
    demand_df.to_excel(writer, sheet_name='Востребованность Отраслей', index=False)
    profession_demand_df.to_excel(writer, sheet_name='Востребованность Профессий', index=False)