import pandas as pd

# Загрузка данных из xlsx-файла
file_path = 'Группированные_данные_по_отраслям.xlsx'
professions_df = pd.read_excel(file_path, sheet_name='Профессии')
positions_df = pd.read_excel(file_path, sheet_name='Должности')
specialties_df = pd.read_excel(file_path, sheet_name='Специальности')

def split_and_convert(column, convert_to_numeric=False):
    """Разделяет значения в строках по '; ' и при необходимости конвертирует в числовые значения."""
    extracted_values = []
    for item in column:
        if isinstance(item, str):
            parts = item.split('; ')
            if convert_to_numeric:
                parts = [pd.to_numeric(part, errors='coerce') for part in parts]
                parts = [part for part in parts if pd.notnull(part)]
            extracted_values.extend(parts)
        else:
            if pd.notnull(item):
                extracted_values.append(item)
    return extracted_values

def aggregate_by_column(dataframe, group_by_col, agg_col, convert_to_numeric=False, count_values=False):
    # Разделение строк и извлечение значений
    extracted_data = []
    for index, row in dataframe.iterrows():
        group_value = row[group_by_col]
        agg_values = row[agg_col]
        split_values = split_and_convert([agg_values], convert_to_numeric)
        for value in split_values:
            extracted_data.append((group_value, value))

    exploded_df = pd.DataFrame(extracted_data, columns=[group_by_col, agg_col])
    if count_values:
        # Группировка по отрасли и агрегация количества записей
        grouped_df = exploded_df.groupby([group_by_col, agg_col]).size().reset_index(name='count')
    else:
        # Группировка по отрасли и агрегация в процентном соотношении
        grouped_df = exploded_df.groupby(group_by_col)[agg_col].value_counts(normalize=True).mul(100).rename(
            'percentage').reset_index()
    return grouped_df

def aggregate_by_age(dataframe, group_by_col, agg_col):
    extracted_data = []
    for index, row in dataframe.iterrows():
        group_value = row[group_by_col]
        agg_values = row[agg_col]
        split_values = split_and_convert([agg_values], convert_to_numeric=True)
        for value in split_values:
            if pd.notnull(value):
                extracted_data.append((group_value, value))

    exploded_df = pd.DataFrame(extracted_data, columns=[group_by_col, agg_col])
    bins = [0, 18, 20, 30, 40, 50, 60, 100]
    labels = ['<18', '18-20', '20-30', '30-40', '40-50', '50-60', '>60']
    exploded_df['age_group'] = pd.cut(exploded_df[agg_col].astype(int), bins=bins, labels=labels, right=False)

    # Добавление диапазонов значений
    def age_range(interval):
        if isinstance(interval, pd.Interval):
            return f'{interval.left}-{interval.right}'
        return interval

    exploded_df['age_range'] = exploded_df['age_group'].apply(age_range)

    grouped_df = exploded_df.groupby([group_by_col, 'age_group', 'age_range'], observed=True)[agg_col].count().groupby(
        level=0).apply(lambda x: 100 * x / x.sum()).rename('percentage').reset_index(level=[1, 2])
    return grouped_df

def aggregate_by_salary(dataframe, group_by_col, agg_col):
    extracted_data = []
    for index, row in dataframe.iterrows():
        group_value = row[group_by_col]
        agg_values = row[agg_col]
        split_values = split_and_convert([agg_values], convert_to_numeric=True)
        for value in split_values:
            if pd.notnull(value):
                extracted_data.append((group_value, value))

    exploded_df = pd.DataFrame(extracted_data, columns=[group_by_col, agg_col])
    bins = [0, 10000, 20000, 30000, 40000, 50000, 100000, 200000, 300000, 400000, 500000, 1000000]
    labels = ['<10k', '10-20k', '20-30k', '30-40k', '40-50k', '50-100k', '100-200k', '200-300k', '300-400k', '400-500k',
              '>500k']
    exploded_df['salary_group'] = pd.cut(exploded_df[agg_col].astype(float), bins=bins, labels=labels, right=False)

    # Добавление диапазонов значений
    def salary_range(interval):
        if isinstance(interval, pd.Interval):
            return f'{interval.left}-{interval.right}'
        return interval

    exploded_df['salary_range'] = exploded_df['salary_group'].apply(salary_range)

    grouped_df = exploded_df.groupby([group_by_col, 'salary_group', 'salary_range'], observed=True)[
        agg_col].count().groupby(
        level=0).apply(lambda x: 100 * x / x.sum()).rename('percentage').reset_index(level=[1, 2])
    return grouped_df

def aggregate_by_work_experience(dataframe, group_by_col, agg_col):
    extracted_data = []
    for index, row in dataframe.iterrows():
        group_value = row[group_by_col]
        agg_values = row[agg_col]
        split_values = split_and_convert([agg_values], convert_to_numeric=True)
        for value in split_values:
            if pd.notnull(value):
                extracted_data.append((group_value, value))

    exploded_df = pd.DataFrame(extracted_data, columns=[group_by_col, agg_col])
    bins = [0, 1, 10, 20, 30, 40, 50, 100, 200, 300, 400, 500, 1000, 2000, 3000, 4000, 5000, 10000, 20000]
    labels = ['<1', '1-10', '10-20', '20-30', '30-40', '40-50', '50-100', '100-200', '200-300', '300-400', '400-500', '500-1000', '1000-2000', '2000-3000', '3000-4000', '4000-5000', '5000-10000', '>10000']
    exploded_df['work_experience_group'] = pd.cut(exploded_df[agg_col].astype(float), bins=bins, labels=labels,
                                                  right=False)

    # Добавление диапазонов значений
    def work_experience_range(interval):
        if isinstance(interval, pd.Interval):
            return f'{interval.left}-{interval.right}'
        return interval

    exploded_df['work_experience_range'] = exploded_df['work_experience_group'].apply(work_experience_range)

    grouped_df = exploded_df.groupby([group_by_col, 'work_experience_group', 'work_experience_range'], observed=True).size().groupby(
        level=0).apply(lambda x: 100 * x / x.sum()).rename('percentage').reset_index(level=[1, 2])
    return grouped_df

# Аггрегации для "Должности"
positions_by_gender = aggregate_by_column(positions_df, 'Отрасль', 'Пол')
positions_by_age = aggregate_by_age(positions_df, 'Отрасль', 'Возраст на момент обращения')
positions_by_unemployment_reason = aggregate_by_column(positions_df, 'Отрасль', 'Основание незанятости', count_values=True)
positions_by_termination_reason = aggregate_by_column(positions_df, 'Отрасль', 'Основание увольнения', count_values=True)
positions_by_salary = aggregate_by_salary(positions_df, 'Отрасль', 'Средний заработок')
positions_by_termination_year = aggregate_by_column(positions_df, 'Отрасль', 'Год увольнения', convert_to_numeric=True, count_values=True)
positions_by_work_experience = aggregate_by_work_experience(positions_df, 'Отрасль', 'Стаж на последней работе')
positions_by_education = aggregate_by_column(positions_df, 'Отрасль', 'Образование', count_values=True)

# Аггрегации для "Специальности"
specialties_by_gender = aggregate_by_column(specialties_df, 'Отрасль', 'Пол')
specialties_by_age = aggregate_by_age(specialties_df, 'Отрасль', 'Возраст на момент обращения')
specialties_by_unemployment_reason = aggregate_by_column(specialties_df, 'Отрасль', 'Основание незанятости', count_values=True)
specialties_by_termination_reason = aggregate_by_column(specialties_df, 'Отрасль', 'Основание увольнения', count_values=True)
specialties_by_salary = aggregate_by_salary(specialties_df, 'Отрасль', 'Средний заработок')
specialties_by_termination_year = aggregate_by_column(specialties_df, 'Отрасль', 'Год увольнения', convert_to_numeric=True, count_values=True)
specialties_by_work_experience = aggregate_by_work_experience(specialties_df, 'Отрасль', 'Стаж на последней работе')
specialties_by_education = aggregate_by_column(specialties_df, 'Отрасль', 'Образование', count_values=True)

# Сохранение результатов в xlsx-файл
output_file_path = 'ЦЗН_анализ.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    positions_by_gender.to_excel(writer, sheet_name='Positions_by_Gender', index=False)
    positions_by_age.to_excel(writer, sheet_name='Positions_by_Age', index=False)
    positions_by_unemployment_reason.to_excel(writer, sheet_name='Positions_by_Unemployment', index=False)
    positions_by_termination_reason.to_excel(writer, sheet_name='Positions_by_Termination', index=False)
    positions_by_salary.to_excel(writer, sheet_name='Positions_by_Salary', index=False)
    positions_by_termination_year.to_excel(writer, sheet_name='Positions_by_Termination_Year', index=False)
    positions_by_work_experience.to_excel(writer, sheet_name='Positions_by_Work_Experience', index=False)
    positions_by_education.to_excel(writer, sheet_name='Positions_by_Education', index=False)

    specialties_by_gender.to_excel(writer, sheet_name='Specialties_by_Gender', index=False)
    specialties_by_age.to_excel(writer, sheet_name='Specialties_by_Age', index=False)
    specialties_by_unemployment_reason.to_excel(writer, sheet_name='Specialties_by_Unemployment', index=False)
    specialties_by_termination_reason.to_excel(writer, sheet_name='Specialties_by_Termination', index=False)
    specialties_by_salary.to_excel(writer, sheet_name='Specialties_by_Salary', index=False)
    specialties_by_termination_year.to_excel(writer, sheet_name='Specialties_by_Termination_Year', index=False)
    specialties_by_work_experience.to_excel(writer, sheet_name='Specialties_by_Work_Experience', index=False)
    specialties_by_education.to_excel(writer, sheet_name='Specialties_by_Education', index=False)

print("Результаты сохранены в файл:", output_file_path)