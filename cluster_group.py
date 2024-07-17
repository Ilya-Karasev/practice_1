import pandas as pd

# Функция для определения отрасли по профессии
def determine_industry(profession):
    profession = profession.lower()
    if any(keyword in profession for keyword in ['врач', 'госпитал', 'акушер', 'бактери', 'псих', 'экол', 'вет', 'санитар', 'гигиен', 'сестра', 'фельдшер', 'дезинфек', 'зубн']):
        return 'Здравоохранение'
    elif any(keyword in profession for keyword in ['учитель', 'профессор', 'академ', 'методист', 'вожат', 'воспит', 'дефектолог', 'доцент', 'инструктор']):
        return 'Образование'
    elif any(keyword in profession for keyword in ['а/слесарь', 'автомехан', 'автослесарь', 'автоэлектр', 'авиационный', 'аэродром', 'води', 'вызыв', 'детейл', 'машинист', 'кондуктор']):
        return 'Транспорт'
    elif any(keyword in profession for keyword in ['программист', 'системный администратор', 'веб', 'кабель', 'контент']):
        return 'IT и связь'
    elif any(keyword in profession for keyword in ['аврхив', 'архив', 'аудит', 'библиограф', 'адвокат', 'аналитик', 'библиоте', 'документ', 'юрис', 'делопроизвод', 'код']):
        return 'Информационные и деловые услуги'
    elif any(keyword in profession for keyword in ['агент', 'бизнес', 'менеджер', 'маркет', 'бухгалтер', 'эконом', 'консультант', 'товар', 'декларант', 'демонстр', 'калькулятор', 'кладовщ', 'комплект', 'курьер']):
        return 'Экономика и торговля'
    elif any(keyword in profession for keyword in ['агро', 'скот', 'охот', 'вес', 'зоотехник', 'глава кфх', 'глава кх', 'лесничий', 'мелиоратор', 'рыбовод', 'животновод', 'гранул', 'дезодаратор', 'дояр', 'земледел', 'зерносуш', 'зоотехник', 'конюх', 'косар']):
        return 'Сельское хозяйство'
    elif any(keyword in profession for keyword in ['администр', 'бригадир', 'ассист', 'инспект', 'директ', 'секрет', 'глава', 'ревизор', 'представитель', 'совет', 'дежур', 'диспетч', 'зав', 'зам', 'звед', 'комендант', 'управ', 'контрол']):
        return 'Администрирование и управление'
    elif any(keyword in profession for keyword in ['аккомп', 'артист', 'архео', 'балет', 'бутафор', 'оператор', 'программы', 'худож', 'дирижер', 'редакт', 'режисс', 'хормейстер', 'дизайнер', 'грим', 'декор', 'дизайн', 'инкруст', 'корреспондент', 'исполнитель', 'кино', 'колор', 'концерт', 'корректор', 'костюм', 'кружев', 'культ']):
        return 'Искусство и культура'
    elif any(keyword in profession for keyword in ['арматур', 'архитект', 'бетонщик', 'асфальт', 'автокран', 'строит', 'грунт', 'дорожный', 'изолировщ', 'кровель']):
        return 'Строительство'
    elif any(keyword in profession for keyword in ['агломер', 'барильет', 'бункер', 'автоклав', 'автоматчик', 'аккум', 'аппарат', 'брошюр', 'аэрограф', 'вальц', 'варщ', 'инженер', 'механ', 'техн', 'энерг', 'волоч', 'вулкан', 'выбив', 'вышив', 'газ', 'гальван', 'гибщ', 'диспетчер', 'конструкт', 'металлург', 'сварщ', 'плавил', 'теплотехник', 'технолог', 'горнов', 'грохот', 'гумм', 'дверев', 'дефектоскоп', 'доводчик', 'долбеж', 'жестянщ', 'заварщ', 'загруз', 'заквас', 'закройщ', 'заливщ', 'засольщ', 'засыпщ', 'заточ', 'зашивальщ', 'зуборезч', 'изготов', 'испытат', 'истоп', 'кабинщ', 'картонаж', 'клейщ', 'клейм', 'клепальщ', 'ковш', 'кондитер', 'конопатч', 'котельщ', 'котл', 'кочегар', 'кристаллизаторщ', 'кузнец', 'купаж']):
        return 'Обрабатывающая промышленность'
    elif any(keyword in profession for keyword in ['кассир', 'бармен', 'вах', 'дискотек', 'водопров', 'водоразд', 'газов', 'гардероб', 'гладильщик', 'горничн', 'гравер', 'груз', 'двор', 'сторож', 'домработ', 'кастелянша', 'курьер', 'кух']):
        return 'Бытовое обслуживание и услуги'
    elif any(keyword in profession for keyword in ['биолог', 'хим', 'лаборант', 'зоолог']):
        return 'Научная деятельность'
    # elif any(keyword in profession for keyword in ['буфет', 'глазировщик']):
    #     return 'Общественное питание'
    elif any(keyword in profession for keyword in ['бурил', 'вальщ', 'взрыв', 'гасил', 'гео', 'гидротехник', 'гидроциклонщик', 'маркшейдер', 'горнорабочий', 'горный мастер', 'дози', 'дробил', 'замерщ', 'каменщ']):
        return 'Добывающая промышленность'
    else:
        return 'Другое'

# Загрузка данных из файлов
professions_file_path = 'ПП_профессии_стандарт_группировка.xlsx'
dolzhnosti_file_path = 'ЦЗН_сгруппированные_должности.xlsx'
specialnosti_file_path = 'ЦЗН_сгруппированные_специальности.xlsx'

professions_df = pd.read_excel(professions_file_path, engine='openpyxl')
dolzhnosti_df = pd.read_excel(dolzhnosti_file_path, engine='openpyxl')
specialnosti_df = pd.read_excel(specialnosti_file_path, engine='openpyxl')

# Применение функции определения отрасли
professions_df['Отрасль'] = professions_df['Профессия'].apply(determine_industry)
dolzhnosti_df['Отрасль'] = dolzhnosti_df['Базовая должность'].apply(determine_industry)
specialnosti_df['Отрасль'] = specialnosti_df['Базовая специальность'].apply(determine_industry)

# Определение числовых столбцов для профессий
numeric_columns_professions = professions_df.select_dtypes(include=[int, float]).columns.tolist()

# Группировка данных по отраслям для профессий
grouped_professions_df = professions_df.groupby('Отрасль')[numeric_columns_professions].sum().reset_index()
grouped_professions_df['Профессии'] = professions_df.groupby('Отрасль')['Профессия'].apply(lambda x: ', '.join(x)).reset_index(drop=True)

# Группировка данных по отраслям для должностей
grouped_dolzhnosti_df = dolzhnosti_df.groupby('Отрасль').agg({
    'Базовая должность': lambda x: '; '.join(x),
    'Пол': lambda x: '; '.join(map(str, x)),
    'Возраст на момент обращения': lambda x: '; '.join(map(str, x)),
    'Основание незанятости': lambda x: '; '.join(x),
    'Основание увольнения': lambda x: '; '.join(map(str, x)),  # Преобразование в строки
    'Средний заработок': lambda x: '; '.join(map(str, x)),  # Преобразование в строки
    'Год увольнения': lambda x: '; '.join(map(str, x)),
    'Стаж на последней работе': lambda x: '; '.join(map(str, x)),
    'Образование': lambda x: '; '.join(x)
}).reset_index()

# Группировка данных по отраслям для специальностей
grouped_specialnosti_df = specialnosti_df.groupby('Отрасль').agg({
    'Базовая специальность': lambda x: '; '.join(x),
    'Пол': lambda x: '; '.join(map(str, x)),
    'Возраст на момент обращения': lambda x: '; '.join(map(str, x)),
    'Основание незанятости': lambda x: '; '.join(x),
    'Основание увольнения': lambda x: '; '.join(map(str, x)),  # Преобразование в строки
    'Средний заработок': lambda x: '; '.join(map(str, x)),  # Преобразование в строки
    'Год увольнения': lambda x: '; '.join(map(str, x)),
    'Стаж на последней работе': lambda x: '; '.join(map(str, x)),
    'Образование': lambda x: '; '.join(x)
}).reset_index()

# Сохранение результата в один файл Excel с разными листами
output_file_path = 'Группированные_данные_по_отраслям.xlsx'
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    grouped_professions_df.to_excel(writer, sheet_name='Профессии', index=False)
    grouped_dolzhnosti_df.to_excel(writer, sheet_name='Должности', index=False)
    grouped_specialnosti_df.to_excel(writer, sheet_name='Специальности', index=False)