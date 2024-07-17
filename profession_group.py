import pandas as pd

# Загрузим данные из Excel файла
file_path = 'ПП_профессии_стандарт_группировка.xlsx'
df = pd.read_excel(file_path, engine='openpyxl')

# Функция для определения отрасли по профессии
def determine_industry(profession):
    if 'врач' in profession.lower() or 'госпитал' in profession.lower() or 'акушер' in profession.lower() or 'бактери' in profession.lower() or 'псих' in profession.lower() or 'экол' in profession.lower() or 'вет' in profession.lower() or 'санитар' in profession.lower() or 'гигиен' in profession.lower() or 'сестра' in profession.lower() or 'фельдшер' in profession.lower() or 'дезинфек' in profession.lower() or 'зубн' in profession.lower():
        return 'Здравоохранение'
    elif 'учитель' in profession.lower() or 'профессор' in profession.lower() or 'академ' in profession.lower() or 'методист' in profession.lower() or 'вожат' in profession.lower() or 'воспит' in profession.lower() or 'дефектолог' in profession.lower() or 'доцент' in profession.lower():
        return 'Образование'
    elif 'а/слесарь' in profession.lower() or 'автомехан' in profession.lower() or 'автослесарь' in profession.lower() or 'автоэлектр' in profession.lower() or 'авиационный' in profession.lower() or 'аэродром' in profession.lower() or 'води' in profession.lower() or 'вызыв' in profession.lower() or 'детейл' in profession.lower() or 'веб' in profession.lower():
        return 'Транспорт'
    elif 'программист' in profession.lower() or 'системный администратор' in profession.lower() or 'веб' in profession.lower():
        return 'IT и связь'
    elif 'аврхив' in profession.lower() or 'архив' in profession.lower() or 'аудит' in profession.lower() or 'библиограф' in profession.lower() or 'адвокат' in profession.lower() or 'аналитик' in profession.lower() or 'библиоте' in profession.lower() or 'документ' in profession.lower() or 'юрис' in profession.lower() or 'делопроизвод' in profession.lower():
        return 'Информационные и деловые услуги'
    elif 'агент' in profession.lower() or 'бизнес' in profession.lower() or 'менеджер' in profession.lower() or 'маркет' in profession.lower() or 'бухгалтер' in profession.lower() or 'эконом' in profession.lower() or 'консультант' in profession.lower() or 'товар' in profession.lower() or 'декларант' in profession.lower() or 'демонстр' in profession.lower():
        return 'Экономика и торговля'
    elif 'агро' in profession.lower() or 'скот' in profession.lower() or 'охот' in profession.lower() or 'вес' in profession.lower() or 'зоотехник' in profession.lower() or 'глава кфх' in profession.lower() or 'глава кх' in profession.lower() or 'лесничий' in profession.lower() or 'мелиоратор' in profession.lower() or 'рыбовод' in profession.lower() or 'гранул' in profession.lower() or 'дезодаратор' in profession.lower() or 'дояр' in profession.lower() or 'животновод' in profession.lower() or 'земледел' in profession.lower() or 'зерносуш' in profession.lower() or 'зоотехник' in profession.lower():
        return 'Сельское хозяйство'
    elif 'администр' in profession.lower() or 'бригадир' in profession.lower() or 'ассист' in profession.lower() or 'инспект' in profession.lower() or 'директ' in profession.lower() or 'секрет' in profession.lower() or 'глава' in profession.lower() or 'ревизор' in profession.lower() or 'представитель' in profession.lower() or 'совет' in profession.lower() or 'дежур' in profession.lower() or 'диспетч' in profession.lower() or 'зав' in profession.lower() or 'зам' in profession.lower() or 'звед' in profession.lower():
        return 'Администрирование и управление'
    elif 'аккомп' in profession.lower() or 'артист' in profession.lower() or 'архео' in profession.lower() or 'балет' in profession.lower() or 'бутафор' in profession.lower() or 'оператор' in profession.lower() or 'программы' in profession.lower() or 'худож' in profession.lower() or 'дирижер' in profession.lower() or 'редакт' in profession.lower() or 'режисс' in profession.lower() or 'хормейстер' in profession.lower() or 'дизайнер' in profession.lower() or 'грим' in profession.lower() or 'декор' in profession.lower() or 'дизайн' in profession.lower():
        return 'Искусство и культура'
    elif 'арматур' in profession.lower() or 'архитект' in profession.lower() or 'бетонщик' in profession.lower() or 'асфальт' in profession.lower() or 'автокран' in profession.lower() or 'строит' in profession.lower() or 'грунт' in profession.lower() or 'дорожный' in profession.lower():
        return 'Строительство'
    elif 'агломер' in profession.lower() or 'барильет' in profession.lower() or 'бункер' in profession.lower() or 'автоклав' in profession.lower() or 'автоматчик' in profession.lower() or 'аккум' in profession.lower() or 'аппарат' in profession.lower() or 'брошюр' in profession.lower() or 'аэрограф' in profession.lower() or 'вальц' in profession.lower() or 'варщ' in profession.lower() or 'инженер' in profession.lower() or 'механ' in profession.lower() or 'техн' in profession.lower() or 'энерг' in profession.lower() or 'волоч' in profession.lower() or 'вулкан' in profession.lower() or 'выбив' in profession.lower() or 'вышив' in profession.lower() or 'газ' in profession.lower() or 'гальван' in profession.lower() or 'гибщ' in profession.lower() or 'диспетчер' in profession.lower() or 'конструкт' in profession.lower() or 'металлург' in profession.lower() or 'сварщ' in profession.lower() or 'плавил' in profession.lower() or 'теплотехник' in profession.lower() or 'технолог' in profession.lower() or 'горнов' in profession.lower() or 'грохот' in profession.lower() or 'гумм' in profession.lower() or 'дверев' in profession.lower() or 'дефектоскоп' in profession.lower() or 'доводчик' in profession.lower() or 'долбеж' in profession.lower() or 'жестянщ' in profession.lower() or 'заварщ' in profession.lower() or 'загруз' in profession.lower() or 'заквас' in profession.lower() or 'закройщ' in profession.lower() or 'заливщ' in profession.lower() or 'засольщ' in profession.lower() or 'засыпщ' in profession.lower() or 'заточ' in profession.lower() or 'зашивальщ' in profession.lower() or 'зуборезч' in profession.lower():
        return 'Обрабатывающая промышленность'
    elif 'кассир' in profession.lower() or 'бармен' in profession.lower() or 'вах' in profession.lower() or 'дискотек' in profession.lower() or 'водопров' in profession.lower() or 'водоразд' in profession.lower() or 'газов' in profession.lower() or 'гардероб' in profession.lower() or 'гладильщик' in profession.lower() or 'горничн' in profession.lower() or 'гравер' in profession.lower() or 'груз' in profession.lower() or 'двор' in profession.lower() or 'сторож' in profession.lower() or 'домработ' in profession.lower():
        return 'Бытовое обслуживание и услуги'
    elif 'биолог' in profession.lower() or 'хим' in profession.lower() or 'лаборант' in profession.lower() or 'зоолог' in profession.lower():
        return 'Научная деятельность'
    elif 'буфет' in profession.lower() or 'глазировщик' in profession.lower():
        return 'Общественное питание'
    elif 'бурил' in profession.lower() or 'вальщ' in profession.lower() or 'взрыв' in profession.lower() or 'гасил' in profession.lower() or 'гео' in profession.lower() or 'гидротехник' in profession.lower() or 'гидроциклонщик' in profession.lower() or 'маркшейдер' in profession.lower() or 'горнорабочий' in profession.lower() or 'горный мастер' in profession.lower() or 'дози' in profession.lower() or 'дробил' in profession.lower() or 'замерщ' in profession.lower():
        return 'Добывающая промышленность'
    else:
        return 'Другое'

# Применим функцию к каждому значению в столбце 'Профессия'
df['Отрасль'] = df['Профессия'].apply(determine_industry)

# Определим столбцы с числовыми данными
numeric_columns = df.select_dtypes(include=[int, float]).columns.tolist()

# Группируем данные по отраслям и суммируем числовые значения
grouped_df = df.groupby('Отрасль')[numeric_columns].sum().reset_index()

# Добавляем колонку с профессиями
grouped_df['Профессии'] = df.groupby('Отрасль')['Профессия'].apply(list).reset_index(drop=True)

# Выводим результат
professions = df['Профессия'].tolist()
industries = df['Отрасль'].tolist()

result = {
    'Профессия': professions,
    'Отрасль': industries
}

# Сохранение результата в новый файл
output_file_path = 'ПП_отрасли_группировка.xlsx'
grouped_df.to_excel(output_file_path, index=False)

print(f"Стандартизированные и сгруппированные данные сохранены в файл {output_file_path}")