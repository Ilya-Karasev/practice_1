import pandas as pd
import re
from keras._tf_keras.keras.preprocessing.text import Tokenizer
from keras._tf_keras.keras.preprocessing.sequence import pad_sequences
from keras._tf_keras.keras.models import Sequential
from keras._tf_keras.keras.layers import Embedding, LSTM, Dense, SpatialDropout1D
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
import numpy as np

# Функция для стандартизации профессии
def standardize_profession(profession):
    if isinstance(profession, str):
        match = re.search(r'[А-Яа-яЁё]', profession)
        if match:
            first_russian_letter_index = match.start()
            profession = profession[first_russian_letter_index:]
            profession = profession[0].upper() + profession[1:]
    return profession

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

# Загрузка таблицы "Потребность персонала"
file_path = 'ПП_группировка.xlsx'  # Укажите путь к вашему файлу Excel
df = pd.read_excel(file_path)

# Применение функции стандартизации к столбцу "Профессия"
df['Профессия'] = df['Профессия'].apply(standardize_profession)

# Преобразование нечисловых значений в столбцах с числовыми данными
numeric_columns = [
    'Количество штатных единиц (2019)', 'Количество штатных единиц (2020)', 'Количество штатных единиц (2021)',
    'Потребность (2022)', 'Потребность (2023)', 'Потребность (2024)', 'Потребность (2025)',
    'Потребность (2026)', 'Потребность (2027)', 'Потребность (2028)', 'Потребность (2029)',
    'Потребность (2030)', 'Потребность (2031)'
]

for col in numeric_columns:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

# Определение отрасли для каждой профессии в основном DataFrame
df['Отрасль'] = df['Профессия'].apply(determine_industry)

# Подготовка данных для обучения
training_data = df[['Профессия', 'Отрасль']]

# Добавление "Другие" для непризнанных профессий
additional_data = pd.DataFrame({'Профессия': ['Другие'], 'Отрасль': ['Другое']})
training_data = pd.concat([training_data, additional_data], ignore_index=True)

# Подготовка данных для обучения модели
tokenizer = Tokenizer()
tokenizer.fit_on_texts(training_data['Профессия'])

X = tokenizer.texts_to_sequences(training_data['Профессия'])
X = pad_sequences(X, maxlen=10)
le = LabelEncoder()
y = le.fit_transform(training_data['Отрасль'])

# Разделение данных на обучающую и тестовую выборки
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Определение модели
model = Sequential()
model.add(Embedding(len(tokenizer.word_index) + 1, 100, input_length=10))
model.add(SpatialDropout1D(0.2))
model.add(LSTM(100, dropout=0.2, recurrent_dropout=0.2))
model.add(Dense(len(np.unique(y)), activation='softmax'))

model.compile(loss='sparse_categorical_crossentropy', optimizer='adam', metrics=['accuracy'])

# Обучение модели
history = model.fit(X_train, y_train, epochs=16, batch_size=32, validation_data=(X_test, y_test), verbose=2)

# Преобразование значений в столбце "Профессия" к строкам
df['Профессия'] = df['Профессия'].astype(str)

# Применение модели для классификации профессий в основном DataFrame
df_sequences = tokenizer.texts_to_sequences(df['Профессия'])
df_sequences_padded = pad_sequences(df_sequences, maxlen=10)
predicted_probs = model.predict(df_sequences_padded)
predicted_classes = np.argmax(predicted_probs, axis=1)

# Преобразование предсказанных классов обратно в названия отраслей
predicted_labels = le.inverse_transform(predicted_classes)

# Проверка на непризнанные профессии и отнесение их к "Другое"
recognized_professions = set(training_data['Профессия'])
df['Отрасль'] = [label if profession in recognized_professions else 'Другое'
                 for profession, label in zip(df['Профессия'], predicted_labels)]

# Группировка данных по отраслям и добавление столбца с профессиями
grouped_df = df.groupby('Отрасль', as_index=False).agg(
    {**{col: 'sum' for col in numeric_columns}, 'Профессия': lambda x: '; '.join(x)}
)

# Загрузка данных из файлов для группировки по должностям и специальностям
dolzhnosti_file_path = 'ЦЗН_сгруппированные_должности.xlsx'
specialnosti_file_path = 'ЦЗН_сгруппированные_специальности.xlsx'

dolzhnosti_df = pd.read_excel(dolzhnosti_file_path, engine='openpyxl')
specialnosti_df = pd.read_excel(specialnosti_file_path, engine='openpyxl')

# Применение функции определения отрасли к должностям и специальностям
dolzhnosti_df['Отрасль'] = dolzhnosti_df['Базовая должность'].apply(determine_industry)
specialnosti_df['Отрасль'] = specialnosti_df['Базовая специальность'].apply(determine_industry)

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
    grouped_df.to_excel(writer, index=False, sheet_name='Профессии')
    grouped_dolzhnosti_df.to_excel(writer, index=False, sheet_name='Должности')
    grouped_specialnosti_df.to_excel(writer, index=False, sheet_name='Специальности')

print(f"Результаты анализа сгруппированных данных сохранены в файл {output_file_path}")