import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html
from dash.dependencies import Input, Output

# Загрузка данных из файла
file_path = 'ЦЗН_анализ.xlsx'
grouped_data_file_path = 'Группированные_данные_по_отраслям.xlsx'

# Чтение данных из всех листов
positions_by_gender = pd.read_excel(file_path, sheet_name='Positions_by_Gender')
positions_by_age = pd.read_excel(file_path, sheet_name='Positions_by_Age')
positions_by_unemployment = pd.read_excel(file_path, sheet_name='Positions_by_Unemployment')
positions_by_termination = pd.read_excel(file_path, sheet_name='Positions_by_Termination')
positions_by_salary = pd.read_excel(file_path, sheet_name='Positions_by_Salary')
positions_by_termination_year = pd.read_excel(file_path, sheet_name='Positions_by_Termination_Year')
positions_by_work_experience = pd.read_excel(file_path, sheet_name='Positions_by_Work_Experience')
positions_by_education = pd.read_excel(file_path, sheet_name='Positions_by_Education')

specialties_by_gender = pd.read_excel(file_path, sheet_name='Specialties_by_Gender')
specialties_by_age = pd.read_excel(file_path, sheet_name='Specialties_by_Age')
specialties_by_unemployment = pd.read_excel(file_path, sheet_name='Specialties_by_Unemployment')
specialties_by_termination = pd.read_excel(file_path, sheet_name='Specialties_by_Termination')
specialties_by_salary = pd.read_excel(file_path, sheet_name='Specialties_by_Salary')
specialties_by_termination_year = pd.read_excel(file_path, sheet_name='Specialties_by_Termination_Year')
specialties_by_work_experience = pd.read_excel(file_path, sheet_name='Specialties_by_Work_Experience')
specialties_by_education = pd.read_excel(file_path, sheet_name='Specialties_by_Education')

professions_data = pd.read_excel(grouped_data_file_path, sheet_name='Профессии')

# Заполнение NaN значений строковыми представлениями
def fill_na_and_convert_to_str(df, columns):
    for col in columns:
        df[col] = df[col].fillna('Не указано').astype(str)

fill_na_and_convert_to_str(positions_by_unemployment, ['Отрасль', 'Основание незанятости'])
fill_na_and_convert_to_str(positions_by_termination, ['Отрасль', 'Основание увольнения'])
fill_na_and_convert_to_str(positions_by_termination_year, ['Отрасль', 'Год увольнения'])
fill_na_and_convert_to_str(positions_by_education, ['Отрасль', 'Образование'])

fill_na_and_convert_to_str(specialties_by_unemployment, ['Отрасль', 'Основание незанятости'])
fill_na_and_convert_to_str(specialties_by_termination, ['Отрасль', 'Основание увольнения'])
fill_na_and_convert_to_str(specialties_by_termination_year, ['Отрасль', 'Год увольнения'])
fill_na_and_convert_to_str(specialties_by_education, ['Отрасль', 'Образование'])

# Создание приложения Dash
app = Dash(__name__)

app.layout = html.Div([
    html.H1("Отраслевые дэшборды"),
    dcc.Tabs([
        dcc.Tab(label='Должности по полу', children=[
            dcc.Dropdown(
                id='positions-gender-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_gender['Отрасль'].unique()],
                value=positions_by_gender['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-gender-graph')
        ]),
        dcc.Tab(label='Должности по возрасту', children=[
            dcc.Dropdown(
                id='positions-age-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_age['Отрасль'].unique()],
                value=positions_by_age['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-age-graph')
        ]),
        dcc.Tab(label='Должности по основанию незанятости', children=[
            dcc.Dropdown(
                id='positions-unemployment-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_unemployment['Отрасль'].unique()],
                value=positions_by_unemployment['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-unemployment-graph')
        ]),
        dcc.Tab(label='Должности по основанию увольнения', children=[
            dcc.Dropdown(
                id='positions-termination-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_termination['Отрасль'].unique()],
                value=positions_by_termination['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-termination-graph')
        ]),
        dcc.Tab(label='Должности по годам увольнения', children=[
            dcc.Dropdown(
                id='positions-termination-year-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_termination_year['Отрасль'].unique()],
                value=positions_by_termination_year['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-termination-year-graph')
        ]),
        dcc.Tab(label='Должности по образованию', children=[
            dcc.Dropdown(
                id='positions-education-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_education['Отрасль'].unique()],
                value=positions_by_education['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-education-graph')
        ]),
        dcc.Tab(label='Должности по среднему заработку', children=[
            dcc.Dropdown(
                id='positions-salary-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_salary['Отрасль'].unique()],
                value=positions_by_salary['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-salary-graph')
        ]),
        dcc.Tab(label='Должности по стажу', children=[
            dcc.Dropdown(
                id='positions-experience-dropdown',
                options=[{'label': i, 'value': i} for i in positions_by_work_experience['Отрасль'].unique()],
                value=positions_by_work_experience['Отрасль'].unique()[0]
            ),
            dcc.Graph(id='positions-experience-graph')
        ]),
    ])
])

@app.callback(
    Output('positions-gender-graph', 'figure'),
    [Input('positions-gender-dropdown', 'value')]
)
def update_gender_chart(selected_industry):
    filtered_data = positions_by_gender[positions_by_gender['Отрасль'] == selected_industry]
    fig = px.pie(filtered_data, values='percentage', names='Пол', title=f'Должности по полу: {selected_industry}')
    return fig

@app.callback(
    Output('positions-age-graph', 'figure'),
    [Input('positions-age-dropdown', 'value')]
)
def update_age_chart(selected_industry):
    filtered_data = positions_by_age[positions_by_age['Отрасль'] == selected_industry]
    fig = px.pie(filtered_data, values='percentage', names='age_group', title=f'Должности по возрасту: {selected_industry}')
    return fig

@app.callback(
    Output('positions-unemployment-graph', 'figure'),
    [Input('positions-unemployment-dropdown', 'value')]
)
def update_unemployment_chart(selected_industry):
    filtered_data = positions_by_unemployment[positions_by_unemployment['Отрасль'] == selected_industry]
    fig = px.bar(filtered_data, x='Основание незанятости', y='count', title=f'Должности по основанию незанятости: {selected_industry}')
    return fig

@app.callback(
    Output('positions-termination-graph', 'figure'),
    [Input('positions-termination-dropdown', 'value')]
)
def update_termination_chart(selected_industry):
    filtered_data = positions_by_termination[positions_by_termination['Отрасль'] == selected_industry]
    fig = px.bar(filtered_data, x='Основание увольнения', y='count', title=f'Должности по основанию увольнения: {selected_industry}')
    return fig

@app.callback(
    Output('positions-termination-year-graph', 'figure'),
    [Input('positions-termination-year-dropdown', 'value')]
)
def update_termination_year_chart(selected_industry):
    filtered_data = positions_by_termination_year[positions_by_termination_year['Отрасль'] == selected_industry]
    fig = px.line(filtered_data, x='Год увольнения', y='count', title=f'Должности по годам увольнения: {selected_industry}')
    return fig

@app.callback(
    Output('positions-education-graph', 'figure'),
    [Input('positions-education-dropdown', 'value')]
)
def update_education_chart(selected_industry):
    filtered_data = positions_by_education[positions_by_education['Отрасль'] == selected_industry]
    fig = px.bar(filtered_data, x='Образование', y='count', title=f'Должности по образованию сотрудников: {selected_industry}')
    return fig

@app.callback(
    Output('positions-salary-graph', 'figure'),
    [Input('positions-salary-dropdown', 'value')]
)
def update_salary_chart(selected_industry):
    filtered_data = positions_by_salary[positions_by_salary['Отрасль'] == selected_industry]
    fig = px.pie(filtered_data, values='percentage', names='salary_group', title=f'Должности по среднему заработку: {selected_industry}')
    return fig

@app.callback(
    Output('positions-experience-graph', 'figure'),
    [Input('positions-experience-dropdown', 'value')]
)
def update_experience_chart(selected_industry):
    filtered_data = positions_by_work_experience[positions_by_work_experience['Отрасль'] == selected_industry]
    fig = px.pie(filtered_data, values='percentage', names='work_experience_group', title=f'Должности по стажу: {selected_industry}')
    return fig

if __name__ == '__main__':
    app.run_server(debug=True)