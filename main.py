import csv

import pandas as pd
import xlwings as xw


# 1.
def load_data():
    reviews_df = pd.read_csv('reviews_sample.csv')
    recipes_df = pd.read_csv('recipes_sample.csv')
    recipes_df = recipes_df[['id', 'name', 'minutes', 'submitted', 'description', 'n_ingredients']]
    return reviews_df, recipes_df


# 2.
def save_sample_to_excel(reviews_df, recipes_df):
    reviews_sample = reviews_df.sample(frac=0.05)
    recipes_sample = recipes_df.sample(frac=0.05)

    with pd.ExcelWriter('recipes.xlsx', engine='openpyxl') as writer:
        recipes_sample.to_excel(writer, sheet_name='Рецепты', index=False)
        reviews_sample.to_excel(writer, sheet_name='Отзывы', index=False)


# 3.
def add_seconds_assign():
    wb = xw.Book('recipes.xlsx')
    sheet = wb.sheets['Рецепты']
    minutes = sheet.range('C2:C' + str(sheet.cells.last_cell.row)).value
    seconds = [m * 60 for m in minutes]
    sheet.range('G1').value = 'seconds_assign'
    sheet.range('G2').options(transpose=True).value = seconds
    wb.save()
    wb.close()


# 4.
def add_seconds_formula():
    wb = xw.Book('recipes.xlsx')
    sheet = wb.sheets['Рецепты']
    last_row = sheet.cells.last_cell.row
    sheet.range('H1').value = 'seconds_formula'
    sheet.range('H2:H' + str(last_row)).formula = '=$C2*60'
    wb.save()
    wb.close()


# 5.
def format_headers():
    wb = xw.Book('recipes.xlsx')
    sheet = wb.sheets['Рецепты']
    headers = sheet.range('A1:H1')
    headers.font.bold = True
    headers.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    wb.save()
    wb.close()


# 6.
def color_minutes_cells():
    wb = xw.Book('recipes.xlsx')
    sheet = wb.sheets['Рецепты']
    minutes = sheet.range('C2:C' + str(sheet.cells.last_cell.row))

    for cell in minutes:
        if cell.value < 5:
            cell.color = (0, 255, 0)
        elif 5 <= cell.value <= 10:
            cell.color = (255, 255, 0)
        else:
            cell.color = (255, 0, 0)

    wb.save()
    wb.close()


# 7.
def add_n_reviews():
    wb = xw.Book('recipes.xlsx')
    sheet_recipes = wb.sheets['Рецепты']
    sheet_reviews = wb.sheets['Отзывы']

    last_row_recipes = sheet_recipes.cells.last_cell.row
    last_row_reviews = sheet_reviews.cells.last_cell.row

    reviews_count_formula = (
        f'=COUNTIF(Отзывы!$C$2:$C${last_row_reviews}, Рецепты!A2)'
    )

    sheet_recipes.range('I1').value = 'n_reviews'
    sheet_recipes.range('I2:I' + str(last_row_recipes)).formula = reviews_count_formula

    wb.save()
    wb.close()


# 8.
def validate():
    wb = xw.Book('recipes.xlsx')
    sheet_reviews = wb.sheets['Отзывы']
    sheet_recipes = wb.sheets['Рецепты']

    reviews_df = pd.DataFrame(sheet_reviews.range('A1').options(expand='table').value[1:],
                              columns=sheet_reviews.range('A1').expand('right').value)
    recipes_df = pd.DataFrame(sheet_recipes.range('A1').options(expand='table').value[1:],
                              columns=sheet_recipes.range('A1').expand('right').value)

    valid_recipe_ids = set(recipes_df['id'].astype(int))

    for i in range(len(reviews_df)):
        row = reviews_df.iloc[i]
        row_num = i + 2
        rating = row['rating']
        recipe_id = row['recipe_id']

        if not (0 <= rating <= 5) or recipe_id not in valid_recipe_ids:
            sheet_reviews.range(f'A{row_num}:F{row_num}').color = (255, 0, 0)

    wb.save()
    wb.close()


# 9.
def load_and_save_model():
    with open('recipes_model.csv', newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        data = list(reader)

    wb = xw.Book('recipes_model.xlsx')
    sheet = wb.sheets.add('Модель')

    sheet.range('A2').value = data

    wb.save()
    wb.close()


# 10.
def add_sql_formula():
    wb = xw.Book('recipes_model.xlsx')
    sheet = wb.sheets['Модель']

    last_row = sheet.range('A1').end('down').row
    for i in range(2, last_row + 1):
        formula = (
            f'=$B{i} & " " & UPPER($C{i}) & IF($F{i}="PK", " PRIMARY KEY", IF($F{i}="FK", " REFERENCES " & $H{i} & "(" & $I{i} & ")", "")) & IF($D{i}="Y" AND $F{i}<>"PK", " NOT NULL", "")'
        )
        sheet.range(f'J{i}').formula = formula

    wb.save()
    wb.close()


# 11.
def style_model_sheet():
    wb = xw.Book('recipes_model.xlsx')
    sheet = wb.sheets['Модель']

    header = sheet.range('A1:J1')
    header.color = (0, 204, 255)
    header.font.bold = True

    sheet.autofit()
    sheet.range('A1:J1').api.AutoFilter()

    wb.save()
    wb.close()


# 12.
def create_statistics():
    wb = xw.Book('recipes_model.xlsx')
    sheet_model = wb.sheets['Модель']

    data = sheet_model.range('A2').expand('table').value
    df = pd.DataFrame(data, columns=sheet_model.range('A1:J1').value)

    entity_counts = df['Рецепт'].value_counts().reset_index()
    entity_counts.columns = ['Entity', 'Count']

    sheet_stats = wb.sheets.add('Статистика')
    sheet_stats.range('A1').value = entity_counts.values.tolist()
    sheet_stats.range('A1').value = ['Entity', 'Count']

    chart = sheet_stats.charts.add(left=300, top=50, width=500, height=300)
    chart.chart_type = 'column_clustered'
    chart.set_source_data(sheet_stats.range('A1').expand())
    chart.api[1].HasTitle = True
    chart.api[1].ChartTitle.Text = 'Number of Attributes per Entity'

    wb.save()
    wb.close()


def main():
    reviews_df, recipes_df = load_data()
    save_sample_to_excel(reviews_df, recipes_df)
    add_seconds_assign()
    add_seconds_formula()
    format_headers()
    color_minutes_cells()
    add_n_reviews()
    validate()
    load_and_save_model()
    add_sql_formula()
    style_model_sheet()
    create_statistics()


if __name__ == "__main__":
    main()
