import xlwings as xw

def write_to_excel(df, filename, sheet_name='Sheet1'):
    df.to_excel(filename, sheet_name=sheet_name, index=False, engine='openpyxl')

def write_df_to_excel_table(df, file_path, sheet_name, table_name):
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    ws = wb.sheets[sheet_name]
    nrows, ncols = df.shape
    rng = ws.range((1, 1), (nrows + 1, ncols))
    for tbl in ws.api.ListObjects:
        if tbl.Name == table_name:
            tbl.Delete()
    ws.api.ListObjects.Add(1, rng.api, None, 1).Name = table_name
    wb.save()
    wb.close()
    app.quit()
    print(f'✅ Данные записаны в файл: {file_path}, лист: {sheet_name}, таблица: {table_name}')
