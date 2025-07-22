# main.py

from file_loader import load_files
from aggregator import aggregate_data
from excel_writer import write_to_excel, write_df_to_excel_table

def main():
    org_folder = r"C:\Users\Public\Finmodel\НачисленияУслугОзон\ИП Закирова Р.Х"
    files_df = load_files(org_folder)
    result_df = aggregate_data(files_df)
    output_path = r"C:\Users\Public\Finmodel\Finmodel.xlsm"
    sheet = 'НачисленияУслугОзон'
    table = 'НачисленияУслугОзонTable'
    # 1. Сохраняем просто DataFrame на лист
    write_to_excel(result_df, output_path, sheet_name=sheet)
    # 2. Преобразуем лист в умную таблицу (Excel Table)
    write_df_to_excel_table(result_df, output_path, sheet, table)

if __name__ == '__main__':
    main()
