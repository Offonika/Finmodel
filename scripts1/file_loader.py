import pandas as pd
import os

def load_files(folder_path):
    dataframes = []
    for fname in os.listdir(folder_path):
        if fname.endswith('.xlsx') or fname.endswith('.csv'):
            fpath = os.path.join(folder_path, fname)
            print(f'📥 Загрузка файла: {fname}')
            df = pd.read_excel(fpath) if fname.endswith('.xlsx') else pd.read_csv(fpath)
            df['organization'] = os.path.basename(folder_path)  # Пример, можно заменить на парсинг имени
            dataframes.append(df)
    return pd.concat(dataframes, ignore_index=True)
