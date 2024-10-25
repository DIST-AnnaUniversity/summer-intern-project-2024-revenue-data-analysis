import pandas as pd
import os
from configparser import ConfigParser

# Load configuration from config.ini
config = ConfigParser()
config.read('config3.ini')

# Read configuration values
folder_path = config.get('Paths', 'folder_path')
base_forecast_sheet = config.get('Sheets', 'base_forecast_sheet')
weekly_forecast_sheet = config.get('Sheets', 'weekly_forecast_sheet')
base_forecast_col = int(config.get('Columns', 'base_forecast_col'))
weekly_forecast_col = int(config.get('Columns', 'weekly_forecast_col'))
project_col = int(config.get('Columns', 'project_col'))
email_col = int(config.get('Columns', 'email_col'))
table = []
for file_name in os.listdir(folder_path):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file_name)
        print(f"Processing {file_path}...")
        df1 = pd.read_excel(file_path, sheet_name=base_forecast_sheet)
        df2 = pd.read_excel(file_path, sheet_name=weekly_forecast_sheet)
        data1 = df1.iloc[17:, base_forecast_col].dropna()
        data2 = df2.iloc[17:, weekly_forecast_col].dropna()
        common_data = set(data1) & set(data2)
        uncommon_data = (set(data1) | set(data2)) - common_data
        for value in uncommon_data:
            base_forecast_rows = data1[data1 == value].index
            for idx in base_forecast_rows:
                project = df1.iloc[idx, project_col]
                mail_id = df1.iloc[idx, email_col]
                base_hours = df1.iloc[idx, base_forecast_col]
                table.append([project, mail_id, base_hours, None, file_name])
        for value in uncommon_data:
            weekly_forecast_rows = data2[data2 == value].index
            for idx in weekly_forecast_rows:
                project = df2.iloc[idx, project_col]
                mail_id = df2.iloc[idx, email_col]
                weekly_hours = df2.iloc[idx, weekly_forecast_col]
                table.append([project, mail_id, None, weekly_hours, file_name])
df_uncommon = pd.DataFrame(table, columns=['Project', 'EmailID', 'Base-Hours', 'Weekly-Hours', 'Source File'])
df_uncommon = df_uncommon.groupby(['Project', 'EmailID', 'Source File'], as_index=False).first()
print(df_uncommon)


