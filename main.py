from data_tools import scan_folder,read_and_align,merge_all,clean_data,export_excel,plot_bar
import os
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# ===== 用户配置区域 =====
FOLDER_PATH = "input"
OUTPUT_PATH = r"output\汇总.xlsx"
STANDARD_COLS = ['name', 'department', 'salary']
# =======================


result = scan_folder(FOLDER_PATH, STANDARD_COLS)
print('Valid files:', result['valid_files'])
print('Issue report:', result['report'])

df_list = []
for file in result['valid_files']:
    df = read_and_align(file, STANDARD_COLS)
    if df is not None:
        df_list.append(df)
print(f'Successfully loaded {len(df_list)} files')

merged_df = merge_all(df_list)
print(f'Merged dataset: {len(merged_df)} rows')
print(merged_df)

clean_data_df = clean_data(merged_df, col='salary')
print(clean_data_df)
export_excel(clean_data_df, OUTPUT_PATH)
output_folder = r"C:\Users\29406\Desktop\Data架构\etl_demo\output"
plot_bar(clean_data_df, output_folder, 'salary', x='name', y='salary', color='department', title='Salary Overview')
