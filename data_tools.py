import pandas as pd
from pathlib import Path
import openpyxl
from openpyxl.styles import PatternFill,Font,Border,Side 
import plotly.express as px

def scan_folder(folder_path, standard_cols, rename_map=None):
    folder = Path(folder_path)
    valid_files = []
    report = []

    if not folder.exists():
        print(f"Folder path not found:{folder_path}")
        return None
    
    for file in folder.glob("*.xls*"):
        try:
            df = pd.read_excel(file)
        except Exception as e:
            report.append({'file': file.name, 'issue': f"Cannot open file:{e}"})
            continue
        if df.empty:
            report.append({'file': file.name, 'issue': "Empty file"})
            continue
        actual_cols = list(df.columns)
        missing_cols = [col for col in standard_cols if col not in actual_cols]
        if missing_cols:
            if rename_map:
                R_name = [r_n for r_n in df.columns if r_n in rename_map]
                if not R_name:
                    report.append({'file': file.name, 'issue': f"Missing columns: {missing_cols}"})
                    continue
                else:
                    df = df.rename(columns=rename_map)
                    actual_cols = list(df.columns)
                    missing_cols = [col for col in standard_cols if col not in actual_cols]
                    if missing_cols:
                        report.append({'file': file.name, 'issue': f"Missing columns: {missing_cols}"})
                        continue
            else:
                report.append({'file': file.name, 'issue': f"missing columns: {missing_cols}"})
                continue
        valid_files.append(str(file))
    
    return {
        "valid_files": valid_files,
        "report": report
    }


def read_and_align(file_path,standard_cols,rename_map=None):
    '''
    尝试读写方式打开excel文件,无法读取的返回报告,可以读取的对空列表返回None,对可处理列表做表头检测。
    file_path : 文件名
    rename_map = [] : 表头映射字典(可选)
    standard_cols : 表头列值
    '''
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Cannot read file {file_path}: {e}")
        return None
    if rename_map:
    #映射字典
        df=df.rename(columns=rename_map)
    if not df.empty:
        try:
            df = df[standard_cols]
            return df
        except Exception as e:
            print(f"{file_path} processing failed: {e}")
            return None
    else:
        print(f"{file_path} is empty")
        return None
    
def merge_all(df_list):
    '''
    对有效文件列表做合并处理
    df_list : 有效文件列表
    ignore_index : 行号重置
    '''
    merged=pd.concat(df_list, ignore_index=True)
    return merged

def clean_data(df,col,subset=None):
    '''
    对列表进行清洗,建立列表副本后进行查重-修改格式 错误值转NaN-删除空白值，最后输出处理报告
    df : 有效列表文件
    col : 格式转换行名
    subset : 可调只读列查重 默认值关闭
    输出处理报告 : 原始{before}行->保留{after}行，删除{before-after}行
    '''
    df=df.copy()
    before=len(df)

    df=df.drop_duplicates(subset=subset)
    df[col]=pd.to_numeric(df[col], errors='coerce')
    df=df.dropna()
    df[col]=df[col].astype(int)
    
    after = len(df)
    print(f'Data cleaned:{before}rows——>{after}rows kept,{before-after}rows removed')
    return df

def export_excel(df,output_path,header_color='4F81BD'):
    '''
    对有效excel列表文件处理,在需要的时候用openpyxl写入模式打开文件,然后美化表格,最后保存到输出文件夹里
    df : 有效列表文件名
    output_path : 输出文件夹
    header_color : 表头颜色 默认蓝色
    '''
    with pd.ExcelWriter(output_path,engine='openpyxl') as writer:
        df.to_excel(writer,index=False,sheet_name='汇总')
    wb=openpyxl.load_workbook(output_path)
    ws=wb.active

    fill=PatternFill(fill_type='solid',fgColor=header_color)
    font=Font(bold=True,color='FFFFFF')
    thin=Side(style='thin')
    border=Border(left=thin,right=thin,top=thin,bottom=thin)
    ws.freeze_panes = 'A2'

    for cell in ws[1]:
        cell.fill=fill
        cell.font=font
        cell.border=border
    for row_index,row in enumerate(ws.iter_rows(min_row=2),1):
        for cell in row:
            cell.border = border
            if row_index % 2:
                cell.fill = PatternFill(fill_type='solid',fgColor='E8F4FD')

    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length,len(str(cell.value)))
        ws.column_dimensions[col_letter].width=max_length+4
    wb.save(output_path)
    print(f'Excel exported:{output_path}')

def plot_bar(df,output_folder,filename,x,y,color=None,title=None):
    '''
    自动绘制柱状图
    '''
    output_path=Path(output_folder)/(filename+ '.html')
    fig=px.bar(df,x=x,y=y,color=color,title=title)
    fig.write_html(output_path)
    print(f'Chart exported:{output_path}')