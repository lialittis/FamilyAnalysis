import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from collections import defaultdict


"""读取excel文件"""

input_path = "./目标/UncleOfSun_choice_data_all.xlsx"
output_path = "./结果/逐时信息统计.xlsx"

df = pd.read_excel(open(input_path,'rb'))
df = df.iloc[: , 1:]

dict_date = {}


"""对每天的数据进行操作，并保存为一个dataframe"""
for index, row in df.iterrows():
    time = row.time
    date = time.date()
    hour = time.hour
    minute = time.minute
    second = time.second
    # check if the date is existed
    if date not in dict_date.keys():
        dict_date[date] = {
                "per_hour":defaultdict(list),
                "per_halfhour":defaultdict(list),
                "per_15mins":defaultdict(list)
        }

    
    dict_date[date]["per_hour"][str(hour)].append(row)
    if minute >= 0 and minute < 30 :
        dict_date[date]["per_halfhour"][str(hour)+":0-29"].append(row)
        if minute < 15:
            dict_date[date]["per_15mins"][str(hour)+":0-14"].append(row)
        else:
            dict_date[date]["per_15mins"][str(hour)+":15-29"].append(row)
    else:
        dict_date[date]["per_halfhour"][str(hour)+":30-59"].append(row)
        if minute < 45:
            dict_date[date]["per_15mins"][str(hour)+":30-44"].append(row)
        else:
            dict_date[date]["per_15mins"][str(hour)+":45-59"].append(row)


"""对df单独保存为一个sheet并保存excel"""
groupof_df_as_output = {}
for date in dict_date:
    day = dict_date[date]
    # for per hour    
    list_dfs_hour = []
    for hour in day["per_hour"]:
        list_of_series = day["per_hour"][hour]
        df_temp = pd.DataFrame(list_of_series)
        df_temp.insert(loc=0,column="分类依据",value=hour)
        list_dfs_hour.append(df_temp)
    groupof_df_as_output[str(date) + "逐时"] = pd.concat(list_dfs_hour)
    # for per half hour
    list_dfs_halfhour = []
    for halfhour in day["per_halfhour"]:
        list_of_series = day["per_halfhour"][halfhour]
        df_temp = pd.DataFrame(list_of_series)
        df_temp.insert(loc=0,column="分类依据",value=halfhour)
        list_dfs_halfhour.append(df_temp)
    groupof_df_as_output[str(date) + "逐半小时"] = pd.concat(list_dfs_halfhour)
    # for per 15 minutes    
    list_dfs_15mins = []
    for mins in day["per_15mins"]:
        list_of_series = day["per_15mins"][mins]
        df_temp = pd.DataFrame(list_of_series)
        df_temp.insert(loc=0,column="分类依据",value=mins)
        list_dfs_15mins.append(df_temp)
    groupof_df_as_output[str(date) + "逐15分钟"] = pd.concat(list_dfs_15mins)


#writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
writer = pd.ExcelWriter(output_path, engine='openpyxl')

for sheet_name in groupof_df_as_output:
    df = groupof_df_as_output[sheet_name]
    print(df)
    df.to_excel(writer,sheet_name=sheet_name)

writer.save()

def saveDFtoWB(df_n,filename):
    # try to store by workbook
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(df_n, index=False, header=True):
        ws.append(r)
    wb.save(filename)












