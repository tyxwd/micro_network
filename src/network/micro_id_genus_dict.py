# 将Excel文件的fungi信息转化成json文件；好像没啥用；…………
import pandas
import json
import os

curPath = os.path.abspath(os.path.dirname(__file__))
# D:\Python\Project\micro_network\src\network
rootPath = curPath[:curPath.find("micro_network") + len("micro_network")]
# D:\Python\Project\micro_network

dataPath = rootPath + "\\data/init_data/"

fungi_data_path = dataPath + "fungi data.xlsx"
fungi_data = pandas.read_excel(fungi_data_path, sheet_name="Sheet1", index_col=0)
fungi_data_columns = fungi_data.columns
fungi_data_index = fungi_data.index


all_fungi_dict = {}
for fungi in fungi_data_index:
    all_fungi_dict[fungi] = dict(zip(fungi_data_columns, fungi_data.loc[fungi]))

all_fungi_dict_json_obj = json.dumps(all_fungi_dict)

all_fungi_dict_file = open(dataPath + "all_fungi_dict.json", "w")
all_fungi_dict_file.writelines(all_fungi_dict_json_obj)
all_fungi_dict_file.close()
