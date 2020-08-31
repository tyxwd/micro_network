# 按genus分类每个OTU，并生成Excel输出；
import pandas
import os
curPath = os.path.abspath(os.path.dirname(__file__))
# D:\Python\Project\micro_network\src\network
rootPath = curPath[:curPath.find("micro_network") + len("micro_network")]
# D:\Python\Project\micro_network

dataPath = rootPath + "\\data/init_data/"

# fungi的genus信息（geus）写错了；
fungi_data = pandas.read_excel(dataPath + "fungi data.xlsx", sheet_name="Sheet1", index_col=0)


# species的fungi信息；
species_fungi_data = pandas.read_excel(dataPath + "fungi data for analysis.xlsx", sheet_name="all", index_col=0)
species_fungi_data_columns = species_fungi_data.columns


# 遍历fungi的index，将每行根据genus生成dict；
geus_id_dict = {}
for column in species_fungi_data_columns:
    # 每个column的geus
    column_genus = fungi_data.loc[column, "geus"]
    if column_genus not in geus_id_dict.keys():
        geus_id_dict[column_genus] = [column]
    else:
        geus_id_dict[column_genus].append(column)


# 遍历genus的dict，生成dataframe，计算每行的和 （species_fungi_data）；
all_Series_dict = {}
# key 是genus；value是该genus对应的OTU
for key, value in geus_id_dict.items():
    temp_Series = species_fungi_data.loc[:, value].sum(axis=1)
    all_Series_dict[key] = temp_Series

new_DataFrame = pandas.DataFrame(all_Series_dict)
new_DataFrame.to_excel("new.xlsx")