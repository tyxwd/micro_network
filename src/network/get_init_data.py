import pandas
import os
import json
import time
import numpy
# 可以计算p值的包
import scipy.stats as stats


t1 = time.time()
print("开始时间：" + str(t1))
curPath = os.path.abspath(os.path.dirname(__file__))
# D:\Python\Project\micro_network\src\network
rootPath = curPath[:curPath.find("micro_network") + len("micro_network")]
# D:\Python\Project\micro_network

dataPath = rootPath + "\\data/init_data/"

# all OTU data
init_data = pandas.read_excel(dataPath + "fungi data for analysis.xlsx", sheet_name="all", index_col=0)

# all genus data; 通过genus分类计算出的每个genus的OTU的数目；
# init_data = pandas.read_excel(dataPath + "fungi data by genus.xlsx", sheet_name="Sheet1", index_col=0)

# fungi的分类、功能信息；
fungi_data = pandas.read_excel(dataPath + "fungi data.xlsx", sheet_name="Sheet1", index_col=0)

# 获取json文件中的dict信息
def get_id_species_dict(path):
    json_file = open(path, "r")
    id_species_dict = json.load(json_file)
    json_file.close()
    return id_species_dict


# 计算p值的方法；
def get_p_value(specie_excel):
    # excel所有的列；
    columns = list(specie_excel.columns)
    # 先用两个list直接装p值，后面直接生成DataFrame，加快速度；
    p_value_lists = []

    for column_index, column in enumerate(columns):
        p_value_list = []
        for col_index, col in enumerate(columns):
            if column_index > col_index:
                # p_value 为1，完全不相关；
                p_value = 1
            else:
                p_value = stats.spearmanr(specie_excel[column], specie_excel[col])[1]
            # p_value_excel[column][col] = p_value
            p_value_list.append(p_value)
        p_value_lists.append(p_value_list)

    # 生成p_value的Excel，为了好看进行转秩
    p_value_excel = pandas.DataFrame(p_value_lists, columns=columns, index=columns).T
    # p_value_excel.to_excel(rootPath + "\\data/out_data/final_OTU_p.xlsx")
    return p_value_excel


# 计算相关系数
def get_corr_coefficient(specie_excel):
    # 计算相关系数，spearman
    corr_excel = specie_excel.corr(method="spearman", min_periods=8)

    p_value_excel = get_p_value(specie_excel)
    # p值大于0.01的改成0
    corr_excel[p_value_excel > 0.01] = 0
    # corr_excel.to_excel(rootPath + "\\data/out_data/final_OTU_corr.xlsx")
    return corr_excel


# 生成边表格
def get_line_excel(corr_excel):
    target_excel_col_index = ["Source", "Target", "Type", "Id", "Label", "timeset", "Weight", "Neg_Pos"]
    # 源Excel的信息
    source_excel_row_index = list(corr_excel.index)
    source_excel_col_index = list(corr_excel.columns)
    # 行数
    source_excel_rows = len(source_excel_col_index)
    # 列数
    source_excel_cols = len(source_excel_col_index)

    # 目标输出Excel
    line_excel = pandas.DataFrame(columns=target_excel_col_index)

    for line in range(source_excel_rows):
        for col in range(source_excel_cols):
            Weight = corr_excel.iloc[line, col]

            # 判断该单元格是否为空，为0：
            if not pandas.isna(Weight) and Weight != 0:
                # 不为空、0之后需要添加的信息
                append_data_list = [source_excel_row_index[line],   #"Source"
                                    source_excel_col_index[col],    #"Target"
                                    "Undirected",                   #"Type"
                                    source_excel_row_index[line],   #"Id"
                                    None,                           #"Label"
                                    None,                           #"timeset"
                                    Weight,                         #"Weight"
                                    "Positive" if Weight > 0 else "Negative"] #"Neg_Pos"
                append_data_dict = dict(zip(target_excel_col_index, append_data_list))

                line_excel = line_excel.append([append_data_dict], ignore_index=False)
    return line_excel


# 生成点表格；
def get_point_excel(line_excel, fungi_data, isGenus=True):
    # point表格的列名
    # columns = ["Id", "Label", "timeset"]

    # 获取所有的point，排除重复值
    all_points = line_excel["Source"].unique()

    point_excel = pandas.DataFrame(index=all_points)
    point_excel["Id"] = all_points
    point_excel["Label"] = all_points
    point_excel["timeset"] = None

    # 如果已经是genus表格，则不需要按geus分类了；
    point_excel["genus"] = None if isGenus else fungi_data['geus']
    point_excel["phyl"] = fungi_data["phyl"]
    point_excel["class"] = fungi_data["class"]
    point_excel["order"] = fungi_data["order"]
    point_excel["family"] = fungi_data["family"]
    return point_excel


if __name__ == '__main__':
    # sample ID 的物种名录
    id_species_dict = get_id_species_dict(dataPath + "id_species_dict.json")

    # 所有不重复的物种名
    all_speices = numpy.unique(list(id_species_dict.values()))
    print(all_speices)

    # 总代码，计算每个species的相关系数矩阵；
    i = 0
    for specie in all_speices:
        # 得到是这个specie的所有ID
        temp = []
        for key, value in id_species_dict.items():
            if value == specie:
                temp.append(key)
        # 第一个物种的dataFrame，'Achyranthes aspera'
        specie_DataFrame = init_data.loc[temp]
        print("当前物种：" + specie)
        # 有的成改成1，没的还是0；
        zero_or_one = specie_DataFrame.copy()
        zero_or_one[zero_or_one > 0] = 1

        # 计算每列的和，返回的是一个series；
        sum_series = zero_or_one.sum()

        # >= 5 的series的index；
        sum_series_index = list(sum_series[sum_series >= 5].index)

        # 最后需要保留的columns
        final_OTU = specie_DataFrame[sum_series_index]
        # final_OTU.to_excel(rootPath + "\\data/out_data/final_OTU.xlsx")
        corr_excel = get_corr_coefficient(final_OTU)
        line_excel = get_line_excel(corr_excel)
        point_excel = get_point_excel(line_excel, fungi_data, False)
        line_excel.to_excel(rootPath + "\\data/out_data/%s_line.xlsx" % specie, index=False)
        point_excel.to_excel(rootPath + "\\data/out_data/%s_point.xlsx" % specie, index=False)
        i += 1
        if i == 5:
            break

    t2 = time.time()
    print("结束时间：" + str(t2))
    print(str(int(t2 - t1)) + "秒")
