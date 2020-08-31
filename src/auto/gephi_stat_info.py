# gephi的统计信息
import re

# 获取剪切板信息
import win32con
import win32clipboard


class Stat_Info:
    # 平均度：每个节点相邻节点的数量；
    # 一个节点的度越大，就意味着该节点在某种意义上越重要；
    average_degree = None

    # 平均聚类系数：
    average_clustering_coefficient = None

    # 平均路径长度：网络中任意两个节点（不包括自身到自身）之间距离的平均值；
    average_path_length = None

    # 网络直径
    diameter = None

    def __init__(self):
        pass

    @staticmethod
    def get_average_degree():
        string = Stat_Info.get_clipboard_info()
        average_degree = re.findall(r"</h2>Average Degree:[\s\S\w]+?<br />", string)[0]\
            .replace("</h2>Average Degree: ", "").replace("<br />", "", -1)
        # self.average_degree = average_degree_html.replace("</h2>Average Degree: ", "").replace("<br /><br />", "")
        return average_degree

    @staticmethod
    def get_average_clustering_coefficient():
        string = Stat_Info.get_clipboard_info()
        average_clustering_coefficient = re.findall(r"</h2>Average Clustering Coefficient:[\s\S\w]+?<br />", string)[0]\
            .replace("</h2>Average Clustering Coefficient: ", "").replace("<br />", "", -1)
        return average_clustering_coefficient

    @staticmethod
    def get_average_path_length():
        string = Stat_Info.get_clipboard_info()
        average_path_length =re.findall(r"<br />Average Path length:[\s\S\w]+?<br />", string)[0].replace("<br />Average Path length: ", "").replace("<br />", "", -1)
        return average_path_length

    # 这个值可以附带着和上面的平均路径长度一起出来的，不需要复制两次了；
    @staticmethod
    def get_diameter():
        string = Stat_Info.get_clipboard_info()
        diameter = re.findall(r"</h2>Diameter:[\s\S\w]+?<br />", string)[0].replace("</h2>Diameter: ", "").replace("<br />", "", -1)
        return diameter

    # 获取剪切板内容
    @staticmethod
    def get_clipboard_info():
        # 先打开剪切板，再获取内容
        win32clipboard.OpenClipboard()
        clipboard_info = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return clipboard_info


if __name__ == '__main__':
    pass
    # print(Stat_Info.get_clipboard_info())
    # print(Stat_Info.get_average_degree())
    # print(Stat_Info.get_average_clustering_coefficient())
    # print(Stat_Info.get_diameter())
    # print(Stat_Info.get_average_path_length())
