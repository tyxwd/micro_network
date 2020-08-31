from pywinauto import keyboard, mouse, application
import time, cv2, numpy
import aircv
# 进行截图
import pyautogui
import gephi_stat_info

import pandas


class Gephi_Auto:
    def __init__(self, gephi_path=r'C:\Program Files\Gephi-0.9.2\bin\gephi64.exe'):
        Gephi_Auto.min_click()
        self.gephi_path = gephi_path
        application.Application(backend="uia").start(self.gephi_path)
        time.sleep(7)
        # 进行一次截图
        self.screenshot()
        self.mouse_click(self.get_positon("概览"))

    def keyboard_click(self, string, sleep_time=0.8):
        keyboard.send_keys(string)
        time.sleep(sleep_time)
        self.screenshot()

    def mouse_click(self, position, sleep_time=0.8, button='left'):
        # 添加鼠标的移动轨迹
        # pyautogui.moveTo(position[0], position[1], duration=0.4 * sleep_time)
        mouse.click(button=button, coords=position)
        time.sleep(sleep_time)
        self.screenshot()

    # 最小化pycharm界面
    @staticmethod
    def min_click():
        mouse.click('left', coords=[1774, 14])

    # 全屏截图的方法
    def screenshot(self, name="刚打开"):
        pyautogui.screenshot().save("./runing_img/%s.jpg" % name)
        self.now_screenshot = cv2.imdecode(numpy.fromfile("./runing_img/%s.jpg" % name, dtype=numpy.uint8), -1)

    # imgsrc=原始图像，imgobj=待查找的图片
    def matchImg(self, imgobj, confidencevalue=0.7):
        imobj = cv2.imdecode(numpy.fromfile(imgobj, dtype=numpy.uint8), -1)  # 读取中文路径及名称
        # 读英文路径
        # imobj = aircv.imread(imgobj)

        match_result = aircv.find_template(self.now_screenshot, imobj, confidencevalue)  # {'confidence': 0.5435812473297119, 'rectangle': ((394, 384), (394, 416), (450, 384), (450, 416)), 'result': (422.0, 400.0)}
        if match_result is not None:
            match_result['shape'] = (self.now_screenshot.shape[1], self.now_screenshot.shape[0])  # 0为高，1为宽
        position = (int(match_result["result"][0]), int(match_result["result"][1]))
        return position
        # {'result': (12.0, 13.0),
        #  'rectangle': ((2, 4), (2, 22), (22, 4), (22, 22)),
        #  'confidence': 0.9873911142349243,
        #  'shape': (1920, 1040)}

    # 获取图片右边位置
    def match_img_right(self, imgobj, confidencevalue=0.7):
        imobj = cv2.imdecode(numpy.fromfile(imgobj, dtype=numpy.uint8), -1)  # 读取中文路径及名称
        # 读英文路径
        # imobj = aircv.imread(imgobj)

        match_result = aircv.find_template(self.now_screenshot, imobj, confidencevalue)
        if match_result is not None:
            match_result['shape'] = (self.now_screenshot.shape[1], self.now_screenshot.shape[0])  # 0为高，1为宽
        position = (int(match_result["rectangle"][3][0] * 0.95),
                    int(match_result["result"][1]))
        return position
    # 获取position
    def get_positon(self, name):
        return self.matchImg("./pictures/%s.png" % name)

    def get_right_position(self, name):
        return self.match_img_right("./pictures/%s.png" % name)

    # 导入点表格
    def import_point_file(self, point_file):
        self.mouse_click(self.get_positon("文件"), 1)
        self.mouse_click(self.get_positon("导入电子表格"), 2)
        self.mouse_click(self.get_positon("桌面"))
        self.mouse_click(self.get_positon("输入框"))
        self.keyboard_click(point_file.replace(" ", "{SPACE}", -1))
        # 按3次回车
        for i in range(3):
            self.keyboard_click("{ENTER}")
        self.mouse_click(self.get_positon("混合的"))
        # 按上方向键
        self.keyboard_click("{UP}")
        self.keyboard_click("{ENTER}")
        self.keyboard_click("{ENTER}")

    # 导入边表格
    def import_line_file(self, line_file):
        self.mouse_click(self.get_positon("文件"), 1)
        self.mouse_click(self.get_positon("导入电子表格"), 2)
        self.mouse_click(self.get_positon("桌面"))
        self.mouse_click(self.get_positon("输入框"))
        self.keyboard_click(line_file.replace(" ", "{SPACE}", -1))
        # 按3次回车
        for i in range(3):
            self.keyboard_click("{ENTER}", 2)
        self.mouse_click(self.get_positon("Append"))
        self.keyboard_click("{ENTER}")

    # 改颜色和大小
    def color_and_size(self):
        self.mouse_click(self.get_positon("Partition"))
        self.mouse_click(self.get_positon("选择一种渲染方式"))
        self.mouse_click(self.get_positon("phyl"))
        self.mouse_click(self.get_positon("应用"))

        self.mouse_click(self.get_positon("point_size"))
        self.mouse_click(self.get_positon("Ranking"))
        self.mouse_click(self.get_positon("选择一种渲染方式"))
        self.mouse_click(self.get_positon("度"))
        self.mouse_click(self.get_positon("应用"))

        self.mouse_click(self.get_positon("边"))
        self.mouse_click(self.get_positon("Partition"))
        self.mouse_click(self.get_positon("选择一种渲染方式"))
        self.mouse_click(self.get_positon("neg_pos"))
        self.mouse_click(self.get_positon("应用"))

    # 获取网络图的相关信息
    def get_gephi_info(self, string="average_degree"):
        if string == "average_degree":
            # 点击统计按钮
            self.mouse_click(self.get_positon("统计"))
            self.mouse_click(self.get_right_position("平均度"))
            self.mouse_click(self.get_positon("复制"))
            average_degree = gephi_stat_info.Stat_Info.get_average_degree()
            self.mouse_click(self.get_positon("关闭"))
            return average_degree
        elif string == "average_clustering_coefficient":
            self.mouse_click(self.get_right_position("平均聚类系数"))
            self.mouse_click(self.get_positon("确定"))
            self.mouse_click(self.get_positon("复制"))
            average_clustering_coefficient = gephi_stat_info.Stat_Info.get_average_clustering_coefficient()
            self.mouse_click(self.get_positon("关闭"))
            return average_clustering_coefficient
        else:
            self.mouse_click(self.get_right_position("平均路径长度"))
            self.mouse_click(self.get_positon("确定"))
            self.mouse_click(self.get_positon("复制"))
            average_path_length = gephi_stat_info.Stat_Info.get_average_path_length()
            diameter = gephi_stat_info.Stat_Info.get_diameter()
            self.mouse_click(self.get_positon("关闭"))
            return average_path_length, diameter

    # 将网络图信息生成字典
    def save_gephi_info(self, specie):
        gephi_info = {"specie": specie,
                      "average_degree": self.get_gephi_info("average_degree"),
                      "average_clustering_coefficient": self.get_gephi_info("average_clustering_coefficient"),
                      "average_path_length": self.get_gephi_info("average_path_length")[0],
                      "diameter": self.get_gephi_info("average_path_length")[1]}
        return gephi_info

    #开始画
    def layout(self):
        self.mouse_click(self.get_positon("过滤"))
        self.mouse_click(self.get_positon("选择一个布局"))
        self.mouse_click(self.get_positon("Fruch"))
        self.mouse_click(self.get_positon("运行"))
        time.sleep(30)
        self.mouse_click(self.get_positon("停止"))

    # 进入预览
    def view_and_save(self, final_name):
        self.mouse_click(self.get_positon("预览"))
        self.mouse_click(self.get_positon("mixed"))
        self.keyboard_click("original")
        self.mouse_click(self.get_positon("刷新"))
        self.mouse_click(self.get_positon("save"), 1)
        # 或许不需要修改类型；
        # self.mouse_click(self.get_positon("文件类型"))
        # self.mouse_click(self.get_positon("svg"))
        # self.mouse_click(self.get_positon("输入框_save"))

        # 需要ctrl+a 全选后删除；
        self.keyboard_click("{VK_CONTROL down}"
                            "{a down}"
                            "{VK_CONTROL up}"
                            "{a up}"
                            "{DEL}")
        self.keyboard_click(final_name.replace(" ", "{SPACE}", -1) + ".svg")
        self.mouse_click(self.get_positon("保存"))
        self.keyboard_click("{ENTER}")

    # 关闭gephi
    @staticmethod
    def close():
        # alt+f4
        keyboard.send_keys("%{F4}")
        time.sleep(0.5)
        keyboard.send_keys("{TAB}")
        time.sleep(0.5)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.5)

    # 增加异常处理
    @staticmethod
    def exception_handling():
        keyboard.send_keys("{ESC}")
        time.sleep(0.5)
        keyboard.send_keys("{ESC}")
        time.sleep(0.5)

    def start(self, point_file, line_file):
        self.import_point_file(point_file)
        self.import_line_file(line_file)
        self.color_and_size()

        # gephi信息
        gephi_info = self.save_gephi_info(point_file.split("_")[0])
        self.layout()
        self.view_and_save(point_file.split("_")[0])
        return gephi_info


if __name__ == '__main__':
    species = ['Achyranthes aspera',
               'Bidens biternata',
               'Bidens pilosa',
               'Borreria latifolia',
               'Borreria stricta',
               'Celosia argentea',
               'Malvastrum coromandelianum',
               'Paspalum conjugatum',
               'Paspalum thunbergii',
               'Physalis angulata',
               'Senna occidentalis',
               'Senna tora',
               'Setaria palmifolia',
               'Setaria viridis',
               'Solanum nigrum',
               'Urena lobata']

    final_gephi_info = {"specie": [],
                      "average_degree": [],
                      "average_clustering_coefficient": [],
                      "average_path_length": [],
                      "diameter": []}
    for specie in species:
        draw_succeed = False
        # 只要没成功就一直画
        while not draw_succeed:
            try:
                gephi = Gephi_Auto()
                gephi_info = gephi.start(specie + "_point.xlsx", specie + "_line.xlsx")
                final_gephi_info["specie"].append(gephi_info["specie"])
                final_gephi_info["average_degree"].append(gephi_info["average_degree"])
                final_gephi_info["average_clustering_coefficient"].append(gephi_info["average_clustering_coefficient"])
                final_gephi_info["average_path_length"].append(gephi_info["average_path_length"])
                final_gephi_info["diameter"].append(gephi_info["diameter"])
                draw_succeed = True
            except:
                # 有异常直接进行关闭
                Gephi_Auto.exception_handling()
            finally:
                Gephi_Auto.close()

    # 保存Excel文件
    gephi_info_excel = pandas.DataFrame(final_gephi_info)
    gephi_info_excel.to_excel("gephi_info.xlsx", index=False)