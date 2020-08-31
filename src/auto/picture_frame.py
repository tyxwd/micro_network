# 分格显示
import matplotlib.pyplot as plt
import matplotlib.image as imgplt


class Picture_Frame:
	m_plot = plt
	image_rows = 4
	image_cols = 2

	def __init__(self):
		super()

	@staticmethod
	def add_image(image_index=1, image_path="保存"):
		image_path = "./pictures/%s.png" % image_path
		Picture_Frame.m_plot.subplot(Picture_Frame.image_rows, Picture_Frame.image_cols, image_index)
		Picture_Frame.m_plot.axis('off')
		image = imgplt.imread(image_path)
		Picture_Frame.add_text_info()
		Picture_Frame.m_plot.imshow(image)

	@staticmethod
	def add_text_info(title="title", text={"x": 1, "y": -1, "string": "text_info"}, font="Times New Roman"):
		Picture_Frame.m_plot.title(title, font=font)
		Picture_Frame.m_plot.text(text["x"], text["y"], text["string"], font=font)

	@staticmethod
	def show():
		Picture_Frame.m_plot.show()

	# 这是没写完整的
	@staticmethod
	def save():
		Picture_Frame.m_plot.savefig()


if __name__ == '__main__':
	for index, image in enumerate(["保存", "停止", "刷新", "桌面"]):
		Picture_Frame.add_image(index + 1, image)
	Picture_Frame.show()