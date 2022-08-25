import projectfunctions as pf

class Settings():
    # 设置参数

    def __int__(self):
        # 文档方向
        self.pagesetup_orientation = 1  # 1表示横向 landscape，0表示纵向 portrait

        # 页边距值
        self.pagesetup.topMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.bottomMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.leftMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.rightMargin = 0 * pf.pt_in_cm(3, 1)

        # 字体大小、名称
        self.fonts = []

        font = {'fontname': '华文新魏', 'fontsize_less3': 140, 'fontsize_more3': 100}  # 字体 华文新魏
        self.fonts.append(font)

        font = {'fontname': '楷体', 'fontsize_less3': 140, 'fontsize_more3': 110}  # 字体 楷体
        self.fonts.append(font)

        font = {'fontname': '仿宋_GB2312', 'fontsize_less3': 140, 'fontsize_more3': 110}  # 字体 仿宋_GB2312
        self.fonts.append(font)

        font = {'fontname': '仿宋', 'fontsize_less3': 140, 'fontsize_more3': 110}  # 字体 仿宋
        self.fonts.append(font)

        font = {'fontname': '宋体', 'fontsize_less3': 140, 'fontsize_more3': 100}  # 字体 宋体
        self.fonts.append(font)

        font = {'fontname': '黑体', 'fontsize_less3': 140, 'fontsize_more3': 110}  # 字体 黑体
        self.fonts.append(font)

        # 席卡横向宽度、纵向高度
        self.tablewidth = {'portrait': 21 * pf.pt_in_cm(3, 1), 'landscape': 20 * pf.pt_in_cm(3, 1)}
        self.tableheight = {'portrait': 9.25 * pf.pt_in_cm(3, 1), 'landscape': 9.25 * pf.pt_in_cm(3, 1)}

        # 是否需要自动适应字体大小
        self.auto_format = True

        # 表格左右边距，根据字数不同进行调整
        self.cellpadding = {'more3': 1 * pf.pt_in_cm(3, 1), 'less3': 1.5 * pf.pt_in_cm(3, 1)}
