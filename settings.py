import projectfunctions as pf

class Settings():
    # 设置参数

    def __int__(self):
        # 文档方向
        self.pagesetup_orientation = 1  # 1表示横向，0表示纵向

        # 页边距值
        self.pagesetup.topMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.bottomMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.leftMargin = 0 * pf.pt_in_cm(3, 1)
        self.pagesetup.rightMargin = 0 * pf.pt_in_cm(3, 1)

        # 字体大小、名称
        self.fontsize_less3 = 140  # 三字及以下的字体大小
        self.fontsize_over3 = 100  # 超过三字的字体大小
        self.fontname = "华文新魏"  # 字体名字

        # 席卡横向宽度、纵向高度
        self.tablewidth = 20 * pf.pt_in_cm(3, 1)
        self.tableheight = 9.25 * pf.pt_in_cm(3, 1)

        # 是否需要自动适应字体大小
        self.auto_format = True
