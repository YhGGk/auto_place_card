import projectfunctions as pf

class Settings():
    # 设置参数

    def __int__(self):
        # 文档方向
        self.pagesetup_orientation = 1  # 1表示横向，0表示纵向

        # 页边距值
        self.pagesetup.TopMargin = 0.1 * pf.pt_in_cm(3, 1)
        self.pagesetup.BottomMargin = 0.1 * pf.pt_in_cm(3, 1)
        self.pagesetup.LeftMargin = 0.1 * pf.pt_in_cm(3, 1)
        self.pagesetup.RightMargin = 0.1 * pf.pt_in_cm(3, 1)

        # 字体大小、名称
        self.fontsize = 150
        self.fontname = "华文新魏"

        # 席卡横向宽度、纵向高度
        self.tablewidth = 20 * pf.pt_in_cm(3, 1)
        self.tableheight = 9.25 * pf.pt_in_cm(3, 1)

        # 是否需要自动适应字体大小
        self.auto_format = True
