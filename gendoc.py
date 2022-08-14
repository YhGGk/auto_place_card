import win32com.client
import projectfunctions as pf

class Gendoc():

    def __init__(self, names=[]):
        self.names = names
        self.number_names = len(self.names)

    def add_document(self):
        # 创建一个新的 后台运行的 空白文档
        self.word = win32com.client.Dispatch('Word.Application')  # 打开word程序
        self.word.Visible = 1  # 是否在后台运行word
        self.word.DisplayAlerts = 0  # 是否显示警告信息
        self.doc = self.word.Documents.Add()  # 创建一个新的word文档

    def doc_layout(self):
        # 整体文档页面设置
        self.doc.PageSetup.Orientation = 0  # 页面设置1为横向2为纵向

        # 页边距设置
        self.doc.PageSetup.TopMargin = 0 * pf.pt_in_cm(3, 1)
        self.doc.PageSetup.BottomMargin = 0 * pf.pt_in_cm(3, 1)
        self.doc.PageSetup.LeftMargin = 0 * pf.pt_in_cm(3, 1)
        self.doc.PageSetup.RightMargin = 0 * pf.pt_in_cm(3, 1)

        self.range = self.doc.Range(0, 0)  # 获取该文档最开头
        self.range.Style.Font.Name = "华文新魏"  # 设置字体
        self.range.Style.Font.Size = 150  # 设置字号

    def generate_table(self):
        # 生成表格并调整列宽
        self.table = self.range.Tables.Add(self.doc.Range(self.range.End, self.range.End), self.number_names, 2)  # 建一张表格

        # 设置列宽
        column = self.table.columns(1)
        column.PreferredWidth = 9.25 * pf.pt_in_cm(3, 1)
        column = self.table.columns(2)
        column.PreferredWidth = 9.25 * pf.pt_in_cm(3, 1)

    def generate_place_cards(self):
        # 为每个人逐个生成席卡
        row_index = 1
        for name in self.names:
            row = self.table.rows(row_index)

            # 设置行格式
            row.HeightRule = 2  # 表格行高为精确值
            row.Height = 20 * pf.pt_in_cm(3, 1)

            # 在左边的单元格填入一个人名，并设置好格式
            cell = self.table.Cell(row_index, 1)
            cell.VerticalAlignment = 0  # 表格内文字与上边框对齐
            cell.Borders.Enable = True  # 单元格所有边框开启
            range_ = cell.Range
            range_.Text = name  # 输入人名
            range_.Orientation = 3  # 文字方向设置为向左
            range_.ParagraphFormat.Alignment = 1  # 文字居中对齐

            # 字体大小自适应
            if len(name) >= len("一二三四"):
                range_.Font.Size = 100
                range_.ParagraphFormat.LineSpacing = 100  # 行距
                range_.ParagraphFormat.LineSpacingRule = 4  # 行距规则：4为固定值 3为最小值

            # 在右边的单元格填入一个人名，并设置好格式
            cell = self.table.Cell(row_index, 2)
            cell.VerticalAlignment = 0  # 表格内文字与上边框对齐
            cell.Borders.Enable = True  # 单元格所有边框开启
            range_ = cell.Range
            range_.Text = name
            range_.Orientation = 2  # 文字方向设置为向右
            range_.ParagraphFormat.Alignment = 1  # 文字居中对齐

            # 字体大小自适应
            if len(name) >= len("一二三四"):
                range_.Font.Size = 100
                range_.ParagraphFormat.LineSpacing = 100  # 行距
                range_.ParagraphFormat.LineSpacingRule = 4  # 行距规则：4为固定值 3为最小值

            # 单元格左右边距
            self.table.RightPadding = 1.5 * pf.pt_in_cm(3, 1)
            self.table.LeftPadding = 1.5 * pf.pt_in_cm(3, 1)

            row_index += 1

    def word_visible(self):
        # 使word前台显示
        self.word.Visible = 1
