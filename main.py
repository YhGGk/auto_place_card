import sys
import ui_auto_place_card

from PyQt5.QtWidgets import QApplication, QMainWindow

#程序UI生成及操作
if __name__ == '__main__':

    # 实例化，传参
    app = QApplication(sys.argv)

    # 创建对象
    mainWindow = QMainWindow()

    # 创建ui，引用ui.py文件中的Ui_Dialog类
    ui = ui_auto_place_card.Ui_Dialog()

    # 调用Ui_Dialog类的setupUi，创建初始组件
    ui.setupUi(mainWindow)

    # 创建窗口
    mainWindow.show()

    # 进入程序的主循环，并通过exit函数确保主循环安全结束(该释放资源的一定要释放)
    sys.exit(app.exec_())
