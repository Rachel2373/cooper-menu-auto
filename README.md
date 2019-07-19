# cooper-menu-auto
为实现cooper菜单自动化努力！

2019-7-2 面对营养成分表相关难题

学习类属性和对象属性

读取excel

————————————————————————

2019-7-3 关于docx.py 可以在mac上运行，在Windows上报错"cannot import MailMerge"

经排查，不安装"mailmerge"库；应安装"docx-mailmerge"

一身轻松，终于可以在宿舍电脑上继续进行了

————————————————————————

2019-7-17 更新GUI

用PyQt5制作的GUI

pycharm无法下载模块包，但可以用cmd pip install下载安装，最后将python文件夹下的模块包直接复制到menuauto pycharm项目下面，解决。

用designer设计时，直接拖拽工具，挨个试，看看功能是什么，最终完成初稿

在pycharm上编辑了将ui转化为代码的工具，但无法运行，需要在代码后面加上

    class MyWindow(QtWidgets.QMainWindow, Ui_MainWindow):
        def __init__(self):
            super(MyWindow, self).__init__()
            self.setupUi(self)


    if __name__ == "__main__":
        import sys

        app = QtWidgets.QApplication(sys.argv)
        myshow = MyWindow()
        myshow.show()
        sys.exit(app.exec_())
    
这样运行后可以调用GUI

面临问题：

1.如何将GUI上文本框LineEdit中的字符获取到后台

2.如何将"ok"触发运行后台运算的过程

————————————————————————

2019年7月19日

重大突破！！！！

将前端与后端连接起来啦！！

用test.docx做实验模板，进行小demo实验

大成功！

接上次更新，给"ok"设置了一个信号，一个槽（新建槽"add()"）

在代码中的class mywindow 这个层级下进行定义

    class MyWindow(QtWidgets. QMainWindow, Ui_MainWindow):
        def __init__(self):
            super(MyWindow, self).__init__()
            self.setupUi(self)

        def add(self):
            year = self.year.displayText()
            month = self.month.displayText()
            day = self.day.displayText()
            lesson_zh = self.lesson_zh.displayText()
            lesson_en = self.lesson_en.displayText()
            meal_zh = self.meal_zh.currentText()
            protein_1_zh = self.protein_1_zh.displayText()

            from mailmerge import MailMerge

            template = 'test.docx'
            document = MailMerge(template)
            ……

最终实现了前后端的连接

但是新问题：

将完整版代码写入后，会引起崩溃。是否是因为内存原因？
