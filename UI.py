# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'UI.ui'
#
# Created: Mon Jul 04 17:15:38 2016
#      by: PyQt5 UI code generator 5.2
#
# WARNING! All changes made in this file will be lost!

from MainBulletin import mainProcess
from PyQt5.QtWidgets import(QFileDialog, QMainWindow, QGraphicsView, QGraphicsScene,
                             QMessageBox,QDesktopWidget,QSizePolicy,QDialog,QWidget,
                            QVBoxLayout,QHBoxLayout,QPushButton,QSpacerItem,QSpinBox,
                            QProgressBar,QMenuBar,QLabel,QTabWidget,QMenu,QStatusBar,
                            QAction,QTextBrowser,QDateEdit,QComboBox,QApplication)
from PyQt5.QtCore import (pyqtSignal, QObject, QTimer,Qt,QProcess,QDate,QMetaObject,
                          QTime,QRect,QCoreApplication,QDateTime)
from PyQt5.QtGui import QPixmap,QIcon
import time,os
from multiprocessing import Process
from shutil import move,copy


cwd = os.getcwd()
class WorkThread(QObject):
    trigger = pyqtSignal()

    def __init__(self):
        super(WorkThread, self).__init__()
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.timeOut)

    def beginRun(self, in_parameters):
        self.p = Process(target=mainProcess, kwargs=in_parameters)
        self.p.start()
        self.timer.start(1000)

    def timeOut(self):
        if not self.p.is_alive():  # 当前子进程完成后
            self.trigger.emit()  # 发出处理完成信息
            self.timer.stop()  # 并停止计时器，否则在同一进程下再次运行会导致过早判断子进程的状态


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.lastyear = int(time.strftime('%Y', time.localtime(time.time()))) - 1
        self.in_parameters = {u'datetime': str(self.lastyear) + u'年',
                              u'target_area': u'绍兴市',
                              u'density_cell': u'10',
                              u'density_class': 10,
                              u'day_cell': u'15',
                              u'day_class': 10,
                              u'out_type': u'tiff'}
        self.setupUi()

    def setupUi(self):
        self.setObjectName("MainWindow")
        self.setFixedSize(1040, 915)
        sizePolicy = QSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        icon = QIcon()
        icon.addPixmap(QPixmap('./resource/weather-thunder.png'),QIcon.Normal, QIcon.Off)
        self.setWindowIcon(icon)
        self.centralwidget = QWidget(self)
        sizePolicy = QSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.centralwidget.sizePolicy().hasHeightForWidth())
        self.centralwidget.setSizePolicy(sizePolicy)
        self.centralwidget.setObjectName("centralwidget")
        self.layoutWidget = QWidget(self.centralwidget)
        self.layoutWidget.setGeometry(QRect(32, 10, 979, 851))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout_5 =QVBoxLayout(self.layoutWidget)
        self.verticalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        spacerItem = QSpacerItem(300, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.datetime_label = QLabel(self.layoutWidget)
        self.datetime_label.setObjectName("datetime_label")
        self.horizontalLayout.addWidget(self.datetime_label)
        self.datetime = QDateEdit(self.layoutWidget)
        self.datetime.setDateTime(QDateTime(QDate(self.lastyear, 1, 1), QTime(0, 0, 0)))
        self.datetime.setObjectName("datetime")
        self.horizontalLayout.addWidget(self.datetime)
        spacerItem1 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        self.target_area_label = QLabel(self.layoutWidget)
        self.target_area_label.setObjectName("target_area_label")
        self.horizontalLayout.addWidget(self.target_area_label)
        self.target_area = QComboBox(self.layoutWidget)
        self.target_area.setObjectName("target_area")
        self.target_area.addItem("")
        self.target_area.addItem("")
        self.target_area.addItem("")
        self.target_area.addItem("")
        self.target_area.addItem("")
        self.target_area.addItem("")
        self.horizontalLayout.addWidget(self.target_area)
        spacerItem2 = QSpacerItem(300, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem2)
        self.verticalLayout_5.addLayout(self.horizontalLayout)
        self.tabWidget = QTabWidget(self.layoutWidget)
        sizePolicy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tabWidget.sizePolicy().hasHeightForWidth())
        self.tabWidget.setSizePolicy(sizePolicy)
        self.tabWidget.setObjectName("tabWidget")
        self.density_tab = QWidget()
        self.density_tab.setObjectName("density_tab")
        self.verticalLayout_3 =QVBoxLayout(self.density_tab)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.verticalLayout_2 =QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.density_cell_label = QLabel(self.density_tab)
        self.density_cell_label.setObjectName("density_cell_label")
        self.horizontalLayout_2.addWidget(self.density_cell_label)
        self.density_cell = QSpinBox(self.density_tab)
        self.density_cell.setProperty("value", 10)
        self.density_cell.setObjectName("density_cell")
        self.horizontalLayout_2.addWidget(self.density_cell)
        spacerItem3 = QSpacerItem(40, 0, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem3)
        self.density_class_label = QLabel(self.density_tab)
        self.density_class_label.setObjectName("density_class_label")
        self.horizontalLayout_2.addWidget(self.density_class_label)
        self.density_class = QSpinBox(self.density_tab)
        self.density_class.setProperty("value", 10)
        self.density_class.setObjectName("density_class")
        self.horizontalLayout_2.addWidget(self.density_class)
        spacerItem4 = QSpacerItem(478, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem4)
        self.density_mxd = QPushButton(self.density_tab)
        self.density_mxd.setObjectName("density_mxd")
        self.horizontalLayout_2.addWidget(self.density_mxd)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.density_view = QGraphicsView(self.density_tab)
        self.density_view.setObjectName("density_view")
        self.verticalLayout_2.addWidget(self.density_view)
        self.verticalLayout_3.addLayout(self.verticalLayout_2)
        self.tabWidget.addTab(self.density_tab, "")
        self.day_tab = QWidget()
        self.day_tab.setObjectName("day_tab")
        self.verticalLayout_4 =QVBoxLayout(self.day_tab)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.verticalLayout =QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout_3 =QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.day_cell_label = QLabel(self.day_tab)
        sizePolicy = QSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.day_cell_label.sizePolicy().hasHeightForWidth())
        self.day_cell_label.setSizePolicy(sizePolicy)
        self.day_cell_label.setObjectName("day_cell_label")
        self.horizontalLayout_3.addWidget(self.day_cell_label)
        self.day_cell = QSpinBox(self.day_tab)
        self.day_cell.setProperty("value", 15)
        self.day_cell.setObjectName("day_cell")
        self.horizontalLayout_3.addWidget(self.day_cell)
        spacerItem5 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem5)
        self.day_class_label = QLabel(self.day_tab)
        self.day_class_label.setObjectName("day_class_label")
        self.horizontalLayout_3.addWidget(self.day_class_label)
        self.day_class = QSpinBox(self.day_tab)
        self.day_class.setProperty("value", 10)
        self.day_class.setObjectName("day_class")
        self.horizontalLayout_3.addWidget(self.day_class)
        spacerItem6 = QSpacerItem(478, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem6)
        self.day_mxd = QPushButton(self.day_tab)
        self.day_mxd.setObjectName("day_mxd")
        self.horizontalLayout_3.addWidget(self.day_mxd)
        self.verticalLayout.addLayout(self.horizontalLayout_3)
        self.day_view = QGraphicsView(self.day_tab)
        self.day_view.setObjectName("day_view")
        self.verticalLayout.addWidget(self.day_view)
        self.verticalLayout_4.addLayout(self.verticalLayout)
        self.tabWidget.addTab(self.day_tab, "")
        self.verticalLayout_5.addWidget(self.tabWidget)
        self.horizontalLayout_4 =QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.progressBar = QProgressBar(self.layoutWidget)
        sizePolicy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.progressBar.sizePolicy().hasHeightForWidth())
        self.progressBar.setSizePolicy(sizePolicy)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.horizontalLayout_4.addWidget(self.progressBar)
        self.execute_button = QPushButton(self.layoutWidget)
        sizePolicy = QSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.execute_button.sizePolicy().hasHeightForWidth())
        self.execute_button.setSizePolicy(sizePolicy)
        self.execute_button.setObjectName("execute_button")
        self.horizontalLayout_4.addWidget(self.execute_button)
        self.verticalLayout_5.addLayout(self.horizontalLayout_4)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(self)
        self.menubar.setGeometry(QRect(0, 0, 1040, 26))
        self.menubar.setObjectName("menubar")
        self.file_menu = QMenu(self.menubar)
        self.file_menu.setObjectName("file_menu")
        self.help_menu = QMenu(self.menubar)
        self.help_menu.setObjectName("help_menu")
        self.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)
        self.action_add_data = QAction(self)
        self.action_add_data.setObjectName("action_add_data")
        self.action_help = QAction(self)
        self.action_help.setObjectName("action_help")
        self.action_about = QAction(self)
        self.action_about.setObjectName("action_about")
        self.action_save_pic = QAction(self)
        self.action_save_pic.setObjectName("action_save_pic")
        self.file_menu.addAction(self.action_add_data)
        self.file_menu.addAction(self.action_save_pic)
        self.help_menu.addAction(self.action_help)
        self.help_menu.addAction(self.action_about)
        self.menubar.addAction(self.file_menu.menuAction())
        self.menubar.addAction(self.help_menu.menuAction())

        self.retranslateUi()
        self.tabWidget.setCurrentIndex(0)
        QMetaObject.connectSlotsByName(self)
        self.center()
        self.show()

        self.target_area.activated[str].connect(self.updateTargetArea)
        self.datetime.dateChanged.connect(self.updateDatetime)
        self.density_cell.valueChanged.connect(self.updateDensityCell)
        self.density_class.valueChanged.connect(self.updateDensityClass)
        self.day_cell.valueChanged.connect(self.updateDayCell)
        self.day_class.valueChanged.connect(self.updateDayClass)

        self.action_add_data.triggered.connect(self.addData)
        self.action_save_pic.triggered.connect(self.savePic)
        self.action_about.triggered.connect(self.showAbout)
        self.action_help.triggered.connect(self.showHelp)
        self.execute_button.clicked.connect(self.execute)
        self.density_mxd.clicked.connect(self.openMxdDensity)
        self.day_mxd.clicked.connect(self.openMxdDay)


        self.density_mxd.setDisabled(True)
        self.day_mxd.setDisabled(True)
        self.action_save_pic.setDisabled(True)

    def execute(self):
        dir = u"E:/Documents/工作/雷电公报/闪电定位原始文本数据/" + self.in_parameters[u'datetime']

        if os.path.exists(dir):
            datafiles = os.listdir(dir)
            datafiles = map(lambda x:os.path.join(dir,x),datafiles)
            self.in_parameters[u'origin_data_path'] = datafiles

        if not self.in_parameters.has_key(u'origin_data_path'):
            message = u"请加载%s的数据" % self.in_parameters[u'datetime']
            msgBox = QMessageBox()
            msgBox.setText(message)
            msgBox.setIcon(QMessageBox.Information)
            icon = QIcon()
            icon.addPixmap(QPixmap('./resource/weather-thunder.png'), QIcon.Normal, QIcon.Off)
            msgBox.setWindowIcon(icon)
            msgBox.setWindowTitle(" ")
            msgBox.exec_()
            return

        self.execute_button.setDisabled(True)
        self.execute_button.setText(u'正在制图中……')
        self.progressBar.setMaximum(0)
        self.progressBar.setMinimum(0)

        self.action_add_data.setDisabled(True)
        self.target_area.setDisabled(True)
        self.datetime.setDisabled(True)
        self.density_cell.setDisabled(True)
        self.density_class.setDisabled(True)
        self.day_cell.setDisabled(True)
        self.day_class.setDisabled(True)

        # for outfile in self.in_parameters[u'origin_data_path']:
        #     infile =
        #     try:
        #         with open(infile, 'w+') as in_f:
        #             for line in in_f:
        #                 line = line.replace(u"：",":")
        #                 in_f.write(line)
        #     except Exception,inst:
        #         print infile

        self.process_thread = WorkThread()
        self.process_thread.trigger.connect(self.finished)
        self.process_thread.beginRun(self.in_parameters)

    def finished(self):

        #绘制闪电密度图
        ##清除上一次的QGraphicsView对象，防止其记录上一次图片结果，影响显示效果
        self.density_view.setAttribute(Qt.WA_DeleteOnClose)
        self.verticalLayout_2.removeWidget(self.density_view)
        size = self.density_view.size()
        self.density_view.close()

        self.density_view = QGraphicsView(self.density_tab)
        self.density_view.setObjectName("density_view")
        self.density_view.resize(size)
        self.verticalLayout_2.addWidget(self.density_view)

        densityPic = ''.join([cwd,u'/bulletinTemp/',
            self.in_parameters[u'datetime'],u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'闪电密度空间分布.tif'])

        scene = QGraphicsScene()
        pixmap_density = QPixmap(densityPic)
        scene.addPixmap(pixmap_density)
        self.density_view.setScene(scene)
        scale = float(self.density_view.width()) / pixmap_density.width()
        self.density_view.scale(scale, scale)

        #绘制雷暴日图
        self.day_view.setAttribute(Qt.WA_DeleteOnClose)
        self.verticalLayout.removeWidget(self.day_view)
        size = self.day_view.size()
        self.day_view.close()

        self.day_view = QGraphicsView(self.day_tab)
        self.day_view.setObjectName("day_view")
        self.day_view.resize(size)
        self.verticalLayout.addWidget(self.day_view)

        dayPic = ''.join([cwd,u'/bulletinTemp/',
            self.in_parameters[u'datetime'],u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'地闪雷暴日空间分布.tif'])

        pixmap_day = QPixmap(dayPic)
        scene = QGraphicsScene()
        scene.addPixmap(pixmap_day)
        self.day_view.resize(self.density_view.width(),self.density_view.height())
        self.day_view.setScene(scene)
        scale = float(self.day_view.width()) / pixmap_day.width()
        self.day_view.scale(scale, scale)

        #处理进度条和执行按钮状态
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.progressBar.setValue(100)
        self.progressBar.setFormat(u'完成!')
        self.execute_button.setDisabled(False)
        self.execute_button.setText(u'执行')
        #改变一些控件的状态
        self.action_add_data.setDisabled(False)
        self.target_area.setDisabled(False)
        self.datetime.setDisabled(False)
        self.density_cell.setDisabled(False)
        self.density_class.setDisabled(False)
        self.day_cell.setDisabled(False)
        self.day_class.setDisabled(False)
        self.density_mxd.setDisabled(False)
        self.day_mxd.setDisabled(False)
        self.action_save_pic.setDisabled(False)

    def addData(self):
        fnames = QFileDialog.getOpenFileNames(self, u'请选择原始的电闪数据',
                                              u'E:/Documents/工作/雷电公报/闪电定位原始文本数据',
                                              'Text files (*.txt);;All(*.*)')

        self.in_parameters[u'origin_data_path'] = fnames[0]

    def savePic(self):
        densityPic = ''.join([cwd,u'/bulletinTemp/',self.in_parameters[u'datetime'],u'/',
            self.in_parameters[u'target_area'],'.gdb',u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'闪电密度空间分布.tif'])

        dayPic = ''.join([cwd,u'/bulletinTemp/',self.in_parameters[u'datetime'],u'/',
            self.in_parameters[u'target_area'],'.gdb',u'/',self.in_parameters[u'datetime'],
            self.in_parameters[u'target_area'],u'地闪雷暴日空间分布.tif'])

        directory = QFileDialog.getExistingDirectory(self,u'请选择图片保存位置',
                                                     u'E:/Documents/工作/雷电公报',
                                    QFileDialog.ShowDirsOnly|QFileDialog.DontResolveSymlinks)

        dest_density = os.path.join(directory,os.path.basename(densityPic))
        dest_day = os.path.join(directory,os.path.basename(dayPic))

        if os.path.isfile(dest_day) or os.path.isfile(dest_density):
            message = u"文件已经存在！"
            msgBox = QMessageBox()
            msgBox.setText(message)
            msgBox.setIcon(QMessageBox.Information)
            icon = QIcon()
            icon.addPixmap(QPixmap("./resource/weather-thunder.png"), QIcon.Normal, QIcon.Off)
            msgBox.setWindowIcon(icon)
            msgBox.setWindowTitle(" ")
            msgBox.exec_()
            return

        move(dayPic,directory)
        move(densityPic,directory)

    def openMxdDay(self):
        program  = u'C:/Program Files (x86)/ArcGIS/Desktop10.3/bin/ArcMap.exe'
        src_dir = ''.join([cwd,u'/data/LightningBulletin.gdb'])
        dest_dir = ''.join([cwd,u"/bulletinTemp/",self.in_parameters[u'datetime'], u"/" , self.in_parameters[u'target_area']])
        src_file = ''.join([self.in_parameters[u'target_area'] , u"地闪雷暴日空间分布模板.mxd"])

        copy(os.path.join(src_dir,src_file),dest_dir)

        arguments = [os.path.join(dest_dir,src_file)]
        self.process = QProcess(self)
        self.process.start(program,arguments)

    def openMxdDensity(self):
        program  = u'C:/Program Files (x86)/ArcGIS/Desktop10.3/bin/ArcMap.exe'
        src_dir = ''.join([cwd,u'/data/LightningBulletin.gdb'])
        dest_dir = ''.join([cwd,u"/bulletinTemp/",self.in_parameters[u'datetime'], u"/" , self.in_parameters[u'target_area']])
        src_file = ''.join([self.in_parameters[u'target_area'] ,u"闪电密度空间分布模板.mxd"])


        copy(os.path.join(src_dir,src_file),dest_dir)

        arguments = [os.path.join(dest_dir,src_file)]
        self.process = QProcess(self)
        self.process.start(program,arguments)

    def showAbout(self):
        self.about = About_Dialog()

    def showHelp(self):
        program  = u'C:/Windows/hh.exe'
        arguments = [''.join([cwd,'/help/help.CHM'])]
        self.process = QProcess(self)
        self.process.start(program,arguments)

    def updateTargetArea(self, area):
        self.in_parameters[u'target_area'] = area

    def updateDatetime(self, date):
        self.in_parameters[u'datetime'] = str(date.year()) + u'年'
        if self.in_parameters.has_key(u'origin_data_path'):
            self.in_parameters.__delitem__(u'origin_data_path')

    def updateDensityCell(self, cell):
        self.in_parameters[u'density_cell'] = str(cell)

    def updateDensityClass(self, nclass):
        self.in_parameters[u'density_class'] = nclass

    def updateDayCell(self, cell):
        self.in_parameters[u'day_cell'] = str(cell)


    def updateDayClass(self, nclass):
        self.in_parameters[u'day_class'] = nclass

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def retranslateUi(self):
        _translate = QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "绍兴防雷中心 雷电公报制图"))
        self.datetime_label.setText(_translate("MainWindow", "年份"))
        self.datetime.setDisplayFormat(_translate("MainWindow", "yyyy"))
        self.target_area_label.setText(_translate("MainWindow", "地区"))
        self.target_area.setItemText(0, _translate("MainWindow", "绍兴市"))
        self.target_area.setItemText(1, _translate("MainWindow", "柯桥区"))
        self.target_area.setItemText(2, _translate("MainWindow", "上虞区"))
        self.target_area.setItemText(3, _translate("MainWindow", "诸暨市"))
        self.target_area.setItemText(4, _translate("MainWindow", "嵊州市"))
        self.target_area.setItemText(5, _translate("MainWindow", "新昌县"))
        self.density_cell_label.setText(_translate("MainWindow", "插值网格大小"))
        self.density_class_label.setText(_translate("MainWindow", "制图分类数目"))
        self.density_mxd.setText(_translate("MainWindow", "ArcGIS文档"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.density_tab), _translate("MainWindow", "电闪密度"))
        self.day_cell_label.setText(_translate("MainWindow", "插值网格大小"))
        self.day_class_label.setText(_translate("MainWindow", "制图分类数目"))
        self.day_mxd.setText(_translate("MainWindow", "ArcGIS文档"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.day_tab), _translate("MainWindow", "雷暴日"))
        self.execute_button.setText(_translate("MainWindow", "执行"))
        self.file_menu.setTitle(_translate("MainWindow", "文件"))
        self.help_menu.setTitle(_translate("MainWindow", "帮助"))
        self.action_add_data.setText(_translate("MainWindow", "加载数据"))
        self.action_help.setText(_translate("MainWindow", "使用说明"))
        self.action_about.setText(_translate("MainWindow", "关于"))
        self.action_save_pic.setText(_translate("MainWindow", "图片另存为"))

class About_Dialog(QDialog):
    def __init__(self):
        super(About_Dialog, self).__init__()
        self.setupUi()

    def setupUi(self):
        self.setObjectName("Dialog")
        self.setFixedSize(464, 257)
        icon = QIcon()
        icon.addPixmap(QPixmap("./resource/weather-thunder.png"), QIcon.Normal, QIcon.Off)
        self.setWindowIcon(icon)
        self.setSizeGripEnabled(False)
        self.setModal(True)
        self.textBrowser = QTextBrowser(self)
        self.textBrowser.setGeometry(QRect(0, 130, 491, 192))
        self.textBrowser.setObjectName("textBrowser")
        self.label = QLabel(self)
        self.label.setGeometry(QRect(0, 0, 641, 131))
        self.label.setText("")
        self.label.setPixmap(QPixmap("./resource/about.jpg"))
        self.label.setScaledContents(True)
        self.label.setObjectName("label")

        self.retranslateUi()
        QMetaObject.connectSlotsByName(self)
        self.show()

    def retranslateUi(self):
        _translate = QCoreApplication.translate
        self.setWindowTitle(_translate("Dialog", "关于"))
        self.textBrowser.setHtml(_translate("Dialog", "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
"<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
"p, li { white-space: pre-wrap; }\n"
"</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:11pt; font-weight:600;\">绍兴防雷中心 雷电公报制图-v1.0</span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt; font-weight:600;\"><br /></p>\n"
"<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:10pt;\">联系：zhufeng314@163.com</span></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-size:10pt;\"><br /></p>\n"
"<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><br /></p>\n"
"<p align=\"center\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-size:8pt;\">©绍兴防雷中心 All rights reserved</span></p></body></html>"))


if __name__ == "__main__":
    import sys

    app = QApplication(sys.argv)
    ui = MainWindow()
    sys.exit(app.exec_())
