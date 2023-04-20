import sys
import minimalmodbus
import configparser
# import qdarkstyle
# from qdarkstyle.light.palette import LightPalette
import qtmodern.styles
import qtmodern.windows
# import qdarktheme
import serial.tools.list_ports
from PyQt5.QtWidgets import QMessageBox, QApplication, QMainWindow, QLabel, QPushButton, QLineEdit
from PyQt5.QtCore import QTimer, Qt, QEvent,QPropertyAnimation, QPoint
import PyQt5.QtGui as QtGui
from PyQt5.QtGui import QIcon
from Table import NewWindow
import sqlite3


class MainWindow(QMainWindow):
    def __init__(self):

        super().__init__()

        # # 使用浅色主题
        # app.setStyleSheet(qdarkstyle.load_stylesheet(
        #     qt_api='pyqt5', palette=LightPalette()))
        qtmodern.styles.light(app)
        # Apply dark theme.
        # qdarktheme.setup_theme("auto")

        # 初始化温度，电压
        self.previous_temperatures = [0.0] * 16

        self.setWindowTitle("空气沿平板强迫流动对流换热实验V1.0")

        # 创建应用程序图标对象
        app_icon = QIcon("logo.jpg")

        # 设置应用程序图标
        self.setWindowIcon(app_icon)

        # 安装事件过滤器
        self.installEventFilter(self)

        # 连接数据库
        self.conn = sqlite3.connect("experiment.db")
        self.c = self.conn.cursor()

        
        # 创建表格         # 清空数据库
        self.c.execute("""CREATE TABLE IF NOT EXISTS data (
                            id INTEGER PRIMARY KEY,
                            voltage FLOAT,
                            current FLOAT,
                            power FLOAT,
                            pressure INTEGER,
                            inlet_temp FLOAT,
                            temperature1 FLOAT,
                            temperature2 FLOAT,
                            temperature3 FLOAT,
                            temperature4 FLOAT,
                            temperature5 FLOAT,
                            temperature6 FLOAT,
                            temperature7 FLOAT,
                            temperature8 FLOAT,
                            temperature9 FLOAT,
                            temperature10 FLOAT,
                            temperature11 FLOAT,
                            temperature12 FLOAT,
                            temperature13 FLOAT,
                            temperature14 FLOAT,
                            temperature15 FLOAT,
                            temperature16 FLOAT
                            )""")
        self.c.execute("DELETE FROM data")
        self.conn.commit()


        # 获取COM口
        comnum = None
        ports = list(serial.tools.list_ports.comports())
        # 输出所有串口的信息
        for port in ports:
            comnum = list(port)[0]

        # 创建Modbus连接
        self.inst = minimalmodbus.Instrument(comnum, 1)
        self.inst.serial.baudrate = 9600
        self.inst.serial.timeout = 1
        self.inst2 = minimalmodbus.Instrument(comnum, 2)
        self.inst2.serial.baudrate = 9600
        self.inst2.serial.timeout = 1
        self.inst3 = minimalmodbus.Instrument(comnum, 3)
        self.inst3.serial.baudrate = 9600
        self.inst3.serial.timeout = 1
        self.inst4 = minimalmodbus.Instrument(comnum, 4)
        self.inst4.serial.baudrate = 9600
        self.inst4.serial.timeout = 1

        # 添加标题
        title_label = QLabel("空气沿平板强迫流动对流换热实验", self)
        title_label.setFont(QtGui.QFont("黑体", 24))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: red;")
        title_label.setGeometry(0, 0, 960, 50)

        # 创建一个QLabel来显示斜杠
        slash_label = QLabel("/", self)
        slash_label.setStyleSheet("color: white; font-size: 30px;")

        # 设置动画
        animation = QPropertyAnimation(slash_label, b"pos")
        animation.setDuration(1000)
        animation.setLoopCount(-1)
        animation.setStartValue(QPoint(-50, 25))
        animation.setEndValue(QPoint(960, 25))
        animation.start()

        # 插入图片
        self.image_label = QLabel(self)
        self.image_label.move(600, 45)
        self.image_label.resize(500, 400)
        pixmap = QtGui.QPixmap("pic.png").scaled(
            self.image_label.size(), Qt.KeepAspectRatio)
        self.image_label.setPixmap(pixmap)

        # 创建标签
        self.label_power_supply = QLabel(self)
        self.label_power_supply.move(30, 400)
        self.label_power_supply.resize(200, 50)
        self.label_power_supply.setText("直流加热电源：")

        self.label_voltage = QLabel(self)
        self.label_voltage.move(170, 400)
        self.label_voltage.resize(200, 50)
        self.label_voltage.setText("V")

        self.label_current = QLabel(self)
        self.label_current.move(260, 400)
        self.label_current.resize(200, 50)
        self.label_current.setText("A")

        self.label_power = QLabel(self)
        self.label_power.move(350, 400)
        self.label_power.resize(200, 50)
        self.label_power.setText("W")

        self.label_Wind_pressure = QLabel(self)
        self.label_Wind_pressure.move(54, 450)
        self.label_Wind_pressure.resize(200, 50)
        self.label_Wind_pressure.setText("风机风压：")

        self.label_pressure = QLabel(self)
        self.label_pressure.move(170, 450)
        self.label_pressure.resize(200, 50)
        self.label_pressure.setText("Pa")

        self.label_inlet_temp = QLabel(self)
        self.label_inlet_temp.move(232, 500)
        self.label_inlet_temp.resize(200, 50)
        self.label_inlet_temp.setText("来流温度：         ℃")

        self.label_room_temp = QLabel(self)
        self.label_room_temp.move(78, 500)
        self.label_room_temp.resize(200, 50)
        self.label_room_temp.setText("室温：         ℃")

        self.label_temps = []
        for i in range(16):
            label_temp = QLabel(self)
            label_temp.resize(200, 20)
            if i < 8:
                label_temp.move(30, 75 + i * 40)
            else:
                k = i-8
                label_temp.move(230, 75 + k * 40)
            if i < 9:
                label_temp.setText("通道%d：            ℃" % (i+1))
            else:
                label_temp.setText("通道%d：           ℃" % (i+1))
            self.label_temps.append(label_temp)

        # 创建只读输入框
        self.lineedit_voltage_value = QLineEdit(self)
        self.lineedit_voltage_value.move(120, 410)
        self.lineedit_voltage_value.resize(40, 30)
        self.lineedit_voltage_value.setText("0.00")
        self.lineedit_voltage_value.setReadOnly(True)
        self.previous_voltage = float(
            self.lineedit_voltage_value.text())   # 初始化电压

        self.lineedit_current_value = QLineEdit(self)
        self.lineedit_current_value.move(210, 410)
        self.lineedit_current_value.resize(40, 30)
        self.lineedit_current_value.setText("0.00")
        self.lineedit_current_value.setReadOnly(True)

        self.lineedit_power_value = QLineEdit(self)
        self.lineedit_power_value.move(300, 410)
        self.lineedit_power_value.resize(40, 30)
        self.lineedit_power_value.setText("0.00")
        self.lineedit_power_value.setReadOnly(True)

        self.lineedit_pressure_value = QLineEdit(self)
        self.lineedit_pressure_value.move(120, 460)
        self.lineedit_pressure_value.resize(40, 30)
        self.lineedit_pressure_value.setText("0")
        self.lineedit_pressure_value.setReadOnly(True)
        self.previous_pressure = float(
            self.lineedit_pressure_value.text())   # 初始化压强

        self.lineedit_inlet_temp_value = QLineEdit(self)
        self.lineedit_inlet_temp_value.move(300, 510)
        self.lineedit_inlet_temp_value.resize(40, 30)
        self.lineedit_inlet_temp_value.setText("0.0")
        self.lineedit_inlet_temp_value.setReadOnly(True)

        self.lineedit_room_temp_value = QLineEdit(self)
        self.lineedit_room_temp_value.move(120, 510)
        self.lineedit_room_temp_value.resize(40, 30)
        self.lineedit_room_temp_value.setText("0.0")
        self.lineedit_room_temp_value.setReadOnly(True)

        self.lineedit_temperature_values = []
        for i in range(16):
            lineedit_temperature_value = QLineEdit(self)
            if i < 8:
                lineedit_temperature_value.move(90, 70 + i*40)
            else:
                k = i-8
                lineedit_temperature_value.move(290, 70 + k*40)
            lineedit_temperature_value.resize(50, 30)
            lineedit_temperature_value.setText("0.0")
            lineedit_temperature_value.setReadOnly(True)
            self.lineedit_temperature_values.append(lineedit_temperature_value)

        # 创建计时器，定时更新数据
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_data)

        # 创建控制按钮
        self.button = QPushButton(self)
        self.button.setText("开始")
        self.button.move(450, 100)
        self.button.clicked.connect(self.control)
        self.button.setStyleSheet(
            "background-color: rgb(173, 216, 230);")

        # 创建提示字符
        self.label_status = QLabel(self)
        self.label_status.move(400, 160)
        self.label_status.resize(200, 60)
        self.label_status.setText("请点击开始按钮\n开始实验")
        self.label_status.setAlignment(Qt.AlignCenter)
        self.label_status.setFont(QtGui.QFont("黑体", 15))
        self.label_status.setStyleSheet("color: red;")

        # 创建采集按钮
        self.button_collect = QPushButton(self)
        self.button_collect.setText("采集")
        self.button_collect.move(450, 250)
        self.button_collect.clicked.connect(self.collect_data)
        self.button_collect.setEnabled(False)
        self.button_collect.setStyleSheet(
            "background-color:rgb(173, 216, 230);")

        # 创建跳转按钮
        self.button2 = QPushButton(self)
        self.button2.setText("数据查询")
        self.button2.move(450, 325)
        self.button2.clicked.connect(self.open_new_window)
        self.button2.setStyleSheet(
            "background-color:rgb(173, 216, 230);")

        # 创建清除数据按钮
        self.button_clear = QPushButton(self)
        self.button_clear.setText("清除数据")
        self.button_clear.move(450, 400)
        self.button_clear.clicked.connect(self.clear_data)
        self.button_clear.setStyleSheet(
            "background-color:rgb(173, 216, 230);")

        # 设置窗口居中显示
        self.setFixedSize(960, 560)
        screen = QtGui.QGuiApplication.primaryScreen()
        screenGeometry = screen.geometry()
        x = int((screenGeometry.width() - self.width()) / 2)
        y = int((screenGeometry.height() - self.height()) / 4)
        self.move(x, y)

    def open_new_window(self):
        self.setEnabled(False)  # 禁用主界面
        self.new_window = NewWindow(self.conn,self)
        self.new_window.exec_()
        self.setEnabled(True)  # 重新启用主界面

    def update_data(self):
        # 读取Modbus数据
        try:
            data = self.inst.read_registers(0, 2, 4)
            data2 = self.inst2.read_registers(4, 1, 3)
            data3 = self.inst3.read_registers(40, 16, 3)
            data4 = self.inst4.read_registers(40, 2, 3)
        except minimalmodbus.NoResponseError as e:
            print("Modbus通信错误：", e)
            return
        voltage = data[0] * 0.003
        current = data[1] * 0.02
        power = voltage * current
        pressure = data2[0]
        temperatures = []
        for i in range(16):
            t0 = data3[i]
            if t0 == 63488:
                t0 = 0
            t = t0 / 10
            temperatures.append(t)
        inlet = data4[0]       # 来流温度
        room = data4[1]        # 室温
        if inlet== 63488:
            inlet = 0
        if room== 63488:
            room = 0
        inlet = inlet / 10
        room = room / 10


        # 更新界面数据
        self.lineedit_voltage_value.setText("{:.2f}".format(voltage))
        self.lineedit_current_value.setText("{:.2f}".format(current))
        self.lineedit_power_value.setText("{:.2f}".format(power))
        self.lineedit_pressure_value.setText(str(pressure))
        for i in range(16):
            self.lineedit_temperature_values[i].setText(
                "{:.1f}".format(temperatures[i]))
        self.lineedit_inlet_temp_value.setText("{:.1f}".format(inlet))
        self.lineedit_room_temp_value.setText("{:.1f}".format(room))

        # 暂停数据
        if (self.previous_voltage == 0):
            self.previous_voltage = voltage
        print(self.previous_voltage, voltage)
        if (self.previous_pressure == 0):
            self.previous_pressure == pressure
        print(self.previous_pressure)

        # 创建一个新的ConfigParser对象并读取配置文件：
        config = configparser.ConfigParser()
        config.read('config.ini')

        # 读取配置文件中的值
        voltage_threshold = config.getint('Thresholds', 'Voltage')
        pressure_threshold = config.getint('Thresholds', 'Pressure')
        temperature_threshold = config.getint('Thresholds','Temperature')

        # 判断是否稳态
        temperature_stable = [
            abs(temperatures[i] - self.previous_temperatures[i]) < temperature_threshold for i in range(16)]
        # 判断是否同一工况，电压或风压的变化情况，与上一次稳态或者实验开始时对比
        voltage_stable = abs(
            voltage - self.previous_voltage) > voltage_threshold
        pressure_stable = abs(
            pressure - self.previous_pressure) > pressure_threshold
        print(voltage_stable)
        print(pressure_stable)

        # 稳态条件
        if all(temperature_stable) and pressure != 0 and (voltage_stable or pressure_stable):
            # 停止更新数据
            self.timer.stop()
            self.button.setText("开始")
            self.button.setStyleSheet(
                "background-color: rgb(173, 216, 230);")     # 浅蓝色
            self.label_status.setText("已暂停\n请采集数据")
            self.label_status.setStyleSheet("color: red;")
            self.previous_temperatures = [0.0] * 16
            # 更新self.previous_pressure 为稳态时的数据
            self.previous_voltage = voltage
            self.previous_pressure == pressure

            # 提示用户采集数据
            reply = QMessageBox.question(
                self, '提示', "已到达稳态，请采集数据",
                QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                self.button_collect.setEnabled(True)
                self.previous_temperatures = [0.0] * 16

        else:
            # 实时更新 previous_temperatures
            self.previous_temperatures = temperatures.copy()
            # self.previous_voltage = voltage

    # 采集按钮
    def collect_data(self):
        # 将数据存入数据库
        voltage = float(self.lineedit_voltage_value.text())
        current = float(self.lineedit_current_value.text())
        power = float(self.lineedit_power_value.text())
        pressure = int(self.lineedit_pressure_value.text())
        inlet = float(self.lineedit_inlet_temp_value.text())
        temperatures = [
            float(self.lineedit_temperature_values[i].text()) for i in range(16)]

        self.c.execute("""INSERT INTO data (voltage, current, power, pressure, inlet_temp,
                                            temperature1, temperature2, temperature3, temperature4, temperature5, temperature6,
                                            temperature7, temperature8, temperature9, temperature10, temperature11, temperature12,
                                            temperature13, temperature14, temperature15, temperature16)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,?)""",
                       (voltage, current, power, pressure,inlet,
                        temperatures[0], temperatures[1], temperatures[2], temperatures[3], temperatures[4], temperatures[5],
                        temperatures[6], temperatures[7], temperatures[8], temperatures[9], temperatures[10], temperatures[11],
                        temperatures[12], temperatures[13], temperatures[14], temperatures[15]))
        self.conn.commit()

        # 禁用采集按钮，以免重复采集
        self.button_collect.setEnabled(False)
        self.label_status.setText("已采集\n请开始实验")
        self.label_status.setStyleSheet("color: red;")

        # 打开数据表格界面
        self.open_new_window()

    # 拖动时进制数据更新，解决卡顿问题
    def eventFilter(self, obj, event):
        if event.type() == QEvent.NonClientAreaMouseButtonPress:
            # 停止计时器
            self.timer.stop()
        elif event.type() == QEvent.NonClientAreaMouseButtonRelease:
            if self.button.text() == "暂停":
                # 启动计时器
                self.timer.start(100)
        return super().eventFilter(obj, event)
        

    # 开关按钮
    def control(self):
        if self.timer.isActive():
            self.timer.stop()
            self.button.setText("开始")
            self.label_status.setText("已暂停")
            self.label_status.setStyleSheet("color: red;")
            self.button.setStyleSheet(
                "background-color: rgb(173, 216, 230);")      # 浅蓝色

        else:
            self.timer.start(100)
            self.button.setText("暂停")
            self.label_status.setText("已开始")
            self.label_status.setStyleSheet("color: blue;")
            self.button.setStyleSheet(
                "background-color: rgb(255, 192, 203);")      # 浅红色

    # 删除数据
    def clear_data(self):
        reply = QMessageBox.question(
            self, '提示', "确认清除所有数据？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.c.execute("DELETE FROM data")
            self.conn.commit()
            QMessageBox.information(
                self, '提示', "已清除所有数据", QMessageBox.Ok)

    # 关闭页面删除数据
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Message',
                                     "确认退出程序？", QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.conn.execute("DELETE FROM data")
            self.conn.commit()
            event.accept()
        else:
            event.ignore()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
