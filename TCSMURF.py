#!/usr/bin/python3
# -*- coding: UTF-8 -*-
import sys
import os
import time
import datetime
import win32com
import socket
from win32com.shell import shell
from win32com.shell import shellcon
from win32com.client import Dispatch
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QWidget, QDesktopWidget, QLabel, QComboBox, QGridLayout, \
    QPushButton, QHBoxLayout, QVBoxLayout, QCheckBox


class Window(QWidget):
    def __init__(self):
        self.cur_time = time.strftime('%m/%d/%Y', time.localtime(time.time()))
        self.cfgfile = r'\\s46file1.cd.intel.com\sdx_eng\sdx_eng\System\Yujiao\TCSMURF\TC_Smurf_Footprint.xlsx'
        self.could_path = r'\\s46file1.cd.intel.com\sdx\event\SDxLogs\User\shaohaoh\Conversion\Xiu2.csv'
        # self.tclog_path = r'E:\TCLogs\PTO.OperationalLog'
        self.tclog_path = r'\\s46file1.cd.intel.com\sdx_eng\sdx_eng\System\Yujiao\TCSMURF\TCLogs'  # local test
        self.local_path = os.getcwd() + '\TCSMURF'
        self.tester_Log_path = r'\d$\HDMT3\logs'
        self.user_name = "SysC"
        self.pwd = "tr@nsf3r"

        # load cfg
        self.alarm_type_list = self.load_cfg()
        self.title_slt = ''
        self.src_day_slt = ''
        self.alarm_time_slt = ''
        super().__init__()
        self.initUI()

    def initUI(self):

        # QGridLayout for AlarmType/StartTime/EndTime
        title = QLabel('AlarmType')
        sch_day = QLabel('SearchDays')
        ala_time = QLabel('AlarmTime')

        self.title_combo = QComboBox()
        for alarm_type in self.alarm_type_list:
            self.title_combo.addItem(str(alarm_type).split(', ')[0])

        sch_cb = QCheckBox()
        sch_cb.setToolTip('Enable/Disable search day')

        self.sch_combo = QComboBox()
        for day in range(1, 20):
            self.sch_combo.addItem(str(day))

        self.ala_combox = QComboBox()

        grid = QGridLayout()
        grid.setSpacing(1)
        grid.addWidget(title, 1, 0)
        grid.addWidget(self.title_combo, 1, 1, 1, 2)

        grid.addWidget(sch_day, 2, 0)
        grid.addWidget(self.sch_combo, 2, 1, 1, 2)
        # grid.addWidget(sch_cb, 2, 3)

        grid.addWidget(ala_time, 3, 0)
        grid.addWidget(self.ala_combox, 3, 1, 1, 2)

        alarm_type = self.title_combo.currentText()
        for type_item in self.alarm_type_list:
            # print(type_item)
            if alarm_type in str(type_item):
                keywords = str(type_item).split(', ')[1]
                break
        alarm_info_list = self.init_alarm_time(alarm_type, keywords)

        self.ala_combox.addItems(alarm_info_list)

        # OK/Quit Button
        ok_btn = QPushButton('OK', self)
        ok_btn.resize(ok_btn.sizeHint())
        ok_btn.clicked.connect(self.buttonClicked)

        qbtn = QPushButton('Quit', self)
        qbtn.resize(qbtn.sizeHint())

        # UI layout
        hbox = QHBoxLayout()
        hbox.addStretch(1)
        hbox.addWidget(ok_btn)
        hbox.addWidget(qbtn)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox)

        glayout = QtWidgets.QVBoxLayout()
        grid_wg = QtWidgets.QWidget()
        btn_wg = QtWidgets.QWidget()
        grid_wg.setLayout(grid)
        btn_wg.setLayout(vbox)

        glayout.addWidget(grid_wg)
        glayout.addWidget(btn_wg)

        self.setLayout(glayout)
        # refresh AlarmTime according to AlarmType and SearchDays, default config = Start Tray + 1 day(SearchDays)
        self.title_combo.currentIndexChanged.connect(self.combox_changed)
        self.sch_combo.currentIndexChanged.connect(self.combox_changed)

        # set window location and size
        self.resize(400, 100)
        self.center()
        # set window title
        hostname = socket.gethostname()
        link = hostname[3:9].upper()
        self.setWindowTitle(link + " - TC SMURF")
        # set window icon
        self.setWindowIcon(QtGui.QIcon('TCSmurf_draft.png'))
        self.show()

    def load_cfg(self):
        common_obj = common_op('False')
        open_excel = common_obj.open_excel(self.cfgfile)
        result = common_obj.get_sheet('cfg', open_excel)
        common_obj.close_excel(open_excel)
        return result

    def center(self):
        # get window
        qr = self.frameGeometry()
        # get screen center
        cp = QDesktopWidget().availableGeometry().center()
        # move window to screen center
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def init_alarm_time(self, alarm_type, keywords, srcday=1):
        alarm_info_list = []
        common_obj = common_op(False)
        # print("inti_alarm_time-srcday=", srcday)
        file_list = os.listdir(self.tclog_path)
        file_list.sort(reverse=True)
        for file in file_list:
            file_path = self.tclog_path + '\\' + file
            last_mdtime = common_obj.get_fileModifyTime(file_path)
            if (common_obj.diff_time_cur(last_mdtime) <= int(srcday) * 24 * 3600 and '.txt' in file) \
                    or 'OperationalLogCache.xml' in file:
                copy_file = self.local_path + '\\' + file
                common_obj.download_file(file_path, copy_file)
                alarm_info_list += common_obj.get_alarm_time(copy_file, alarm_type, keywords)

        return alarm_info_list

    def combox_changed(self):
        self.get_ui()
        # print('title_slt={0}'.format(self.title_slt))
        # print('src_day_slt={0}'.format(self.src_day_slt))
        # print('alarm_time_slt={0}'.format(self.alarm_time_slt))
        for type_item in self.alarm_type_list:
            # print(type_item)
            if self.title_slt in str(type_item):
                keywords = str(type_item).split(', ')[1]
                break
        alarm_info_list = self.init_alarm_time(self.title_slt, keywords, self.src_day_slt)

        self.ala_combox.clear()
        self.ala_combox.addItems(alarm_info_list)

    def get_ui(self):
        self.title_slt = self.title_combo.currentText()
        self.src_day_slt = self.sch_combo.currentText()
        self.alarm_time_slt = self.ala_combox.currentText()
        # print(self.title_slt + "\\" + self.src_day_slt + '\\' + self.alarm_time_slt)

    def buttonClicked(self):
        # sender = self.sender()
        common_obj = common_op(False)
        # try:
        self.get_ui()
        if 'EOT' in self.title_slt.upper():
            alarm_date = self.alarm_time_slt.split(',')[0].strip()
            slot = self.alarm_time_slt.split(',')[1].strip()
            common_obj.get_tester_cell(alarm_date, slot, self.local_path)
            tester_ip = common_obj.to_ip(slot)
            remote_com_path = "\\\\" + tester_ip + self.tester_Log_path + "\\commonhdmt"
            # map network
            # common_obj.map_network(remote_com_path, self.user_name, self.pwd, True)
            EOT_obj = EOT(alarm_date, slot)
            EOT_obj.new_get_ComLog(remote_com_path, alarm_date)
        elif 'Start Tray' in self.title_slt:
            print("--------------------------------")
            print("Add code here for start tray....")
            print("--------------------------------")


# common operations for different alarm
class common_op():
    def __init__(self, visibility):
        self.excelApp = win32com.client.Dispatch('Excel.Application')
        self.excelApp.DisplayAlerts = visibility

    def open_excel(self, file_path):
        return self.excelApp.Workbooks.Open(file_path)

    def close_excel(self, open_file):
        open_file.Close(SaveChanges=1)

    def get_sheet(self, sheet, open_file):
        sht = open_file.Worksheets(sheet)
        if 'cfg'.lower() in sheet.lower():
            row = 2
            col = 1
        cell_value = sht.Cells(row, col).Value
        reslist = []
        while cell_value:
            keywords = sht.Cells(row, col + 1).Value
            reslist.append(cell_value + ', ' + keywords)
            row += 1
            cell_value = sht.Cells(row, col).Value
        return reslist

    def get_cell(self, sheet, row, col, open_file):
        sht = open_file.Worksheets(sheet)
        return sht.Cells(row, col).Value

    def TimeStampToTime(self, timestamp):
        timeStruct = time.localtime(timestamp)
        return time.strftime('%Y/%m/%d %H:%M:%S', timeStruct)

    def get_fileCreateTime(self, filePath):
        t = os.path.getctime(filePath)
        return self.TimeStampToTime(t)

    def get_fileModifyTime(self, file_path):
        t = os.path.getmtime(file_path)
        return self.TimeStampToTime(t)

    # return seconds = now - str_time
    def diff_time_cur(self, str_time):
        cur_time = time.strftime('%Y/%m/%d %H:%M:%S', time.localtime(time.time()))
        now = datetime.datetime.strptime(cur_time, '%Y/%m/%d %H:%M:%S')
        format_time = datetime.datetime.strptime(str_time, '%Y/%m/%d %H:%M:%S')
        delta = now - format_time
        diff_time = delta.days * 24 * 3600 + delta.seconds
        return diff_time

    # return seconds, days = file_time - alarm_time
    def sub_Time(self, file_time, alarm_time):
        then = datetime.datetime.strptime(alarm_time, '%Y/%m/%d %H:%M:%S')
        now = datetime.datetime.strptime(file_time, '%Y/%m/%d %H:%M:%S')
        delta = now - then
        sub_time = delta.days * 24 * 3600 + delta.seconds
        return sub_time

    def download_file(self, src_path, des_path):
        shell.SHFileOperation(
            (0, shellcon.FO_COPY, src_path,
             des_path,
             shellcon.FOF_NOCONFIRMATION + shellcon.FOF_NOCONFIRMMKDIR + shellcon.FOF_ALLOWUNDO, None,
             None)
        )

    def comp(self, x, y, path):
        x = path + "\\" + x
        y = path + "\\" + y
        x_mtime = self.get_fileModifyTime(x)
        y_mtime = self.get_fileModifyTime(y)

        if x_mtime < y_mtime:
            return 1
        elif x_mtime > y_mtime:
            return -1
        else:
            return 0


    # get alarm time and cell id according to alarm type
    def get_alarm_time(self, log_file, alarm_type, keywords):
        keywords = '\'' + keywords + '\' changed to On'
        alarm_inf_list = []
        if 'EOT' in str(alarm_type).upper():
            with open(log_file) as log_content:
                for line in log_content:
                    if keywords in line:
                        # print(line)
                        alarm_time = line.split('AlarmComponent:')[0]
                        shelf = line.split('\'')[1][0:4]
                        alarm_inf_list.append(alarm_time + ', ' + shelf)
        return alarm_inf_list

    # get tester id and cell id according to slot&alarm_time in EIB log
    # keywords: SystemConfigurationRequestCompleted
    # def get_tester_cell(self, alarm_time, slot, local_path):
    #     # eib_path = r"E:\TCLogs\EIB"
    #     # local test
    #     eib_path = r"\\s46file1.cd.intel.com\sdx_eng\sdx_eng\System\Yujiao\TCSMURF\TCLogs\EIB"
    #     file_list = os.listdir(eib_path)
    #     file_list.sort(reverse=True)
    #     for file in file_list:
    #         file_path = eib_path + "\\" + file
    #         copy_file = local_path + '\\' + file
    #         file_mtime = self.get_fileModifyTime(file_path)
    #         sub_time = self.sub_Time(file_mtime, alarm_time)
    #         if 0 <= sub_time:
    #             self.download_file(file_path, copy_file)
    #             with open(copy_file) as log_conetent:
    #                 for line in log_conetent:
    #                     if



    # return tester ip ("10.250.0.1~20") according to slot(A101~E401)
    def to_ip(self, slot):
        post_base = 0
        if "A" in slot.upper():
            post_base = 0
        elif "B" in slot.upper():
            post_base = 4
        elif "C" in slot.upper():
            post_base = 8
        elif "D" in slot.upper():
            post_base = 12
        elif "E" in slot.upper():
            post_base = 16
        post_ip = post_base + int(slot[1])
        ip = "10.250.0." + str(post_ip)
        # print("ip{0}=".format(ip))
        return ip

    def map_network(self, network, user_name, pwd, flag):
        # if flag = True: map network, else(False): delete network
        if flag:
            os.system('net use ' + network + ' \"' + pwd + '\" /user:' + user_name + " /persistent:no")
        else:
            os.system('net use /delete /y ' + network)

    def download_Log(self, cloud_path, local_path):
        component_path = local_path + '\component_temp.csv'
        self.download_file(cloud_path, component_path)


class EOT():
    def __init__(self, alarm_date, slot):
        self.currentTime = time.strftime('%m/%d/%Y', time.localtime(time.time()))
        self.eot_folder = os.getcwd() + 'TCSMURF\EOT\\' + self.currentTime.replace('/', '.')
        self.alarm_date = alarm_date
        self.slot = slot
        self.cellID = None
        self.HDMTid = None
        self.XIU = None
        self.before_down_die_id = None
        self.before_down_die_msg = None
        self.down_die_id = None
        self.com_log_msg = None
        self.com_log_rc = None
        self.TP = None
        self.TOS_version = None
        self.SOC = None
        self.up_time = None
        self.up_info = None
        self.detail = None

    # def new_get_ComLog(self, remote_com_path):
    #     try:
    #         range_time = 25 * 3600
    #         near_file_list = []
    #         common_obj = common_op(False)
    #         for log in remote_com_path:
    #             log_path = remote_com_path + "\\" + log
    #             log_name = log
    #             log_mtime = common_obj.get_fileModifyTime(log_path)
    #             sub_time = common_obj.sub_Time(log_mtime, self.alarm_date)
    #             if 0 <= sub_time <= range_time:
    #                 near_file_list.append(log_path, log_name)
    #
    #         # copy common log to local eot\comLog folder
    #         if len(near_file_list) != 0:
    #             for comlog in near_file_list:
    #                 file_path = log_path + "\\" + comlog
    #                 copy_file = self.eot_folder + "\\" + comlog
    #                 common_obj.download_file(file_path, copy_file)
    #                 Flag_comlog = self.new_find_comlog_infor(self.eot_folder + '\comLog' + "\\" + self.HDMTid)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    UI = Window()
    # UI.get_ui()
    # print(UI.title_slt)
    # print(UI.src_day_slt)
    # print(UI.alarm_time_slt)
    sys.exit(app.exec_())
