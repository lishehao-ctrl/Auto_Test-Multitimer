# 运行环境
# cd C:\Users\15038\Desktop\HardWare\mm_report 
# python -m venv venv
# .\venv\Scripts\Activate.ps1
# python issues.py
# pip install requests PyGithub pyinstaller
# pyinstaller --onefile --name mm_test.exe equips_v0.py

"""equip driver Module deal with base equips operation.
  new equips are encouraged to be inherited from bATEinst_base
"""

import time
import traceback
from datetime import datetime, timedelta, timezone
import re
import struct
import math
import os
import sys
import scipy.interpolate as intpl 
from pyvisa import constants as pyconst
import pyvisa as visa
from pyvisa import errors
import serial
import numpy as np
from scipy.io import savemat,loadmat
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import StringVar
from tkinter import filedialog
import scipy.io as sio
import ctypes
from ctypes import wintypes
import tkinter.font as font
from tkinter import filedialog, scrolledtext

class bATEinst_Exception(Exception):pass

class bATEinst_base(object):
    Equip_Type = "None"
    Model_Supported = ["None"]
    bATEinst_doEvent = None
    isRunning = False
    RequestStop = False
    VisaRM = None   
    
    def __init__(self, name=""):
        self.Name = name
        self.VisaAddress = None
        self.Inst = None

    def __del__(self):
        self.close()

    def set_error(self, ss):
        raise(bATEinst_Exception("Equip %s error:\n%s" % (self.Name, ss)))
        
    @staticmethod
    def open_VisaRM():
        if not bATEinst_base.VisaRM:
            bATEinst_base.VisaRM = visa.ResourceManager()
        return bATEinst_base.VisaRM
    
    def isvalid(self):
        return True if self.VisaAddress else False
    
    def inst_open(self):
        if not self.Inst:
            if not self.VisaAddress:
                self.set_error("Equip Address has not been set!")
            try:
                self.Inst = bATEinst_base.open_VisaRM().open_resource(self.VisaAddress)
            except:
                time.sleep(2)
                if not self.Inst:
                    self.Inst = bATEinst_base.open_VisaRM().open_resource(self.VisaAddress)
        return self.Inst
        
    def inst_close(self):
        if self.Inst:
            try:
                self.Inst.close()
            except:
                pass
            self.Inst = None
            
    def callback_after_open(self): pass
    
    def set_visa_timeout_value(self,tmo):
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_TMO_VALUE,tmo)
        
    def check_open(self):
        if not self.Inst:
            try:
                self.inst_open()
            except Exception as e:
                self.set_error("Can not open address:%s\ninfo:%s" %(str(self.VisaAddress),str(e)))
            self.callback_after_open()
        return self.Inst
        
    def close(self):
        try:
            if self.Inst:
                self.inst_close()
        except:
            pass
        self.Inst = None
    
    def read(self):
        self.check_open()
        try:
            ss = self.Inst.read()
        except Exception as e:
            self.set_error("read error\n info:" +str(e))
        return ss
    
    def write(self, ss):
        self.check_open()
        try:
            if isinstance(ss,list):
                for k in ss:
                    self.Inst.write(ss)
            else:
                self.Inst.write(ss)
        except Exception as e:
            self.set_error("Write error\n info:" +str(e))

    def query(self, ss):
        self.write(ss)
        return self.read()

    def write_raw(self, vv: list):
        self.check_open()
        if isinstance(vv, list):
            vv = bytes(vv)
        self.Inst.write_raw(vv)

    def read_raw(self, n: int):
        self.check_open()
        ss = self.Inst.read_bytes(n)
        return ss

    def write_block(self, v):
        self.write_raw(list(("#8%08d" % len(v)).encode()) + v)

    def read_block(self,cmd=None):
        if cmd:
            self.write(cmd)
        ss = self.read_raw(2)
        if ss[0] != b'#'[0]:
            self.set_error("Equip read block error")
        sz = self.read_raw(int(ss[1])-48)
        n = int(bytes(sz).decode())
        return self.read_raw(n)
        
    def delay (self, sec):
        time.sleep(sec)
        
    def x_write(self, vvs, chx=""):
        if isinstance(vvs, str):
            vvs = vvs.splitlines()
        res = []
        for cc in vvs:
            cc = cc.strip()
            if not cc:
                continue
            cc.replace(r"\$CHX\$", chx)
            if re.match(r"\$WAIT *= *(\d+) *\$",cc):
                self.delay(int(re.match(r"\$WAIT *= *(\d+) *\$", cc).group(1))/1000)
            else:
                if cc.find("?") >= 0:
                    res += [self.query(cc)]
                else:
                    self.write(cc)
        return res

    def is_number(self, str):
        try:
            if str=='NaN':
                return False
            float(str)
            return True
        except ValueError:
            return False
     
    # fn_relative change the fn to path related to the current path
    def fn_relative(self, fn,sub_folder=None):
        if os.path.isabs(fn):
            return fn
        else:
            # 关键修改：区分打包环境和开发环境
            if getattr(sys, 'frozen', False):
                # 打包后：获取EXE所在目录
                hd = os.path.dirname(sys.executable)
            else:
                # 开发环境：原逻辑
                hd, _ = os.path.split(os.path.realpath(__file__))
            
            # 拼接完整路径
            if sub_folder is None:
                fn_full = os.path.realpath(os.path.join(hd, fn))
            else:
                fn_full = os.path.realpath(os.path.join(hd, sub_folder, fn))
            
            # 确保目录存在（新增：避免路径不存在导致保存失败）
            os.makedirs(os.path.dirname(fn_full), exist_ok=True)
            return fn_full

    def get_filelist(self, fpp,modes=".py" ):
        Filelist = []
        for home, dirs, files in os.walk(fpp ):
            for filename in files:
                if filename.lower().endswith(modes):
                    Filelist.append(os.path.join(home, filename))
        return Filelist

    def load_cal_cable_loss(self,fn,freq_unit_rate=1e6, domain='V'):        
        """ freq_unit_rate: used to make sure all units is consistent inside the driver,default 1e6
        """
        if isinstance(freq_unit_rate, str):
            freq_unit_rate = {"MHz":1e6,"KHz":1e3,"Hz":1}[freq_unit_rate]
            
        if self.is_number(fn):
            freq = [0,1e9]
            loss = [float(fn)]*2 #[10**(float(fn/20))]*2
        else:
            with open(self.fn_relative(fn,"calibration"), "rt") as fid:
                xys = fid.readlines()
            freq = [float(k.split("\t")[0])*freq_unit_rate for k in xys]
            loss = [float(k.split("\t")[1]) for k in xys] 
            loss = [10**(k/20) for k in loss] if domain == 'V' else loss
        return intpl.interp1d(freq,loss,bounds_error=False, fill_value="extrapolate")
    
    def load_matfile(self, fn):
        return loadmat(fn)
    
    def save_matfile(self, fn, mm):        
        try:
            # 先检查mm的数据类型（新增：提前排查格式问题）
            # self._check_mat_data(mm)
            # 执行保存
            savemat(fn, mm)
            print(f"mat文件保存成功：{fn}")  # 明确提示成功
            return True
        except Exception as e:
            # 打印详细错误（关键：定位问题）
            print(f"mat文件保存失败：{str(e)}")
            print("错误堆栈：\n", traceback.format_exc())
            return False
        
    def _check_mat_data(self, data_dict):
        for key, value in data_dict.items():
            if isinstance(value, (list, np.ndarray)):
                # 检查列表/数组是否有混合类型
                types = set(type(x) for x in value) if isinstance(value, list) else {value.dtype.type}
                if len(types) > 1:
                    raise ValueError(f"数据'{key}'包含混合类型：{types}，无法保存为mat文件")
class instAWG(bATEinst_base):
    Equip_Type = "awg"
    AWG_MODE_DC = "DC"
    AWG_MODE_SIN = "SIN"

    def set_output(self, on=True):
        self.set_error("Function not implemented")

    def set_amplitude(self, v):
        self.set_error("Function not implemented")

    def set_offset(self, v):
        self.set_error("Function not implemented")

    def set_dc(self, v):
        self.set_mode("DC")
        self.set_offset(v)

    def set_mode(self, mode):
        self.set_error("Function not implemented")

    def set_freq(self, freq):
        self.set_error("Function not implemented")

    def set_impedance(self, z):
        self.set_error("Function not implemented")

    def set_offset_quick(self, v):
        self.set_offset(v)

class instMultimeter(bATEinst_base):
    Equip_Type = "mm"

    MM_MODE_V  = "V"
    MM_MODE_I  = "I"
    MM_MODE_R  = "R"
    MM_MODE_R4 = "R4"
    MM_RANGE_AUTO = "AUTO"

    MM_AC = "AC"
    MM_DC = "DC"

    def __init__(self, name=""):
        super().__init__(name)
        self.current_mode = None
        self.current_mode_for_equip = None
        self.current_ac_dc = None
        self.current_range = None

    def set_mode(self, mode=MM_MODE_V, ac_dc = MM_AC):
        if len(mode) <= 2:
            mode = "VOLT" if mode == 'V' else "CURR" if mode == 'I' else None
            if mode == None: raise ValueError('模式不符合要求')
        self.x_write(["CONF:%s:%s" %(mode, ac_dc), "*OPC?"] )
        self.current_mode = mode
        self.current_ac_dc = ac_dc

    def measure(self):
        return float(self.x_write(f"MEAS:{self.current_mode}:{self.current_ac_dc}? {self.current_range}")[0])

    def measure_quick(self):
        return self.measure()

    def set_range(self, rng=MM_RANGE_AUTO):
        self.x_write([f"CONF:{self.current_mode}:{self.current_ac_dc} {rng}", "*OPC?"])
        self.current_range = rng

    def set_speed(self, speed):
        self.set_error("Function not implemented")

    def capture_waveform(self):
        self.set_error("Function not implemented")
        return []

    def measure_i(self):
        self.set_mode(self.MM_MODE_I)
        return self.measure()

    def measure_v(self):
        self.set_mode(self.MM_MODE_V)
        return self.measure()

    def measure_r(self):
        self.set_mode(self.MM_MODE_R)
        return self.measure()

class instKS_34461A(instMultimeter):
    def __init__(self, name="", visa_address=""):
        super().__init__(name)
        # self.VisaAddressSocket = f"TCPIP::{self.ip_address}::5025::SOCKET" 
        # self.VisaAddressUsb    = "USB::0x2A8D::0x1301::MY60055883::INSTR"
        self.VisaAddress   = visa_address
        self.sleep_time    = None
        self.time_dur      = None
        self.time_dur_unit = None

class UI(tk.Tk):
    # label of users' input data types
    data_type_mode          = "Mode"
    data_type_ac_dc         = "直流电交流电"
    data_type_range         = "Range"
    data_type_sleep_time    = "间隔时间"
    data_type_time_dur      = "监测时间"
    data_type_time_dur_unit = "时长单位"
    data_type_usb_lan       = "USB/LAN"
    data_type_visa_address  = "Visa Address"
    data_type_user_input_btn_process = "结束"

    # error message templates
    user_input_non_positive_alert    = "请输入正数！！！"
    user_input_wrong_type_alert      = "请输入数字！！！"
    user_input_miss_sleep_time_alert = "请输入触发间隔时间！！！"
    user_input_miss_time_dur_alert   = "请输入监测时长！！！"
    user_input_miss_visa_address     = "请输入Visa地址！！！"
    user_input_miss_ip_address       = "请输入ip地址！！！"

    # default values for functions
    default_mode          = "VOLT"
    default_ac_dc         = "AC"
    default_range         = "AUTO"
    default_time_dur_unit = "秒"
    default_usb_lan       = "USB"
    default_fn            = "Test_File.mat"
    show_selection_text_font = ("Microsoft YaHei UI", 18)
    default_text_font        = ("Microsoft YaHei UI", 10)
    update_frequency         = 100

    # message
    lable_for_show_selection   = "你选中了："
    lable_for_mode_input       = "请选择想要测量的Mode"
    lable_for_ac_dc_input      = "请选择想要测量的AC/DC"
    lable_for_range_input      = "请选择想要测量的Range"
    lable_for_sleep_time_input = "请输入想要的触发间隔时间(秒)"
    lable_for_time_dur_input   = "请输入想要的监测时间(秒/分/时)"
    lable_for_usb_lan          = "请选择设备连接模式"
    lable_for_visa_address     = "请输入设备visa地址"
    lable_for_ip_address       = "请输入设备ip地址   "
    label_for_filedialog_title = "选择保存路径和文件名"

    # mappings
    AC = "AC"
    DC = "DC"
    VOLTAGE = "VOLT"
    CURRENT = "CURR"
    time_unit_second = "秒"
    time_unit_minute = "分钟"
    time_unit_hour   = "小时"
    usb = "USB"
    lan = "LAN"
    auto_detect = "自动识别"

    def __init__(self):
        super().__init__()
        self.change_font()
        self.screen_wdith = self.winfo_screenwidth()
        self.screen_height = self.winfo_screenheight()
        self.title("输入控制界面")

        fn = self.default_fn
        default_filepath = instKS_34461A.fn_relative(self, fn=fn)
        dir_path = os.path.dirname(default_filepath)
        new_file_name = datetime.now().strftime("%Y%m%d_%H_%M_%S") + "_" + fn
        self.file_path = os.path.join(dir_path, new_file_name)

    def change_font(self):
        self.default_font = font.nametofont("TkDefaultFont")
        self.default_font.configure(family=self.default_text_font[0], size=self.default_text_font[1])
        self.text_font = font.nametofont("TkTextFont")
        self.text_font.configure(family=self.default_text_font[0], size=self.default_text_font[1])

    def get_insts(self):
        rm = visa.ResourceManager()
        insts = rm.list_resources()
        return insts

    # ui to get ip address
    def generat_ui(self):
        self.usb_lan = tk.StringVar(value = self.default_usb_lan)
        self.var_usb_visa_address = tk.StringVar()
        self.var_visa_address = None

        self.lb_usb_lan = tk.Label(self, text = self.lable_for_usb_lan)
        self.lb_usb_lan.pack(anchor=tk.W)

        self.frame_usb_lan = tk.Frame(self)
        self.frame_usb_lan.pack(side=tk.LEFT, pady=5)

        self.frame_usb = tk.Frame(self)
        self.frame_usb.pack(anchor=tk.W)
    
        self.rb_btn_usb = tk.Radiobutton(self.frame_usb, text=self.auto_detect, variable=self.usb_lan, value=self.usb, command=lambda: self.show_selected(self.data_type_usb_lan))
        self.rb_btn_usb.pack(side=tk.LEFT)
        
        self.lb_usb_visa_address = tk.Label(self.frame_usb, text=self.lable_for_visa_address)
        self.lb_usb_visa_address.pack(side=tk.LEFT)

        self.cmb_usb_visa_address = ttk.Combobox(self.frame_usb, textvariable=self.var_usb_visa_address, width=40)
        self.cmb_usb_visa_address['values'] = self.get_insts()
        self.cmb_usb_visa_address.pack(side=tk.LEFT, padx=10)
        if self.get_insts():  # 确保列表非空
            self.cmb_usb_visa_address.set(self.get_insts()[0])

        self.btn_refresh = tk.Button(self.frame_usb, text="刷新", command=self.refresh_insts)
        self.btn_refresh.pack(side=tk.LEFT)

        self.frame_lan = tk.Frame(self)
        self.frame_lan.pack(anchor=tk.W)

        self.rb_btn_lan = tk.Radiobutton(self.frame_lan, text=self.lan, variable=self.usb_lan, value=self.lan, command=lambda: self.show_selected(self.data_type_usb_lan))
        self.rb_btn_lan.pack(side=tk.LEFT)

        self.lb_lan_visa_address = tk.Label(self.frame_lan, text=self.lable_for_ip_address)
        self.lb_lan_visa_address.pack(side=tk.LEFT)

        self.txt_lan_visa_address_1 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_1.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_2 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_2.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_3 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_3.pack(side=tk.LEFT, padx=(10, 0))
        self.lb_lan_visa_address_1 = tk.Label(self.frame_lan, text=".")
        self.lb_lan_visa_address_1.pack(side=tk.LEFT)
        self.txt_lan_visa_address_4 = tk.Text(self.frame_lan, width=10, height=1)
        self.txt_lan_visa_address_4.pack(side=tk.LEFT, padx=(10, 0))

        # data collections for users' input
        self.var_mode = tk.StringVar(value = self.default_mode)
        self.var_ac_dc = tk.StringVar(value = self.default_ac_dc)
        self.var_range = tk.StringVar(value = self.default_range)
        self.var_sleep_time = None
        self.var_time_dur = None
        self.var_time_dur_unit = tk.StringVar(value = self.default_time_dur_unit)

        self.frame_mode_text = tk.Frame(self)
        self.frame_mode_text.pack(anchor=tk.W, pady=5)

        self.lb_mode = tk.Label(self.frame_mode_text, text=self.lable_for_mode_input)
        self.lb_mode.pack(side=tk.LEFT)

        self.frame_mode = tk.Frame(self)
        self.frame_mode.pack(pady = 5, anchor=tk.W)

        self.rd_btn_mode_1 = tk.Radiobutton(self.frame_mode, text=self.VOLTAGE, variable=self.var_mode, value=self.VOLTAGE, command=lambda: self.show_selected(self.data_type_mode))
        self.rd_btn_mode_1.pack(side=tk.LEFT)
        self.rd_btn_mode_2 = tk.Radiobutton(self.frame_mode, text=self.CURRENT, variable=self.var_mode, value=self.CURRENT, command=lambda: self.show_selected(self.data_type_mode))
        self.rd_btn_mode_2.pack(side=tk.LEFT)

        self.lb_ac_dc = tk.Label(self, text = self.lable_for_ac_dc_input)
        self.lb_ac_dc.pack(pady = 10, anchor=tk.W)

        self.frame_ac_dc = tk.Frame(self)
        self.frame_ac_dc.pack(pady = 5, anchor=tk.W)

        self.rd_btn_ac_dc_1 = tk.Radiobutton(self.frame_ac_dc, text = self.AC, variable = self.var_ac_dc, value = self.AC, command=lambda: self.show_selected(self.data_type_ac_dc))
        self.rd_btn_ac_dc_1.pack(side=tk.LEFT)
        self.rd_btn_ac_dc_2 = tk.Radiobutton(self.frame_ac_dc, text = self.DC, variable = self.var_ac_dc, value = self.DC, command=lambda: self.show_selected(self.data_type_ac_dc))
        self.rd_btn_ac_dc_2.pack(side=tk.LEFT)

        self.lb_range = tk.Label(self, text = self.lable_for_range_input)
        self.lb_range.pack(pady = 10, anchor=tk.W)

        self.frame_range = tk.Frame(self)
        self.frame_range.pack(pady = 5, anchor=tk.W)

        self.rd_btn_range_1 = tk.Radiobutton(self.frame_range, text = "AUTO", variable = self.var_range, value = "AUTO", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_1.pack(side=tk.LEFT)
        self.rd_btn_range_2 = tk.Radiobutton(self.frame_range, text = "0.1", variable = self.var_range, value = "0.1", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_2.pack(side=tk.LEFT)
        self.rd_btn_range_3 = tk.Radiobutton(self.frame_range, text = "1", variable = self.var_range, value = "1", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_3.pack(side=tk.LEFT)
        self.rd_btn_range_4 = tk.Radiobutton(self.frame_range, text = "10", variable = self.var_range, value = "10", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_4.pack(side=tk.LEFT)
        self.rd_btn_range_5 = tk.Radiobutton(self.frame_range, text = "100", variable = self.var_range, value = "100", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_5.pack(side=tk.LEFT)
        self.rd_btn_range_6 = tk.Radiobutton(self.frame_range, text = "1000", variable = self.var_range, value = "1000", command=lambda: self.show_selected(self.data_type_range))
        self.rd_btn_range_6.pack(side=tk.LEFT)
        self.rd_btn_range_7 = tk.Radiobutton(self.frame_range, text = "10", variable = self.var_range, value = "10", command=lambda: self.show_selected(self.data_type_range))

        self.lb_sleep = tk.Label(self, text=self.lable_for_sleep_time_input)
        self.lb_sleep.pack(pady = 10, anchor=tk.W)

        self.frame_sleep = tk.Frame(self)
        self.frame_sleep.pack(pady = 5, anchor=tk.W)

        self.txt_sleep = tk.Text(self.frame_sleep, width=10, height=1)
        self.txt_sleep.pack(side=tk.LEFT, padx=5)

        self.lb_time_dur = tk.Label(self, text=self.lable_for_time_dur_input)
        self.lb_time_dur.pack(pady = 10, anchor=tk.W)

        self.frame_time_dur = tk.Frame(self)
        self.frame_time_dur.pack(pady = 5, anchor=tk.W)

        self.txt_time_dur = tk.Text(self.frame_time_dur, width=10, height=1)
        self.txt_time_dur.pack(side=tk.LEFT, padx=5)
        self.rd_btn_time_dur_unit_1 = tk.Radiobutton(self.frame_time_dur, text=self.time_unit_second, variable = self.var_time_dur_unit, value = self.time_unit_second, command=lambda: self.show_selected(self.data_type_time_dur_unit))
        self.rd_btn_time_dur_unit_1.pack(side=tk.LEFT)
        self.rd_btn_time_dur_unit_2 = tk.Radiobutton(self.frame_time_dur, text=self.time_unit_minute, variable = self.var_time_dur_unit, value = self.time_unit_minute, command=lambda: self.show_selected(self.data_type_time_dur_unit))
        self.rd_btn_time_dur_unit_2.pack(side=tk.LEFT)
        self.rd_btn_time_dur_unit_3 = tk.Radiobutton(self.frame_time_dur, text=self.time_unit_hour, variable = self.var_time_dur_unit, value = self.time_unit_hour, command=lambda: self.show_selected(self.data_type_time_dur_unit))
        self.rd_btn_time_dur_unit_3.pack(side=tk.LEFT)

        self.lb_show_selected = tk.Label(self, text=self.lable_for_show_selection, font=(self.show_selection_text_font))
        self.lb_show_selected.pack(anchor=tk.W, padx =5)

        self.frame_btn_control = tk.Frame(self)
        self.frame_btn_control.pack(pady=5, anchor=tk.W)

        self.btn_file_path = tk.Button(self.frame_btn_control, width=20, height=2, text = "文件保存地址", command=lambda: self.get_filepath())
        self.btn_file_path.pack(side=tk.LEFT, padx=5)

        self.btn_begin_test = tk.Button(self.frame_btn_control, width=20, height=2, text = "开始测试", command=self.begin_measure)        
        self.btn_begin_test.pack(side=tk.LEFT, padx=5)

        self.btn_terminate_test = tk.Button(self.frame_btn_control, width=20, height=2, text = "停止测试", command=self.terminate)

        self.btn_exit = tk.Button(self.frame_btn_control, width=20, height=2, text = "退出程序", command=sys.exit)
        self.btn_exit.pack(side=tk.LEFT, padx=5)

        self.show_terminal()

    def refresh_insts(self):
        self.cmb_usb_visa_address['values'] = self.get_insts()
        if self.get_insts():  # 确保列表非空
            self.cmb_usb_visa_address.set(self.get_insts()[0])
    
    # insert the terminal module into the UI
    def show_terminal(self):
        self.text_area = scrolledtext.ScrolledText(
            self,
            wrap=tk.WORD,
            bg="black",    
            fg="white",   
            font=self.default_text_font
        )
        self.text_area.pack(padx=10, pady=5, anchor=tk.W, fill=tk.BOTH, expand=True)

        sys.stdout = TerminalRedirector(text_widget=self.text_area)
        sys.stderr = TerminalRedirector(text_widget=self.text_area)
        
    # match the range selections for Voltage setting
    def show_remained_V(self):
        self.rd_btn_range_1.config(text = "AUTO", variable = self.var_range, value = "AUTO")
        self.rd_btn_range_2.config(text = "0.1", variable = self.var_range, value = "0.1")
        self.rd_btn_range_3.config(text = "1", variable = self.var_range, value = "1")
        self.rd_btn_range_4.config(text = "10", variable = self.var_range, value = "10")
        self.rd_btn_range_5.config(text = "100", variable = self.var_range, value = "100")
        self.rd_btn_range_6.config(text = "1000", variable = self.var_range, value = "1000")
        self.rd_btn_range_7.pack_forget()

    # match the range selections for Current setting
    def show_remained_I(self):
        self.rd_btn_range_1.config(text = "AUTO", variable = self.var_range, value = "AUTO")
        self.rd_btn_range_2.config(text = "0.0001", variable = self.var_range, value = "0.0001")
        self.rd_btn_range_3.config(text = "0.001", variable = self.var_range, value = "0.001")
        self.rd_btn_range_4.config(text = "0.01", variable = self.var_range, value = "0.01")
        self.rd_btn_range_5.config(text = "1", variable = self.var_range, value = "1")
        self.rd_btn_range_6.config(text = "3", variable = self.var_range, value = "3")
        self.rd_btn_range_7.pack(side=tk.LEFT, padx=0)
    
    def get_filepath(self, fn = "Test_File.mat"):
        # open the filedialog to allow users browse and select file path for saving
        self.file_path = filedialog.asksaveasfilename(
            defaultextension=".mat",
            filetypes=[("MAT files", "*.mat"), ("All files", "*.*")],
            initialfile=fn,
            title=self.label_for_filedialog_title
        )

    # display users' input and update the input into data collections
    def show_selected(self, data_type):
        var = None
        if data_type == self.data_type_mode:
            var = self.var_mode.get()
            if var == self.VOLTAGE: 
                self.show_remained_V()
            elif var == self.CURRENT: 
                self.show_remained_I()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_ac_dc:
            var = self.var_ac_dc.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_range:
            var = self.var_range.get()
            if not var == instMultimeter.MM_RANGE_AUTO: var = float(var)
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_sleep_time or data_type == self.data_type_time_dur:
            var_sleep_time = self.txt_sleep.get("1.0", tk.END).strip()
            var_time_dur = self.txt_time_dur.get("1.0", tk.END).strip()
            try:
                num_sleep_time = float(var_sleep_time)
                num_time_dur   = float(var_time_dur)
                if num_sleep_time <= 0 or num_time_dur <= 0:
                    var = self.user_input_non_positive_alert
                    self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
                else:
                    self.var_sleep_time = num_sleep_time
                    self.var_time_dur = num_time_dur
            except:
                if var_sleep_time and var_time_dur:
                    var = self.user_input_wrong_type_alert
                else:
                    var = []
                    if not var_sleep_time: var.append(self.user_input_miss_sleep_time_alert)
                    if not var_time_dur: var.append(self.user_input_miss_time_dur_alert)
                self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_time_dur_unit: 
            var = self.var_time_dur_unit.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_usb_lan:
            var = self.usb_lan.get()
            self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")
        elif data_type == self.data_type_visa_address:
            var_usb = self.var_usb_visa_address.get().replace(" ","")
            var_lan = (
                self.txt_lan_visa_address_1.get("1.0", tk.END).replace(" ", "").rstrip('\n') + "." + 
                self.txt_lan_visa_address_2.get("1.0", tk.END).replace(" ", "").rstrip('\n') + "." +
                self.txt_lan_visa_address_3.get("1.0", tk.END).replace(" ", "").rstrip('\n') + "." +
                self.txt_lan_visa_address_4.get("1.0", tk.END).replace(" ", "").rstrip('\n')
            )

            usb_selected = self.usb_lan.get() == self.usb
            lan_selected = self.usb_lan.get() == self.lan

            self.lan_visa_address_typed = (self.txt_lan_visa_address_1.get("1.0", tk.END).rstrip('\n') and
                                           self.txt_lan_visa_address_2.get("1.0", tk.END).rstrip('\n') and
                                           self.txt_lan_visa_address_3.get("1.0", tk.END).rstrip('\n') and
                                           self.txt_lan_visa_address_4.get("1.0", tk.END).rstrip('\n') and
                                           lan_selected)
            
            self.usb_visa_address_typed = var_usb and usb_selected

            if self.lan_visa_address_typed or self.usb_visa_address_typed:
                self.var_visa_address = f"TCPIP0::{var_lan}::INSTR" if lan_selected else var_usb
                
            else:
                if usb_selected:
                    var = self.user_input_miss_visa_address
                else:
                    var = self.user_input_miss_ip_address
                # self.var_visa_address = None
                self.lb_show_selected.config(text=f"{self.lable_for_show_selection}{var}")

    # send users' input to the main function 
    def get_data(self, data_type):
        if   data_type == self.data_type_mode:
            return self.var_mode.get()
        elif data_type == self.data_type_ac_dc:
            return self.var_ac_dc.get()
        elif data_type == self.data_type_range:
            if self.var_range.get() == 0: return instMultimeter.MM_RANGE_AUTO 
            else: return self.var_range.get()
        elif data_type == self.data_type_sleep_time:
            return self.var_sleep_time
        elif data_type == self.data_type_time_dur:
            return self.var_time_dur
        elif data_type == self.data_type_time_dur_unit:
            return self.var_time_dur_unit.get()
        elif data_type == self.data_type_visa_address:
            return self.var_visa_address
        
    # convert the user time input in seconds
    def cal_run_time(self, time_unit, time_dur):

        time_in_second = 0

        # convert the time into seconds based on the unit
        if time_unit == self.time_unit_second: 
            time_in_second = timedelta(seconds=time_dur).total_seconds()
        elif time_unit == self.time_unit_minute: 
            time_in_second = timedelta(minutes=time_dur).total_seconds()
        else:
            time_in_second = timedelta(hours=time_dur).total_seconds()

        return time_in_second
        
     # check whether all settings are setted to deteine if the test should be statred  
    def begin_measure(self):

        self.show_selected(self.data_type_sleep_time)
        self.show_selected(self.data_type_time_dur)
        self.show_selected(self.data_type_visa_address)

        if self.var_sleep_time and self.var_time_dur and self.var_visa_address:
            self.is_terminated = False

            self.btn_terminate_test.pack(side=tk.LEFT, padx=5)
            self.btn_exit.pack_forget()
            self.btn_begin_test.pack_forget()
            self.btn_file_path.pack_forget()

            # get and save locally the user input from UI
            self.saved_visa_address  = self.get_data(UI.data_type_visa_address)
            self.saved_mode_input    = self.get_data(UI.data_type_mode)
            self.saved_ac_dc_input   = self.get_data(UI.data_type_ac_dc)
            self.saved_range_input   = self.get_data(UI.data_type_range)
            self.saved_sleep_time    = self.get_data(UI.data_type_sleep_time)
            self.saved_time_dur      = self.get_data(UI.data_type_time_dur)
            self.saved_time_dur_unit = self.get_data(UI.data_type_time_dur_unit)

            # open the instrument
            mt = instKS_34461A(visa_address=self.saved_visa_address)
            mt.inst_open()

            # set the configuration by the mode/AC/DC/range from users' input
            mt.set_mode(self.saved_mode_input, self.saved_ac_dc_input)
            mt.set_range(self.saved_range_input)

            print("主程序开始处理")

            # calculate the runtime for the test
            total_runtime = self.cal_run_time(self.saved_time_dur_unit, self.saved_time_dur)
            self.time_start = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            # record the start time for the test
            start_time = time.time()

            # count the number of times of data saving
            count = 0

            # test data collections
            self.time_stamps = []         # the storing time for each piece data
            self.power_data  = []         # the power meaured

            self.time_stamps_path = None
            self.power_data_path = None

            is_delete_first_measure = False

            while True:

                # automaticlly terminate the test in respect to the pre-set runtime or the user terminates the test
                time_since_start = time.time() - start_time
                if  time_since_start >= total_runtime or self.is_terminated:
                    self.save_mat_file()
                    print(f"数据采集结束 程序已运行{time_since_start:.2f}{self.time_unit_second}")
                    break

                if count - 100 >= 0 and count % 100 == 0:
                    self.save_mat_file()

                # update the test data collections
                self.time_stamps.append(time_since_start)
                self.time_measure_start = time.time()
                power = mt.measure()
                self.power_data.append(power)
                count += 1
                current_time = datetime.now().strftime("%m.%d %H:%M:%S")
                self.update()

                if not is_delete_first_measure:
                    is_delete_first_measure = 1
                    start_time += time_since_start
                    time_since_start = 0
                
                print(f"[{current_time}] 执行任务{count}次...{power}")

                self.update_during_sleep(start_time, total_runtime, count, self.saved_sleep_time)

            self.btn_file_path.pack(side=tk.LEFT, padx=5)
            self.btn_begin_test.pack(side=tk.LEFT, padx=5)
            self.btn_exit.pack(side=tk.LEFT, padx=5)
            self.btn_terminate_test.pack_forget()

            mt.close()

    # terminate the test
    def terminate(self):
        self.is_terminated = True
        self.btn_file_path.pack(side=tk.LEFT, padx=5)
        self.btn_begin_test.pack(side=tk.LEFT, padx=5)
        self.btn_exit.pack(side=tk.LEFT, padx=5)
        self.btn_terminate_test.pack_forget()

    # enable the UI to receive the user input during the sleep time
    def update_during_sleep(self, start_time, total_runtime, count, sleep_time):
        interval = 1/self.update_frequency
        sleep_count = int(sleep_time*self.update_frequency)
        time_end = self.time_measure_start + sleep_time
        for i in range(sleep_count):
            self.update()
            time_since_start = time.time() - start_time
            if time_since_start >= total_runtime or self.is_terminated: 
                break

            if time_since_start > (count * sleep_time):
                break

            current_time = time.time()
            remaining_time = time_end - current_time 
            if remaining_time < interval:
                if remaining_time > 0: time.sleep(interval)
                break

            time.sleep(interval)
    
    # save the mat file
    def save_mat_file(self):
        # convert the data into np(more applicaple for MatLab)

        mat_var_time_stamps = "time_stamps"
        mat_var_power = "power"
        mat_var_config = "configuration"

        config = [f"{self.data_type_visa_address}: " + self.saved_visa_address, 
                  f"{self.data_type_mode}: "         + self.saved_mode_input, 
                  f"{self.data_type_ac_dc}: "        + self.saved_ac_dc_input, 
                  f"{self.data_type_range}: "        + self.saved_range_input, 
                  f"{self.data_type_sleep_time}: "   + str(self.saved_sleep_time) + self.default_time_dur_unit,
                  f"{self.data_type_time_dur}: "     + str(self.saved_time_dur) + self.saved_time_dur_unit,
                  "开始时间: "                        + self.time_start,
                  "保存时间: "                        + datetime.now().strftime("%Y-%m-%d %H:%M:%S")]

        data_to_save = {
            mat_var_time_stamps: np.array(self.time_stamps, dtype=np.float64),
            mat_var_power: np.array(self.power_data, dtype=np.float64),
            mat_var_config: np.array(config)
        }

        # save the file into the pre-set path
        instKS_34461A.save_matfile(self, fn=self.file_path, mm=data_to_save)
    
class TerminalRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.originial_stdout = sys.stdout

    def write(self, string):
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)

class instOSC_DS1104(bATEinst_base):
    Model_Supported = ["DS1104"]
    
    def __init__(self):
        super().__init__("osc")
        self.VisaAddress="USB0::0x05E6::0x2450::04542396::INSTR"
        
    def callback_after_open(self):
        self.set_visa_timeout_value(10000)
    
    def set_x(self, xscale=None, offset=None):
        self.x_write([":TIM:SCAL %6e" %xscale, "*OPC?"] if xscale else [] 
                    +[":TIM:OFFS %6e" %offset, "*OPC?"] if offset is not None else [] 
        )

    def set_y(self,ch, yscale=None, yoffset=None):
        self.x_write([":CHAN%d:SCAL %6e" %(ch, yscale), "*OPC?"] if yscale else [] 
                    +[":CHAN%d:OFFS %6e" %(ch, yoffset), "*OPC?"] if yoffset is not None else [] 
        )
    
    def measure(self):
        self.x_write([":STOP", "*OPC?", "SING","$WAIT=1"
        ":TRIG:STAT?$LOOP UNTIL = STOP in 20"]
        )
    def start(self):
        self.x_write([":RUN"])
    
    def load_setup(self, fn):
        self.x_write([":LOAD:SET '%s'" % fn, "*OPC?"])
        
    def save_image(self, fn):
        self.x_write([":STORage:IMAGe:TYPE PNG", "*OPC?"])
        self.write(":DISPlay:DATA?")
        dd = self.read_block()
        with open(fn, "wb") as fid:
            fid.write(dd)

    def save_waveform(self, fn):
        res = [k.strip() == "1" for k in self.x_write([":CHAN1:DISP?", ":CHAN2:DISP?", ":CHAN3:DISP?", ":CHAN4:DISP?"])]
        chs = []
        for k in range(len(res)):
            if res[k]:
                chs.append(k + 1)
        vvs = []
        for ch in chs:
            self.x_write([":WAV:SOUR CHAN%d" % ch,
                          ":WAV:MODE RAW", ":WAV:FORM BYTE",
                          ])
            point, av, xinc, xor, xref, yinc, yor, yref = [float(k) for k in
                                                           self.x_write(":WAV:PRE?")[0].split(",")[2:]]
            dd = []
            for st in range(1, int(point) + 1, 125000):
                self.x_write(
                    [":WAV:STAR %d" % st, ":WAV:STOP %d" % (st + 125000 - 1 if point >= st + 125000 - 1 else point)])
                self.write(":WAV:DATA?")
                dd += self.read_block()
            vvs.append([(k - yor - yref) * yinc for k in dd])
        pp = min([len(k) for k in vvs])
        with open(fn, "wt") as fid:
            for tt in range(pp):
                fid.write("\t".join(["%g" % ((tt - xor - xref) * xinc)]
                                     + ["%g" % vvs[k][tt] for k in range(len(chs))]) + "\n")
                           
class instOSC_MDO34(bATEinst_base):
    Model_Supported = ["MDO34"]
    
    def __init__(self):
        super().__init__("osc")
        self.VisaAddress="USB::0x0699::0x052C::C050152::INSTR"
        
    
    def set_x(self, xscale=None, offset=None):
        self.x_write([":TIM:SCAL %6e" %xscale, "*OPC?"] if xscale else [] 
                    +[":TIM:OFFS %6e" %offset, "*OPC?"] if offset is not None else [] 
        )

    def set_y(self,ch, yscale=None, yoffset=None):
        self.x_write([":CHAN%d:SCAL %6e" %(ch, yscale), "*OPC?"] if yscale else [] 
                    +[":CHAN%d:OFFS %6e" %(ch, yoffset), "*OPC?"] if yoffset is not None else [] 
        )
    
    def measure(self):
        self.x_write([":STOP", "*OPC?", "SING","$WAIT=1$"] )
        
    def start(self):
        self.x_write([":RUN"])
    
    def load_setup(self, fn):
        self.x_write([":LOAD:SET '%s'" % fn, "*OPC?"])
        
    def save_image(self, fn):
        self.x_write([":SAV:IMAG:FILEF PNG", "SAV:IMAG "+fn, "*OPC?"])
        # dd = self.read_block()
        # with open(fn, "wb") as fid:
        #     fid.write(dd)

    def save_waveform(self, fn):
        self.x_write([":WFMO:ENC BIN",":WFMO:BN_FMT RI",":WFMO:BYT_O MSB",":WFMO:BYT_N 1",":DAT:SOU CH1",
                      ":DAT:START 1",":DAT:STOP 20000000","*OPC?"])
        pre = self.x_write(":WFMO?")[0].split(";")
        n = int(pre[6])
        x_inc = float(pre[10])
        x_off = float(pre[11])
        y_inc = float(pre[14])
        y_off = float(pre[15])
        n_block = 200000
        with open(fn, "wb") as fid:
            hd = struct.pack("5d",n, x_inc,x_off,y_inc,y_off)
            fid.write(hd)
            for k in range(1,n,n_block):
                self.x_write([":DAT:START %d " % k,":DAT:STOP %d" % (k+n_block-1 if (k+n_block-1) <=n else n ),"*OPC?"])
                self.write("CURV?")
                dd = self.read_block()
                fid.write(dd)
            
class instOSC_DHO1204(bATEinst_base):
    Model_Supported = ["DHO1204"]
    
    def __init__(self):
        super().__init__("osc")
        self.VisaAddress="USB::0x1AB1::0x0610::HDO1B254200719::INSTR"
    
    def callback_after_open(self):
        self.set_visa_timeout_value(8000)
     
    def set_x(self, xscale=None, offset=None):
        self.x_write([":TIM:SCAL %6e" %xscale, "*OPC?"] if xscale else [] 
                    +[":TIM:OFFS %6e" %offset, "*OPC?"] if offset is not None else [] 
        )

    def set_y(self,ch, yscale=None, yoffset=None):
        self.x_write([":CHAN%d:SCAL %6e" %(ch, yscale), "*OPC?"] if yscale else [] 
                    +[":CHAN%d:OFFS %6e" %(ch, yoffset), "*OPC?"] if yoffset is not None else [] 
        )
    
    def measure(self):
        self.x_write([":STOP", "*OPC?", "SING","$WAIT=1$"] )
        
    def start(self):
        self.x_write([":RUN"])
        
    def stop(self):
        self.x_write([":STOP"])
        
    def load_setup(self, fn):
        self.x_write([F":LOAD:SET '{fn}'", "*OPC?"])
        
    def save_image(self, fn):
        #self.x_write([":SAV:IMAG:FILEF PNG", "SAV:IMAG "+fn, "*OPC?"])
        self.x_write([":SAV:IMAG:FILEF PNG", "*OPC?"])
        dd = self.read_block(":SAVE:IMAGe:DATA?")
        with open(fn, "wb") as fid:
             fid.write(dd)
    
    def set_acquire(self,depth=None, mode =None):
        if depth:
            self.x_write([F":ACQ:MDEP {depth:.0f}"])
        
    def save_waveform(self, fn,waves=None):
        if waves is None:
            chs = [(ch+1) if self.query(F":CHAN{ch+1}:DISP?").startswith("1")  else None for ch in range(4)]
            chs = [k for k in chs if k]
            mm = {}
            for ch in chs:
                wv = self.read_waveform(ch)
                mm = {**mm, **wv}
            mm["channels"] =  chs
        else:
            if isinstance(waves,dict):
                mm = waves
            elif isinstance(waves, list):
                mm = {}
                for wv in waves:
                    mm = {**mm, **wv}
                chs = [k["channels"][0] for k in waves]
                mm["channels"] =  chs   
        mm.pop("scale",None) 
        mm.pop("data",None)     
        savemat(fn, mm,appendmat=False)
    
    def read_waveform(self,ch):
        self.x_write([":STOP","*OPC?",F":WAV:SOUR CHAN{ch}",":WAV:MODE RAW", ":WAV:FORM BYTE","*OPC?"])
        point = int(round(float(self.x_write([":ACQuire:MDEP?"])[0].strip())))
        self.x_write([":WAV:STAR 1", F":WAV:STOP {point}"])
        pre =  self.x_write(":WAV:PRE?")[0]
        ff, tt, point, count, xinc, xor, xref, yinc, yor, yref = [float(k) for k in pre.split(",")]

        mm = {"xscale":[xinc,xor, xref], "xinc":xinc,
        "point":point,"format":ff,"count":count ,"type":tt,
        "channels":[ch],
        "time":str(datetime.fromtimestamp(time.time()))}
        yscale = [yinc, yor,yref]

        mm[F"ch{ch}scale"]=yscale
        dd = []
        max_size=125000 #125000
        for st in range(1, int(point) + 1, max_size):
            self.x_write([F":WAV:STAR {st}", F":WAV:STOP {(st + max_size - 1 if point >= st + max_size - 1 else point)}","*OPC?"])
            self.write(":WAV:DATA?")
            dd += self.read_block()

        ddar =  np.array(dd,dtype=np.uint8)
        mm[F"ch{ch}data"] = ddar
        mm["scale"]  =yscale
        mm["data"]  = ddar
        return mm
    
    def raw2float(self,raw,scale=None):
        if isinstance(raw,dict):
            scale = raw["scale"]
            raw = raw["data"]
        return (np.array(raw)-(scale[1]+scale[2]))*scale[0]
    
    def test(self):
        self.start()
        time.sleep(5)
        self.stop()
        plt.plot(self.raw2float(self.read_waveform(1)))
        plt.show()
        #self.save_waveform(r"d:\work\tt\dd.mat")
                
class instSW_CP2102(bATEinst_base):
    Model_Supported = ["3000072"]
    
    def __init__(self):
        super().__init__("sw")
        self.VisaAddress="COM7" #"ASRL7::INSTR"
        self.get_cal_amp = None
        self.current_freq = 0
    
    def inst_open(self):
        rr=re.match(r"COM(\d+)", self.VisaAddress)
        if rr:
            self.VisaAddress =f"ASRL{rr[1]}::INSTR"
        return super().inst_open()
    
    def callback_after_open(self):
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_ASRL_RTS_STATE,0)
    
    def set_sw(self, on=None):
        """_summary_

        Be careful: the switch's default status is ON, when open, the status will be Off.   
        when Power on,  VI_ATTR_ASRL_RTS_STATE==0, on ==1 --> LED on, switch to rf spliter (outside of sw)
             Power off, VI_ATTR_ASRL_RTS_STATE==1, on ==0 --> LED off, switch to AWG ( inside of sw)
        """
        self.check_open()
        if isinstance(on, str):
            on = (on.lower() != "awg")
        self.Inst.set_visa_attribute(pyconst.VI_ATTR_ASRL_RTS_STATE,0 if on else 1)
        time.sleep(0.2)
        #self.close()
        #
    def test(self):
        self.set_sw(1)
        time.sleep(1)
        self.set_sw(0)
           
class instSG_DSG836(bATEinst_base):
    Model_Supported = ["DSG836"]
    
    def __init__(self):
        super().__init__("sg")
        self.VisaAddress="USB::0x1AB1::0x099C::DSG8M253400109::INSTR"
        self.get_cal_amp:intpl.interp1d= None
        self.current_freq = 0
        
    def calib_level(self, val):
        if self.get_cal_amp:
            return val/ math.pow(10,float(self.get_cal_amp(self.current_freq))/20)
        else:
            return val
        
    def set_freq(self, freq):
        self.x_write([":FREQ %.2f" %freq, "*OPC?"] )
        self.current_freq = freq

    def set_amp_v(self,amp_v):
        self.x_write([":LEV %.6fV" % self.calib_level(amp_v), "*OPC?"])
    
    def set_on(self,on=True):
        self.x_write([":OUTP %d" % (1 if on else 0), "*OPC?"])
        
    def set_lf_freq(self, freq):
        self.x_write([":LFO:FREQ %.2f" %freq, "*OPC?"] )
        self.current_freq = freq

    def set_lf_amp_v(self,amp_v):
        self.x_write([":LFO:LEV %.6fV" % amp_v, "*OPC?"])
    
    def set_lf_shape(self,shape="SINE"):
        """ shape: SINE, SQU"""
        self.x_write([":LFO:SHAP SINE"])
        
    def set_lf_on(self,on=True):
        self.x_write([":LFO %d" % (1 if on else 0), "*OPC?"])
         
class instAWG_DG4102(bATEinst_base):
    Model_Supported = ["DG4102"]
    
    def __init__(self):
        super().__init__("awg")
        self.ch = 1 # 1, 2
        self.VisaAddress="USB::0x1AB1::0x0641::DG4E245103182::INSTRs"
        self.get_cal_level= None  # a interpolate function table.
        self.freqs = [0,0]
        self.levels = None
        
    def callback_after_open(self):
        pass
        #self.set_reset()

    def calib_level(self, ch, val,freq = None):
        if self.get_cal_level:
            freq = self.freqs[ch-1] if freq is None else freq
            return val/float(self.get_cal_level[ch-1](freq)) 
        else:
            return val

    def sel_chan(self,ch):
        self.ch = ch
                  
    class MODE:
        SIN = 1
        DC = 0
        PULSE = 2
        SQU = 3
    CH_ALL = 0
    
    def set_freq(self,freq, ch=None):
        if  isinstance(freq, list):
            for ch,vv in enumerate(freq):
                self.x_write([":SOUR%d:FREQ %f" %(ch+1,vv), "*OPC?"] )
                self.freqs[ch] = vv
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:FREQ %f" %(ch,freq), "*OPC?"] )
                self.freqs[ch-1] = freq
                
    def ch2chs(self, ch):
        chs = self.ch if ch is None else ch
        if not isinstance(chs, list):
            chs = [chs]
        if chs == []:
            chs = [1,2]
        return chs     
    
    def set_reset(self):
        self.x_write(["*RST", "*OPC?", ":OUTP1:IMP INF",":OUTP2:IMP INF"])

    def set_mode(self, mode = MODE.SIN, ch=None):
        if not isinstance(mode, str):
            mode = "PULSE" if mode ==2 else "DC" if mode ==0 else "SQU" if mode ==3 else "SIN"
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:APPL:%s" %(ch,mode), "*OPC?"] )
            
    def set_sine_mode(self,freq=1e8,amp=0.01, ch=None):   
        self.set_mode(self.MODE.SIN,ch)
        self.set_freq(freq,ch)
        self.set_amp(amp,ch)
        self.set_offset(0,ch)
        self.set_on(True, ch)     
             
    def set_dc_mode(self,dc=0,ch=None):   
        self.set_mode(self.MODE.SIN,ch)
        self.set_freq(1e-6,ch)
        self.set_amp(1e-3,ch)
        self.set_offset(dc,ch)
        self.set_on(True, ch)
        
    def set_phase(self, ph, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:%s" %(ch,ph), "*OPC?"] )
            
    def phase_sync(self,ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:INIT" %(ch), "*OPC?"] )
                             
    def set_amp(self,amp,ch=None):
        if  isinstance(amp, list):
            for ch,vv in enumerate(amp):
                self.x_write([":SOUR%d:VOLT:AMPL %.4f" %(ch+1,self.calib_level(ch+1, vv)), "*OPC?"] )
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT:AMPL %.4f" %(ch,self.calib_level(ch, amp)), "*OPC?"] )
   
    def set_burst_phase(self,ph, ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:BURS:PHAS %.4f" %(ch,ph), "*OPC?"] )
            
    def set_offset(self, v, ch = None):
        if  isinstance(v, list):
            for ch,vv in enumerate(v):
                self.x_write([":SOUR%d:VOLT:OFFS %.4f" %(ch+1,self.calib_level(ch+1, vv, 0)), "*OPC?"] )
        else:          
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT:OFFS %.4f" %(ch,self.calib_level(ch, v, 0)), "*OPC?"] )
        
    def set_on(self,on=True,ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":OUTP%d %d" %(ch, (1 if on else 0)), "*OPC?"])
    
    def set_data_rate_test(self,afreq=200e3, bfreq=100, bursts=500,level=3.3):
        self.x_write(["*RST","*OPC?",
                      ":OUTP1:IMP INF",":OUTP2:IMP INF",
                      ":SOUR1:APPL:SQU %f, %.2f, %.2f,300" % (afreq,level/2,level/4),
                      ":SOUR2:APPL:SQU %f, %.2f, %.2f,0" % (bfreq,level/2,level/4),
                      "*OPC?",
                      ":SOUR2:TRIG:SOUR MAN",
                      ":SOUR2:BURS:MODE TRIG","*OPC?",
                      ":SOUR2:BURS:NCYC %d" % bursts,"*OPC?",
                      ":SOUR2:BURS ON",
                      ":OUTP1 1",
                      ":OUTP2 1", 
                      "*OPC?",
                      ])

    def fire_burst_manul_trigger(self,ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:BURS:TRIG" % ch,"*OPC?"])
            
    def reset(self):
        self.x_write([":SYST:PRES DEF","*OPC?",":OUTP1:IMP INF",":OUTP2:IMP INF",
        "*OPC?",":SOUR1:APPL:SIN", ":SOUR2:APPL:SIN", "*OPC?" ])
        
class instAWG_DG852(instAWG_DG4102):
    Model_Supported = ["DG852"]
    
    def __init__(self):
        super().__init__()
        self.VisaAddress="USB::0x1AB1::0x0646::DG8R262900659::INSTR"

    def set_reset(self):
        self.x_write(["*RST","$WAIT=1500$", "*OPC?", ":OUTP1:LOAD 50",":OUTP2:LOAD 50"])
        
    def set_amp(self,amp,ch=None):
        if  isinstance(amp, list):
            for ch,vv in enumerate(amp):
                self.x_write([":SOUR%d:VOLT %.4f" %(ch+1,self.calib_level(ch+1, vv)), "*OPC?"] )
        else:
            for ch in self.ch2chs(ch):
                self.x_write([":SOUR%d:VOLT %.4f" %(ch,self.calib_level(ch, amp)), "*OPC?"] )
                
    def phase_sync(self,ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([":SOUR%d:PHAS:SYNC" %(ch), "*OPC?"] )
  
    def set_data_rate_test(self,afreq=200e3, bfreq=100, bursts=500,level=3.3):
        self.x_write(["*RST","$WAIT=1500$", "*OPC?",
                      ":OUTP1:LOAD 50",":OUTP2:LOAD 50",
                      ":SOUR1:APPL:SQU %f, %.2f, %.2f,300" % (afreq,level/2,level/4),
                      ":SOUR2:APPL:SQU %f, %.2f, %.2f,0" % (bfreq,level/2,level/4),
                      "*OPC?"])
        self.x_write([":SOUR2:TRIG:SOUR MAN",
                      ":SOUR2:BURS:MODE TRIG","*OPC?",
                      F":SOUR2:BURS:NCYC {bursts}","*OPC?",
                      ":SOUR2:BURS:STAT ON",
                      ":OUTP1 1",
                      ":OUTP2 1", 
                      "*OPC?",
                      ])
    def fire_burst_manul_trigger(self,ch=None):
        for ch in self.ch2chs(ch):
            self.x_write([F":TRIG{ch}","*OPC?"])
            
    def reset(self):
        self.x_write(["*RST","$WAIT=1500$", "*OPC?",":OUTP1:LOAD 50",":OUTP2:LOAD 50",
        "*OPC?",":SOUR1:APPL:SIN", ":SOUR2:APPL:SIN", "*OPC?" ])

class instDC_KA3003P(bATEinst_base):
    Model_Supported = ["KA3003P"]
    
    def __init__(self):
        super().__init__("dc")
        self.VisaAddress="ASRL7::INSTR"
        
    def measure_v(self):
        return float(self.x_write(["VOUT1?"] )[0])
        
    def measure_i(self):
        try:
            return float(self.x_write(["IOUT1?"] )[0])
        except:
            return float(self.x_write(["IOUT1?"] )[0])
    
    def measure_iv(self):
        return (self.measure_i(),self.measure_v())
        
    def set_v(self,v):
        self.x_write(["VSET1:%.2fV" % v])
        
    def set_i(self,v):
        self.x_write(["ISET1:%.3fV" % v])
    
    def set_on(self,on=True):
        self.x_write(["OUT%d" % (1 if on else 0)])
        time.sleep(0.5)
        
    def test(self):
        self.set_v(3.3)
        self.set_i(0.2)
        self.set_on(1)
        i = self.measure_i()
        v = self.measure_v()
        print((i,v))
        self.set_on(0)

class instTrigger(bATEinst_base):
    Model_Supported = [""]
    
    def __init__(self):
        super().__init__("pwm")
        self.VisaAddress="COM4"
        
    def measure_v(self):
        return float(self.x_write(["VOUT1?"] )[0])
    
    def inst_open(self):
        if not self.Inst:
            if not self.VisaAddress:
                self.set_error("Equip Address has not been set!")
            try:
                self.Inst = serial.Serial(self.VisaAddress, 115200, timeout=1)
            except serial.SerialException:
                time.sleep(2)
                self.Inst = serial.Serial(self.VisaAddress, 115200, timeout=1)
        return self.Inst
    
    def trigger(self, afreq,bfreq, csize):
        acycles = round(1e8/afreq)-1
        bcycles = round(afreq/bfreq)
        absize = acycles + (bcycles<<16)
        self.send(self.CMD_CCSIZE, 0x0000_0000 )
        self.send(self.CMD_ABSIZE,absize)
        self.send(self.CMD_CCSIZE, csize+0x8000_0000 )

    def wait_done(self,maxdelay=10):
        st = time.time()
        while time.time()-st < maxdelay:
            v = self.send(self.CMD_STATE, 0x0000_0000 )
            if ( v &0x03 == 2):
                break
            time.sleep(0.1)
        return  time.time()-st
    
    CMD_ABSIZE =0
    CMD_CCSIZE =1
    CMD_STATE = 2
    
    def send(self,cmd, value):
        self.check_open()
        for _ in range(2):
            if self.Inst.in_waiting>0:
                self.Inst.read(self.Inst.in_waiting)
            self.Inst.write(list("SV".encode("utf-8"))+[0,cmd]+list(struct.pack ("I",value)))
            ret=self.Inst.read(8)
            if len(ret) == 8 and ret[0] =='s': break
            self.Inst.write([0]*256)
            time.sleep(0.5)
        if len(ret)!=8:
            self.set_error("return error")
        return struct.unpack("II",ret)[1]
    
def py_code_clean():
	pass
      