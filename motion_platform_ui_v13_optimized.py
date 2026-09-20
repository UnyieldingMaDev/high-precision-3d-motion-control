# -*- coding: utf-8 -*-
"""
三维运动平台 Python UI 控制程序 V13 优化版
------------------------------------------------
这个版本尽量仿照厂家 Motion 软件的“控制”界面布局：

1. 当前位置
2. 相对运动
3. 绝对运动
4. 点动
5. 速度滑条
6. S型扫描轨迹设置
7. 示波器采集联动 + 实时波形显示
8. 通讯日志：显示原始通讯代码，并自动用中文解释含义

通信方式：
Python 程序作为 TCP Client，连接厂家软件，例如 127.0.0.1:8888。
厂家软件需要先打开，并开启 TCP/Remote 通讯服务。

注意：
厂家给的通讯指令集中没有真正的“连续点动 Jog”指令。
V4 的“按住点动”改为安全分段点动：
按住鼠标时，程序一段一段发送较小的 Move;Rel；
松开鼠标后，不再发送新的 Move，并立即发送 Stop。
这样不会因为一次性发送了很长的 Move;Rel 而导致松开鼠标后平台仍继续运动很久。
真正安全急停仍然应使用设备上的硬件急停按钮。
"""

import socket
import threading
import queue
import time
import os
import csv
import json
import math
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from xml.sax.saxutils import escape
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


AXIS_NAME = {
    0: "X",
    1: "Y",
    2: "Z",
}

DEFAULT_SOFT_LIMITS = {
    0: (-600.0, 0.0),
    1: (-420.0, 200.0),
    2: (-380.0, 210.0),
}

POSITION_POLL_INTERVAL_MS = 100
MOVE_POLL_INTERVAL_S = 0.1
MOVE_POSITION_TOLERANCE_MM = 0.01

XLSX_MAX_DATA_ROWS = 1_048_575  # Excel 总行数上限减去表头
SCAN_SUMMARY_FIELDS = (
    "point_index", "line_index", "point_in_line", "plane", "span_axis", "step_axis",
    "x_mm", "y_mm", "z_mm", "vpp_v", "fs_hz", "sample_len", "xlsx_file",
)


def scan_grid_index(point_in_line, points_per_line, direction):
    """把 S 型往返的物理采集顺序映射为坐标递增的数据列。"""
    return point_in_line - 1 if direction > 0 else points_per_line - point_in_line


def format_duration(seconds):
    seconds = max(0, int(round(float(seconds))))
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)
    return "{:02d}:{:02d}:{:02d}".format(hours, minutes, seconds)


def _xlsx_cell(ref, value):
    """生成一个最小 XLSX 单元格；只使用 Python 标准库。"""
    if isinstance(value, str):
        return '<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'.format(
            ref, escape(value)
        )
    try:
        number = float(value)
    except (TypeError, ValueError):
        return '<c r="{}" t="inlineStr"><is><t>{}</t></is></c>'.format(
            ref, escape(str(value))
        )
    if not math.isfinite(number):
        return '<c r="{}"/>'.format(ref)
    return '<c r="{}"><v>{:.17g}</v></c>'.format(ref, number)


def _write_xlsx_sheet(archive, path, rows):
    with archive.open(path, "w") as stream:
        stream.write(
            b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        )
        for row_index, row in enumerate(rows, start=1):
            cells = []
            for column_index, value in enumerate(row, start=1):
                column = ""
                n = column_index
                while n:
                    n, remainder = divmod(n - 1, 26)
                    column = chr(65 + remainder) + column
                cells.append(_xlsx_cell("{}{}".format(column, row_index), value))
            stream.write(
                ('<row r="{}">{}</row>'.format(row_index, "".join(cells))).encode("utf-8")
            )
        stream.write(b"</sheetData></worksheet>")


def write_matlab_xlsx(filepath, metadata_pairs, time_s, wave_v):
    """写入 MATLAB 分析器需要的“元数据/波形数据”双工作表 XLSX。"""
    sample_count = min(len(time_s), len(wave_v))
    if sample_count > XLSX_MAX_DATA_ROWS:
        raise ValueError(
            "波形点数 {} 超过 XLSX 单表上限 {}，请降低采集点数".format(
                sample_count, XLSX_MAX_DATA_ROWS
            )
        )

    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    temp_path = filepath + ".tmp"

    def metadata_rows():
        yield ("Metadata", "Value")
        yield from metadata_pairs

    def waveform_rows():
        yield ("Time_s", "Voltage_V")
        yield from zip(time_s[:sample_count], wave_v[:sample_count])

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "[Content_Types].xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                '</Types>',
            )
            archive.writestr(
                "_rels/.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '</Relationships>',
            )
            archive.writestr(
                "xl/workbook.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="元数据" sheetId="1" r:id="rId1"/>'
                '<sheet name="波形数据" sheetId="2" r:id="rId2"/></sheets></workbook>',
            )
            archive.writestr(
                "xl/_rels/workbook.xml.rels",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                '</Relationships>',
            )
            archive.writestr(
                "xl/styles.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
                '<fills count="2"><fill><patternFill patternType="none"/></fill>'
                '<fill><patternFill patternType="gray125"/></fill></fills>'
                '<borders count="1"><border/></borders>'
                '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
                '<cellXfs count="1"><xf xfId="0"/></cellXfs>'
                '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                '</styleSheet>',
            )
            _write_xlsx_sheet(archive, "xl/worksheets/sheet1.xml", metadata_rows())
            _write_xlsx_sheet(archive, "xl/worksheets/sheet2.xml", waveform_rows())
        os.replace(temp_path, filepath)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


class RigolScope:
    """RIGOL 示波器 VISA 控制类。

    逻辑来自你发来的 rigol_scope.py：
    connect() 通过 pyvisa 打开 VISA 资源并查询 *IDN?；
    acquire_screen_norm() 使用 :WAVeform:DATA? 读取屏幕/运行波形；
    acquire_stop_mode() 先 :STOP 再用 RAW 模式读取大点数，最后恢复 :RUN。
    """

    def __init__(self):
        self.rm = None
        self.inst = None
        self.is_connected = False
        self.idn = ""
        self._screen_cache_key = None
        self._screen_preamble = None
        self._screen_time_key = None
        self._screen_time_s = None

    def connect(self, resource: str):
        try:
            try:
                import numpy  # noqa: F401
                import pyvisa
            except ImportError as exc:
                raise RuntimeError("缺少示波器依赖，请安装 numpy 和 pyvisa，并安装 NI-VISA/RIGOL VISA 驱动") from exc
            self.rm = pyvisa.ResourceManager()
            self.inst = self.rm.open_resource(resource)
            try:
                self.inst.chunk_size = 1000000
            except Exception:
                pass
            self.inst.timeout = 5000
            _idn = self.inst.query("*IDN?").strip()
            if "RIGOL" not in _idn.upper():
                raise RuntimeError(f"连接的设备不是 RIGOL 示波器: {_idn}")
            self.is_connected = True
            self.idn = _idn
            return _idn
        except Exception as e:
            self.disconnect()
            raise RuntimeError(f"示波器连接失败: {str(e)}") from e

    def disconnect(self):
        try:
            if self.inst is not None:
                self.inst.close()
        except Exception:
            pass
        finally:
            self.inst = None
            try:
                if self.rm is not None:
                    self.rm.close()
            except Exception:
                pass
            self.rm = None
            self.is_connected = False
            self.idn = ""
            self._screen_cache_key = None
            self._screen_preamble = None
            self._screen_time_key = None
            self._screen_time_s = None

    def write(self, cmd: str):
        if not self.is_connected:
            raise RuntimeError("Scope not connected")
        self.inst.write(cmd)

    def query(self, cmd: str) -> str:
        if not self.is_connected:
            raise RuntimeError("Scope not connected")
        return self.inst.query(cmd).strip()

    def query_status_summary(self) -> str:
        def q(cmd):
            try:
                return self.query(cmd)
            except Exception:
                return "N/A"
        run = q(":TRIGger:STATus?")
        sr = q(":ACQuire:SRATe?")
        md = q(":ACQuire:MDEPth?")
        return f"TRIG: {run}; SRATE: {sr}; MDEPTH: {md}"

    @staticmethod
    def _parse_tmc(raw: bytes) -> bytes:
        if not raw:
            return b""
        if raw[:1] == b"#" and len(raw) >= 3:
            try:
                n_digits = int(chr(raw[1]))
                data_length = int(raw[2:2+n_digits].decode("ascii", errors="ignore") or "0")
                start = 2 + n_digits
                end = start + data_length
                return raw[start:end]
            except Exception:
                return raw
        nl = raw.find(b"\n")
        if nl != -1:
            return raw[nl+1:]
        return raw

    def acquire_screen_norm(self, channel: int = 1, points: int | None = None,
                            time_start_us: float = None, time_end_us: float = None,
                            use_cache: bool = False):
        import numpy as np
        if not self.is_connected:
            raise RuntimeError("Scope not connected")
        try:
            ch = int(channel)
            acquire_points = int(points) if points is not None else 1000
            cache_key = (ch, acquire_points)
            # ponytail: 扫描期间时基/垂直档位必须保持不变；每次新扫描前会清空缓存。
            if use_cache and self._screen_cache_key == cache_key:
                preamble = self._screen_preamble
            else:
                self.write(f":WAVeform:SOURce CHANnel{ch}")
                self.write(":WAVeform:MODE MAXimum")
                self.write(":WAVeform:FORMat BYTE")
                self.write(f":WAVeform:POINts {acquire_points}")
                pre = self.query(":WAVeform:PREamble?")
                preamble = np.fromstring(pre, sep=",")
                if preamble.size < 10:
                    raise RuntimeError(f"Invalid PREamble: got {preamble.size} values, expected at least 10")
                if use_cache:
                    self._screen_cache_key = cache_key
                    self._screen_preamble = preamble
                else:
                    self._screen_cache_key = None
                    self._screen_preamble = None
                    self._screen_time_key = None
                    self._screen_time_s = None

            xincrement = float(preamble[4])
            xorigin = float(preamble[5])
            xreference = float(preamble[6])
            yincrement = float(preamble[7])
            yorigin = float(preamble[8])
            yreference = float(preamble[9])
            pts_reported = int(preamble[2])
            self.inst.write(":WAVeform:DATA?")
            raw = self.inst.read_raw()
            payload = self._parse_tmc(raw)
            if len(payload) == 0:
                raise RuntimeError("No waveform data received")

            data_u8 = np.frombuffer(payload, dtype=np.uint8)
            sample_len = int(data_u8.size)
            wave_v = (data_u8.astype(np.float64) - yreference) * yincrement + yorigin
            time_key = (cache_key, sample_len, xincrement, xorigin, xreference)
            if use_cache and self._screen_time_key == time_key:
                time_s = self._screen_time_s
            else:
                time_s = ((np.arange(sample_len, dtype=np.float64) - xreference) * xincrement + xorigin)
                if use_cache:
                    self._screen_time_key = time_key
                    self._screen_time_s = time_s

            if acquire_points > 0 and acquire_points < sample_len:
                wave_v = wave_v[:acquire_points]
                time_s = time_s[:acquire_points]
                sample_len = acquire_points

            fs = 1.0 / xincrement if xincrement > 0 else 0.0
            meta = {
                "xincrement": xincrement,
                "xorigin": xorigin,
                "xreference": xreference,
                "yincrement": yincrement,
                "yorigin": yorigin,
                "yreference": yreference,
                "points_reported": pts_reported,
                "points_actual": sample_len,
            }
            return wave_v.astype(np.float32), float(fs), int(sample_len), time_s.astype(np.float64), meta
        except Exception as e:
            raise RuntimeError(f"波形采集失败: {str(e)}") from e

    def acquire_stop_mode(self, channel: int = 1, points: int = 2000):
        import numpy as np
        if not self.is_connected:
            raise RuntimeError("Scope not connected")
        try:
            ch = int(channel)
            acquire_points = int(points)
            self.write(":STOP")
            time.sleep(0.1)
            self.write(f":WAVeform:SOURce CHANnel{ch}")
            self.write(":WAVeform:MODE RAW")
            self.write(":WAVeform:FORMat BYTE")
            self.write(f":WAVeform:POINts {acquire_points}")
            self.write(":WAVeform:STARt 1")
            self.write(f":WAVeform:STOP {acquire_points}")

            pre = self.query(":WAVeform:PREamble?")
            preamble = np.fromstring(pre, sep=",")
            if preamble.size < 10:
                raise RuntimeError(f"Invalid PREamble: got {preamble.size} values")
            xincrement = float(preamble[4])
            xorigin = float(preamble[5])
            xreference = float(preamble[6])
            yincrement = float(preamble[7])
            yorigin = float(preamble[8])
            yreference = float(preamble[9])
            pts_reported = int(preamble[2])
            self.inst.write(":WAVeform:DATA?")
            raw = self.inst.read_raw()
            payload = self._parse_tmc(raw)
            if len(payload) == 0:
                raise RuntimeError("No waveform data received")

            data_u8 = np.frombuffer(payload, dtype=np.uint8)
            sample_len = int(data_u8.size)
            wave_v = (data_u8.astype(np.float64) - yreference) * yincrement + yorigin
            time_s = ((np.arange(sample_len, dtype=np.float64) - xreference) * xincrement + xorigin)

            if acquire_points > 0 and acquire_points < sample_len:
                wave_v = wave_v[:acquire_points]
                time_s = time_s[:acquire_points]
                sample_len = acquire_points
            fs = 1.0 / xincrement if xincrement > 0 else 0.0
            meta = {
                "xincrement": xincrement,
                "xorigin": xorigin,
                "xreference": xreference,
                "yincrement": yincrement,
                "yorigin": yorigin,
                "yreference": yreference,
                "points_reported": pts_reported,
                "points_actual": sample_len,
            }
            return wave_v.astype(np.float32), float(fs), int(sample_len), time_s.astype(np.float64), meta
        except Exception as e:
            raise RuntimeError(f"STOP模式采集失败: {str(e)}") from e
        finally:
            self._screen_cache_key = None
            self._screen_preamble = None
            self._screen_time_key = None
            self._screen_time_s = None
            try:
                self.write(":RUN")
            except Exception:
                pass

    def set_average_mode(self, count: int = 4):
        if not self.is_connected:
            return
        self.write(":ACQuire:TYPE AVERage")
        self.write(f":ACQuire:AVERage {int(count)}")

    def set_normal_mode(self):
        if not self.is_connected:
            return
        self.write(":ACQuire:TYPE NORMal")


class MotionClient:
    """负责和厂家软件进行 TCP 通讯"""

    def __init__(self):
        self.sock = None
        self.host = ""
        self.port = 0
        self.timeout = 5.0
        self.command_lock = threading.Lock()
        self.send_lock = threading.Lock()
        self.rx_buffer = b""
        self.stop_requested = threading.Event()

    @property
    def connected(self):
        return self.sock is not None

    def connect(self, host, port, timeout=5.0):
        self.disconnect()

        self.host = str(host).strip()
        self.port = int(port)
        self.timeout = float(timeout)

        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.settimeout(self.timeout)
            s.connect((self.host, self.port))
        except Exception:
            s.close()
            raise
        self.sock = s
        self.rx_buffer = b""
        self.stop_requested.clear()

    def disconnect(self):
        if self.sock is not None:
            try:
                self.sock.close()
            except Exception:
                pass
            self.sock = None
            self.rx_buffer = b""
            self.stop_requested.set()

    def clear_stop_request(self):
        self.stop_requested.clear()

    def _recv_line(self, sock, deadline):
        while b"\n" not in self.rx_buffer:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TimeoutError("等待回复超时")
            sock.settimeout(remaining)
            chunk = sock.recv(4096)
            if not chunk:
                raise ConnectionError("连接已经断开")
            self.rx_buffer += chunk
            if len(self.rx_buffer) > 65536:
                raise RuntimeError("厂家软件回复超过 64 KB，协议可能已失步")
        line, self.rx_buffer = self.rx_buffer.split(b"\n", 1)
        return line.rstrip(b"\r").decode("ascii", errors="ignore").strip()

    @staticmethod
    def _reply_matches(cmd, reply):
        key = str(cmd).split(";", 1)[0].strip().lower()
        lower = str(reply).strip().lower()
        expected = {
            "remote": "remoteok",
            "quitremote": "quitremoteok",
            "move": "moveok",
            "getposition": "position;",
            "stop": "stopok",
            "gohome": "gohomeok",
        }.get(key)
        if expected is None:
            return True
        return lower.startswith(expected)

    def send_command(self, cmd, timeout=None):
        """
        发送一条指令，并等待一行回复。
        厂家通讯指令通常以 \r\n 结尾。
        """
        if self.sock is None:
            raise RuntimeError("尚未连接厂家软件")

        if timeout is None:
            timeout = self.timeout

        cmd = str(cmd).strip()
        if not cmd:
            raise ValueError("指令不能为空")

        command_key = cmd.split(";", 1)[0].strip().lower()
        full_cmd = (cmd + "\r\n").encode("ascii", errors="ignore")

        with self.command_lock:
            sock = self.sock
            if sock is None:
                raise RuntimeError("尚未连接厂家软件")
            with self.send_lock:
                # 必须在发送锁内检查，保证 Stop 之后绝不会再发出新的 Move。
                if command_key == "move" and self.stop_requested.is_set():
                    raise InterruptedError("停止状态已锁存，拒绝发送新的运动指令")
                sock.sendall(full_cmd)

            deadline = time.monotonic() + float(timeout)
            ignored = 0
            while True:
                try:
                    reply = self._recv_line(sock, deadline)
                except socket.timeout as exc:
                    raise TimeoutError("等待回复超时：" + cmd) from exc
                if self._reply_matches(cmd, reply):
                    return reply
                if reply.lower() == "stopok" and self.stop_requested.is_set():
                    raise InterruptedError("运动已由 Stop 指令终止")
                ignored += 1
                if ignored > 20:
                    raise RuntimeError("连续收到不匹配回复，TCP 协议已失步")

    def send_no_wait(self, cmd):
        """发送指令但不等待回复；厂家 Move 指令实际不会返回可用的完成回复。"""
        cmd = str(cmd).strip()
        if not cmd:
            raise ValueError("指令不能为空")
        command_key = cmd.split(";", 1)[0].strip().lower()
        full_cmd = (cmd + "\r\n").encode("ascii", errors="ignore")

        with self.command_lock:
            sock = self.sock
            if sock is None:
                raise RuntimeError("尚未连接厂家软件")
            with self.send_lock:
                if command_key == "move" and self.stop_requested.is_set():
                    raise InterruptedError("停止状态已锁存，拒绝发送新的运动指令")
                sock.sendall(full_cmd)

    def stop_no_wait(self):
        """
        尽快发送 Stop，不等待回复。
        用于松开点动按钮或紧急停止。
        """
        if self.sock is None:
            raise RuntimeError("尚未连接厂家软件")
        try:
            self.stop_requested.set()
            with self.send_lock:
                self.sock.sendall(b"Stop\r\n")
        except Exception:
            self.disconnect()
            raise


class MotionApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Motion - Python 控制端 V13 优化版")
        self.geometry("1280x820")
        self.minsize(1050, 720)

        try:
            self.option_add("*Font", ("Microsoft YaHei UI", 10))
        except Exception:
            pass

        self.client = MotionClient()
        self.scope = RigolScope()
        self.scope_lock = threading.Lock()  # 防止实时显示、单次采集、扫描采集同时访问示波器
        self.log_queue = queue.Queue()
        self.ui_queue = queue.Queue()
        self.motion_operation_lock = threading.Lock()
        self.software_stop_event = threading.Event()
        self.position_lock = threading.Lock()
        self.current_position = {0: None, 1: None, 2: None}
        self.position_poll_busy = False
        self.position_poll_error_logged = False
        self.connection_attempt_lock = threading.Lock()

        # 当前坐标显示
        self.pos_vars = {
            0: tk.StringVar(value="连接中..."),
            1: tk.StringVar(value="连接中..."),
            2: tk.StringVar(value="连接中..."),
        }

        # 相对运动距离
        self.rel_vars = {
            0: tk.StringVar(value="10.000"),
            1: tk.StringVar(value="10.000"),
            2: tk.StringVar(value="10.000"),
        }

        # 绝对目标位置
        self.abs_vars = {
            0: tk.StringVar(value="0.000"),
            1: tk.StringVar(value="0.000"),
            2: tk.StringVar(value="0.000"),
        }

        self.speed_var = tk.DoubleVar(value=10)
        self.speed_label_var = tk.StringVar(value="10%")
        self.jog_distance_var = tk.StringVar(value="0.5")
        self.auto_query_var = tk.BooleanVar(value=True)
        self.soft_limit_min_vars = {
            axis: tk.StringVar(value=self.num_text(DEFAULT_SOFT_LIMITS[axis][0]))
            for axis in (0, 1, 2)
        }
        self.soft_limit_max_vars = {
            axis: tk.StringVar(value=self.num_text(DEFAULT_SOFT_LIMITS[axis][1]))
            for axis in (0, 1, 2)
        }

        self.hold_event = None
        self.hold_thread = None
        self.hold_info = None
        self.jog_run_id = 0

        # 扫描轨迹参数
        self.scan_stop_event = threading.Event()
        self.scan_thread = None
        self.scan_running = False
        self.scan_plane_var = tk.StringVar(value="XY")
        self.scan_span_axis_var = tk.StringVar(value="X")
        self.scan_step_axis_var = tk.StringVar(value="Y")
        self.scan_direction_var = tk.StringVar(value="自动")
        self.scan_step_direction_var = tk.StringVar(value="自动")
        self.scan_start_position = None  # {0: X, 1: Y, 2: Z}，开始扫描时记录，用于扫描后返回起点
        self.last_scan_params = None
        self.scan_span_var = tk.StringVar(value="10.0")
        self.scan_span_step_var = tk.StringVar(value="1.0")
        self.scan_line_step_var = tk.StringVar(value="1.0")
        self.scan_lines_var = tk.StringVar(value="10")
        self.scan_dwell_var = tk.StringVar(value="0.02")
        self.scan_preview_var = tk.StringVar(value="扫描预览：未计算")
        self.scan_progress_var = tk.StringVar(value="扫描状态：未开始")
        self.scan_relative_position_var = tk.StringVar(value="相对扫描起点：尚未记录")
        self.scan_relative_target_vars = {
            axis: tk.StringVar(value="0.000") for axis in (0, 1, 2)
        }
        self.scan_session_dir = None
        self.scan_summary_stream = None
        self.scan_summary_writer = None

        # 示波器/采集参数
        self.scope_visa_var = tk.StringVar(value="USB0::0x1AB1::0x0515::YOUR_SERIAL_NUMBER::INSTR")
        self.scope_channel_var = tk.StringVar(value="1")
        self.scope_points_k_var = tk.StringVar(value="1")
        self.scope_avg_var = tk.StringVar(value="无")
        self.scope_mode_var = tk.StringVar(value="自动")
        self.scope_status_var = tk.StringVar(value="示波器：未连接")
        self.save_dir_var = tk.StringVar(value=os.path.join(os.path.expanduser("~"), "motion_scope_data"))
        self.scan_acquire_scope_var = tk.BooleanVar(value=True)
        self.save_npy_var = tk.BooleanVar(value=True)
        self.save_each_point_var = tk.BooleanVar(value=False)
        self.scan_live_plot_var = tk.BooleanVar(value=False)

        # 示波器实时波形显示
        self.scope_live_running = False
        self.scope_live_busy = False
        self.scope_live_interval_var = tk.StringVar(value="500")  # ms
        self.wave_info_var = tk.StringVar(value="波形：未采集")
        self.last_waveform = None  # (time_s, wave_v, fs, n, vpp)

        self._build_ui()

        # 即使鼠标松开时不在按钮上，也尽量捕获全局松开事件，防止点动不停。
        self.bind_all("<ButtonRelease-1>", self.stop_jog_hold)

        self._poll_log_queue()
        self._poll_ui_queue()
        self.after(POSITION_POLL_INTERVAL_MS, self._poll_position)
        self.after(250, lambda: self.on_connect(automatic=True))
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ============================================================
    # UI
    # ============================================================
    def _build_ui(self):
        # V10：整体内容放进 Canvas + 垂直滚动条。
        # 窗口高度不够时，可以用右侧滚动条或鼠标滚轮向下查看。
        outer = ttk.Frame(self)
        outer.pack(fill=tk.BOTH, expand=True)

        self.main_canvas = tk.Canvas(outer, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)

        self.main_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        main = ttk.Frame(self.main_canvas, padding=10)
        self.main_window_id = self.main_canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_frame_configure(event=None):
            self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

        def _on_canvas_configure(event):
            self.main_canvas.itemconfigure(self.main_window_id, width=event.width)

        main.bind("<Configure>", _on_frame_configure)
        self.main_canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            if event.delta:
                self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _on_linux_wheel_up(event):
            self.main_canvas.yview_scroll(-3, "units")

        def _on_linux_wheel_down(event):
            self.main_canvas.yview_scroll(3, "units")

        self.bind_all("<MouseWheel>", _on_mousewheel)
        self.bind_all("<Button-4>", _on_linux_wheel_up)
        self.bind_all("<Button-5>", _on_linux_wheel_down)

        self._build_connection_area(main)
        self._build_control_area(main)
        self._build_scope_area(main)
        self._build_trajectory_area(main)
        self._build_raw_and_log_area(main)

    def _build_connection_area(self, parent):
        frame = ttk.LabelFrame(parent, text="连接厂家软件 TCP 服务", padding=8)
        frame.pack(fill=tk.X, pady=(0, 8))

        ttk.Label(frame, text="Host:").grid(row=0, column=0, padx=5, pady=4, sticky=tk.E)
        self.host_var = tk.StringVar(value="127.0.0.1")
        ttk.Entry(frame, textvariable=self.host_var, width=18).grid(row=0, column=1, padx=5, pady=4)

        ttk.Label(frame, text="Port:").grid(row=0, column=2, padx=5, pady=4, sticky=tk.E)
        self.port_var = tk.StringVar(value="8888")
        ttk.Entry(frame, textvariable=self.port_var, width=10).grid(row=0, column=3, padx=5, pady=4)

        ttk.Label(frame, text="超时(s):").grid(row=0, column=4, padx=5, pady=4, sticky=tk.E)
        self.timeout_var = tk.StringVar(value="5")
        ttk.Entry(frame, textvariable=self.timeout_var, width=8).grid(row=0, column=5, padx=5, pady=4)

        ttk.Button(frame, text="连接", command=self.on_connect).grid(row=0, column=6, padx=8, pady=4)
        ttk.Button(frame, text="断开", command=self.on_disconnect).grid(row=0, column=7, padx=5, pady=4)

        ttk.Button(frame, text="进入远程 Remote", command=self.enter_remote).grid(row=0, column=8, padx=(20, 5), pady=4)
        ttk.Button(frame, text="退出远程", command=self.quit_remote).grid(row=0, column=9, padx=5, pady=4)

        self.status_var = tk.StringVar(value="未连接")
        ttk.Label(frame, textvariable=self.status_var).grid(row=0, column=10, padx=15, pady=4, sticky=tk.W)

        frame.columnconfigure(11, weight=1)

    def _build_control_area(self, parent):
        frame = ttk.LabelFrame(parent, text="控制", padding=12)
        frame.pack(fill=tk.X)

        # 表头
        header_font = ("Microsoft YaHei UI", 15, "bold")
        ttk.Label(frame, text="当前方位", font=header_font).grid(row=0, column=1, columnspan=2, pady=(0, 8))
        ttk.Label(frame, text="相对运动", font=header_font).grid(row=0, column=3, columnspan=3, pady=(0, 8))
        ttk.Label(frame, text="绝对运动", font=header_font).grid(row=0, column=6, columnspan=2, pady=(0, 8))
        ttk.Label(frame, text="点动", font=header_font).grid(row=0, column=8, columnspan=2, pady=(0, 8))

        for r, axis in enumerate([0, 1, 2], start=1):
            axis_name = AXIS_NAME[axis]

            ttk.Label(frame, text=axis_name, font=("Microsoft YaHei UI", 13, "bold")).grid(
                row=r, column=0, padx=(5, 16), pady=8, sticky=tk.E
            )

            # 当前位置
            pos_entry = ttk.Entry(frame, textvariable=self.pos_vars[axis], width=16, justify="right", state="readonly")
            pos_entry.grid(row=r, column=1, padx=4, pady=8)
            ttk.Label(frame, text="mm").grid(row=r, column=2, padx=(0, 18), pady=8, sticky=tk.W)

            # 相对运动：单击一次移动设定距离
            minus_rel = ttk.Button(frame, text="-", width=4, command=lambda a=axis: self.relative_move(a, -1))
            minus_rel.grid(row=r, column=3, padx=3, pady=8)

            rel_spin = tk.Spinbox(
                frame,
                textvariable=self.rel_vars[axis],
                from_=0,
                to=100000,
                increment=0.1,
                width=10,
                justify="right",
                format="%.3f",
            )
            rel_spin.grid(row=r, column=4, padx=3, pady=8)

            plus_rel = ttk.Button(frame, text="+", width=4, command=lambda a=axis: self.relative_move(a, 1))
            plus_rel.grid(row=r, column=5, padx=(3, 25), pady=8)

            # 绝对运动
            abs_spin = tk.Spinbox(
                frame,
                textvariable=self.abs_vars[axis],
                from_=-100000,
                to=100000,
                increment=0.1,
                width=10,
                justify="right",
                format="%.3f",
            )
            abs_spin.grid(row=r, column=6, padx=3, pady=8)

            ttk.Button(frame, text="Go", width=4, command=lambda a=axis: self.absolute_move(a)).grid(
                row=r, column=7, padx=(3, 25), pady=8
            )

            # 点动：按住连续移动，松开停止
            jog_minus = tk.Button(frame, text="-", width=4, height=1)
            jog_minus.grid(row=r, column=8, padx=5, pady=8)
            jog_minus.bind("<ButtonPress-1>", lambda event, a=axis: self.start_jog_hold(a, -1))
            jog_minus.bind("<ButtonRelease-1>", self.stop_jog_hold)
            jog_minus.bind("<Leave>", self.stop_jog_hold)

            jog_plus = tk.Button(frame, text="+", width=4, height=1)
            jog_plus.grid(row=r, column=9, padx=5, pady=8)
            jog_plus.bind("<ButtonPress-1>", lambda event, a=axis: self.start_jog_hold(a, 1))
            jog_plus.bind("<ButtonRelease-1>", self.stop_jog_hold)
            jog_plus.bind("<Leave>", self.stop_jog_hold)

        # 速度滑条
        ttk.Label(frame, text="速度", font=("Microsoft YaHei UI", 13, "bold")).grid(
            row=4, column=0, padx=5, pady=(25, 10), sticky=tk.E
        )

        speed_scale = ttk.Scale(
            frame,
            from_=1,
            to=100,
            orient=tk.HORIZONTAL,
            variable=self.speed_var,
            command=self.on_speed_change,
        )
        speed_scale.grid(row=4, column=1, columnspan=8, padx=8, pady=(25, 10), sticky=tk.EW)
        ttk.Label(frame, textvariable=self.speed_label_var, width=6).grid(
            row=4, column=9, padx=5, pady=(25, 10), sticky=tk.W
        )

        # 点动参数和自动查询
        jog_option_frame = ttk.Frame(frame)
        jog_option_frame.grid(row=5, column=1, columnspan=9, sticky=tk.W, pady=(0, 8))

        ttk.Label(jog_option_frame, text="点动分段距离(mm/段):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(jog_option_frame, textvariable=self.jog_distance_var, width=8).pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(
            jog_option_frame,
            text="运动中实时刷新位置（100 ms，固定开启）",
            variable=self.auto_query_var,
            state="disabled",
        ).pack(side=tk.LEFT)

        limit_frame = ttk.Frame(frame)
        limit_frame.grid(row=6, column=0, columnspan=10, sticky=tk.W, pady=(2, 8))
        ttk.Label(
            limit_frame,
            text="软件限位(mm，请按设备实际安全行程设置):",
            foreground="#c00000",
        ).pack(side=tk.LEFT, padx=(5, 10))
        for axis in (0, 1, 2):
            ttk.Label(limit_frame, text="{}".format(AXIS_NAME[axis])).pack(side=tk.LEFT, padx=(8, 2))
            ttk.Entry(limit_frame, textvariable=self.soft_limit_min_vars[axis], width=8).pack(side=tk.LEFT)
            ttk.Label(limit_frame, text="～").pack(side=tk.LEFT, padx=2)
            ttk.Entry(limit_frame, textvariable=self.soft_limit_max_vars[axis], width=8).pack(side=tk.LEFT)

        # 状态区域
        status_frame = ttk.Frame(parent, padding=(0, 8, 0, 0))
        status_frame.pack(fill=tk.X)

        self.card_status_var = tk.StringVar(value="控制卡：未连接")
        self.estop_status_var = tk.StringVar(value="急 停：未知")
        self.error_status_var = tk.StringVar(value="轴错误：未知")

        ttk.Label(status_frame, textvariable=self.card_status_var).pack(side=tk.LEFT, padx=(0, 30))
        ttk.Label(status_frame, textvariable=self.estop_status_var).pack(side=tk.LEFT, padx=(0, 30))
        ttk.Label(status_frame, textvariable=self.error_status_var).pack(side=tk.LEFT, padx=(0, 30))

        stop_btn = tk.Button(
            status_frame,
            text="停止运动",
            bg="#ffdddd",
            fg="#d7191c",
            activebackground="#ffbbbb",
            font=("Microsoft YaHei UI", 14, "bold"),
            command=self.emergency_stop,
            width=12,
        )
        stop_btn.pack(side=tk.RIGHT, padx=5)

        for col in range(10):
            frame.columnconfigure(col, weight=0)
        frame.columnconfigure(4, weight=1)
        frame.columnconfigure(6, weight=1)

    def _build_scope_area(self, parent):
        """示波器连接和采集设置。"""
        frame = ttk.LabelFrame(parent, text="示波器与采集联动", padding=10)
        frame.pack(fill=tk.X, pady=(8, 0))

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="VISA地址:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row1, textvariable=self.scope_visa_var, width=48).pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)
        ttk.Button(row1, text="连接示波器", command=self.on_scope_connect).pack(side=tk.LEFT, padx=4)
        ttk.Button(row1, text="断开", command=self.on_scope_disconnect).pack(side=tk.LEFT, padx=4)
        ttk.Button(row1, text="查询状态", command=self.on_scope_status).pack(side=tk.LEFT, padx=4)
        ttk.Label(row1, textvariable=self.scope_status_var).pack(side=tk.LEFT, padx=(12, 0))

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="通道:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(row2, textvariable=self.scope_channel_var, values=["1", "2", "3", "4"], width=5, state="readonly").pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="点数(K):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row2, textvariable=self.scope_points_k_var, width=8).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="平均:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(row2, textvariable=self.scope_avg_var, values=["无", "2次", "4次", "8次", "16次", "32次", "64次"], width=8, state="readonly").pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(row2, text="模式:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(row2, textvariable=self.scope_mode_var, values=["自动", "屏幕MAX", "STOP-RAW"], width=10, state="readonly").pack(side=tk.LEFT, padx=(0, 12))

        ttk.Button(row2, text="单次采集并保存", command=self.on_single_acquire).pack(side=tk.LEFT, padx=(8, 12))

        save_options = ttk.Frame(frame)
        save_options.pack(fill=tk.X, pady=3)
        ttk.Checkbutton(save_options, text="扫描每点采集", variable=self.scan_acquire_scope_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(save_options, text="连续 NPY 数据集（推荐）", variable=self.save_npy_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(save_options, text="扫描时刷新波形（慢）", variable=self.scan_live_plot_var).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(save_options, text="每点保存 MATLAB XLSX（慢）", variable=self.save_each_point_var).pack(side=tk.LEFT)

        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="保存路径:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row3, textvariable=self.save_dir_var, width=60).pack(side=tk.LEFT, padx=(0, 8), fill=tk.X, expand=True)
        ttk.Button(row3, text="选择", command=self.browse_save_dir).pack(side=tk.LEFT, padx=4)
        ttk.Label(row3, text="说明：每点先配置采集、再等待和保存；使用平均时等待时间应覆盖平均所需触发次数。").pack(side=tk.LEFT, padx=(12, 0))

        # 实时波形显示区域
        wave_frame = ttk.LabelFrame(frame, text="示波器实时波形显示", padding=8)
        wave_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        wave_ctrl = ttk.Frame(wave_frame)
        wave_ctrl.pack(fill=tk.X, pady=(0, 6))

        self.btn_scope_live_start = ttk.Button(wave_ctrl, text="开始实时显示", command=self.start_scope_live)
        self.btn_scope_live_start.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_scope_live_stop = ttk.Button(wave_ctrl, text="停止实时显示", command=self.stop_scope_live)
        self.btn_scope_live_stop.pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(wave_ctrl, text="刷新间隔(ms):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(wave_ctrl, textvariable=self.scope_live_interval_var, width=8).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Label(wave_ctrl, textvariable=self.wave_info_var).pack(side=tk.LEFT, padx=(10, 0))

        self.wave_canvas = tk.Canvas(
            wave_frame,
            height=260,
            bg="#0b1020",
            highlightthickness=1,
            highlightbackground="#94a3b8"
        )
        self.wave_canvas.pack(fill=tk.X, expand=True)

        self.clear_waveform_plot("等待采集波形...")

    def _build_trajectory_area(self, parent):
        """
        S 型扫描轨迹设置。
        参照你发来的程序：行程轴负责一行内往返扫描，行数轴负责每行结束后的换行。
        """
        frame = ttk.LabelFrame(parent, text="扫描轨迹设置（S型轨迹）", padding=10)
        frame.pack(fill=tk.X, pady=(8, 0))

        # 第一行：扫描面与轴选择
        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=3)

        ttk.Label(row1, text="扫描面:").pack(side=tk.LEFT, padx=(0, 5))
        plane_cb = ttk.Combobox(row1, textvariable=self.scan_plane_var, values=["XY", "XZ", "YZ"], width=6, state="readonly")
        plane_cb.pack(side=tk.LEFT, padx=(0, 15))
        plane_cb.bind("<<ComboboxSelected>>", lambda event: self.on_scan_plane_changed())

        ttk.Label(row1, text="行程轴:").pack(side=tk.LEFT, padx=(0, 5))
        span_axis_cb = ttk.Combobox(row1, textvariable=self.scan_span_axis_var, values=["X", "Y", "Z"], width=6, state="readonly")
        span_axis_cb.pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row1, text="行数轴/步距轴:").pack(side=tk.LEFT, padx=(0, 5))
        step_axis_cb = ttk.Combobox(row1, textvariable=self.scan_step_axis_var, values=["X", "Y", "Z"], width=6, state="readonly")
        step_axis_cb.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(row1, text="首行方向:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(
            row1,
            textvariable=self.scan_direction_var,
            values=["自动", "正向（小→大）", "反向（大→小）"],
            width=16,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row1, text="点位等待(s):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row1, textvariable=self.scan_dwell_var, width=8).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row1, text="第三轴保持不动。").pack(side=tk.LEFT, padx=(10, 0))

        # 第二行：参数输入
        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=3)

        ttk.Label(row2, text="行程(mm):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row2, textvariable=self.scan_span_var, width=9).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row2, text="行内步距(mm):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row2, textvariable=self.scan_span_step_var, width=9).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row2, text="行步距(mm):").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row2, textvariable=self.scan_line_step_var, width=9).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row2, text="行数:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(row2, textvariable=self.scan_lines_var, width=7).pack(side=tk.LEFT, padx=(0, 15))

        ttk.Label(row2, text="换行方向:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Combobox(
            row2,
            textvariable=self.scan_step_direction_var,
            values=["自动", "正向（小→大）", "反向（大→小）"],
            width=16,
            state="readonly",
        ).pack(side=tk.LEFT)

        # 第三行：扫描控制按钮
        # V6 中“开始扫描”放在预览文字同一行，窗口较窄时会被挤出可视区域。
        # V7 将按钮单独放一行，并把“开始扫描”做成明显的绿色大按钮。
        row3_btn = ttk.Frame(frame)
        row3_btn.pack(fill=tk.X, pady=(8, 4))

        self.btn_start_scan = tk.Button(
            row3_btn,
            text="▶ 开始S型扫描",
            bg="#d9f7d9",
            fg="#087a08",
            activebackground="#bff0bf",
            font=("Microsoft YaHei UI", 11, "bold"),
            command=self.start_trajectory_scan,
            width=16,
        )
        self.btn_start_scan.pack(side=tk.LEFT, padx=(0, 10))

        self.btn_stop_scan = tk.Button(
            row3_btn,
            text="■ 停止扫描",
            bg="#ffe5e5",
            fg="#c00000",
            activebackground="#ffcccc",
            font=("Microsoft YaHei UI", 11, "bold"),
            command=self.stop_trajectory_scan,
            width=12,
        )
        self.btn_stop_scan.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(row3_btn, text="刷新预览", command=self.update_scan_preview).pack(side=tk.LEFT, padx=(0, 10))

        self.btn_return_start = tk.Button(
            row3_btn,
            text="↩ 回到扫描起点",
            bg="#e8f0ff",
            fg="#1d4ed8",
            activebackground="#dbeafe",
            font=("Microsoft YaHei UI", 10, "bold"),
            command=self.return_to_scan_start,
            width=14,
        )
        self.btn_return_start.pack(side=tk.LEFT, padx=(0, 10))

        # 第四行：预览和状态
        row4_preview = ttk.Frame(frame)
        row4_preview.pack(fill=tk.X, pady=(4, 0))

        ttk.Label(row4_preview, textvariable=self.scan_preview_var).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Label(row4_preview, textvariable=self.scan_progress_var).pack(side=tk.LEFT, padx=(0, 15))

        relative_row = ttk.Frame(frame)
        relative_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(
            relative_row,
            textvariable=self.scan_relative_position_var,
            foreground="#1d4ed8",
        ).pack(side=tk.LEFT)

        relative_target_row = ttk.Frame(frame)
        relative_target_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(relative_target_row, text="移动到相对扫描起点：").pack(side=tk.LEFT, padx=(0, 5))
        for axis in (0, 1, 2):
            ttk.Label(relative_target_row, text=AXIS_NAME[axis]).pack(side=tk.LEFT, padx=(8, 2))
            ttk.Entry(
                relative_target_row,
                textvariable=self.scan_relative_target_vars[axis],
                width=9,
            ).pack(side=tk.LEFT)
        ttk.Label(relative_target_row, text="mm").pack(side=tk.LEFT, padx=(5, 10))
        ttk.Button(
            relative_target_row,
            text="移动",
            command=self.move_to_scan_relative_position,
        ).pack(side=tk.LEFT)

        # 参数改变时自动刷新预览
        for var in [
            self.scan_plane_var,
            self.scan_span_axis_var,
            self.scan_step_axis_var,
            self.scan_direction_var,
            self.scan_step_direction_var,
            self.scan_span_var,
            self.scan_span_step_var,
            self.scan_line_step_var,
            self.scan_lines_var,
            self.scan_dwell_var,
        ]:
            var.trace_add("write", lambda *_: self.update_scan_preview())

        self.update_scan_preview()

    def _build_raw_and_log_area(self, parent):
        raw_frame = ttk.LabelFrame(parent, text="原始指令测试", padding=8)
        raw_frame.pack(fill=tk.X, pady=(8, 0))

        self.raw_cmd_var = tk.StringVar(value="GetPosition;0,1,2")
        ttk.Entry(raw_frame, textvariable=self.raw_cmd_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(raw_frame, text="发送原始指令", command=self.send_raw_command).pack(side=tk.RIGHT)

        log_frame = ttk.LabelFrame(parent, text="通讯日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scroll.set)

    # ============================================================
    # 日志与中文解释
    # ============================================================
    def log(self, text):
        now = time.strftime("%H:%M:%S")
        self.log_queue.put("[{}] {}\n".format(now, text))

    def _poll_log_queue(self):
        try:
            while True:
                text = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, text)
                self.log_text.see(tk.END)
        except queue.Empty:
            pass

        self.after(80, self._poll_log_queue)

    def post_ui(self, func, *args, **kwargs):
        self.ui_queue.put((func, args, kwargs))

    def _poll_ui_queue(self):
        try:
            while True:
                func, args, kwargs = self.ui_queue.get_nowait()
                func(*args, **kwargs)
        except queue.Empty:
            pass
        self.after(40, self._poll_ui_queue)

    def describe_command(self, cmd):
        cmd = str(cmd).strip()
        parts = cmd.split(";")
        key = parts[0].lower()

        if key == "remote":
            return "进入远程控制模式，厂家软件开始侦听外部指令。"

        if key == "quitremote":
            return "退出远程控制模式，之后可继续用厂家软件手动控制。"

        if key == "stop":
            return "停止当前运动。"

        if key == "gohome":
            return "执行机械归零。根据当前指令集，GoHome 是所有轴归零。"

        if key == "getposition":
            if len(parts) >= 2:
                axes = self.axes_text(parts[1])
                return "查询 {} 轴当前绝对位置，单位 mm。".format(axes)
            return "查询当前位置。"

        if key == "move" and len(parts) >= 5:
            mode = parts[1]
            speed = parts[2]
            axes = self.parse_axes(parts[3])
            values = self.parse_float_list(parts[4])
            mode_text = "相对运动" if mode.lower() == "rel" else "绝对运动" if mode.lower() == "abs" else mode

            details = []
            for i, axis in enumerate(axes):
                if i < len(values):
                    name = AXIS_NAME.get(axis, str(axis))
                    if mode.lower() == "rel":
                        details.append("{}轴移动 {} mm".format(name, values[i]))
                    else:
                        details.append("{}轴到达 {} mm".format(name, values[i]))

            return "{}，速度 {}%，{}。".format(mode_text, speed, "，".join(details))

        return "自定义/未知指令，请参考通讯指令集。"

    def describe_response(self, text):
        text = str(text).strip()
        lower = text.lower()

        if lower == "remoteok":
            return "厂家软件已进入远程控制模式。"

        if lower == "quitremoteok":
            return "厂家软件已退出远程控制模式。"

        if lower == "stopok":
            return "停止指令已执行。"

        if lower == "moveok":
            return "移动指令完成，平台已运动到指定位置。"

        if lower == "gohomeok":
            return "机械归零完成。"

        if lower.startswith("position;"):
            desc = self.parse_position_description(text)
            if desc:
                return "返回当前位置：{}。".format(desc)
            return "返回当前位置数据。"

        if not text:
            return "没有收到回复。"

        return "厂家软件返回：{}。".format(text)

    def axes_text(self, axes_str):
        axes = self.parse_axes(axes_str)
        if not axes:
            return axes_str
        return "/".join(AXIS_NAME.get(a, str(a)) for a in axes)

    def parse_axes(self, axes_str):
        result = []
        for item in str(axes_str).split(","):
            item = item.strip()
            if not item:
                continue
            try:
                result.append(int(item))
            except ValueError:
                pass
        return result

    def parse_float_list(self, value_str):
        result = []
        for item in str(value_str).split(","):
            item = item.strip()
            if not item:
                continue
            try:
                result.append(float(item))
            except ValueError:
                pass
        return result

    def parse_position_description(self, text):
        parts = str(text).strip().split(";")
        if len(parts) < 3:
            return ""

        axes = self.parse_axes(parts[1])
        values = self.parse_float_list(parts[2])

        pairs = []
        for i, axis in enumerate(axes):
            if i < len(values):
                name = AXIS_NAME.get(axis, str(axis))
                pairs.append("{} = {:.4f} mm".format(name, values[i]))

        return "，".join(pairs)

    def update_position_from_response(self, text):
        text = str(text).strip()
        if not text.lower().startswith("position;"):
            return

        parts = text.split(";")
        if len(parts) < 3:
            return

        axes = self.parse_axes(parts[1])
        values = self.parse_float_list(parts[2])

        updates = {}
        for i, axis in enumerate(axes):
            if i < len(values) and axis in self.pos_vars:
                updates[axis] = float(values[i])
        if updates:
            with self.position_lock:
                self.current_position.update(updates)
                position = dict(self.current_position)
            for axis, value in updates.items():
                self.post_ui(self.pos_vars[axis].set, "{:.4f}".format(value))
            self.post_ui(self.update_scan_relative_position, position)
        return updates

    def update_scan_relative_position(self, position=None):
        if not self.scan_start_position:
            self.scan_relative_position_var.set("相对扫描起点：尚未记录")
            return
        if position is None:
            with self.position_lock:
                position = dict(self.current_position)
        if any(position.get(axis) is None for axis in (0, 1, 2)):
            self.scan_relative_position_var.set("相对扫描起点：当前坐标无效")
            return
        relative = {
            axis: float(position[axis]) - float(self.scan_start_position[axis])
            for axis in (0, 1, 2)
        }
        self.scan_relative_position_var.set(
            "相对扫描起点：ΔX={:+.4f}  ΔY={:+.4f}  ΔZ={:+.4f} mm".format(
                relative[0], relative[1], relative[2]
            )
        )

    # ============================================================
    # 通讯执行
    # ============================================================
    def run_worker(self, title, func, show_error=True):
        def worker():
            try:
                func()
            except Exception as exc:
                msg = "{}失败：{}".format(title, exc)
                self.log("!! " + msg)
                if show_error:
                    self.post_ui(messagebox.showerror, "错误", msg)

        threading.Thread(target=worker, daemon=True).start()

    def run_motion_worker(self, title, func, stop_on_error=True):
        if not self.motion_operation_lock.acquire(blocking=False):
            messagebox.showwarning("运动忙", "已有运动或扫描任务正在执行，请先停止并等待任务结束")
            return
        self.software_stop_event.clear()
        self.client.clear_stop_request()

        def worker():
            try:
                func()
            except InterruptedError as exc:
                self.log("运动已停止：{}".format(exc))
            except Exception as exc:
                msg = "{}失败：{}".format(title, exc)
                self.log("!! " + msg)
                if stop_on_error:
                    self.stop_after_motion_error(msg)
                self.post_ui(messagebox.showerror, "错误", msg)
            finally:
                self.motion_operation_lock.release()

        threading.Thread(target=worker, daemon=True).start()

    def send_and_log(self, cmd, timeout=None):
        if str(cmd).strip().lower().startswith("move;"):
            raise RuntimeError("Move 指令必须通过实时位置监控流程发送")
        self.log(">> {}".format(cmd + r"\r\n"))
        self.log("   中文说明：{}".format(self.describe_command(cmd)))

        reply = self.client.send_command(cmd, timeout=timeout)

        self.log("<< {}".format(reply))
        self.log("   中文说明：{}".format(self.describe_response(reply)))

        self.update_position_from_response(reply)
        return reply

    def move_and_wait(self, cmd, target_position, timeout, quiet=False):
        """发送异步 Move，并通过实际坐标实时刷新和判断到位。"""
        if self.software_stop_event.is_set():
            raise InterruptedError("停止状态已锁存")

        if not quiet:
            self.log(">> {}".format(cmd + r"\r\n"))
            self.log("   中文说明：{}".format(self.describe_command(cmd)))
        self.client.send_no_wait(cmd)

        deadline = time.monotonic() + float(timeout)
        last_position = None
        while True:
            if self.software_stop_event.is_set():
                raise InterruptedError("运动已由 Stop 指令终止")

            remaining = deadline - time.monotonic()
            if remaining <= 0:
                actual = "，".join(
                    "{}={:.4f}".format(AXIS_NAME[axis], last_position[axis])
                    for axis in target_position
                    if last_position and axis in last_position
                ) or "无有效坐标"
                raise TimeoutError("运动到位超时；当前 {}".format(actual))

            try:
                last_position = self.query_position(
                    quiet=True, timeout=min(1.0, remaining)
                )
            except TimeoutError:
                continue

            if all(
                abs(last_position[axis] - target) <= MOVE_POSITION_TOLERANCE_MM
                for axis, target in target_position.items()
            ):
                if not quiet:
                    self.log("<< 已到达目标位置：{}".format(
                        "，".join(
                            "{}={:.4f} mm".format(AXIS_NAME[axis], last_position[axis])
                            for axis in target_position
                        )
                    ))
                return last_position

            time.sleep(MOVE_POLL_INTERVAL_S)

    def query_position(self, quiet=False, timeout=5):
        cmd = "GetPosition;0,1,2"
        if not quiet:
            self.log(">> {}".format(cmd + r"\r\n"))
            self.log("   中文说明：查询 X/Y/Z 轴当前绝对位置。")
        reply = self.client.send_command(cmd, timeout=timeout)
        if not quiet:
            self.log("<< {}".format(reply))
            self.log("   中文说明：{}".format(self.describe_response(reply)))
        position = self.parse_position_values(reply)
        if not all(axis in position and math.isfinite(position[axis]) for axis in (0, 1, 2)):
            raise RuntimeError("厂家位置回复不完整：{}".format(reply))
        self.update_position_from_response(reply)
        return position

    def _set_position_text(self, text):
        with self.position_lock:
            self.current_position = {0: None, 1: None, 2: None}
        for var in self.pos_vars.values():
            var.set(text)
        self.update_scan_relative_position()

    def _poll_position(self):
        self.after(POSITION_POLL_INTERVAL_MS, self._poll_position)
        if (
            not self.client.connected
            or self.position_poll_busy
            or self.connection_attempt_lock.locked()
            or self.motion_operation_lock.locked()
        ):
            return

        self.position_poll_busy = True

        def worker():
            try:
                self.query_position(quiet=True, timeout=1.0)
                self.position_poll_error_logged = False
            except InterruptedError:
                pass
            except ConnectionError as exc:
                self.client.disconnect()
                if not self.position_poll_error_logged:
                    self.position_poll_error_logged = True
                    self.log("!! 实时位置连接中断：{}".format(exc))
                self.post_ui(self.status_var.set, "连接已断开")
                self.post_ui(self.card_status_var.set, "控制卡：未连接")
                self.post_ui(self._set_position_text, "连接中断")
            except Exception as exc:
                if not self.position_poll_error_logged:
                    self.position_poll_error_logged = True
                    self.log("!! 实时位置刷新暂停：{}".format(exc))
                    self.post_ui(self._set_position_text, "通信异常")
            finally:
                self.position_poll_busy = False

        threading.Thread(target=worker, daemon=True).start()

    def require_connected(self):
        if not self.client.connected:
            raise RuntimeError("请先连接厂家软件 TCP 服务")

    # ============================================================
    # 输入获取
    # ============================================================
    def get_float(self, var, name):
        try:
            value = float(var.get())
        except Exception:
            raise ValueError("{} 必须是数字".format(name))
        if not math.isfinite(value):
            raise ValueError("{} 必须是有限数字".format(name))
        return value

    def get_speed(self):
        speed = float(self.speed_var.get())
        if not math.isfinite(speed) or speed <= 0:
            raise ValueError("速度必须大于 0")
        if speed > 100:
            raise ValueError("速度不能超过 100%")
        return speed

    def get_soft_limits(self):
        limits = {}
        for axis in (0, 1, 2):
            lo = self.get_float(self.soft_limit_min_vars[axis], "{}轴软件下限".format(AXIS_NAME[axis]))
            hi = self.get_float(self.soft_limit_max_vars[axis], "{}轴软件上限".format(AXIS_NAME[axis]))
            if lo >= hi:
                raise ValueError("{}轴软件下限必须小于上限".format(AXIS_NAME[axis]))
            limits[axis] = (lo, hi)
        return limits

    def check_soft_limits(self, position, limits, context="目标位置"):
        for axis, value in position.items():
            lo, hi = limits[axis]
            if not math.isfinite(float(value)) or not (lo <= float(value) <= hi):
                raise ValueError(
                    "{}超出软件限位：{}={:.4f} mm，允许范围 [{:.4f}, {:.4f}] mm".format(
                        context, AXIS_NAME[axis], float(value), lo, hi
                    )
                )

    def request_software_stop(self, reason):
        self.software_stop_event.set()
        self.scan_stop_event.set()
        if self.hold_event is not None:
            self.hold_event.set()
        self.jog_run_id += 1
        if self.client.connected:
            self.client.stop_no_wait()
            self.log(">> Stop\\r\\n")
            self.log("   中文说明：{}；已锁存停止状态，不再发送新的运动指令。".format(reason))

    def stop_after_motion_error(self, reason):
        try:
            self.request_software_stop(reason)
        except Exception as exc:
            self.log("!! 运动异常后发送 Stop 失败：{}".format(exc))

    def on_speed_change(self, value):
        try:
            speed = round(float(value))
            self.speed_label_var.set("{}%".format(speed))
        except Exception:
            pass

    # ============================================================
    # 按钮事件
    # ============================================================
    def on_connect(self, automatic=False):
        if automatic and self.client.connected:
            return
        if self.motion_operation_lock.locked():
            if not automatic:
                messagebox.showwarning("运动忙", "运动任务尚未结束，不能重新连接")
            return
        if not self.connection_attempt_lock.acquire(blocking=False):
            if not automatic:
                messagebox.showwarning("连接中", "正在连接厂家软件，请稍候")
            return
        try:
            host = self.host_var.get().strip()
            port = int(self.port_var.get().strip())
            timeout = float(self.timeout_var.get().strip())
            if not host or not (1 <= port <= 65535) or not math.isfinite(timeout) or timeout <= 0:
                raise ValueError("Host、Port 或超时设置无效")
        except Exception as exc:
            self.connection_attempt_lock.release()
            if automatic:
                self.status_var.set("自动连接失败")
                self._set_position_text("未连接")
            else:
                messagebox.showerror("连接参数错误", str(exc))
            return

        self.status_var.set("正在连接 {}:{}...".format(host, port))
        self._set_position_text("连接中...")

        def task():
            try:
                self.client.connect(host, port, timeout=timeout)
                self.client.clear_stop_request()
                self.send_and_log("Remote", timeout=5)
                self.query_position()

                self.post_ui(self.status_var.set, "已连接 {}:{}（远程模式）".format(host, port))
                self.post_ui(self.card_status_var.set, "控制卡：已连接厂家软件")
                self.post_ui(self.estop_status_var.set, "急 停：请以厂家软件/硬件状态为准")
                self.post_ui(self.error_status_var.set, "轴错误：请以厂家软件状态为准")
                self.log("已连接厂家软件并进入 Remote 模式，实时位置刷新已启动。")
            except Exception:
                self.client.disconnect()
                self.post_ui(self.status_var.set, "自动连接失败" if automatic else "连接失败")
                self.post_ui(self.card_status_var.set, "控制卡：未连接")
                self.post_ui(self._set_position_text, "未连接")
                raise
            finally:
                self.connection_attempt_lock.release()

        self.run_worker("自动连接" if automatic else "连接", task, show_error=not automatic)

    def on_disconnect(self):
        if self.connection_attempt_lock.locked():
            messagebox.showwarning("连接中", "正在连接厂家软件，请稍候后再断开")
            return
        try:
            if self.motion_operation_lock.locked() and self.client.connected:
                self.request_software_stop("断开连接前停止运动")
        except Exception as exc:
            self.log("!! 断开前发送 Stop 失败：{}".format(exc))
        finally:
            self.client.disconnect()
            self._set_position_text("未连接")
            self.status_var.set("未连接")
            self.card_status_var.set("控制卡：未连接")
            self.log("已断开连接")

    def enter_remote(self):
        if self.motion_operation_lock.locked():
            messagebox.showwarning("运动忙", "运动任务执行期间不能切换远程模式")
            return
        def task():
            self.require_connected()
            self.client.clear_stop_request()
            self.send_and_log("Remote", timeout=5)
        self.run_worker("进入远程模式", task)

    def quit_remote(self):
        if self.motion_operation_lock.locked():
            messagebox.showwarning("运动忙", "运动任务执行期间不能切换远程模式")
            return
        def task():
            self.require_connected()
            self.client.clear_stop_request()
            self.send_and_log("QuitRemote", timeout=5)
        self.run_worker("退出远程模式", task)

    def send_raw_command(self):
        cmd = self.raw_cmd_var.get().strip()
        if not cmd:
            messagebox.showwarning("提示", "请输入原始指令")
            return
        if self.motion_operation_lock.locked():
            messagebox.showwarning("运动忙", "运动或扫描期间不能发送原始指令")
            return
        if cmd.split(";", 1)[0].strip().lower() in {"move", "gohome", "stop"}:
            messagebox.showwarning("安全限制", "交付版禁止通过原始指令发送运动/归零/停止命令，请使用专用按钮")
            return

        def task():
            self.require_connected()
            self.client.clear_stop_request()
            self.send_and_log(cmd, timeout=60)

        self.run_worker("发送原始指令", task)

    def get_position(self):
        if self.motion_operation_lock.locked():
            messagebox.showwarning("运动忙", "运动任务执行期间位置由程序自动查询")
            return
        def task():
            self.require_connected()
            self.client.clear_stop_request()
            self.query_position()
        self.run_worker("查询位置", task)

    def relative_move(self, axis, sign):
        try:
            distance = self.get_float(self.rel_vars[axis], "{}轴相对运动距离".format(AXIS_NAME[axis]))
            distance = abs(distance) * sign
            speed = self.get_speed()
            limits = self.get_soft_limits()
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        def task():
            self.require_connected()
            position = self.query_position()
            target = position[axis] + distance
            self.check_soft_limits({axis: target}, limits, "相对运动目标")
            cmd = "Move;Rel;{};{};{}".format(self.num_text(speed), axis, self.num_text(distance))
            self.move_and_wait(cmd, {axis: target}, timeout=120)

        self.run_motion_worker("{}轴相对运动".format(AXIS_NAME[axis]), task)

    def absolute_move(self, axis):
        try:
            target = self.get_float(self.abs_vars[axis], "{}轴绝对目标位置".format(AXIS_NAME[axis]))
            speed = self.get_speed()
            limits = self.get_soft_limits()
            self.check_soft_limits({axis: target}, limits, "绝对运动目标")
            cmd = "Move;Abs;{};{};{}".format(self.num_text(speed), axis, self.num_text(target))
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        def task():
            self.require_connected()
            self.query_position()  # 同步并清理可能残留的 Stop/Move 回复
            self.move_and_wait(cmd, {axis: target}, timeout=120)

        self.run_motion_worker("{}轴绝对运动".format(AXIS_NAME[axis]), task)

    def emergency_stop(self):
        try:
            self.require_connected()
            self.request_software_stop("用户按下停止运动")
            if self.scan_running:
                self.scan_progress_var.set("扫描状态：正在停止...")
            self.log("   真正急停请优先使用硬件急停按钮。")
        except Exception as exc:
            self.log("!! 发送 Stop 失败：{}".format(exc))
            messagebox.showerror("错误", "发送 Stop 失败：{}".format(exc))

    # ============================================================
    # 示波器采集
    # ============================================================
    def browse_save_dir(self):
        path = filedialog.askdirectory(title="选择数据保存文件夹")
        if path:
            self.save_dir_var.set(path)

    # ============================================================
    # 示波器实时波形显示
    # ============================================================
    def clear_waveform_plot(self, message="等待采集波形..."):
        if not hasattr(self, "wave_canvas"):
            return
        c = self.wave_canvas
        c.delete("all")
        w = max(600, c.winfo_width() or 1000)
        h = int(c["height"]) if str(c["height"]).isdigit() else 260
        c.create_rectangle(0, 0, w, h, fill="#0b1020", outline="")
        c.create_text(w // 2, h // 2, text=message, fill="#cbd5e1", font=("Microsoft YaHei UI", 12, "bold"))

    def start_scope_live(self):
        if not self.scope.is_connected:
            messagebox.showwarning("提示", "请先连接示波器")
            return
        if self.scope_live_running:
            return
        if self.scan_running:
            messagebox.showwarning("扫描运行中", "扫描采集期间不能启动实时波形显示")
            return
        self.scope_live_running = True
        self.log("示波器实时波形显示已启动。")
        self._schedule_scope_live(0)

    def stop_scope_live(self):
        self.scope_live_running = False
        self.scope_live_busy = False
        self.log("示波器实时波形显示已停止。")

    def _get_live_interval_ms(self):
        try:
            v = int(float(self.scope_live_interval_var.get()))
        except Exception:
            v = 500
        return max(200, min(v, 5000))

    def _schedule_scope_live(self, delay_ms=None):
        if not self.scope_live_running:
            return
        if delay_ms is None:
            delay_ms = self._get_live_interval_ms()
        self.after(delay_ms, self._scope_live_tick)

    def _scope_live_tick(self):
        if not self.scope_live_running:
            return
        if self.scope_live_busy:
            self._schedule_scope_live(self._get_live_interval_ms())
            return
        if not self.scope.is_connected:
            self.scope_live_running = False
            self.wave_info_var.set("波形：示波器未连接")
            return

        try:
            scope_config = self.get_scope_config()
        except Exception as exc:
            self.scope_live_running = False
            self.wave_info_var.set("波形：参数错误")
            messagebox.showerror("示波器参数错误", str(exc))
            return

        self.scope_live_busy = True

        def worker():
            try:
                wave_v, fs, n, time_s, meta = self.acquire_scope_waveform(scope_config)
                vpp = float(max(wave_v) - min(wave_v)) if len(wave_v) else 0.0
                self.last_waveform = (time_s, wave_v, fs, n, vpp)
                self.post_ui(self.update_waveform_plot, time_s, wave_v, fs, n, vpp, "实时波形")
            except Exception as exc:
                self.post_ui(self.wave_info_var.set, "波形：实时采集失败")
                self.log("!! 实时波形采集失败：{}".format(exc))
            finally:
                def done():
                    self.scope_live_busy = False
                    self._schedule_scope_live(self._get_live_interval_ms())
                self.post_ui(done)

        threading.Thread(target=worker, daemon=True).start()

    def update_waveform_plot(self, time_s, wave_v, fs, n, vpp, title="波形"):
        """
        在 Tk Canvas 上绘制波形。这里不用 matplotlib/pyqtgraph，避免额外依赖。
        """
        if not hasattr(self, "wave_canvas"):
            return

        c = self.wave_canvas
        c.delete("all")
        c.update_idletasks()

        w = max(700, c.winfo_width())
        h = max(220, c.winfo_height())

        margin_l = 62
        margin_r = 22
        margin_t = 28
        margin_b = 42

        plot_w = max(100, w - margin_l - margin_r)
        plot_h = max(80, h - margin_t - margin_b)

        # 背景
        c.create_rectangle(0, 0, w, h, fill="#0b1020", outline="")
        c.create_rectangle(margin_l, margin_t, w - margin_r, h - margin_b, fill="#111827", outline="#334155")

        # 数据转换
        try:
            ys = [float(v) for v in wave_v]
            xs = [float(t) * 1e6 for t in time_s]  # µs
        except Exception:
            self.clear_waveform_plot("波形数据解析失败")
            return

        if not ys or not xs:
            self.clear_waveform_plot("没有波形数据")
            return

        # 保证长度一致
        m = min(len(xs), len(ys))
        xs = xs[:m]
        ys = ys[:m]

        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys), max(ys)

        if abs(x_max - x_min) < 1e-15:
            x_max = x_min + 1.0
        if abs(y_max - y_min) < 1e-12:
            y_min -= 0.5
            y_max += 0.5

        # 给 Y 轴留边
        pad_y = (y_max - y_min) * 0.12
        y_min -= pad_y
        y_max += pad_y

        # 网格
        for i in range(6):
            x = margin_l + plot_w * i / 5
            c.create_line(x, margin_t, x, margin_t + plot_h, fill="#1e293b")
            xv = x_min + (x_max - x_min) * i / 5
            c.create_text(x, h - margin_b + 18, text="{:.2f}".format(xv), fill="#94a3b8", font=("Consolas", 9))

        for i in range(5):
            y = margin_t + plot_h * i / 4
            c.create_line(margin_l, y, margin_l + plot_w, y, fill="#1e293b")
            yv = y_max - (y_max - y_min) * i / 4
            c.create_text(margin_l - 8, y, text="{:.3g}".format(yv), fill="#94a3b8", anchor=tk.E, font=("Consolas", 9))

        # 0V 参考线
        if y_min <= 0 <= y_max:
            y0 = margin_t + (y_max - 0) / (y_max - y_min) * plot_h
            c.create_line(margin_l, y0, margin_l + plot_w, y0, fill="#475569", dash=(4, 4))

        # 降采样，避免点太多卡顿
        max_draw_points = max(300, plot_w)
        step = max(1, int(m / max_draw_points))
        pts = []
        for i in range(0, m, step):
            px = margin_l + (xs[i] - x_min) / (x_max - x_min) * plot_w
            py = margin_t + (y_max - ys[i]) / (y_max - y_min) * plot_h
            pts.extend([px, py])

        if len(pts) >= 4:
            c.create_line(*pts, fill="#38bdf8", width=1.5)

        # 标题和轴标签
        c.create_text(margin_l, 12, text=title, fill="#e2e8f0", anchor=tk.W, font=("Microsoft YaHei UI", 11, "bold"))
        c.create_text(w - margin_r, 12, text="Vpp={:.6g} V   fs={:.3f} MSa/s   N={}".format(
            vpp, fs / 1e6 if fs else 0.0, n
        ), fill="#a7f3d0", anchor=tk.E, font=("Consolas", 10))

        c.create_text(margin_l + plot_w / 2, h - 10, text="时间 / µs", fill="#cbd5e1", font=("Microsoft YaHei UI", 10))
        c.create_text(18, margin_t + plot_h / 2, text="电压 / V", fill="#cbd5e1", angle=90, font=("Microsoft YaHei UI", 10))

        self.wave_info_var.set("波形：{}，Vpp={:.6g} V，采样率={:.3f} MSa/s，点数={}".format(
            title, vpp, fs / 1e6 if fs else 0.0, n
        ))

    def get_scope_points(self):
        k = self.get_float(self.scope_points_k_var, "采集点数(K)")
        if k <= 0:
            raise ValueError("采集点数必须大于 0")
        points = int(round(k * 1000))
        if points > XLSX_MAX_DATA_ROWS:
            raise ValueError("采集点数不能超过 {}，否则无法保存为 MATLAB 兼容 XLSX".format(XLSX_MAX_DATA_ROWS))
        return points

    def get_scope_channel(self):
        ch = int(self.scope_channel_var.get())
        if ch < 1 or ch > 4:
            raise ValueError("通道必须是 1~4")
        return ch

    def get_scope_config(self):
        return {
            "channel": self.get_scope_channel(),
            "points": self.get_scope_points(),
            "mode": self.scope_mode_var.get().strip(),
            "average": self.scope_avg_var.get().strip(),
        }

    def apply_scope_average(self, avg):
        avg = str(avg).strip()
        if avg == "无":
            self.scope.set_normal_mode()
            return
        count = int(avg.replace("次", ""))
        self.scope.set_average_mode(count)

    def on_scope_connect(self):
        if self.scan_running:
            messagebox.showwarning("扫描运行中", "扫描期间不能重新连接示波器")
            return
        resource = self.scope_visa_var.get().strip()
        if not resource:
            messagebox.showwarning("提示", "请输入示波器 VISA 地址")
            return

        def task():
            idn = self.scope.connect(resource)
            self.post_ui(self.scope_status_var.set, "示波器：已连接")
            self.post_ui(self.clear_waveform_plot, "示波器已连接，点击“开始实时显示”或“单次采集并保存”。")
            self.log("示波器连接成功：{}".format(idn))

        self.run_worker("连接示波器", task)

    def on_scope_disconnect(self):
        if self.scan_running:
            messagebox.showwarning("扫描运行中", "请先停止并等待扫描结束后再断开示波器")
            return
        self.stop_scope_live()

        def task():
            with self.scope_lock:
                self.scope.disconnect()
            self.post_ui(self.scope_status_var.set, "示波器：未连接")
            self.post_ui(self.clear_waveform_plot, "示波器已断开")
            self.log("示波器已断开。")

        self.run_worker("断开示波器", task)

    def on_scope_status(self):
        def task():
            if not self.scope.is_connected:
                raise RuntimeError("示波器未连接")
            with self.scope_lock:
                status = self.scope.query_status_summary()
            self.log("示波器状态：{}".format(status))
        self.run_worker("查询示波器状态", task)

    def prepare_scope_acquisition(self, config):
        with self.scope_lock:
            if not self.scope.is_connected:
                raise RuntimeError("示波器未连接")
            self.apply_scope_average(config["average"])
            self.scope._screen_cache_key = None
            self.scope._screen_preamble = None
            self.scope._screen_time_key = None
            self.scope._screen_time_s = None

    def acquire_scope_waveform(self, config, apply_settings=True, cache_config=False):
        """根据界面参数采集一次波形。"""
        if not self.scope.is_connected:
            raise RuntimeError("示波器未连接")

        ch = int(config["channel"])
        pts = int(config["points"])
        mode = str(config["mode"]).strip()

        # 示波器 VISA 通讯不能多个线程同时访问，因此这里加锁。
        with self.scope_lock:
            if not self.scope.is_connected:
                raise RuntimeError("示波器未连接")
            if apply_settings:
                self.apply_scope_average(config["average"])
            if mode == "STOP-RAW" or (mode == "自动" and pts > 1000):
                return self.scope.acquire_stop_mode(channel=ch, points=pts)
            return self.scope.acquire_screen_norm(
                channel=ch, points=pts, use_cache=cache_config
            )

    def write_wave_xlsx(self, filepath, metadata_pairs, time_s, wave_v):
        write_matlab_xlsx(filepath, metadata_pairs, time_s, wave_v)

    def append_summary_csv(self, session_dir, row_dict):
        if self.scan_summary_writer is not None:
            self.scan_summary_writer.writerow(row_dict)
            return
        os.makedirs(session_dir, exist_ok=True)
        path = os.path.join(session_dir, "scan_summary.csv")
        exists = os.path.exists(path)
        with open(path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=SCAN_SUMMARY_FIELDS)
            if not exists:
                writer.writeheader()
            writer.writerow(row_dict)

    def get_current_position_dict(self):
        with self.position_lock:
            position = dict(self.current_position)
        if any(position[axis] is None or not math.isfinite(position[axis]) for axis in (0, 1, 2)):
            raise RuntimeError("当前位置无效，请先成功查询 X/Y/Z 坐标")
        return position

    def get_current_position_tuple(self):
        position = self.get_current_position_dict()
        return position[0], position[1], position[2]

    def acquire_and_save_scan_point(self, p, point_index, line_idx, point_in_line):
        """S型扫描到达每个点后采集示波器并保存。"""
        if not self.scan_session_dir:
            raise RuntimeError("扫描保存文件夹未创建")
        if not self.scope.is_connected:
            raise RuntimeError("示波器未连接")

        acquire_started_at = time.monotonic()
        wave_v, fs, n, time_s, meta = self.acquire_scope_waveform(
            p["scope_config"], apply_settings=False, cache_config=True
        )
        acquire_seconds = time.monotonic() - acquire_started_at
        x, y, z = self.get_current_position_tuple()
        vpp = float(max(wave_v) - min(wave_v)) if len(wave_v) else 0.0
        report_point = (
            point_index == 1
            or point_index == p.get("total_points")
            or point_index % p.get("ui_update_stride", 1) == 0
        )
        self.last_waveform = (time_s, wave_v, fs, n, vpp)
        if report_point and p.get("live_plot"):
            self.post_ui(
                self.update_waveform_plot,
                time_s, wave_v, fs, n, vpp,
                "扫描点 {} / {}".format(point_index, p.get("total_points", "?")),
            )
        ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
        filename = (
            "波形数据_{:05d}_L{:03d}_P{:03d}_X={:.4f}_Y={:.4f}_Z={:.4f}_Vpp={:.6f}.xlsx".format(
                point_index, line_idx, point_in_line, x, y, z, vpp
            )
        )
        filepath = os.path.join(self.scan_session_dir, filename)

        metadata_pairs = [
            ("X位置(mm)", x),
            ("Y位置(mm)", y),
            ("Z位置(mm)", z),
            ("采样率(Hz)", fs),
            ("采样点数", n),
            ("Vpp(V)", vpp),
            ("示波器通道", p["scope_config"]["channel"]),
        ]

        save_started_at = time.monotonic()
        if p["save_each_point"]:
            self.write_wave_xlsx(filepath, metadata_pairs, time_s, wave_v)
        else:
            filepath = ""

        self.append_summary_csv(self.scan_session_dir, {
            "point_index": point_index,
            "line_index": line_idx,
            "point_in_line": point_in_line,
            "plane": p.get("plane", ""),
            "span_axis": p.get("span_axis", ""),
            "step_axis": p.get("step_axis", ""),
            "x_mm": x,
            "y_mm": y,
            "z_mm": z,
            "vpp_v": vpp,
            "fs_hz": fs,
            "sample_len": n,
            "xlsx_file": filepath,
        })
        save_seconds = time.monotonic() - save_started_at
        if report_point:
            self.log("采集完成：第 {}/{} 点，Vpp={:.6f} V，采样率={:.3f} MSa/s，点数={}，采集 {:.3f}s，保存 {:.3f}s。".format(
                point_index, p.get("total_points", "?"), vpp, fs / 1e6 if fs else 0, n,
                acquire_seconds, save_seconds,
            ))
            if filepath:
                self.log("波形已保存：{}".format(filepath))
        return {
            "wave_v": wave_v,
            "time_s": time_s,
            "fs": fs,
            "sample_len": n,
            "vpp": vpp,
            "position": (x, y, z),
            "acquire_seconds": acquire_seconds,
            "save_seconds": save_seconds,
        }

    def on_single_acquire(self):
        if self.motion_operation_lock.locked():
            messagebox.showwarning("运动忙", "请等待运动或扫描任务结束后再执行单次采集")
            return
        try:
            scope_config = self.get_scope_config()
            save_dir = self.save_dir_var.get().strip() or os.getcwd()
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        def task():
            if not self.scope.is_connected:
                raise RuntimeError("请先连接示波器")
            os.makedirs(save_dir, exist_ok=True)
            if self.client.connected:
                self.query_position()
            wave_v, fs, n, time_s, meta = self.acquire_scope_waveform(scope_config)
            x, y, z = self.get_current_position_tuple()
            vpp = float(max(wave_v) - min(wave_v)) if len(wave_v) else 0.0
            self.last_waveform = (time_s, wave_v, fs, n, vpp)
            self.post_ui(self.update_waveform_plot, time_s, wave_v, fs, n, vpp, "单次采集波形")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            filename = "波形数据_{}_X={:.4f}_Y={:.4f}_Z={:.4f}_Vpp={:.6f}.xlsx".format(ts, x, y, z, vpp)
            filepath = os.path.join(save_dir, filename)
            metadata_pairs = [
                ("X位置(mm)", x),
                ("Y位置(mm)", y),
                ("Z位置(mm)", z),
                ("采样率(Hz)", fs),
                ("采样点数", n),
                ("Vpp(V)", vpp),
                ("示波器通道", scope_config["channel"]),
            ]
            self.write_wave_xlsx(filepath, metadata_pairs, time_s, wave_v)
            self.log("单次采集完成：Vpp={:.6f} V，采样率={:.3f} MSa/s，点数={}。".format(vpp, fs / 1e6 if fs else 0, n))
            self.log("波形已保存：{}".format(filepath))
        self.run_motion_worker("单次采集", task, stop_on_error=False)

    # ============================================================
    # 扫描轨迹
    # ============================================================
    def axis_to_id(self, axis_name):
        axis_name = str(axis_name).upper().strip()
        mapping = {"X": 0, "Y": 1, "Z": 2}
        if axis_name not in mapping:
            raise ValueError("未知轴：{}".format(axis_name))
        return mapping[axis_name]

    def on_scan_plane_changed(self):
        """
        根据扫描面自动设置行程轴和步距轴。
        XY：X 行程，Y 换行；XZ：X 行程，Z 换行；YZ：Y 行程，Z 换行。
        """
        plane = self.scan_plane_var.get().strip().upper()
        mapping = {
            "XY": ("X", "Y"),
            "XZ": ("X", "Z"),
            "YZ": ("Y", "Z"),
        }
        if plane in mapping:
            span_axis, step_axis = mapping[plane]
            self.scan_span_axis_var.set(span_axis)
            self.scan_step_axis_var.set(step_axis)
        self.update_scan_preview()

    def parse_position_values(self, text):
        """
        解析厂家返回的 Position;0,1,2;x,y,z，返回 {0:x,1:y,2:z}
        """
        result = {}
        parts = str(text).strip().split(";")
        if len(parts) < 3 or parts[0].lower() != "position":
            return result

        axes = self.parse_axes(parts[1])
        values = self.parse_float_list(parts[2])
        for i, axis in enumerate(axes):
            if i < len(values):
                result[axis] = values[i]
        return result

    def calculate_scan_segments(self, span, step):
        """
        把总行程拆成若干段。比如 span=10, step=3，会得到 [3,3,3,1]。
        这样最后一个点能够真正覆盖到 10 mm 的终点。
        """
        span = float(span)
        step = abs(float(step))
        if span <= 0:
            return []
        if step <= 0:
            raise ValueError("行内步距必须大于 0")

        full_segments = int(math.floor(span / step))
        if full_segments > 200_000:
            raise ValueError("单行扫描点数过多，请增大行内步距")
        segments = [step] * full_segments
        remainder = span - full_segments * step
        if remainder > max(1e-12, span * 1e-12):
            segments.append(remainder)
        return segments

    @staticmethod
    def choose_scan_direction(axis, start, distance, limits, preference="自动"):
        """按用户选择确定方向；自动模式优先正向。"""
        lo, hi = limits[axis]
        start = float(start)
        distance = float(distance)
        tolerance = max(1e-9, abs(start) * 1e-12, abs(distance) * 1e-12)
        preference = str(preference).strip()
        if preference == "正向（小→大）":
            if start + distance <= hi + tolerance:
                return 1
            raise ValueError(
                "{}轴选择了正向扫描，但从 {:.4f} mm 到 {:.4f} mm 超出上限 {:.4f} mm".format(
                    AXIS_NAME[axis], start, start + distance, hi
                )
            )
        if preference == "反向（大→小）":
            if start - distance >= lo - tolerance:
                return -1
            raise ValueError(
                "{}轴选择了反向扫描，但从 {:.4f} mm 到 {:.4f} mm 超出下限 {:.4f} mm".format(
                    AXIS_NAME[axis], start, start - distance, lo
                )
            )
        if preference != "自动":
            raise ValueError("未知的扫描方向：{}".format(preference))
        if start + distance <= hi + tolerance:
            return 1
        if start - distance >= lo - tolerance:
            return -1
        raise ValueError(
            "{}轴从当前位置 {:.4f} mm 出发，正负方向都放不下 {:.4f} mm 行程；允许范围 [{:.4f}, {:.4f}] mm".format(
                AXIS_NAME[axis], start, distance, lo, hi
            )
        )

    def get_scan_params(self):
        plane = self.scan_plane_var.get().strip().upper()
        span_axis = self.scan_span_axis_var.get().strip().upper()
        step_axis = self.scan_step_axis_var.get().strip().upper()
        span_direction_mode = self.scan_direction_var.get().strip()
        step_direction_mode = self.scan_step_direction_var.get().strip()

        if plane not in ["XY", "XZ", "YZ"]:
            raise ValueError("扫描面必须是 XY、XZ 或 YZ")

        if span_axis == step_axis:
            raise ValueError("行程轴和行数轴不能相同")

        plane_axes = set(list(plane))
        if span_axis not in plane_axes or step_axis not in plane_axes:
            raise ValueError("当前扫描面为 {}，行程轴和步距轴必须都在该平面内".format(plane))

        span = self.get_float(self.scan_span_var, "行程")
        span_step = self.get_float(self.scan_span_step_var, "行内步距")
        line_step = self.get_float(self.scan_line_step_var, "行步距")
        lines_value = self.get_float(self.scan_lines_var, "行数")
        if not float(lines_value).is_integer():
            raise ValueError("行数必须是整数")
        lines = int(lines_value)
        dwell = self.get_float(self.scan_dwell_var, "点位等待")

        if span <= 0:
            raise ValueError("行程必须大于 0")
        if span_step <= 0:
            raise ValueError("行内步距必须大于 0")
        if line_step <= 0:
            raise ValueError("行步距必须大于 0")
        if lines <= 0:
            raise ValueError("行数必须大于 0")
        if dwell < 0:
            raise ValueError("点位等待不能小于 0")

        segments = self.calculate_scan_segments(span, span_step)
        points_per_line = len(segments) + 1
        total_points = points_per_line * lines
        total_moves = len(segments) * lines + max(0, lines - 1)
        if total_points > 200_000:
            raise ValueError("扫描总点数超过 200000，请增大步距或减少行数")

        return {
            "plane": plane,
            "span_axis": span_axis,
            "step_axis": step_axis,
            "span_direction_mode": span_direction_mode,
            "step_direction_mode": step_direction_mode,
            "span": span,
            "span_step": span_step,
            "line_step": line_step,
            "lines": lines,
            "dwell": dwell,
            "segments": segments,
            "points_per_line": points_per_line,
            "total_points": total_points,
            "total_moves": total_moves,
        }

    def update_scan_preview(self):
        try:
            p = self.get_scan_params()
            # 这里只估算等待时间，不包含真实运动时间
            wait_time = p["total_points"] * p["dwell"]
            self.scan_preview_var.set(
                "扫描预览：{} 面；{}轴{}，行程 {:.3f} mm；{}轴{}换行 {:.3f} mm；{} 行 × 每行 {} 点 = {} 点，运动指令 {} 次，点位等待约 {:.1f}s".format(
                    p["plane"], p["span_axis"], p["span_direction_mode"], p["span"],
                    p["step_axis"], p["step_direction_mode"], p["line_step"],
                    p["lines"], p["points_per_line"], p["total_points"], p["total_moves"], wait_time
                )
            )
        except Exception as exc:
            self.scan_preview_var.set("扫描预览：参数有误 - {}".format(exc))

    def start_trajectory_scan(self):
        try:
            self.require_connected()
            if self.scan_running:
                messagebox.showwarning("提示", "扫描正在进行中")
                return

            p = self.get_scan_params()
            p["speed"] = self.get_speed()
            p["acquire_scope"] = bool(self.scan_acquire_scope_var.get())
            p["soft_limits"] = self.get_soft_limits()
            p["save_dir"] = self.save_dir_var.get().strip() or os.getcwd()
            p["save_npy"] = bool(self.save_npy_var.get())
            p["save_each_point"] = bool(self.save_each_point_var.get())
            p["live_plot"] = bool(self.scan_live_plot_var.get())
            p["scope_visa"] = self.scope_visa_var.get().strip()
            if p["acquire_scope"] and not self.scope.is_connected:
                raise RuntimeError("已勾选扫描每点采集，但示波器未连接")
            p["scope_config"] = self.get_scope_config() if p["acquire_scope"] else None
            if (
                p["acquire_scope"]
                and p["scope_config"]["average"] != "无"
                and p["dwell"] <= 0
            ):
                raise ValueError("使用示波器平均模式时，点位等待必须大于 0")
            if p["acquire_scope"]:
                os.makedirs(p["save_dir"], exist_ok=True)
                if p["save_npy"]:
                    expected_bytes = (
                        p["total_points"] * p["scope_config"]["points"] * 4
                        + p["total_points"] * 3 * 8
                        + p["total_points"]
                        + p["scope_config"]["points"] * 8
                    )
                    free_bytes = shutil.disk_usage(p["save_dir"]).free
                    reserve_bytes = max(512 * 1024**2, expected_bytes // 20)
                    if expected_bytes + reserve_bytes > free_bytes:
                        raise RuntimeError(
                            "NPY 数据集预计需要 {:.2f} GB，当前可用 {:.2f} GB，请更换保存目录或减少采集点数".format(
                                expected_bytes / 1024**3, free_bytes / 1024**3
                            )
                        )
                    p["expected_npy_bytes"] = expected_bytes
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        ok = messagebox.askyesno(
            "确认开始扫描",
            "即将从当前位置开始 S 型扫描：\n\n"
            "扫描面：{}\n"
            "行程轴：{}，首行方向：{}\n"
            "行数轴：{}，换行方向：{}\n"
            "行程：{} mm，行内步距：{} mm\n"
            "行步距：{} mm，行数：{}\n"
            "总点数：{}\n"
            "每点采集示波器：{}\n\n"
            "连续 NPY 数据集：{}\n"
            "逐点 MATLAB XLSX：{}\n\n"
            "程序会在扫描开始前记录当前位置，扫描后可点击“回到扫描起点”。\n"
            "自动模式会根据当前位置和软件限位选择方向。\n"
            "请确认已在厂家软件中设置好零点、限位和安全区域。\n是否开始？".format(
                p["plane"], p["span_axis"], p["span_direction_mode"],
                p["step_axis"], p["step_direction_mode"],
                self.num_text(p["span"]), self.num_text(p["span_step"]),
                self.num_text(p["line_step"]), p["lines"], p["total_points"],
                "是" if p.get("acquire_scope") else "否",
                "是" if p.get("acquire_scope") and p.get("save_npy") else "否",
                "是" if p.get("acquire_scope") and p.get("save_each_point") else "否",
            )
        )
        if not ok:
            return

        if not self.motion_operation_lock.acquire(blocking=False):
            messagebox.showwarning("运动忙", "已有运动或扫描任务正在执行")
            return

        self.stop_scope_live()
        self.scan_stop_event.clear()
        self.software_stop_event.clear()
        self.client.clear_stop_request()
        self.scan_running = True
        self.last_scan_params = dict(p)
        self.scan_progress_var.set("扫描状态：运行中，预计剩余时间正在测算")
        self.scan_thread = threading.Thread(target=self._trajectory_scan_loop, args=(p,), daemon=True)
        self.scan_thread.start()

    def stop_trajectory_scan(self):
        if not self.scan_running:
            return
        self.scan_progress_var.set("扫描状态：正在停止...")
        try:
            self.request_software_stop("用户请求停止扫描")
        except Exception as exc:
            self.log("!! 停止扫描时发送 Stop 失败：{}".format(exc))

    def return_to_scan_start(self):
        """
        扫描结束/停止后，回到扫描开始前记录的 X/Y/Z 位置。
        """
        try:
            self.require_connected()
            if self.scan_running:
                messagebox.showwarning("提示", "扫描正在进行中，请先停止或等待扫描完成")
                return

            if not self.scan_start_position:
                messagebox.showwarning("提示", "还没有记录扫描起点。请先开始一次扫描，程序会自动记录起点。")
                return

            pos = self.scan_start_position
            x = float(pos[0])
            y = float(pos[1])
            z = float(pos[2])
            speed = self.get_speed()
            limits = self.get_soft_limits()
            self.check_soft_limits(pos, limits, "扫描起点")
        except Exception as exc:
            messagebox.showerror("错误", str(exc))
            return

        ok = messagebox.askyesno(
            "确认返回扫描起点",
            "平台将以当前速度 {}% 回到扫描起点：\n\n"
            "X = {:.4f} mm\nY = {:.4f} mm\nZ = {:.4f} mm\n\n"
            "是否继续？".format(self.num_text(speed), x, y, z)
        )
        if not ok:
            return

        def task():
            self.require_connected()
            self.log("准备返回扫描起点。")
            current = self.query_position()
            if self.last_scan_params:
                order = [
                    self.axis_to_id(self.last_scan_params["span_axis"]),
                    self.axis_to_id(self.last_scan_params["step_axis"]),
                ]
                order += [axis for axis in (0, 1, 2) if axis not in order]
            else:
                order = [0, 1, 2]
            for axis in order:
                target = float(pos[axis])
                if abs(current[axis] - target) <= 1e-9:
                    continue
                cmd = "Move;Abs;{};{};{}".format(
                    self.num_text(speed), axis, self.num_text(target)
                )
                current = self.move_and_wait(cmd, {axis: target}, timeout=300)

        self.run_motion_worker("返回扫描起点", task)

    def move_to_scan_relative_position(self):
        """把用户输入的扫描起点相对坐标换算为绝对坐标后移动。"""
        try:
            self.require_connected()
            if self.scan_running:
                raise RuntimeError("扫描正在进行，请先停止或等待扫描完成")
            if not self.scan_start_position:
                raise RuntimeError("还没有记录扫描起点，请先开始一次扫描")
            relative = {
                axis: self.get_float(
                    self.scan_relative_target_vars[axis],
                    "相对扫描起点{}坐标".format(AXIS_NAME[axis]),
                )
                for axis in (0, 1, 2)
            }
            target = {
                axis: float(self.scan_start_position[axis]) + relative[axis]
                for axis in (0, 1, 2)
            }
            speed = self.get_speed()
            self.check_soft_limits(target, self.get_soft_limits(), "相对扫描起点目标")
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        if not messagebox.askyesno(
            "确认移动",
            "相对扫描起点：ΔX={:+.4f}, ΔY={:+.4f}, ΔZ={:+.4f} mm\n"
            "换算绝对目标：X={:.4f}, Y={:.4f}, Z={:.4f} mm\n\n是否继续？".format(
                relative[0], relative[1], relative[2],
                target[0], target[1], target[2],
            ),
        ):
            return

        def task():
            current = self.query_position()
            if self.last_scan_params:
                order = [
                    self.axis_to_id(self.last_scan_params["span_axis"]),
                    self.axis_to_id(self.last_scan_params["step_axis"]),
                ]
                order += [axis for axis in (0, 1, 2) if axis not in order]
            else:
                order = [0, 1, 2]
            for axis in order:
                if abs(current[axis] - target[axis]) <= MOVE_POSITION_TOLERANCE_MM:
                    continue
                cmd = "Move;Abs;{};{};{}".format(
                    self.num_text(speed), axis, self.num_text(target[axis])
                )
                current = self.move_and_wait(cmd, {axis: target[axis]}, timeout=300)
            self.log("已到达扫描起点相对坐标：ΔX={:+.4f}, ΔY={:+.4f}, ΔZ={:+.4f} mm。".format(
                relative[0], relative[1], relative[2]
            ))

        self.run_motion_worker("移动到扫描起点相对坐标", task)

    def _trajectory_scan_loop(self, p):
        span_axis_id = self.axis_to_id(p["span_axis"])
        step_axis_id = self.axis_to_id(p["step_axis"])
        speed = float(p.get("speed", 10))
        point_index = 0
        completed_points = 0
        move_index = 0
        total_acquire_seconds = 0.0
        total_save_seconds = 0.0
        total_move_seconds = 0.0
        result_status = "failed"
        scan_started_at = None
        self.scan_session_dir = None
        npy_waveforms = None
        npy_valid = None
        npy_positions = None
        dataset_paths = {}
        p["ui_update_stride"] = max(1, p["total_points"] // 500)
        line_update_stride = max(1, p["lines"] // 100)
        self.scan_summary_stream = None
        self.scan_summary_writer = None

        try:
            self.log("开始 S 型扫描：扫描面={}，行程轴={}，行数轴={}，总点数={}，运动指令={}。".format(
                p["plane"], p["span_axis"], p["step_axis"], p["total_points"], p["total_moves"]
            ))
            self.log("说明：首行方向={}，后续逐行反向；换行方向={}。".format(
                p["span_direction_mode"], p["step_direction_mode"]
            ))

            # 坐标不可靠时禁止继续扫描，避免保存错误的空间位置。
            start_pos = self.query_position()
            self.scan_start_position = dict(start_pos)
            self.post_ui(self.update_scan_relative_position, dict(start_pos))
            span_axis_id = self.axis_to_id(p["span_axis"])
            step_axis_id = self.axis_to_id(p["step_axis"])
            span_direction = self.choose_scan_direction(
                span_axis_id, start_pos[span_axis_id], p["span"], p["soft_limits"],
                p["span_direction_mode"],
            )
            step_travel = (p["lines"] - 1) * p["line_step"]
            step_direction = self.choose_scan_direction(
                step_axis_id, start_pos[step_axis_id], step_travel, p["soft_limits"],
                p["step_direction_mode"],
            )
            p["span_direction"] = span_direction
            p["step_direction"] = step_direction
            scan_end = dict(start_pos)
            scan_end[span_axis_id] += span_direction * p["span"]
            scan_end[step_axis_id] += step_direction * step_travel
            self.check_soft_limits(start_pos, p["soft_limits"], "扫描起点")
            self.check_soft_limits(scan_end, p["soft_limits"], "扫描终点")
            self.log("已记录并校验扫描范围：起点 X={:.4f}, Y={:.4f}, Z={:.4f} mm。".format(
                start_pos[0], start_pos[1], start_pos[2]
            ))
            self.log("扫描方向：{}轴{}向，{}轴{}向。".format(
                p["span_axis"], "+" if span_direction > 0 else "-",
                p["step_axis"], "+" if step_direction > 0 else "-",
            ))
            if p.get("acquire_scope"):
                self.prepare_scope_acquisition(p["scope_config"])

            # 如果勾选了每点采集，创建本次扫描的数据文件夹
            if p.get("acquire_scope"):
                root_dir = p["save_dir"]
                ts_session = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
                self.scan_session_dir = os.path.join(root_dir, "scan_{}_{}".format(p.get("plane", ""), ts_session))
                os.makedirs(self.scan_session_dir, exist_ok=True)
                params_path = os.path.join(self.scan_session_dir, "scan_params.json")
                with open(params_path, "w", encoding="utf-8") as f:
                    json.dump({
                        "scan_params": {k: v for k, v in p.items() if k != "segments"},
                        "segments": p.get("segments", []),
                        "scope_visa": p.get("scope_visa", ""),
                        "scope_config": p.get("scope_config"),
                        "scan_start_position": self.scan_start_position,
                    }, f, ensure_ascii=False, indent=2)
                self.scan_summary_stream = open(
                    os.path.join(self.scan_session_dir, "scan_summary.csv"),
                    "w", newline="", encoding="utf-8-sig",
                )
                self.scan_summary_writer = csv.DictWriter(
                    self.scan_summary_stream, fieldnames=SCAN_SUMMARY_FIELDS
                )
                self.scan_summary_writer.writeheader()
                self.log("本次扫描数据保存目录：{}".format(self.scan_session_dir))
                if p.get("save_npy"):
                    import numpy as np
                    cumulative = [0.0]
                    for segment in p["segments"]:
                        cumulative.append(cumulative[-1] + segment)
                    span_mm = np.sort(
                        start_pos[span_axis_id]
                        + span_direction * np.asarray(cumulative, dtype=np.float64)
                    )
                    step_mm = (
                        start_pos[step_axis_id]
                        + step_direction * p["line_step"] * np.arange(p["lines"], dtype=np.float64)
                    )
                    dataset_paths.update({
                        "span_mm": os.path.join(self.scan_session_dir, "span_mm.npy"),
                        "step_mm": os.path.join(self.scan_session_dir, "step_mm.npy"),
                        "time_s": os.path.join(self.scan_session_dir, "time_s.npy"),
                        "waveforms_v": os.path.join(self.scan_session_dir, "waveforms_v.npy"),
                        "valid": os.path.join(self.scan_session_dir, "valid.npy"),
                        "positions_mm": os.path.join(self.scan_session_dir, "positions_mm.npy"),
                        "info": os.path.join(self.scan_session_dir, "dataset_info.json"),
                    })
                    np.save(dataset_paths["span_mm"], span_mm)
                    np.save(dataset_paths["step_mm"], step_mm)
                    self.log(
                        "已启用连续 NPY 数据集，预计上限 {:.2f} GB。".format(
                            p.get("expected_npy_bytes", 0) / 1024**3
                        )
                    )

            scan_started_at = time.monotonic()
            for line_idx in range(1, p["lines"] + 1):
                if self.scan_stop_event.is_set():
                    break

                direction = span_direction if line_idx % 2 == 1 else -span_direction
                dir_text = "正向" if direction > 0 else "反向"
                report_line = (
                    line_idx == 1
                    or line_idx == p["lines"]
                    or line_idx % line_update_stride == 0
                )
                if report_line:
                    self.log("=== 第 {}/{} 行：{}轴{}扫描 ===".format(
                        line_idx, p["lines"], p["span_axis"], dir_text
                    ))

                # 每一行第一个点是当前点
                for point_in_line in range(1, p["points_per_line"] + 1):
                    if self.scan_stop_event.is_set():
                        break

                    point_index += 1
                    point_data = None
                    report_point = (
                        point_index == 1
                        or point_index == p["total_points"]
                        or point_index % p["ui_update_stride"] == 0
                    )
                    if report_point:
                        self.post_ui(
                            self.scan_progress_var.set,
                            "扫描状态：第 {}/{} 点（第 {} 行，第 {} 点）".format(
                                point_index, p["total_points"], line_idx, point_in_line
                            ),
                        )
                        self.log("扫描点：总第 {}/{} 点；第 {} 行第 {} 点。".format(
                            point_index, p["total_points"], line_idx, point_in_line
                        ))

                    # 到达扫描点后的等待时间：这里给后续采集/稳定预留
                    if p["dwell"] > 0:
                        t0 = time.time()
                        while time.time() - t0 < p["dwell"]:
                            if self.scan_stop_event.is_set():
                                break
                            time.sleep(0.02)

                    # 到达点位后采集示波器波形并保存
                    if p.get("acquire_scope") and not self.scan_stop_event.is_set():
                        try:
                            point_data = self.acquire_and_save_scan_point(
                                p, point_index, line_idx, point_in_line
                            )
                            total_acquire_seconds += point_data["acquire_seconds"]
                            total_save_seconds += point_data["save_seconds"]
                            if p.get("save_npy"):
                                import numpy as np
                                if npy_waveforms is None:
                                    sample_len = int(point_data["sample_len"])
                                    npy_waveforms = np.lib.format.open_memmap(
                                        dataset_paths["waveforms_v"], mode="w+", dtype=np.float32,
                                        shape=(p["lines"], p["points_per_line"], sample_len),
                                    )
                                    npy_valid = np.lib.format.open_memmap(
                                        dataset_paths["valid"], mode="w+", dtype=np.bool_,
                                        shape=(p["lines"], p["points_per_line"]),
                                    )
                                    npy_positions = np.lib.format.open_memmap(
                                        dataset_paths["positions_mm"], mode="w+", dtype=np.float64,
                                        shape=(p["lines"], p["points_per_line"], 3),
                                    )
                                    npy_valid[:] = False
                                    npy_positions[:] = np.nan
                                    np.save(
                                        dataset_paths["time_s"],
                                        np.asarray(point_data["time_s"], dtype=np.float64),
                                    )
                                    with open(dataset_paths["info"], "w", encoding="utf-8") as f:
                                        json.dump({
                                            "shape": [p["lines"], p["points_per_line"], sample_len],
                                            "dtype": "float32",
                                            "unit": "V",
                                            "axis_order": [p["step_axis"], p["span_axis"], "sample"],
                                            "span_coordinates": "span_mm.npy (ascending)",
                                            "step_coordinates": "step_mm.npy (scan order)",
                                            "positions": "positions_mm.npy (X,Y,Z actual queried position)",
                                            "valid_mask": "valid.npy",
                                        }, f, ensure_ascii=False, indent=2)
                                if int(point_data["sample_len"]) != npy_waveforms.shape[2]:
                                    raise RuntimeError(
                                        "示波器返回点数从 {} 变为 {}，无法继续写入固定形状 NPY 数据集".format(
                                            npy_waveforms.shape[2], point_data["sample_len"]
                                        )
                                    )
                                grid_index = scan_grid_index(
                                    point_in_line, p["points_per_line"], direction
                                )
                                npy_waveforms[line_idx - 1, grid_index] = point_data["wave_v"]
                                npy_positions[line_idx - 1, grid_index] = point_data["position"]
                                npy_valid[line_idx - 1, grid_index] = True
                                if (completed_points + 1) % 20 == 0:
                                    npy_waveforms.flush()
                                    npy_positions.flush()
                                    npy_valid.flush()
                        except Exception as exc:
                            raise RuntimeError(
                                "第 {}/{} 点采集失败：{}".format(point_index, p["total_points"], exc)
                            ) from exc
                    if not self.scan_stop_event.is_set():
                        completed_points += 1
                        elapsed_s = time.monotonic() - scan_started_at
                        seconds_per_point = elapsed_s / completed_points
                        eta_s = seconds_per_point * (p["total_points"] - completed_points)
                        if report_point or completed_points == p["total_points"]:
                            phase_text = ""
                            if p.get("acquire_scope"):
                                phase_text = "，采集 {:.2f}s，保存 {:.2f}s".format(
                                    total_acquire_seconds / completed_points,
                                    total_save_seconds / completed_points,
                                )
                            if move_index:
                                phase_text += "，运动 {:.2f}s/次".format(
                                    total_move_seconds / move_index
                                )
                            self.post_ui(
                                self.scan_progress_var.set,
                                "扫描状态：已完成 {}/{} 点，平均 {:.2f}s/点{}，预计剩余 {}".format(
                                    completed_points, p["total_points"], seconds_per_point,
                                    phase_text, format_duration(eta_s),
                                ),
                            )
                        if completed_points % 20 == 0 and self.scan_summary_stream is not None:
                            self.scan_summary_stream.flush()

                    # 行内最后一个点不再继续前进
                    if point_in_line < p["points_per_line"]:
                        distance = direction * p["segments"][point_in_line - 1]
                        cmd = "Move;Rel;{};{};{}".format(
                            self.num_text(speed), span_axis_id, self.num_text(distance)
                        )
                        move_index += 1
                        if report_point:
                            self.log("扫描运动 {}/{}：行内位移 {}{} mm。".format(
                                move_index, p["total_moves"], p["span_axis"],
                                self.num_text(distance)
                            ))
                        current = self.get_current_position_dict()
                        target = current[span_axis_id] + distance
                        self.check_soft_limits(
                            {span_axis_id: target}, p["soft_limits"], "扫描行内目标"
                        )
                        move_started_at = time.monotonic()
                        self.move_and_wait(
                            cmd, {span_axis_id: target}, timeout=180,
                            quiet=not report_point,
                        )
                        total_move_seconds += time.monotonic() - move_started_at

                if self.scan_stop_event.is_set():
                    break

                # 换行：不是最后一行才移动行数轴
                if line_idx < p["lines"]:
                    line_distance = step_direction * p["line_step"]
                    cmd = "Move;Rel;{};{};{}".format(
                        self.num_text(speed), step_axis_id, self.num_text(line_distance)
                    )
                    move_index += 1
                    if report_line:
                        self.log("扫描运动 {}/{}：换行 {}{} mm。".format(
                            move_index, p["total_moves"], p["step_axis"], self.num_text(line_distance)
                        ))
                    current = self.get_current_position_dict()
                    target = current[step_axis_id] + line_distance
                    self.check_soft_limits(
                        {step_axis_id: target}, p["soft_limits"], "扫描换行目标"
                    )
                    move_started_at = time.monotonic()
                    self.move_and_wait(
                        cmd, {step_axis_id: target}, timeout=180,
                        quiet=not report_line,
                    )
                    total_move_seconds += time.monotonic() - move_started_at

            if self.scan_stop_event.is_set():
                self.log("扫描已停止。")
                result_status = "stopped"
                self.post_ui(self.scan_progress_var.set, "扫描状态：已停止")
            else:
                self.log("扫描完成。")
                result_status = "completed"
                elapsed_s = time.monotonic() - scan_started_at
                avg_point_s = elapsed_s / completed_points if completed_points else 0.0
                phase_text = ""
                if p.get("acquire_scope") and completed_points:
                    phase_text = "，采集 {:.2f}s，保存 {:.2f}s".format(
                        total_acquire_seconds / completed_points,
                        total_save_seconds / completed_points,
                    )
                if move_index:
                    phase_text += "，运动 {:.2f}s/次".format(total_move_seconds / move_index)
                self.post_ui(
                    self.scan_progress_var.set,
                    "扫描状态：已完成，用时 {}，平均 {:.2f}s/点{}".format(
                        format_duration(elapsed_s), avg_point_s, phase_text
                    ),
                )

        except InterruptedError as exc:
            result_status = "stopped"
            self.log("扫描已停止：{}".format(exc))
            self.post_ui(self.scan_progress_var.set, "扫描状态：已停止")
        except Exception as exc:
            self.log("!! 扫描失败：{}".format(exc))
            self.stop_after_motion_error("扫描异常：{}".format(exc))
            self.post_ui(self.scan_progress_var.set, "扫描状态：失败")
            self.post_ui(messagebox.showerror, "扫描失败", str(exc))
        finally:
            for dataset in (npy_waveforms, npy_positions, npy_valid):
                if dataset is not None:
                    try:
                        dataset.flush()
                    except Exception as exc:
                        self.log("!! 刷新 NPY 数据集失败：{}".format(exc))
            if self.scan_summary_stream is not None:
                try:
                    self.scan_summary_stream.close()
                except Exception as exc:
                    self.log("!! 关闭扫描摘要文件失败：{}".format(exc))
                self.scan_summary_stream = None
                self.scan_summary_writer = None
            if self.scan_session_dir:
                try:
                    with open(os.path.join(self.scan_session_dir, "scan_result.json"), "w", encoding="utf-8") as f:
                        json.dump({
                            "status": result_status,
                            "points_completed": completed_points,
                            "points_expected": p.get("total_points", 0),
                            "elapsed_seconds": (
                                time.monotonic() - scan_started_at if scan_started_at else 0.0
                            ),
                            "average_seconds": {
                                "per_point": (
                                    (time.monotonic() - scan_started_at) / completed_points
                                    if scan_started_at and completed_points else 0.0
                                ),
                                "acquire": (
                                    total_acquire_seconds / completed_points
                                    if completed_points else 0.0
                                ),
                                "save": (
                                    total_save_seconds / completed_points
                                    if completed_points else 0.0
                                ),
                                "move": total_move_seconds / move_index if move_index else 0.0,
                            },
                            "datasets": dataset_paths,
                            "finished_at": datetime.now().isoformat(timespec="seconds"),
                        }, f, ensure_ascii=False, indent=2)
                except Exception as exc:
                    self.log("!! 写入扫描完成状态失败：{}".format(exc))
            self.scan_running = False
            self.scan_stop_event.set()
            self.motion_operation_lock.release()

    # ============================================================
    # 按住点动
    # ============================================================
    def start_jog_hold(self, axis, sign):
        """
        安全分段点动：
        按住鼠标时，每一段运动完成后，只有鼠标仍然按住，才会发送下一段。
        松开鼠标后立刻置停止标志，并发送 Stop，不再继续发送新的 Move。
        """
        if self.hold_event is not None and not self.hold_event.is_set():
            return

        try:
            self.require_connected()
            segment = abs(self.get_float(self.jog_distance_var, "点动分段距离"))
            if segment <= 0:
                raise ValueError("点动分段距离必须大于 0")
            if segment > 5:
                raise ValueError("点动分段距离过大，建议不要超过 5 mm")
            speed = self.get_speed()
            limits = self.get_soft_limits()
            auto_query = bool(self.auto_query_var.get())
        except Exception as exc:
            messagebox.showerror("输入错误", str(exc))
            return

        if not self.motion_operation_lock.acquire(blocking=False):
            messagebox.showwarning("运动忙", "已有运动或扫描任务正在执行")
            return

        self.software_stop_event.clear()
        self.client.clear_stop_request()

        self.jog_run_id += 1
        run_id = self.jog_run_id

        self.hold_event = threading.Event()
        self.hold_info = {
            "axis": axis,
            "sign": sign,
            "segment": segment,
            "speed": speed,
            "run_id": run_id,
            "limits": limits,
            "auto_query": auto_query,
        }

        direction = "+" if sign > 0 else "-"
        self.log("开始按住点动：{}{}，分段距离 {} mm，速度 {}%。".format(
            AXIS_NAME[axis], direction, self.num_text(segment), self.num_text(speed)
        ))
        self.log("说明：V4 不再发送一次很长的 Move；而是按住时分段发送，松开后不再发送下一段，并立即 Stop。")

        self.hold_thread = threading.Thread(target=self._jog_hold_segment_loop, args=(run_id,), daemon=True)
        self.hold_thread.start()

    def stop_jog_hold(self, event=None):
        """
        鼠标松开/离开按钮时调用。
        立刻停止继续分段，并尽快发 Stop。
        """
        if self.hold_event is None or self.hold_event.is_set():
            return

        try:
            self.request_software_stop("鼠标已松开，停止点动")
        except Exception as exc:
            self.log("!! 松开点动时发送 Stop 失败：{}".format(exc))

    def _jog_hold_segment_loop(self, run_id):
        info = dict(self.hold_info or {})
        axis = info.get("axis", 0)
        sign = info.get("sign", 1)
        segment = float(info.get("segment", 0.5))
        speed = float(info.get("speed", 10))
        limits = info.get("limits", {})
        auto_query = bool(info.get("auto_query", True))
        count = 0

        try:
            self.query_position()
            while True:
                # 每一段发送前都检查：鼠标是否仍按住，是否还是当前这次点动
                if self.hold_event is None or self.hold_event.is_set():
                    break
                if run_id != self.jog_run_id:
                    break

                distance = sign * segment
                position = self.get_current_position_dict()
                target = position[axis] + distance
                self.check_soft_limits(
                    {axis: target}, limits, "点动目标"
                )

                cmd = "Move;Rel;{};{};{}".format(
                    self.num_text(speed),
                    axis,
                    self.num_text(distance),
                )

                self.log("{}轴按住点动{}向：第 {} 段，相对移动 {} mm。".format(
                    AXIS_NAME.get(axis, axis),
                    "正" if sign > 0 else "负",
                    count + 1,
                    self.num_text(abs(distance)),
                ))

                self.move_and_wait(cmd, {axis: target}, timeout=60)
                count += 1

                # 给 UI 松开事件一点时间，不要立即发送下一段
                time.sleep(0.08)

        except Exception as exc:
            self.log("!! 点动运动中断：{}".format(exc))
            self.stop_after_motion_error("点动异常：{}".format(exc))
        finally:
            self.log("点动结束：共发送 {} 段相对运动。".format(count))

            if self.client.connected and auto_query:
                time.sleep(0.15)
                try:
                    self.query_position()
                except Exception as exc:
                    self.log("!! 点动结束后查询位置失败：{}".format(exc))

            if self.hold_event is not None:
                self.hold_event.set()
            self.motion_operation_lock.release()

    # ============================================================
    # 工具
    # ============================================================
    def num_text(self, value):
        try:
            value = float(value)
        except Exception:
            return str(value)

        if abs(value - int(value)) < 1e-9:
            return str(int(value))
        return ("{:.6f}".format(value)).rstrip("0").rstrip(".")

    def on_close(self):
        try:
            self.scope_live_running = False
            self.scan_stop_event.set()
            self.software_stop_event.set()
            if self.hold_event is not None:
                self.hold_event.set()
            if self.client.connected:
                try:
                    self.client.stop_no_wait()
                except Exception:
                    pass
            self.client.disconnect()
            if self.scope_lock.acquire(timeout=2):
                try:
                    self.scope.disconnect()
                finally:
                    self.scope_lock.release()
        except Exception:
            pass
        self.destroy()


def _self_test():
    import numpy as np

    assert format_duration(0) == "00:00:00"
    assert format_duration(65) == "00:01:05"
    assert format_duration(3661) == "01:01:01"
    assert MotionApp.calculate_scan_segments(None, 10, 3) == [3.0, 3.0, 3.0, 1.0]
    assert [scan_grid_index(i, 4, 1) for i in range(1, 5)] == [0, 1, 2, 3]
    assert [scan_grid_index(i, 4, -1) for i in range(1, 5)] == [3, 2, 1, 0]
    assert MotionApp.choose_scan_direction(0, 0, 10, DEFAULT_SOFT_LIMITS) == -1
    assert MotionApp.choose_scan_direction(1, -100, 10, DEFAULT_SOFT_LIMITS) == 1
    assert MotionApp.choose_scan_direction(2, 210, 10, DEFAULT_SOFT_LIMITS) == -1
    assert MotionApp.choose_scan_direction(
        1, -100, 10, DEFAULT_SOFT_LIMITS, "正向（小→大）"
    ) == 1
    assert MotionApp.choose_scan_direction(
        1, -100, 10, DEFAULT_SOFT_LIMITS, "反向（大→小）"
    ) == -1
    assert MotionClient._reply_matches("Move;Rel;10;0;1", "MoveOK")
    assert MotionClient._reply_matches("GetPosition;0,1,2", "Position;0,1,2;1,2,3")

    class FakeScopeInstrument:
        def __init__(self):
            self.writes = []
            self.queries = []

        def write(self, cmd):
            self.writes.append(cmd)

        def query(self, cmd):
            self.queries.append(cmd)
            return "0,0,2,1,0.000001,0,0,0.1,2,100"

        def read_raw(self):
            return b"#12" + bytes([100, 110]) + b"\n"

    fake_scope = RigolScope()
    fake_scope.inst = FakeScopeInstrument()
    fake_scope.is_connected = True
    first_wave = fake_scope.acquire_screen_norm(1, 2, use_cache=True)[0]
    second_wave = fake_scope.acquire_screen_norm(1, 2, use_cache=True)[0]
    assert np.allclose(first_wave, [2.0, 3.0])
    assert np.allclose(second_wave, first_wave)
    assert fake_scope.inst.queries.count(":WAVeform:PREamble?") == 1
    assert fake_scope.inst.writes.count(":WAVeform:DATA?") == 2

    blocked_left, blocked_right = socket.socketpair()
    blocked_client = MotionClient()
    blocked_client.sock = blocked_left
    blocked_client.stop_requested.set()
    try:
        blocked_client.send_no_wait("Move;Rel;10;0;1")
        raise AssertionError("停止状态下不应发送 Move")
    except InterruptedError:
        pass
    blocked_right.settimeout(0.05)
    try:
        assert not blocked_right.recv(1024), "停止状态下发出了 Move"
    except socket.timeout:
        pass
    blocked_client.disconnect()
    blocked_right.close()

    async_left, async_right = socket.socketpair()
    async_client = MotionClient()
    async_client.sock = async_left
    async_client.send_no_wait("Move;Rel;10;0;1")
    assert async_right.recv(1024) == b"Move;Rel;10;0;1\r\n"
    async_client.disconnect()
    async_right.close()

    class FakeMotionClient:
        def __init__(self):
            self.sent = []

        def send_no_wait(self, cmd):
            self.sent.append(cmd)

    fake_app = object.__new__(MotionApp)
    fake_app.client = FakeMotionClient()
    fake_app.software_stop_event = threading.Event()
    fake_app.log = lambda text: None
    fake_app.describe_command = lambda cmd: "测试运动"
    positions = iter([{0: 0.4, 1: 0.0, 2: 0.0}, {0: 1.0, 1: 0.0, 2: 0.0}])
    fake_app.query_position = lambda quiet=False, timeout=5: next(positions)
    final_position = fake_app.move_and_wait(
        "Move;Abs;10;0;1", {0: 1.0}, timeout=1
    )
    assert final_position[0] == 1.0
    assert fake_app.client.sent == ["Move;Abs;10;0;1"]

    left, right = socket.socketpair()
    client = MotionClient()
    client.sock = left

    def fake_server():
        right.recv(1024)
        right.sendall(b"StopOK\r\nPosition;0,1,2;1,2,3\r\n")

    thread = threading.Thread(target=fake_server)
    thread.start()
    assert client.send_command("GetPosition;0,1,2", timeout=1) == "Position;0,1,2;1,2,3"
    thread.join(timeout=1)
    client.disconnect()
    right.close()
    with tempfile.TemporaryDirectory() as folder:
        path = os.path.join(folder, "test.xlsx")
        write_matlab_xlsx(
            path,
            [("X位置(mm)", 1), ("Y位置(mm)", 2), ("Z位置(mm)", 3)],
            [0, 1e-6],
            [0.1, -0.1],
        )
        with zipfile.ZipFile(path) as archive:
            workbook = archive.read("xl/workbook.xml").decode("utf-8")
            assert "元数据" in workbook and "波形数据" in workbook
    print("SELF_TEST_OK")


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        _self_test()
    else:
        app = MotionApp()
        app.mainloop()
