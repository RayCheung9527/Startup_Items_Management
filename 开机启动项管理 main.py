# -*- coding: utf-8 -*-
"""
Windows 开机启动管理器
===========================================================
功能：
1. 扫描当前用户 / 所有用户 Startup 启动文件夹
2. 扫描 HKCU/HKLM 的 Run、RunOnce
3. 扫描 Windows 自动启动服务
4. 正确解析带空格、带参数的启动命令
5. 启动项图标提取
6. 双击启动项定位真实 EXE 文件
7. 当前用户可删除的启动项直接删除
8. HKLM、所有用户 Startup、系统服务显示但默认只读
9. 拖放 .exe / .lnk 到窗口即可加入当前用户 Startup
10. 新增“打开启动文件夹”按钮，一键打开 Startup 目录
===========================================================
依赖：
    pip install pywin32 pillow tkinterdnd2
"""

import os
import re
import sys
import traceback
import subprocess
import shutil
import winreg

import win32com.client
import win32service
import win32gui
import win32ui
import win32con

from PIL import Image, ImageTk
from tkinter import Tk, ttk, messagebox, Button, Label, Frame
from tkinterdnd2 import DND_FILES, TkinterDnD


# =========================================================
# 错误日志
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
LOG_FILE = os.path.join(BASE_DIR, "error.log")


def log_error(msg):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
            f.flush()
    except Exception:
        pass


# =========================================================
# 路径 / 环境变量
# =========================================================

CURRENT_USER_STARTUP = os.path.join(
    os.getenv("APPDATA", ""),
    "Microsoft",
    "Windows",
    "Start Menu",
    "Programs",
    "Startup"
)

ALL_USERS_STARTUP = os.path.join(
    os.getenv("PROGRAMDATA", ""),
    "Microsoft",
    "Windows",
    "Start Menu",
    "Programs",
    "Startup"
)

os.makedirs(CURRENT_USER_STARTUP, exist_ok=True)


# =========================================================
# 图标缓存
# =========================================================

_icon_cache = {}
ICON_SIZE = 16


# =========================================================
# 启动命令解析
# =========================================================

def extract_exe_path(cmd_line):
    """
    从 Windows 启动命令中尽可能可靠地提取 EXE 路径。

    支持：
        "C:\\Program Files\\App\\App.exe" /startup
        C:\\Program Files\\App\\App.exe /startup
        %ProgramFiles%\\App\\App.exe /startup
        C:\\App\\App.exe
        App.exe /startup

    特别修复：
        C:\\Program Files (x86)\\Internet Download Manager\\IDMan.exe /onboot

    不再使用 split()[0]，避免把带空格的路径截断。
    """

    if cmd_line is None:
        return ""

    cmd_line = str(cmd_line).strip()

    if not cmd_line:
        return ""

    # 去掉首尾空白，并展开环境变量
    cmd_line = os.path.expandvars(cmd_line).strip()

    # -----------------------------------------------------
    # 1. 标准带引号：
    # "C:\Program Files\App\App.exe" /startup
    # -----------------------------------------------------
    if cmd_line.startswith('"'):
        match = re.match(r'^"([^"]+)"', cmd_line)

        if match:
            path = match.group(1).strip()
            path = os.path.expandvars(path)

            if path:
                return os.path.normpath(path)

    # -----------------------------------------------------
    # 2. 不带引号路径：
    # C:\Program Files\App\App.exe /startup
    #
    # 关键：寻找命令中的第一个 .exe
    # -----------------------------------------------------
    match = re.match(
        r'^(.*?\.exe)(?:\s+.*)?$',
        cmd_line,
        re.IGNORECASE
    )

    if match:
        path = match.group(1).strip().strip('"')
        path = os.path.expandvars(path)

        if path:
            return os.path.normpath(path)

    # -----------------------------------------------------
    # 3. 可能只是：
    # App.exe /startup
    # -----------------------------------------------------
    parts = cmd_line.split()

    if parts:
        candidate = os.path.expandvars(parts[0]).strip('"')

        # 绝对 / 相对路径
        if os.path.isfile(candidate):
            return os.path.normpath(candidate)

        # PATH 中查找
        found = shutil.which(candidate)

        if found:
            return os.path.normpath(found)

        return candidate

    return ""


def split_startup_command(cmd_line):
    """
    返回：
        exe_path, arguments

    例如：
        C:\\Program Files\\App\\App.exe /silent /startup

    返回：
        C:\\Program Files\\App\\App.exe
        /silent /startup
    """

    if not cmd_line:
        return "", ""

    raw = str(cmd_line).strip()
    expanded = os.path.expandvars(raw)

    exe_path = extract_exe_path(expanded)

    if not exe_path:
        return "", raw

    # 尽可能从原命令中切掉 exe 部分
    if expanded.startswith('"'):
        prefix = '"' + exe_path + '"'
        if expanded.lower().startswith(prefix.lower()):
            args = expanded[len(prefix):].strip()
            return exe_path, args

    # 不带引号的 exe
    pos = expanded.lower().find(exe_path.lower())

    if pos == 0:
        args = expanded[len(exe_path):].strip()
        return exe_path, args

    return exe_path, ""


def path_exists(path):
    if not path:
        return False

    try:
        return os.path.isfile(os.path.expandvars(path))
    except Exception:
        return False


# =========================================================
# 快捷方式解析
# =========================================================

def resolve_shortcut(lnk_path):
    """
    返回快捷方式信息：
        {
            target: 目标程序
            arguments: 参数
            working_directory: 工作目录
            icon_location: 图标路径
        }
    """

    result = {
        "target": "",
        "arguments": "",
        "working_directory": "",
        "icon_location": "",
    }

    if not lnk_path or not os.path.isfile(lnk_path):
        return result

    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)

        result["target"] = os.path.expandvars(shortcut.TargetPath or "")
        result["arguments"] = shortcut.Arguments or ""
        result["working_directory"] = os.path.expandvars(
            shortcut.WorkingDirectory or ""
        )
        result["icon_location"] = os.path.expandvars(
            shortcut.IconLocation or ""
        )

    except Exception as e:
        log_error(f"解析快捷方式失败：{lnk_path}\n{traceback.format_exc()}")

    return result


# =========================================================
# 图标
# =========================================================

def parse_icon_location(icon_location):
    """
    解析：
        C:\\xxx\\xxx.exe,0
        C:\\xxx\\xxx.dll,3
    """

    if not icon_location:
        return ""

    value = os.path.expandvars(str(icon_location)).strip()

    # IconLocation 通常是：
    # path,index
    match = re.match(r'^(.*?),(?:-?\d+)$', value)

    if match:
        return match.group(1).strip().strip('"')

    return value.strip('"')


def extract_icon_from_file(file_path, size=ICON_SIZE):
    """
    从 exe/dll/ico 等文件提取图标。
    返回 PIL Image。
    """

    if not file_path:
        return None

    file_path = os.path.expandvars(str(file_path)).strip().strip('"')

    if not os.path.isfile(file_path):
        return None

    cache_key = (os.path.normcase(os.path.abspath(file_path)), size)

    if cache_key in _icon_cache:
        return _icon_cache[cache_key]

    hicon = None
    mem_dc = None
    hdc = None
    hbmp = None

    try:
        # -------------------------------------------------
        # 优先使用 ExtractIcon
        # -------------------------------------------------
        hicon = win32gui.ExtractIcon(
            0,
            file_path,
            0
        )

        if not hicon:
            return None

        screen_dc = win32gui.GetDC(0)
        hdc = win32ui.CreateDCFromHandle(screen_dc)

        mem_dc = hdc.CreateCompatibleDC()

        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(
            hdc,
            size,
            size
        )

        mem_dc.SelectObject(hbmp)

        win32gui.DrawIconEx(
            mem_dc.GetHandleOutput(),
            0,
            0,
            hicon,
            size,
            size,
            0,
            None,
            win32con.DI_NORMAL
        )

        bmpinfo = hbmp.GetInfo()
        bmpstr = hbmp.GetBitmapBits(True)

        img = Image.frombuffer(
            "RGB",
            (
                bmpinfo["bmWidth"],
                bmpinfo["bmHeight"]
            ),
            bmpstr,
            "raw",
            "BGRX",
            0,
            1
        )

        if img.size != (size, size):
            img = img.resize(
                (size, size),
                Image.Resampling.LANCZOS
            )

        # copy 一份，避免底层 bitmap 生命周期影响 PIL
        img = img.copy()

        _icon_cache[cache_key] = img

        return img

    except Exception as e:
        log_error(
            f"图标提取失败：{file_path}\n"
            f"{traceback.format_exc()}"
        )
        return None

    finally:
        try:
            if hicon:
                win32gui.DestroyIcon(hicon)
        except Exception:
            pass

        try:
            if mem_dc:
                mem_dc.DeleteDC()
        except Exception:
            pass

        try:
            if hdc:
                hdc.DeleteDC()
        except Exception:
            pass

        try:
            if hbmp:
                hbmp.DeleteObject()
        except Exception:
            pass


def get_file_icon_image(file_path, size=ICON_SIZE):
    """
    智能获取启动项图标。

    支持：
        exe
        lnk
        启动命令行
        环境变量
    """

    if not file_path:
        return None

    original = os.path.expandvars(str(file_path).strip())

    # -----------------------------------------------------
    # .lnk
    # -----------------------------------------------------
    if original.lower().endswith(".lnk") and os.path.isfile(original):

        info = resolve_shortcut(original)

        # 快捷方式本身指定了图标
        icon_file = parse_icon_location(
            info.get("icon_location", "")
        )

        if icon_file and os.path.isfile(icon_file):
            img = extract_icon_from_file(icon_file, size)

            if img:
                return img

        # 否则使用目标程序图标
        target = info.get("target", "")

        if target:
            img = extract_icon_from_file(target, size)

            if img:
                return img

        # 最后尝试快捷方式本身
        return extract_icon_from_file(original, size)

    # -----------------------------------------------------
    # 已经是实际文件
    # -----------------------------------------------------
    if os.path.isfile(original):
        return extract_icon_from_file(original, size)

    # -----------------------------------------------------
    # 启动命令：
    # C:\Program Files\App\App.exe /startup
    # -----------------------------------------------------
    exe_path = extract_exe_path(original)

    if exe_path:
        if os.path.isfile(exe_path):
            return extract_icon_from_file(exe_path, size)

        found = shutil.which(exe_path)

        if found:
            return extract_icon_from_file(found, size)

    return None


def get_photo_image(file_path, size=ICON_SIZE):
    img = get_file_icon_image(file_path, size)

    if img:
        return ImageTk.PhotoImage(img)

    return None


# =========================================================
# 注册表
# =========================================================

def enum_registry_values(hive, subkey):
    items = []

    try:
        key = winreg.OpenKey(
            hive,
            subkey,
            0,
            winreg.KEY_READ
        )

        index = 0

        while True:
            try:
                name, value, value_type = winreg.EnumValue(
                    key,
                    index
                )

                items.append(
                    (
                        name,
                        value,
                        value_type
                    )
                )

                index += 1

            except OSError:
                break

        winreg.CloseKey(key)

    except OSError:
        pass

    return items


def delete_registry_value(hive, subkey, name):
    try:
        key = winreg.OpenKey(
            hive,
            subkey,
            0,
            winreg.KEY_SET_VALUE
        )

        winreg.DeleteValue(key, name)
        winreg.CloseKey(key)

        return True

    except Exception as e:
        log_error(
            f"删除注册表值失败：{subkey}\\{name}\n"
            f"{traceback.format_exc()}"
        )
        return False


# =========================================================
# Windows 自动启动服务
# =========================================================

def get_auto_start_services():
    services = []
    scm_handle = None

    try:
        scm_handle = win32service.OpenSCManager(
            None,
            None,
            win32service.SC_MANAGER_ENUMERATE_SERVICE
        )

        services_list = win32service.EnumServicesStatus(
            scm_handle,
            win32service.SERVICE_WIN32,
            win32service.SERVICE_STATE_ALL
        )

        for service in services_list:
            try:
                service_name = service[0]
                display_name = service[1]

                config = win32service.QueryServiceConfig(
                    scm_handle,
                    service_name
                )

                start_type = config[1]

                if start_type != win32service.SERVICE_AUTO_START:
                    continue

                bin_path = config[3] if len(config) > 3 else ""

                services.append({
                    "display_name": display_name,
                    "source": "系统服务（自动启动）",
                    "location": bin_path or "",
                    "raw_path": service_name,
                    "delete_action": None,
                    "is_registry": False,
                    "is_service": True,
                    "exe_path": extract_exe_path(bin_path),
                    "arguments": split_startup_command(bin_path)[1],
                })

            except Exception:
                continue

    except Exception as e:
        log_error(
            f"服务扫描失败：\n{traceback.format_exc()}"
        )

    finally:
        try:
            if scm_handle:
                win32service.CloseServiceHandle(scm_handle)
        except Exception:
            pass

    # 排序，让列表更稳定
    services.sort(
        key=lambda x: x.get("display_name", "").lower()
    )

    return services


# =========================================================
# 启动项扫描
# =========================================================

def make_registry_item(
    name,
    value,
    hive_name,
    subkey,
    writable=False
):
    exe_path, arguments = split_startup_command(value)

    hive = (
        winreg.HKEY_CURRENT_USER
        if hive_name == "HKCU"
        else winreg.HKEY_LOCAL_MACHINE
    )

    return {
        "display_name": str(name),
        "source": f"注册表 {hive_name}\\{subkey.split(chr(92))[-1]}",
        "location": str(value),
        "raw_path": (
            f"计算机\\HKEY_CURRENT_USER\\{subkey}"
            if hive_name == "HKCU"
            else
            f"计算机\\HKEY_LOCAL_MACHINE\\{subkey}"
        ),
        "delete_action": (
            lambda h=hive, s=subkey, n=name:
                delete_registry_value(h, s, n)
        ) if writable else None,
        "is_registry": True,
        "is_service": False,
        "exe_path": exe_path,
        "arguments": arguments,
    }


def get_all_startup_items():
    items = []

    # =====================================================
    # 1. 当前用户 Startup
    # =====================================================

    if os.path.exists(CURRENT_USER_STARTUP):

        try:
            filenames = os.listdir(CURRENT_USER_STARTUP)
        except Exception:
            filenames = []

        for fname in sorted(filenames, key=str.lower):

            if not fname.lower().endswith(".lnk"):
                continue

            full_path = os.path.join(
                CURRENT_USER_STARTUP,
                fname
            )

            info = resolve_shortcut(full_path)

            target = info.get("target", "") or full_path

            items.append({
                "display_name": fname[:-4],
                "source": "启动文件夹（当前用户）",
                "location": target,
                "raw_path": full_path,
                "delete_action": (
                    lambda p=full_path:
                    os.remove(p)
                ),
                "is_registry": False,
                "is_service": False,
                "is_shortcut": True,
                "shortcut_path": full_path,
                "exe_path": (
                    extract_exe_path(target)
                    if target
                    else ""
                ),
                "arguments": info.get("arguments", ""),
                "icon_path": info.get("icon_location", ""),
            })

    # =====================================================
    # 2. 所有用户 Startup
    # =====================================================

    if os.path.exists(ALL_USERS_STARTUP):

        try:
            filenames = os.listdir(ALL_USERS_STARTUP)
        except Exception:
            filenames = []

        for fname in sorted(filenames, key=str.lower):

            if not fname.lower().endswith(".lnk"):
                continue

            full_path = os.path.join(
                ALL_USERS_STARTUP,
                fname
            )

            info = resolve_shortcut(full_path)

            target = info.get("target", "") or full_path

            items.append({
                "display_name": fname[:-4],
                "source": "启动文件夹（所有用户）",
                "location": target,
                "raw_path": full_path,
                "delete_action": None,
                "is_registry": False,
                "is_service": False,
                "is_shortcut": True,
                "shortcut_path": full_path,
                "exe_path": (
                    extract_exe_path(target)
                    if target
                    else ""
                ),
                "arguments": info.get("arguments", ""),
                "icon_path": info.get("icon_location", ""),
            })

    # =====================================================
    # 3. HKCU Run
    # =====================================================

    hkcu_run = (
        r"Software\Microsoft\Windows\CurrentVersion\Run"
    )

    for name, value, _ in enum_registry_values(
        winreg.HKEY_CURRENT_USER,
        hkcu_run
    ):
        items.append(
            make_registry_item(
                name,
                value,
                "HKCU",
                hkcu_run,
                writable=True
            )
        )

    # =====================================================
    # 4. HKLM Run
    # =====================================================

    hklm_run = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
    )

    for name, value, _ in enum_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        hklm_run
    ):
        items.append(
            make_registry_item(
                name,
                value,
                "HKLM",
                hklm_run,
                writable=False
            )
        )

    # =====================================================
    # 5. HKCU RunOnce
    # =====================================================

    hkcu_runonce = (
        r"Software\Microsoft\Windows\CurrentVersion\RunOnce"
    )

    for name, value, _ in enum_registry_values(
        winreg.HKEY_CURRENT_USER,
        hkcu_runonce
    ):
        items.append(
            make_registry_item(
                name,
                value,
                "HKCU",
                hkcu_runonce,
                writable=True
            )
        )

    # =====================================================
    # 6. HKLM RunOnce
    # =====================================================

    hklm_runonce = (
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
    )

    for name, value, _ in enum_registry_values(
        winreg.HKEY_LOCAL_MACHINE,
        hklm_runonce
    ):
        items.append(
            make_registry_item(
                name,
                value,
                "HKLM",
                hklm_runonce,
                writable=False
            )
        )

    # =====================================================
    # 7. 自动启动服务
    # =====================================================

    items.extend(get_auto_start_services())

    return items


# =========================================================
# 添加启动项
# =========================================================

def add_startup_item(file_path):
    """
    把 EXE / LNK 添加到当前用户 Startup 文件夹。
    """

    if not file_path:
        return False

    file_path = os.path.abspath(
        os.path.expandvars(file_path)
    )

    if not os.path.isfile(file_path):
        return False

    base_name = os.path.splitext(
        os.path.basename(file_path)
    )[0]

    shortcut_path = os.path.join(
        CURRENT_USER_STARTUP,
        base_name + ".lnk"
    )

    counter = 1

    while os.path.exists(shortcut_path):
        shortcut_path = os.path.join(
            CURRENT_USER_STARTUP,
            f"{base_name}_{counter}.lnk"
        )
        counter += 1

    try:
        shell = win32com.client.Dispatch(
            "WScript.Shell"
        )

        shortcut = shell.CreateShortCut(
            shortcut_path
        )

        shortcut.TargetPath = file_path

        if file_path.lower().endswith(".exe"):
            shortcut.WorkingDirectory = os.path.dirname(
                file_path
            )

        shortcut.save()

        return True

    except Exception:
        log_error(
            f"添加启动项失败：{file_path}\n"
            f"{traceback.format_exc()}"
        )
        return False


# =========================================================
# GUI
# =========================================================

class App:

    def __init__(self, root):

        self.root = root

        root.title("开机启动管理")
        root.geometry("1000x560")
        root.minsize(800, 420)
        root.resizable(True, True)

        # -------------------------------------------------
        # 拖放
        # -------------------------------------------------

        root.drop_target_register(DND_FILES)
        root.dnd_bind(
            "<<Drop>>",
            self.on_drop
        )

        # -------------------------------------------------
        # 顶部区域：标题 + 状态（左上角）
        # -------------------------------------------------

        top_frame = Frame(root)
        top_frame.pack(fill="x", padx=10, pady=5)

        Label(
            top_frame,
            text="开机启动项列表（双击定位真实文件 / 文件夹）",
            font=("Arial", 12)
        ).pack(side="left")

        status_frame = Frame(root)
        status_frame.pack(fill="x", padx=10, pady=(0, 5))

        self.status = Label(
            status_frame,
            text="拖放 .exe 或 .lnk 文件添加开机启动",
            fg="blue"
        )
        self.status.pack(side="left")

        # -------------------------------------------------
        # Treeview
        # -------------------------------------------------

        frame = Frame(root)
        frame.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=5
        )

        self.tree = ttk.Treeview(
            frame,
            columns=(
                "name",
                "source",
                "location"
            ),
            show="tree headings"
        )

        self.tree.heading(
            "#0",
            text="图标"
        )

        self.tree.heading(
            "name",
            text="软件名"
        )

        self.tree.heading(
            "source",
            text="来源"
        )

        self.tree.heading(
            "location",
            text="启动位置"
        )

        self.tree.column(
            "#0",
            width=45,
            minwidth=45,
            anchor="center",
            stretch=False
        )

        self.tree.column(
            "name",
            width=160,
            minwidth=100
        )

        self.tree.column(
            "source",
            width=180,
            minwidth=130
        )

        self.tree.column(
            "location",
            width=550,
            minwidth=300
        )

        self.tree.bind(
            "<Double-Button-1>",
            self.on_double_click
        )

        # -------------------------------------------------
        # 滚动条
        # -------------------------------------------------

        vsb = ttk.Scrollbar(
            frame,
            orient="vertical",
            command=self.tree.yview
        )

        hsb = ttk.Scrollbar(
            root,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=vsb.set,
            xscrollcommand=hsb.set
        )

        self.tree.pack(
            side="left",
            fill="both",
            expand=True
        )

        vsb.pack(
            side="right",
            fill="y"
        )

        hsb.pack(
            fill="x",
            padx=10
        )

        # -------------------------------------------------
        # 按钮（三个，均匀分布）
        # -------------------------------------------------

        btn_frame = Frame(root)
        btn_frame.pack(
            fill="x",
            pady=8,
            padx=20
        )

        # 三个按钮均匀分布
        self.btn_remove = Button(
            btn_frame,
            text="取消开机启动（选中项）",
            command=self.remove_selected,
            width=20
        )
        self.btn_remove.pack(side="left", expand=True, padx=5)

        self.btn_open_folder = Button(
            btn_frame,
            text="打开启动文件夹",
            command=self.open_startup_folder,
            width=18
        )
        self.btn_open_folder.pack(side="left", expand=True, padx=5)

        self.btn_refresh = Button(
            btn_frame,
            text="刷新列表",
            command=self.refresh_list,
            width=15
        )
        self.btn_refresh.pack(side="left", expand=True, padx=5)

        # -------------------------------------------------
        # 数据
        # -------------------------------------------------

        self.current_items = []
        self.id_to_idx = {}
        self.icon_images = []

        self.refresh_list()

    # =====================================================
    # 刷新
    # =====================================================

    def refresh_list(self):

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.id_to_idx.clear()
        self.icon_images.clear()

        self.current_items = get_all_startup_items()

        if not self.current_items:

            self.tree.insert(
                "",
                "end",
                values=(
                    "（无启动项）",
                    "",
                    ""
                )
            )

            self.status.config(
                text="没有找到任何开机启动项",
                fg="orange"
            )

            return

        for idx, item in enumerate(
            self.current_items
        ):

            iid = str(idx)

            # -------------------------------------------------
            # 优先使用专门的图标路径
            # 其次使用 exe_path
            # 最后使用 location
            # -------------------------------------------------

            icon_source = ""

            if item.get("icon_path"):
                icon_source = parse_icon_location(
                    item.get("icon_path")
                )

            if (
                not icon_source
                or not os.path.isfile(
                    os.path.expandvars(icon_source)
                )
            ):
                icon_source = item.get(
                    "exe_path",
                    ""
                )

            if not icon_source:
                icon_source = item.get(
                    "location",
                    ""
                )

            img = get_photo_image(
                icon_source,
                size=ICON_SIZE
            )

            if img:
                self.icon_images.append(img)
                icon = img
            else:
                icon = ""

            # -------------------------------------------------
            # 插入 Treeview
            # -------------------------------------------------

            self.tree.insert(
                "",
                "end",
                iid=iid,
                image=icon,
                values=(
                    item.get(
                        "display_name",
                        ""
                    ),
                    item.get(
                        "source",
                        ""
                    ),
                    item.get(
                        "location",
                        ""
                    )
                )
            )

            self.id_to_idx[iid] = idx

        self.status.config(
            text=f"共 {len(self.current_items)} 个启动项    |    拖放 .exe 或 .lnk 文件添加开机启动",
            fg="green"
        )

    # =====================================================
    # 打开启动文件夹
    # =====================================================

    def open_startup_folder(self):
        """打开当前用户的 Startup 文件夹"""
        try:
            os.startfile(CURRENT_USER_STARTUP)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开启动文件夹：{e}")

    # =====================================================
    # 双击
    # =====================================================

    def on_double_click(self, event):

        sel = self.tree.selection()

        if not sel:
            return

        iid = sel[0]

        idx = self.id_to_idx.get(iid)

        if idx is None:
            return

        item = self.current_items[idx]

        try:

            # -------------------------------------------------
            # 服务
            # -------------------------------------------------

            if item.get("is_service", False):

                exe_path = item.get(
                    "exe_path",
                    ""
                )

                if exe_path and os.path.isfile(
                    os.path.expandvars(exe_path)
                ):
                    self.open_in_explorer(
                        exe_path
                    )
                    return

                msg = (
                    f"该项是一个系统服务："
                    f"{item.get('display_name', '')}\n\n"
                    f"服务名："
                    f"{item.get('raw_path', '')}\n\n"
                    f"是否打开服务管理器？"
                )

                if messagebox.askyesno(
                    "系统服务",
                    msg
                ):
                    os.startfile(
                        "services.msc"
                    )

                return

            # -------------------------------------------------
            # 注册表
            # -------------------------------------------------

            if item.get(
                "is_registry",
                False
            ):

                exe_path = item.get(
                    "exe_path",
                    ""
                )

                # 再次解析，防止旧数据结构
                if not exe_path:
                    exe_path = extract_exe_path(
                        item.get(
                            "location",
                            ""
                        )
                    )

                if exe_path:
                    exe_path = os.path.expandvars(
                        exe_path
                    )

                if exe_path and os.path.isfile(
                    exe_path
                ):
                    self.open_in_explorer(
                        exe_path
                    )
                    return

                # PATH 再查一次
                if exe_path:
                    found = shutil.which(
                        exe_path
                    )

                    if found:
                        self.open_in_explorer(
                            found
                        )
                        return

                msg = (
                    f"无法定位可执行文件："
                    f"{exe_path or '未解析到 EXE'}"
                    f"\n\n"
                    f"启动命令："
                    f"{item.get('location', '')}"
                    f"\n\n"
                    f"是否打开注册表编辑器以查看该键值？"
                )

                if messagebox.askyesno(
                    "定位失败",
                    msg
                ):
                    self.open_registry_editor(
                        item
                    )

                return

            # -------------------------------------------------
            # Startup 文件夹
            # -------------------------------------------------

            shortcut_path = item.get(
                "shortcut_path",
                ""
            )

            if shortcut_path and os.path.isfile(
                shortcut_path
            ):
                # 优先选中 Startup 中的快捷方式
                self.open_in_explorer(
                    shortcut_path
                )
                return

            target_path = item.get(
                "location",
                ""
            )

            if target_path:
                target_path = os.path.expandvars(
                    target_path
                )

            if target_path and os.path.exists(
                target_path
            ):
                self.open_in_explorer(
                    target_path
                )
                return

            fallback = item.get(
                "raw_path",
                ""
            )

            if fallback and os.path.exists(
                fallback
            ):
                self.open_in_explorer(
                    fallback
                )
                return

            messagebox.showwarning(
                "提示",
                "无法定位文件路径"
            )

        except Exception as e:

            log_error(
                f"双击定位失败：\n"
                f"{traceback.format_exc()}"
            )

            messagebox.showerror(
                "错误",
                f"操作失败：{e}"
            )

    # =====================================================
    # Explorer 定位
    # =====================================================

    @staticmethod
    def open_in_explorer(path):

        path = os.path.abspath(
            os.path.expandvars(path)
        )

        if not os.path.exists(path):
            return False

        try:
            subprocess.Popen(
                [
                    "explorer.exe",
                    "/select,",
                    path
                ]
            )
            return True

        except Exception:
            try:
                os.startfile(
                    os.path.dirname(path)
                )
                return True
            except Exception:
                return False

    # =====================================================
    # 打开注册表
    # =====================================================

    @staticmethod
    def open_registry_editor(item):

        try:
            subprocess.Popen(
                [
                    "regedit.exe"
                ]
            )
            return True

        except Exception:
            try:
                os.startfile(
                    "regedit.exe"
                )
                return True
            except Exception:
                return False

    # =====================================================
    # 删除启动项
    # =====================================================

    def remove_selected(self):

        sel = self.tree.selection()

        if not sel:

            messagebox.showwarning(
                "警告",
                "请先选择一个启动项！"
            )

            return

        iid = sel[0]

        idx = self.id_to_idx.get(iid)

        if idx is None:
            return

        item = self.current_items[idx]

        name = item.get(
            "display_name",
            ""
        )

        source = item.get(
            "source",
            ""
        )

        if not messagebox.askyesno(
            "确认",
            f"确定取消“{name}”吗？\n"
            f"来源：{source}"
        ):
            return

        delete_action = item.get(
            "delete_action"
        )

        if delete_action is None:

            messagebox.showinfo(
                "提示",
                "该项需要管理员权限或不能直接删除，请手动处理。"
            )

            return

        try:

            result = delete_action()

            if result is False:

                messagebox.showerror(
                    "失败",
                    f"无法删除“{name}”。"
                )

                return

            messagebox.showinfo(
                "成功",
                f"已取消“{name}”"
            )

            self.refresh_list()

        except Exception as e:

            log_error(
                f"删除启动项失败：\n"
                f"{traceback.format_exc()}"
            )

            messagebox.showerror(
                "错误",
                f"删除失败：{e}"
            )

    # =====================================================
    # 拖放
    # =====================================================

    def on_drop(self, event):

        raw = event.data

        files = []

        try:
            # tkinterdnd2 的 splitlist 能正确处理：
            # {C:\Program Files\App\App.exe}
            # 这类带空格路径
            try:
                dropped = self.root.tk.splitlist(
                    raw
                )
            except Exception:
                dropped = [raw]

            for part in dropped:

                part = part.strip()

                if (
                    part.startswith("{")
                    and part.endswith("}")
                ):
                    part = part[1:-1]

                part = os.path.expandvars(
                    part
                )

                if os.path.isfile(part):
                    files.append(part)

        except Exception:

            log_error(
                f"解析拖放文件失败：\n"
                f"{traceback.format_exc()}"
            )

        added = 0

        for file_path in files:

            ext = os.path.splitext(
                file_path
            )[1].lower()

            if ext not in (
                ".exe",
                ".lnk"
            ):
                continue

            if add_startup_item(
                file_path
            ):
                added += 1

        if added:

            self.refresh_list()

            messagebox.showinfo(
                "添加成功",
                f"成功添加 {added} 个启动项。"
            )

        else:

            self.status.config(
                text=(
                    "未添加任何项，请拖入 .exe 或 .lnk"
                ),
                fg="orange"
            )


# =========================================================
# 主程序
# =========================================================

def main():

    try:

        root = TkinterDnD.Tk()

        app = App(root)

        root.mainloop()

    except Exception as e:

        err_msg = (
            f"程序启动失败：{str(e)}\n"
            f"{traceback.format_exc()}"
        )

        log_error(err_msg)

        try:
            from tkinter import Tk, messagebox

            tmp = Tk()
            tmp.withdraw()

            messagebox.showerror(
                "启动失败",
                "程序启动失败，请查看 error.log 文件。\n\n"
                + str(e)
            )

            tmp.destroy()

        except Exception:
            pass

        sys.exit(1)


if __name__ == "__main__":
    main()
