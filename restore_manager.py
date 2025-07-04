import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import subprocess
import sys
import ctypes
import os
import winreg
import re
from datetime import datetime

AUTO_RESTORE_CONF = "auto_restore.conf"

# 检查是否以管理员身份运行
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# 创建还原点
def create_restore_point(name):
    try:
        cmd = [
            'powershell',
            '-Command',
            f'Checkpoint-Computer -Description "{name}" -RestorePointType "MODIFY_SETTINGS"'
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            # 检查输出内容和错误内容，判断是否因策略未创建
            combined_output = (result.stdout + result.stderr).lower()
            keywords = [
                "已存在", "间隔", "already exists", "interval", "no more than one restore point", "每24小时"
            ]
            if any(kw in combined_output for kw in keywords):
                return False, "还原点创建失败：系统策略限制，24小时内只能创建一个同类型还原点。"
            return True, "还原点创建成功！"
        else:
            return False, result.stderr
    except Exception as e:
        return False, str(e)

# 获取还原点列表
def get_restore_points():
    try:
        cmd = [
            'powershell',
            '-Command',
            'Get-ComputerRestorePoint | Select-Object SequenceNumber, Description, CreationTime | ConvertTo-Json'
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            import json
            points = json.loads(result.stdout)
            # 兼容单个还原点时返回dict
            if isinstance(points, dict):
                points = [points]
            return points
        else:
            return []
    except Exception as e:
        return []

# 还原到指定还原点
def restore_to_point(seq_num):
    try:
        cmd = [
            'powershell',
            '-Command',
            f'Restore-Computer -RestorePoint {seq_num}'
        ]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            return True, "系统将还原并重启。"
        else:
            return False, result.stderr
    except Exception as e:
        return False, str(e)

def set_auto_restore_point(seq_num):
    with open(AUTO_RESTORE_CONF, "w", encoding="utf-8") as f:
        f.write(str(seq_num))

def get_auto_restore_point():
    if os.path.exists(AUTO_RESTORE_CONF):
        with open(AUTO_RESTORE_CONF, "r", encoding="utf-8") as f:
            return f.read().strip()
    return None

def clear_auto_restore_point():
    if os.path.exists(AUTO_RESTORE_CONF):
        os.remove(AUTO_RESTORE_CONF)

def add_to_startup():
    exe_path = os.path.abspath(sys.argv[0])
    key = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0, winreg.KEY_SET_VALUE
    )
    winreg.SetValueEx(key, "WindowsRestoreManager", 0, winreg.REG_SZ, exe_path)
    winreg.CloseKey(key)

def auto_restore_if_needed():
    seq_num = get_auto_restore_point()
    if seq_num:
        ok, msg = restore_to_point(seq_num)
        clear_auto_restore_point()
        if ok:
            subprocess.run(['shutdown', '/r', '/t', '5'])
        else:
            messagebox.showerror("自动还原失败", msg)
        sys.exit(0)

# 删除还原点（通过 CreationTime 匹配 Shadow Copy）
def delete_restore_point_by_time(creation_time):
    try:
        cmd = ['vssadmin', 'list', 'shadows']
        result = subprocess.run(cmd, capture_output=True, text=True, encoding='gbk', errors='ignore')
        if result.returncode != 0:
            return False, result.stderr
        pattern = re.compile(r'Shadow Copy ID: {(.*?)}.*?Creation Time: (.*?)\r?\n', re.S)
        matches = pattern.findall(result.stdout)
        # 还原点时间转为datetime
        def parse_restore_time(s):
            try:
                return datetime.strptime(s.split('.')[0], '%Y%m%d%H%M%S')
            except:
                return None
        # Shadow Copy时间格式如: 2024-07-03 07:15:20
        def parse_shadow_time(s):
            try:
                return datetime.strptime(s.strip().split('.')[0], '%Y-%m-%d %H:%M:%S')
            except:
                return None
        target_time = parse_restore_time(str(creation_time))
        for shadow_id, shadow_time in matches:
            shadow_dt = parse_shadow_time(shadow_time)
            if shadow_dt and target_time and abs((shadow_dt - target_time).total_seconds()) < 120:
                # 时间相差2分钟以内就认为是同一个
                del_cmd = ['vssadmin', 'delete', 'shadows', f'/Shadow={{{shadow_id}}}', '/quiet']
                del_result = subprocess.run(del_cmd, capture_output=True, text=True, encoding='gbk', errors='ignore')
                if del_result.returncode == 0:
                    return True, "还原点删除成功！"
                else:
                    return False, del_result.stderr
        return False, "未找到对应的 Shadow Copy，无法删除。\n注意：不是所有还原点都能通过此方式删除。"
    except Exception as e:
        return False, str(e)

def add_task_scheduler_startup():
    import getpass
    exe_path = os.path.abspath(sys.argv[0])
    task_name = "WindowsRestoreManager"
    user_name = getpass.getuser()
    # 检查是否已存在同名任务，存在则先删除
    subprocess.run(f'schtasks /Delete /TN {task_name} /F', shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    cmd = [
        'schtasks', '/Create',
        '/SC', 'ONLOGON',
        '/RL', 'HIGHEST',
        '/TN', task_name,
        '/TR', f'"{exe_path}"',
        '/F',
        '/RU', user_name
    ]
    result = subprocess.run(' '.join(cmd), shell=True)
    return result.returncode == 0

# 主界面
class RestoreManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows还原点管理工具")
        self.root.geometry("600x550")

        # 创建还原点部分
        frame_create = tk.Frame(root)
        frame_create.pack(pady=10)
        tk.Label(frame_create, text="还原点名称:").pack(side=tk.LEFT)
        self.entry_name = tk.Entry(frame_create, width=30)
        self.entry_name.pack(side=tk.LEFT, padx=5)
        tk.Button(frame_create, text="创建还原点", command=self.create_point).pack(side=tk.LEFT)

        # 还原点列表
        self.tree = ttk.Treeview(root, columns=("序号", "名称", "创建时间"), show="headings")
        self.tree.heading("序号", text="序号")
        self.tree.heading("名称", text="名称")
        self.tree.heading("创建时间", text="创建时间")
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)

        # 还原按钮
        tk.Button(root, text="还原到选中还原点", command=self.restore_point).pack(pady=5)
        # 自动还原按钮
        tk.Button(root, text="设为自动还原（关机时自动还原到选中还原点）", command=self.set_auto_restore).pack(pady=5)
        # 设置开机自启动按钮
        tk.Button(root, text="设置本程序开机自启动", command=self.set_startup).pack(pady=5)
        # 删除还原点按钮
        tk.Button(root, text="删除选中还原点", command=self.delete_point).pack(pady=5)
        # 刷新按钮
        tk.Button(root, text="刷新还原点列表", command=self.refresh_points).pack(pady=5)

        self.refresh_points()

    def create_point(self):
        name = self.entry_name.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入还原点名称！")
            return
        ok, msg = create_restore_point(name)
        if ok:
            messagebox.showinfo("成功", msg)
            self.refresh_points()
        else:
            messagebox.showerror("失败", msg)

    def refresh_points(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        points = get_restore_points()
        for pt in points:
            self.tree.insert('', 'end', values=(pt['SequenceNumber'], pt['Description'], pt['CreationTime']))

    def restore_point(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择一个还原点！")
            return
        seq_num = self.tree.item(selected[0])['values'][0]
        if messagebox.askyesno("确认", f"确定要还原到序号为{seq_num}的还原点吗？\n系统将自动重启！"):
            ok, msg = restore_to_point(seq_num)
            if ok:
                messagebox.showinfo("成功", msg)
                # 重启电脑
                subprocess.run(['shutdown', '/r', '/t', '5'])
            else:
                messagebox.showerror("失败", msg)

    def set_auto_restore(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择一个还原点！")
            return
        seq_num = self.tree.item(selected[0])['values'][0]
        set_auto_restore_point(seq_num)
        messagebox.showinfo("成功", f"已设置关机时自动还原到序号为{seq_num}的还原点。\n请确保本程序已设置为开机自启动。")

    def set_startup(self):
        try:
            if add_task_scheduler_startup():
                messagebox.showinfo("成功", "已通过计划任务设置本程序为开机自启动（管理员权限）！")
            else:
                add_to_startup()
                messagebox.showwarning("部分成功", "计划任务创建失败，已尝试用注册表自启动（无管理员权限）。如需自动还原请手动用管理员权限运行。")
        except Exception as e:
            messagebox.showerror("失败", f"设置开机自启动失败：{e}")

    def delete_point(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择一个还原点！")
            return
        item = self.tree.item(selected[0])
        seq_num = item['values'][0]
        creation_time = item['values'][2]
        if not messagebox.askyesno("确认", f"确定要删除序号为{seq_num}的还原点吗？\n此操作不可恢复！"):
            return
        ok, msg = delete_restore_point_by_time(creation_time)
        if ok:
            messagebox.showinfo("成功", msg)
            self.refresh_points()
        else:
            messagebox.showerror("失败", msg)

if __name__ == "__main__":
    if not is_admin():
        messagebox.showerror("权限不足", "请以管理员身份运行本程序！")
        sys.exit(1)
    auto_restore_if_needed()
    root = tk.Tk()
    app = RestoreManagerApp(root)
    root.mainloop() 