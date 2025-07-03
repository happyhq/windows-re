import tkinter as tk
from tkinter import messagebox, simpledialog, ttk
import subprocess
import sys
import ctypes

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

# 主界面
class RestoreManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows还原点管理工具")
        self.root.geometry("600x400")

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

if __name__ == "__main__":
    if not is_admin():
        messagebox.showerror("权限不足", "请以管理员身份运行本程序！")
        sys.exit(1)
    root = tk.Tk()
    app = RestoreManagerApp(root)
    root.mainloop() 