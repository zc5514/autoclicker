import tkinter as tk
from tkinter import messagebox, simpledialog
import threading
import time
import win32api
import win32con
import win32gui
import keyboard
import json
import os

class AutoClickerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Clicker")
        self.click_interval = tk.IntVar(value=50)  # 默认间隔50ms
        self.click_position = (0, 0)
        self.is_clicking = False
        self.window_handle = None
        self.window_title = ""
        self.hotkey = "F3"  # 默认快捷键
        self.positions = []  # 存储坐标和句柄的列表
        self.save_file = "autoclicker_data.json"

        # 创建GUI元素
        self.create_widgets()
        # 监听默认快捷键
        keyboard.add_hotkey(self.hotkey, self.toggle_clicking)
        # 加载之前保存的数据
        self.load_data()

    def create_widgets(self):
        self.root.option_add("*Font", "Arial 10")
        self.root.option_add("*Background", "#f0f0f0")
        
        # 点击间隔
        tk.Label(self.root, text="点击间隔 (ms):").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.interval_entry = tk.Entry(self.root, textvariable=self.click_interval)
        self.interval_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # 点击位置
        tk.Label(self.root, text="点击位置:").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.position_label = tk.Label(self.root, text="未选择", relief="sunken", bd=1)
        self.position_label.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # 选择点击位置按钮
        select_button = tk.Button(self.root, text="选择点击位置", command=self.select_position, height=2, width=15)
        select_button.grid(row=2, column=0, padx=10, pady=10, sticky="ew")

        # 启动/停止点击按钮
        self.toggle_button = tk.Button(self.root, text="启动/停止点击", command=self.toggle_clicking, height=2, width=15)
        self.toggle_button.grid(row=2, column=1, padx=10, pady=10, sticky="e")

        # 窗口句柄
        tk.Label(self.root, text="窗口句柄:").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.handle_label = tk.Label(self.root, text="未获取", relief="sunken", bd=1)
        self.handle_label.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # 窗口标题
        tk.Label(self.root, text="窗口标题:").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.title_label = tk.Label(self.root, text="未获取", relief="sunken", bd=1)
        self.title_label.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        # 列表框
        self.listbox = tk.Listbox(self.root, width=40, height=6, relief="sunken", bd=1, bg="white")
        self.listbox.grid(row=5, column=0, columnspan=2, rowspan=4, padx=10, pady=5, sticky="nsew")
        self.listbox.bind('<<ListboxSelect>>', self.on_listbox_select)
        self.listbox.bind('<Double-Button-1>', self.on_listbox_double_click)

        # 添加位置按钮
        add_button = tk.Button(self.root, text="添加位置", command=self.add_position)
        add_button.grid(row=5, column=2, padx=10, pady=5, sticky="nsew")

        # 删除位置按钮
        delete_button = tk.Button(self.root, text="删除位置", command=self.delete_position)
        delete_button.grid(row=6, column=2, padx=10, pady=5, sticky="nsew")

        # 编辑位置按钮
        edit_button = tk.Button(self.root, text="编辑位置", command=self.edit_position)
        edit_button.grid(row=7, column=2, padx=10, pady=5, sticky="nsew")

        # 编辑句柄按钮
        edit_handle_button = tk.Button(self.root, text="编辑句柄", command=self.edit_handle)
        edit_handle_button.grid(row=8, column=2, padx=10, pady=5, sticky="nsew")

        # 快捷键输入
        tk.Label(self.root, text="快捷键:").grid(row=9, column=0, padx=10, pady=5, sticky="w")
        self.hotkey_entry = tk.Entry(self.root, textvariable=tk.StringVar(self.root, value=self.hotkey))
        self.hotkey_entry.grid(row=9, column=1, padx=10, pady=5, sticky="ew")
        self.hotkey_entry.bind('<Return>', self.update_hotkey)

        # 设置快捷键按钮
        set_hotkey_button = tk.Button(self.root, text="设置快捷键", command=self.update_hotkey)
        set_hotkey_button.grid(row=9, column=2, padx=10, pady=5, sticky="ew")

        # 状态标签
        self.status_label = tk.Label(self.root, text="按 F3 启动/停止点击", fg="blue")
        self.status_label.grid(row=10, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

    def select_position(self):
        messagebox.showinfo("提示", "请移动鼠标到目标位置并按下回车")
        keyboard.wait('enter')
        self.click_position = win32api.GetCursorPos()
        self.position_label.config(text=f"{self.click_position}")

        self.window_handle = win32gui.WindowFromPoint(self.click_position)
        self.window_title = win32gui.GetWindowText(self.window_handle)
        self.handle_label.config(text=f"{self.window_handle}")
        self.title_label.config(text=self.window_title)

    def start_clicking(self):
        while self.is_clicking:
            if self.window_handle:
                self.send_click(self.window_handle, self.click_position)
            time.sleep(self.click_interval.get() / 1000.0)

    def send_click(self, hwnd, position):
        window_rect = win32gui.GetWindowRect(hwnd)
        x = position[0] - window_rect[0]
        y = position[1] - window_rect[1]
        
        win32api.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, (y << 16) | x)
        time.sleep(0.01)
        win32api.PostMessage(hwnd, win32con.WM_LBUTTONUP, 0, (y << 16) | x)

    def toggle_clicking(self):
        if self.is_clicking:
            self.is_clicking = False
            self.status_label.config(text="点击已停止", fg="red")
            self.toggle_button.config(text="启动点击")
        else:
            self.is_clicking = True
            self.status_label.config(text="点击进行中...", fg="green")
            self.toggle_button.config(text="停止点击")
            threading.Thread(target=self.start_clicking).start()

    def add_position(self):
        position = f"位置: {self.click_position}, 句柄: {self.window_handle}, 标题: {self.window_title}"
        self.listbox.insert(tk.END, position)
        self.positions.append((self.window_handle, self.click_position, self.window_title))

    def delete_position(self):
        selected = self.listbox.curselection()
        for index in reversed(selected):
            self.listbox.delete(index)
            self.positions.pop(index)

    def edit_position(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个位置")
            return
        index = selected[0]
        handle, position, title = self.positions[index]
        new_position = simpledialog.askstring("编辑位置", "输入新坐标 (x,y):")
        if new_position:
            new_position = tuple(map(int, new_position.split(',')))
            self.positions[index] = (handle, new_position, title)
            self.listbox.delete(index)
            self.listbox.insert(index, f"位置: {new_position}, 句柄: {handle}, 标题: {title}")

    def edit_handle(self):
        selected = self.listbox.curselection()
        if not selected:
            messagebox.showwarning("警告", "请先选择一个位置")
            return
        index = selected[0]
        handle, position, title = self.positions[index]
        new_handle = simpledialog.askinteger("编辑句柄", "输入新的窗口句柄:")
        if new_handle:
            self.positions[index] = (new_handle, position, title)
            self.listbox.delete(index)
            self.listbox.insert(index, f"位置: {position}, 句柄: {new_handle}, 标题: {title}")

    def on_listbox_select(self, event):
        selected = self.listbox.curselection()
        if selected:
            index = selected[0]
            handle, position, title = self.positions[index]
            self.click_position = position
            self.window_handle = handle
            self.window_title = title
            self.position_label.config(text=f"{self.click_position}")
            self.handle_label.config(text=f"{self.window_handle}")
            self.title_label.config(text=self.window_title)

    def on_listbox_double_click(self, event):
        self.toggle_clicking()

    def update_hotkey(self, event=None):
        new_hotkey = self.hotkey_entry.get()
        keyboard.remove_hotkey(self.hotkey)
        self.hotkey = new_hotkey
        keyboard.add_hotkey(self.hotkey, self.toggle_clicking)
        messagebox.showinfo("快捷键设置", f"新的快捷键: {self.hotkey}")

    def save_data(self):
        data = {
            "click_interval": self.click_interval.get(),
            "positions": self.positions
        }
        with open(self.save_file, "w") as f:
            json.dump(data, f)

    def load_data(self):
        if os.path.exists(self.save_file):
            with open(self.save_file, "r") as f:
                data = json.load(f)
                self.click_interval.set(data.get("click_interval", 50))
                self.positions = data.get("positions", [])
                for handle, position, title in self.positions:
                    self.listbox.insert(tk.END, f"位置: {position}, 句柄: {handle}, 标题: {title}")

    def on_close(self):
        self.save_data()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AutoClickerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
