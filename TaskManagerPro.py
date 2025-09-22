import tkinter as tk
from tkinter import ttk, messagebox
import psutil
import win32gui
import win32process
import win32con
import win32ui
from PIL import Image, ImageTk
import webbrowser, subprocess, os

# Globals
process_cache = {}
icon_cache = {}
item_map = {}
WM_CLOSE = win32con.WM_CLOSE
font_size = 11
search_text = ""
placeholder_text = "Search Program..."

# ---------------- ToolTip ----------------
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show_tip)
        widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 40
        y = self.widget.winfo_rooty() + 30
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tw, text=self.text, background="#333", foreground="white",
            relief="solid", borderwidth=1, font=("Segoe UI", 10)
        )
        label.pack(ipadx=6, ipady=3)

    def hide_tip(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None

# ---------------- Process Handling ----------------
def enum_windows_callback(hwnd, windows):
    try:
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                windows.append((hwnd, pid, title))
    except Exception:
        pass
    return True

def get_open_windows():
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    seen = set()
    result = []
    for hwnd, pid, title in windows:
        if hwnd not in seen:
            seen.add(hwnd)
            result.append((hwnd, pid, title))
    return result

def get_app_icon(exe_path):
    try:
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        hicon = None
        if small and len(small) > 0:
            hicon = small[0]
        elif large and len(large) > 0:
            hicon = large[0]
        if not hicon:
            return None

        ico_x, ico_y = 32, 32
        hdc_screen = win32gui.GetDC(0)
        dc = win32ui.CreateDCFromHandle(hdc_screen)
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(dc, ico_x, ico_y)
        memdc = dc.CreateCompatibleDC()
        oldbmp = memdc.SelectObject(bmp)

        win32gui.DrawIconEx(memdc.GetSafeHdc(), 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)

        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        image = Image.frombuffer("RGBA", (bmpinfo["bmWidth"], bmpinfo["bmHeight"]), bmpstr, "raw", "BGRA", 0, 1)
        image = image.resize((20, 20), Image.LANCZOS)
        photo = ImageTk.PhotoImage(image)

        memdc.SelectObject(oldbmp)
        memdc.DeleteDC()
        dc.DeleteDC()
        win32gui.ReleaseDC(0, hdc_screen)
        try:
            win32gui.DestroyIcon(hicon)
        except Exception:
            pass

        return photo
    except Exception:
        return None

def update_apps():
    selection = tree.selection()
    selected_pids = [item_map[iid][1] for iid in selection if iid in item_map]

    for row in tree.get_children():
        tree.delete(row)
    item_map.clear()

    windows = get_open_windows()
    current_pids = set()

    for hwnd, pid, title in windows:
        try:
            proc = process_cache.get(pid)
            if proc is None:
                proc = psutil.Process(pid)
                process_cache[pid] = proc
                proc.cpu_percent(None)
                cpu = 0.0
            else:
                cpu = proc.cpu_percent(None)

            try:
                mem = proc.memory_info().rss / (1024 * 1024)
            except Exception:
                mem = 0.0

            try:
                name = proc.name()
            except Exception:
                name = title if title else "Unknown"

            if search_text and search_text.lower() not in name.lower():
                continue

            icon = icon_cache.get(pid)
            if icon is None:
                exe = None
                try:
                    exe = proc.exe()
                except Exception:
                    exe = None
                if exe:
                    icon = get_app_icon(exe)
                    if icon:
                        icon_cache[pid] = icon

            if icon:
                iid = tree.insert("", "end", text="", image=icon,
                                  values=(name, f"{cpu:.1f} %", f"{mem:.1f} MB", pid, title))
            else:
                iid = tree.insert("", "end", text="",
                                  values=(name, f"{cpu:.1f} %", f"{mem:.1f} MB", pid, title))

            item_map[iid] = (hwnd, pid)
            current_pids.add(pid)

            if pid in selected_pids:
                tree.selection_set(iid)

        except Exception:
            continue

    for cached_pid in list(process_cache.keys()):
        if cached_pid not in current_pids:
            process_cache.pop(cached_pid, None)

    root.after(1000, update_apps)

# ---------------- Actions ----------------
def close_selected():
    selection = tree.selection()
    if not selection:
        return
    for iid in selection:
        pair = item_map.get(iid)
        if not pair:
            continue
        hwnd, pid = pair
        try:
            if hwnd and win32gui.IsWindow(hwnd):
                win32gui.PostMessage(hwnd, WM_CLOSE, 0, 0)
            root.after(700, lambda h=hwnd, p=pid, i=iid: check_closed(h, p, i))
        except Exception:
            pass

def check_closed(hwnd, pid, iid):
    try:
        if hwnd and win32gui.IsWindow(hwnd):
            return
        p = psutil.Process(pid)
        if p.is_running():
            do_force = messagebox.askyesno("Force terminate", f"Process {pid} did not close. Force terminate?")
            if do_force:
                p.terminate()
    except Exception:
        pass
    update_apps()

def zoom_in():
    global font_size
    font_size += 1
    style.configure("Custom.Treeview", font=("Segoe UI", font_size))

def zoom_out():
    global font_size
    if font_size > 6:
        font_size -= 1
        style.configure("Custom.Treeview", font=("Segoe UI", font_size))

def on_mousewheel_zoom(event):
    if event.delta > 0:
        zoom_in()
    else:
        zoom_out()

def on_double_click(event):
    item = tree.selection()
    if not item:
        return
    iid = item[0]
    hwnd, pid = item_map.get(iid, (None, None))
    if pid:
        try:
            proc = psutil.Process(pid)
            exe = proc.exe()
            if exe and os.path.exists(exe):
                subprocess.Popen(f'explorer /select,"{exe}"')
        except Exception:
            pass

def on_right_click(event):
    iid = tree.identify_row(event.y)
    if iid:
        tree.selection_set(iid)
        context_menu.post(event.x_root, event.y_root)

def open_linkedin():
    webbrowser.open("https://www.linkedin.com/in/mohamed-fouad0")

def open_github():
    webbrowser.open("https://github.com/Mohamed-Fouad0")

# ---------------- Search Entry ----------------
def on_entry_focus_in(event):
    if search_entry.get() == placeholder_text:
        search_entry.delete(0, "end")
        search_entry.config(fg="black")

def on_entry_focus_out(event):
    if not search_entry.get():
        search_entry.insert(0, placeholder_text)
        search_entry.config(fg="gray")

def on_search(event=None):
    global search_text
    text = search_entry.get()
    if text == placeholder_text:
        search_text = ""
    else:
        search_text = text
    update_apps()

# ---------------- UI ----------------
root = tk.Tk()
root.title("Task Manager Pro For Desktop Running Apps")
root.geometry("1100x780")
root.configure(bg="#434343")

style = ttk.Style(root)
style.theme_use("default")
style.configure("Custom.Treeview",
                background="#434343",
                fieldbackground="#434343",
                foreground="white",
                rowheight=28,
                font=("Segoe UI", font_size))
style.configure("Custom.Treeview.Heading",
                background="#333333",
                foreground="white",
                font=("Segoe UI", 11, "bold"),
                borderwidth=1,
                relief="solid")
style.map("Custom.Treeview", background=[("selected", "#555555")])

# ---------------- Search Bar ----------------
search_frame = tk.Frame(root, bg="#434343")
search_frame.pack(fill="x", padx=6, pady=(6, 0))

search_entry = tk.Entry(search_frame, font=("Segoe UI", 12), fg="gray")
search_entry.pack(fill="x", expand=True, padx=5, pady=5)
search_entry.insert(0, placeholder_text)

search_entry.bind("<FocusIn>", on_entry_focus_in)
search_entry.bind("<FocusOut>", on_entry_focus_out)
search_entry.bind("<KeyRelease>", on_search)

# ---------------- TreeView ----------------
columns = ("Program Name", "CPU", "RAM", "PID", "Window Title")
tree = ttk.Treeview(root, columns=columns, show="tree headings", style="Custom.Treeview")
tree.heading("#0", text="Icon")
tree.column("#0", width=40, anchor="center")
tree.heading("Program Name", text="Program Name")
tree.column("Program Name", width=200)
tree.heading("CPU", text="CPU")
tree.column("CPU", width=80, anchor="center")
tree.heading("RAM", text="RAM")
tree.column("RAM", width=100, anchor="center")
tree.heading("PID", text="PID")
tree.column("PID", width=80, anchor="center")
tree.heading("Window Title", text="Window Title")
tree.column("Window Title", width=450)
tree.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

tree.bind("<Double-1>", on_double_click)
tree.bind("<Button-3>", on_right_click)

# ---------------- Context Menu ----------------
context_menu = tk.Menu(root, tearoff=0, bg="#333", fg="white", activebackground="#555")
context_menu.add_command(label="End Task", command=close_selected)
context_menu.add_command(label="Open File Location", command=lambda: on_double_click(None))

# ---------------- Buttons Frame ----------------
btn_frame = tk.Frame(root, bg="#333333", height=60)
btn_frame.pack(fill="x", padx=6, pady=6)

btn_frame.columnconfigure(0, weight=1)
btn_frame.columnconfigure(1, weight=2)
btn_frame.columnconfigure(2, weight=1)

def add_hover_effect(btn, color, hover_color):
    btn.config(bg=color, activebackground=hover_color, cursor="hand2")

btn_zoom_out = tk.Button(btn_frame, text="Zoom Out", command=zoom_out,
                         fg="white", font=("Segoe UI", 12, "bold"))
btn_zoom_out.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
add_hover_effect(btn_zoom_out, "#ff9800", "#e68900")
ToolTip(btn_zoom_out, "Zoom Out")

btn_end = tk.Button(btn_frame, text="End Task", command=close_selected,
                    fg="white", font=("Segoe UI", 12, "bold"))
btn_end.grid(row=0, column=1, sticky="nsew", padx=4, pady=4)
add_hover_effect(btn_end, "#e53935", "#c62828")
ToolTip(btn_end, "Click Delete")

btn_zoom_in = tk.Button(btn_frame, text="Zoom In", command=zoom_in,
                        fg="white", font=("Segoe UI", 12, "bold"))
btn_zoom_in.grid(row=0, column=2, sticky="nsew", padx=4, pady=4)
add_hover_effect(btn_zoom_in, "#4caf50", "#388e3c")
ToolTip(btn_zoom_in, "Zoom In")

# ---------------- Footer ----------------
footer_frame = tk.Frame(root, bg="#333333", height=60)
footer_frame.pack(fill="x", pady=(0, 10))

footer_frame.columnconfigure(0, weight=2)
footer_frame.columnconfigure(1, weight=4)
footer_frame.columnconfigure(2, weight=4)

lbl_dev = tk.Label(footer_frame, text="Developer:", fg="white", bg="#333333",
                   font=("Segoe UI", 16, "bold"))
lbl_dev.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

btn_linkedin = tk.Button(footer_frame, text="LinkedIn", command=open_linkedin,
                         fg="white", font=("Segoe UI", 14, "bold"))
btn_linkedin.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
add_hover_effect(btn_linkedin, "#0077b5", "#005f8e")
ToolTip(btn_linkedin, "Go To Developer's LinkedIn")

btn_github = tk.Button(footer_frame, text="GitHub", command=open_github,
                       fg="white", font=("Segoe UI", 14, "bold"))
btn_github.grid(row=0, column=2, sticky="nsew", padx=10, pady=10)
add_hover_effect(btn_github, "#24292e", "#000000")
ToolTip(btn_github, "Go To Developer's Github")

# ---------------- Bindings ----------------
root.bind("<Control-plus>", lambda e: zoom_in())
root.bind("<Control-minus>", lambda e: zoom_out())
root.bind("<Control-q>", lambda e: close_selected())
root.bind("<Delete>", lambda e: close_selected())
root.bind("<Control-MouseWheel>", on_mousewheel_zoom)

# Run
update_apps()
root.mainloop()
