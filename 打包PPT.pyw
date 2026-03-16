import sys
import os
import win32com.client
import pythoncom
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import threading

# 全局变量，存放用户右键选中的多张图片/视频路径
selected_files_list = []

def import_media_to_ppt(target_data, output_filename="output.pptx"):
    pythoncom.CoInitialize()
    media_files = []
    img_exts = ['.png', '.jpg', '.jpeg', '.bmp']
    vid_exts = ['.mp4', '.avi', '.mov', '.wmv']
    
    # 如果传来的是一个文件夹路径，扫描它
    if isinstance(target_data, str) and os.path.isdir(target_data):
        folder_path = target_data
        print(f">>> 开始扫描文件夹: {folder_path}\n")
        output_dir = folder_path
        for file in os.listdir(folder_path):
            ext = os.path.splitext(file)[1].lower()
            if ext in img_exts or ext in vid_exts:
                if "diff_result" not in file:
                    media_files.append(os.path.join(folder_path, file))
    # 如果传来的是一组文件列表，直接用它们
    elif isinstance(target_data, list):
        print(f">>> 接收到您右键选中的 {len(target_data)} 个指定文件...\n")
        output_dir = os.path.dirname(target_data[0]) if target_data else os.path.expanduser("~")
        for file_path in target_data:
            ext = os.path.splitext(file_path)[1].lower()
            if ext in img_exts or ext in vid_exts:
                media_files.append(file_path)
            
    if not media_files:
        print("❌ 错误：没有找到任何有效的图片或视频文件！")
        pythoncom.CoUninitialize()
        return
        
    print(f"✅ 共确认 {len(media_files)} 个媒体文件，正在启动 PowerPoint 进行排版...")

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Add(WithWindow=False) 
        
        first_file = media_files[0]
        first_ext = os.path.splitext(first_file)[1].lower()
        temp_slide = presentation.Slides.Add(1, 12)

        if first_ext in img_exts:
            shp = temp_slide.Shapes.AddPicture(FileName=first_file, LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=-1, Height=-1)
        else:
            shp = temp_slide.Shapes.AddMediaObject2(FileName=first_file, LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=-1, Height=-1)

        orig_w = shp.Width
        orig_h = shp.Height
        temp_slide.Delete() 

        ratio = orig_w / orig_h if orig_h != 0 else 16/9
        presentation.PageSetup.SlideHeight = 540
        presentation.PageSetup.SlideWidth = 540 * ratio
        
        slide_w = presentation.PageSetup.SlideWidth
        slide_h = presentation.PageSetup.SlideHeight
        print(f"    [画布自适应] 已探测首个文件比例，PPT 画布尺寸已锁定！")
        
        for index, file_path in enumerate(media_files):
            ext = os.path.splitext(file_path)[1].lower()
            slide = presentation.Slides.Add(index + 1, 12)
            
            print(f" -> 正在等比最大化导入: {os.path.basename(file_path)}")
            
            if ext in img_exts:
                shape = slide.Shapes.AddPicture(FileName=file_path, LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=-1, Height=-1)
            else:
                shape = slide.Shapes.AddMediaObject2(FileName=file_path, LinkToFile=False, SaveWithDocument=True, Left=0, Top=0, Width=-1, Height=-1)
                shape.AnimationSettings.PlaySettings.PlayOnEntry = True
                shape.AnimationSettings.PlaySettings.LoopUntilStopped = True
                
            shape.LockAspectRatio = -1
            ratio_w = slide_w / shape.Width
            ratio_h = slide_h / shape.Height
            min_ratio = min(ratio_w, ratio_h)
            
            shape.Width = shape.Width * min_ratio
            shape.Left = (slide_w - shape.Width) / 2
            shape.Top = (slide_h - shape.Height) / 2
                
        out_abs_path = os.path.abspath(os.path.join(output_dir, output_filename))
        presentation.SaveAs(out_abs_path)
        presentation.Close()
        powerpoint.Quit()
        
        print(f"\n🎉 成功！已按原始比例完美排版 {len(media_files)} 个文件。")
        print(f"📁 PPT 已自动保存在同目录下: {out_abs_path}")
        
    except Exception as e:
        print(f"\n❌ 导入失败: {e}")
        try: powerpoint.Quit()
        except: pass
    finally:
        pythoncom.CoUninitialize()

class RedirectText(object):
    def __init__(self, text_ctrl): self.output = text_ctrl
    def write(self, string):
        self.output.insert(tk.END, string)
        self.output.see(tk.END)
    def flush(self): pass

def start_gui():
    global selected_files_list
    root = tk.Tk()
    root.title("🎬 媒体等比例排版神器 (支持多选文件版)")
    root.geometry("650x450")
    root.configure(bg="#f0f0f0")

    frame_top = tk.Frame(root, bg="#f0f0f0")
    frame_top.pack(fill="x", padx=15, pady=15)

    tk.Label(frame_top, text="要处理的目标 (文件夹路径 或 已选中的多文件):", bg="#f0f0f0", font=("微软雅黑", 10)).pack(anchor="w", pady=(0, 5))
    path_entry = tk.Entry(frame_top, font=("微软雅黑", 10))
    path_entry.pack(side="left", fill="x", expand=True, ipady=4)

    # 处理传入的参数
    if len(sys.argv) > 1:
        if os.path.isdir(sys.argv[1]):
            # 传进来的是文件夹
            path_entry.insert(0, sys.argv[1])
        else:
            # 传进来的是一堆选中的文件 (通过发送到菜单)
            selected_files_list = sys.argv[1:]
            path_entry.insert(0, f"✅ 已成功获取您右键选中的 {len(selected_files_list)} 个文件！")
            path_entry.config(state="readonly", fg="green")
    else:
        path_entry.insert(0, os.path.join(os.path.expanduser("~"), "Desktop", "PPT"))

    def browse_folder():
        global selected_files_list
        folder = filedialog.askdirectory()
        if folder:
            path_entry.config(state="normal", fg="black")
            path_entry.delete(0, tk.END)
            path_entry.insert(0, folder)
            selected_files_list = [] # 清空特定文件列表，改为扫描这个新文件夹

    tk.Button(frame_top, text="选文件夹", font=("微软雅黑", 9), command=browse_folder).pack(side="right", padx=(10, 0), ipadx=5)

    def run_task():
        if selected_files_list:
            target_data = selected_files_list
        else:
            target_data = path_entry.get()
            if not os.path.exists(target_data):
                messagebox.showerror("错误", "路径不存在！")
                return
                
        btn_start.config(state="disabled", text="正在拼命合成排版中...")
        log_area.delete(1.0, tk.END)
        
        def target_func():
            import_media_to_ppt(target_data)
            btn_start.config(state="normal", text="🚀 立即一键生成 PPT")
            
        threading.Thread(target=target_func, daemon=True).start()

    btn_start = tk.Button(root, text="🚀 立即一键生成 PPT", bg="#98FB98", fg="black", font=("微软雅黑", 12, "bold"), command=run_task)
    btn_start.pack(pady=5, ipadx=20, ipady=5)

    tk.Label(root, text="执行日志:", bg="#f0f0f0", font=("微软雅黑", 10)).pack(anchor="w", padx=15)
    log_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, font=("Consolas", 9), bg="#1e1e1e", fg="#00ff00")
    log_area.pack(padx=15, pady=(0, 15), fill="both", expand=True)

    sys.stdout = RedirectText(log_area)
    sys.stderr = RedirectText(log_area)

    root.mainloop()

if __name__ == "__main__":
    # 如果检测到是选中文件后通过“发送到”传进来的
    if len(sys.argv) > 1:
        # 开启静默模式：隐藏主窗口
        root = tk.Tk()
        root.withdraw()
        # 自动判断是文件夹还是多个碎文件
        target_data = sys.argv[1] if os.path.isdir(sys.argv[1]) else sys.argv[1:]
        # 直接开始排版
        import_media_to_ppt(target_data)
        # 干完活后弹窗汇报
        messagebox.showinfo("✅ 打包完成", "媒体文件已按原比例完美铺满并打包为 PPT！\n\n请在素材同目录下查看 output.pptx 文件。")
        sys.exit()
    else:
        # 正常双击打开显示界面
        start_gui()
