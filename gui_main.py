import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import re
import subprocess
import threading
import time
from pathlib import Path

# 确保 docxlatex 安装在当前激活的环境中
try:
    from docxlatex import Document
except ImportError:
    messagebox.showerror("环境错误", "❌ 找不到 'docxlatex' 模块。\n请确保在终端运行了: pip install docxlatex")
    sys.exit(1)

# 全局变量：使用字典来追踪文件信息 { treeview_item_id: Path_object }
file_registry = {}
last_output_dir = None
is_converting = False  # 转换状态锁


# ==========================================
# 核心提取逻辑 (单文件处理)
# ==========================================
def extract_single_file(docx_path: Path, output_format: str) -> bool:
    if not docx_path.exists():
        return False
    output_path = docx_path.parent / f"{docx_path.stem}_extracted.txt"
    try:
        doc = Document(str(docx_path))
        full_text = doc.get_text()
    except Exception:
        return False

    try:
        with output_path.open('w', encoding='utf-8') as f:
            if output_format == 'full':
                f.write(full_text)
            elif output_format == 'split':
                block_pattern = re.compile(r"\$\$(.*?)\$\$", re.DOTALL)
                inline_pattern = re.compile(r"(?<!\$)\$(?!\$)(.*?)(?<!\$)\$(?!\$)", re.DOTALL)
                
                blocks = block_pattern.findall(full_text)
                text_no_blocks = block_pattern.sub("", full_text)
                inlines = inline_pattern.findall(text_no_blocks)
                
                equations = []
                for b in blocks:
                    equations.append({'type': 'display/block', 'content': b.strip()})
                for i in inlines:
                    equations.append({'type': 'inline', 'content': i.strip()})
                
                f.write(f"Source Document: {docx_path.name}\n")
                f.write(f"Total Equations Found: {len(equations)}\n")
                f.write("=" * 40 + "\n\n")
                
                for idx, eq in enumerate(equations, start=1):
                    f.write(f"--- EQUATION {idx} ({eq['type']}) ---\n")
                    f.write(f"{eq['content']}\n\n")
        return True
    except IOError:
        return False


# ==========================================
# 多线程精准追踪批量转换
# ==========================================
def batch_conversion_worker(output_format):
    global last_output_dir, is_converting
    is_converting = True
    
    items = list(file_registry.keys())
    total_files = len(items)
    success_count = 0
    
    progress_bar['maximum'] = total_files
    progress_bar['value'] = 0
    
    for idx, item_id in enumerate(items, start=1):
        docx_path = file_registry[item_id]
        
        # 1. 改变当前文件的状态为 "⏳ 正在转换..." 并高亮显示
        file_tree.item(item_id, values=(docx_path.name, "⏳ 正在转换..."))
        file_tree.see(item_id)  # 自动滚动到当前处理的行
        
        # 执行转换
        success = extract_single_file(docx_path, output_format)
        
        # 2. 根据结果更新该行的状态
        if success:
            file_tree.item(item_id, values=(docx_path.name, "✅ 转换成功"))
            success_count += 1
        else:
            file_tree.item(item_id, values=(docx_path.name, "❌ 转换失败"))
            
        last_output_dir = docx_path.parent
        progress_bar['value'] = idx
        time.sleep(0.1)  # 留出平滑观察时间

    # 收尾
    progress_bar['value'] = total_files
    btn_open_folder.config(state=tk.NORMAL)
    btn_select.config(state=tk.NORMAL)
    btn_delete.config(state=tk.NORMAL)
    btn_convert.config(state=tk.NORMAL)
    is_converting = False
    
    messagebox.showinfo("完成", f"🎉 批量转换结束！\n\n成功：{success_count} 个\n失败：{total_files - success_count} 个")


# ==========================================
# UI 交互触发事件
# ==========================================
def select_files():
    if is_converting: return
    file_paths = filedialog.askopenfilenames(
        title="选择 Word 文档 (.docx)", 
        filetypes=[("Word 文件", "*.docx")]
    )
    if file_paths:
        global last_output_dir
        for p in file_paths:
            path_obj = Path(p)
            # 检查是否已经添加过，避免重复添加
            if path_obj not in file_registry.values():
                # 插入到表格组件中，初始状态为 "⏳ 等待中"
                item_id = file_tree.insert("", tk.END, values=(path_obj.name, "⏳ 等待中"))
                # 将表格的行 ID 和文件路径绑定
                file_registry[item_id] = path_obj
        
        last_output_dir = Path(file_paths[0]).parent
        btn_open_folder.config(state=tk.NORMAL)
        update_window_title()

def delete_selected_files():
    """取消/删除选中的文件"""
    if is_converting:
        messagebox.showwarning("警告", "正在转换中，无法删除队列文件！")
        return
        
    selected_items = file_tree.selection()  # 获取鼠标选中的行（支持多选删除）
    if not selected_items:
        messagebox.showinfo("提示", "请先在下方列表中点击选中要去除的文件。")
        return
        
    for item_id in selected_items:
        file_tree.delete(item_id)      # 从界面删除
        file_registry.pop(item_id, None)  # 从数据核心删除
        
    update_window_title()
    if not file_registry:
        progress_bar['value'] = 0

def update_window_title():
    # 动态显示队列里的文件数量
    root.title(f"Word LaTeX 公式批量提取器 (当前队列: {len(file_registry)}个文件)")

def start_batch_conversion():
    if not file_registry:
        messagebox.showwarning("提示", "队列中没有文件，请先添加文件！")
        return
    
    fmt_display = format_var.get()
    output_format = 'full' if "完整文本" in fmt_display else 'split'
    
    # 锁定所有控制按钮
    btn_select.config(state=tk.DISABLED)
    btn_delete.config(state=tk.DISABLED)
    btn_convert.config(state=tk.DISABLED)
    
    # 开启线程做具体工作
    threading.Thread(
        target=batch_conversion_worker, 
        args=(output_format,), 
        daemon=True
    ).start()

def open_in_finder():
    if last_output_dir and last_output_dir.exists():
        subprocess.run(["open", str(last_output_dir)])


# ==========================================
# UI 界面布局
# ==========================================
root = tk.Tk()
root.title("Word LaTeX 公式批量提取器")
root.geometry("600x500") 

# 1. 顶部控制栏
frame_top = tk.Frame(root, pady=10)
frame_top.pack(fill=tk.X)

btn_select = tk.Button(frame_top, text=" 📂 添加文件 ", command=select_files, padx=8, pady=4)
btn_select.pack(side=tk.LEFT, padx=15)

btn_delete = tk.Button(frame_top, text=" ❌ 移除所选文件 ", command=delete_selected_files, fg="#FF3B30", padx=8, pady=4)
btn_delete.pack(side=tk.LEFT, padx=5)

# 2. 核心：文件列表表格 (Treeview)
frame_list = tk.Frame(root, padx=15)
frame_list.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

# 滚动条
scrollbar = ttk.Scrollbar(frame_list)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# 定义两列：文件名、转换状态
file_tree = ttk.Treeview(frame_list, columns=("name", "status"), show="headings", yscrollcommand=scrollbar.set)
file_tree.heading("name", text="文件名称")
file_tree.heading("status", text="转换进度/状态")
file_tree.column("name", width=380, anchor=tk.W)
file_tree.column("status", width=160, anchor=tk.CENTER)
file_tree.pack(fill=tk.BOTH, expand=True)

scrollbar.config(command=file_tree.yview)

# 3. 底部进度条与配置
frame_bottom = tk.Frame(root, pady=10)
frame_bottom.pack(fill=tk.X)

progress_bar = ttk.Progressbar(frame_bottom, orient="horizontal", mode="determinate")
progress_bar.pack(fill=tk.X, padx=15, pady=5)

# 格式选择与操作按钮并排
frame_ops = tk.Frame(frame_bottom, pady=5)
frame_ops.pack(fill=tk.X, padx=15)

format_var = tk.StringVar()
format_dropdown = ttk.Combobox(frame_ops, textvariable=format_var, state="readonly", width=22)
format_dropdown['values'] = ("full (包含上下文完整文本)", "split (仅隔离公式列表)")
format_dropdown.current(0)
format_dropdown.pack(side=tk.LEFT, pady=5)

btn_open_folder = tk.Button(frame_ops, text=" 🔍 打开输出文件夹 ", command=open_in_finder, state=tk.DISABLED, padx=5, pady=3)
btn_open_folder.pack(side=tk.RIGHT, padx=5)

btn_convert = tk.Button(frame_ops, text=" 🚀 开始批量提取 ", command=start_batch_conversion, bg="#007AFF", fg="white", font=("Arial", 11, "bold"), padx=10, pady=3)
btn_convert.pack(side=tk.RIGHT, padx=5)

root.mainloop()