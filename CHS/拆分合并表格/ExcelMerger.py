import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

class ExcelMerger:
    def __init__(self, root):
        self.root = root
        self.root.title("工作表合并工具")
        self.root.geometry("700x600")  # 进一步增加界面高度
        
        # 数据存储变量
        self.file_paths = []
        self.selected_sheets = {}  # 存储每个文件选中的工作表 {file_path: [sheet1, sheet2...]}
        self.output_path = ""
        
        # 创建UI
        self.create_widgets()
        
    def create_widgets(self):
        # 顶部文件选择区域
        frame_files = ttk.LabelFrame(self.root, text="选择Excel文件")
        frame_files.pack(fill="x", padx=10, pady=5)
        
        btn_add = ttk.Button(frame_files, text="添加文件", command=self.add_files)
        btn_add.pack(side="left", padx=5, pady=5)
        
        btn_remove = ttk.Button(frame_files, text="移除选中", command=self.remove_file)
        btn_remove.pack(side="left", padx=5, pady=5)
        
        btn_clear = ttk.Button(frame_files, text="清空列表", command=self.clear_files)
        btn_clear.pack(side="left", padx=5, pady=5)
        
        # 中间文件和工作表列表区域
        frame_lists = ttk.Frame(self.root)
        frame_lists.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 左侧文件列表
        ttk.Label(frame_lists, text="文件列表").pack(anchor="w")
        self.file_listbox = tk.Listbox(frame_lists, selectmode=tk.SINGLE, height=8)
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=(0,5))
        self.file_listbox.bind('<<ListboxSelect>>', self.on_file_select)
        
        # 右侧工作表列表（带复选框）
        ttk.Label(frame_lists, text="可选择的工作表（勾选要合并的）").pack(anchor="w")
        self.sheets_frame = ttk.Frame(frame_lists)
        self.sheets_frame.pack(side="right", fill="both", expand=True)
        
        # 为工作表列表添加滚动条
        self.sheets_scrollbar = ttk.Scrollbar(self.sheets_frame)
        self.sheets_scrollbar.pack(side="right", fill="y")
        
        self.sheets_canvas = tk.Canvas(self.sheets_frame, yscrollcommand=self.sheets_scrollbar.set)
        self.sheets_canvas.pack(side="left", fill="both", expand=True)
        self.sheets_scrollbar.config(command=self.sheets_canvas.yview)
        
        self.sheets_inner_frame = ttk.Frame(self.sheets_canvas)
        self.sheets_canvas_window = self.sheets_canvas.create_window((0, 0), window=self.sheets_inner_frame, anchor="nw")
        
        # 绑定滚动事件
        self.sheets_inner_frame.bind("<Configure>", self.on_sheets_configure)
        self.sheets_canvas.bind("<Configure>", self.on_canvas_configure)
        
        # 选项区域
        frame_options = ttk.LabelFrame(self.root, text="合并选项")
        frame_options.pack(fill="x", padx=10, pady=5)
        
        # 表头行号选择
        ttk.Label(frame_options, text="表头行号:").pack(side="left", padx=5)
        self.header_row = tk.StringVar(value="1")
        ttk.Entry(frame_options, textvariable=self.header_row, width=5).pack(side="left", padx=5)
        
        # 是否包含文件名
        self.include_filename = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_options, text="包含来源文件名", variable=self.include_filename).pack(side="left", padx=10)
        
        # 是否包含工作表名
        self.include_sheetname = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_options, text="包含来源工作表名", variable=self.include_sheetname).pack(side="left", padx=10)
        
        # 输出区域
        frame_output = ttk.LabelFrame(self.root, text="输出设置")
        frame_output.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(frame_output, text="输出文件:").pack(side="left", padx=5)
        self.output_entry = ttk.Entry(frame_output)
        self.output_entry.pack(side="left", fill="x", expand=True, padx=5)
        
        btn_browse = ttk.Button(frame_output, text="浏览", command=self.browse_output)
        btn_browse.pack(side="left", padx=5)
        
        # 合并按钮区域
        frame_buttons = ttk.Frame(self.root)
        frame_buttons.pack(pady=15)
        
        # 带标题行合并按钮
        btn_merge_with_header = ttk.Button(
            frame_buttons, 
            text="按表头行号合并", 
            command=lambda: self.merge_sheets(has_header=True), 
            style="Accent.TButton"
        )
        btn_merge_with_header.pack(side="left", padx=10)
        
        # 无标题行合并按钮
        btn_merge_without_header = ttk.Button(
            frame_buttons, 
            text="直接拼接", 
            command=lambda: self.merge_sheets(has_header=False), 
            style="Accent.TButton"
        )
        btn_merge_without_header.pack(side="left", padx=10)
        
        # 样式设置
        self.root.style = ttk.Style()
        self.root.style.configure("Accent.TButton", font=("Arial", 10, "bold"))
        
    def on_sheets_configure(self, event):
        """更新画布滚动区域"""
        self.sheets_canvas.configure(scrollregion=self.sheets_canvas.bbox("all"))
        
    def on_canvas_configure(self, event):
        """调整内部框架宽度以匹配画布"""
        self.sheets_canvas.itemconfig(self.sheets_canvas_window, width=event.width)
        
    def add_files(self):
        files = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls;*.xlsm")]
        )
        for file in files:
            if file not in self.file_paths:
                self.file_paths.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
                
                # 获取该文件的所有工作表并默认全部选中
                try:
                    excel_file = pd.ExcelFile(file)
                    all_sheets = excel_file.sheet_names
                    self.selected_sheets[file] = all_sheets.copy()
                except:
                    self.selected_sheets[file] = []
                
    def remove_file(self):
        selected = self.file_listbox.curselection()
        if not selected:
            return
        # 获取选中的索引
        idx = selected[0]
        # 删除对应的文件路径和选择状态
        file_path = self.file_paths.pop(idx)
        del self.selected_sheets[file_path]
        # 从列表中删除
        self.file_listbox.delete(idx)
        # 清空工作表选择区域
        self.clear_sheets_frame()
                
    def clear_files(self):
        self.file_paths = []
        self.selected_sheets = {}
        self.file_listbox.delete(0, tk.END)
        self.clear_sheets_frame()
        
    def clear_sheets_frame(self):
        """清空工作表选择区域"""
        for widget in self.sheets_inner_frame.winfo_children():
            widget.destroy()
                
    def on_file_select(self, event):
        """当选择文件时，显示该文件的所有工作表供选择（默认全选）"""
        selected = self.file_listbox.curselection()
        if not selected:
            return
            
        # 清空之前的工作表选择
        self.clear_sheets_frame()
        
        # 获取选中的文件路径
        idx = selected[0]
        file_path = self.file_paths[idx]
        
        try:
            # 获取该文件的所有工作表
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            
            # 为每个工作表创建复选框（默认选中）
            self.sheet_vars = {}
            for sheet_name in sheet_names:
                var = tk.BooleanVar()
                # 默认选中所有工作表
                var.set(True)
                # 创建复选框
                cb = ttk.Checkbutton(
                    self.sheets_inner_frame,
                    text=sheet_name,
                    variable=var,
                    command=lambda s=sheet_name, v=var: self.update_selected_sheets(file_path, s, v)
                )
                cb.pack(anchor="w", padx=2, pady=1)
                self.sheet_vars[sheet_name] = var
                
        except Exception as e:
            ttk.Label(self.sheets_inner_frame, text=f"无法读取工作表: {str(e)}", foreground="red").pack()
                
    def update_selected_sheets(self, file_path, sheet_name, var):
        """更新选中的工作表列表"""
        if var.get():
            if sheet_name not in self.selected_sheets[file_path]:
                self.selected_sheets[file_path].append(sheet_name)
        else:
            if sheet_name in self.selected_sheets[file_path]:
                self.selected_sheets[file_path].remove(sheet_name)
                
    def browse_output(self):
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="保存合并结果"
        )
        if output_file:
            self.output_path = output_file
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, output_file)
    
    def merge_sheets(self, has_header):
        # 检查是否有选中的文件
        if not self.file_paths:
            messagebox.showerror("错误", "请先添加要合并的Excel文件")
            return
            
        # 检查是否有选中的工作表
        has_selected = False
        for sheets in self.selected_sheets.values():
            if sheets:
                has_selected = True
                break
        if not has_selected:
            messagebox.showerror("错误", "请至少选择一个工作表进行合并")
            return
            
        # 检查输出路径
        if not self.output_path:
            messagebox.showerror("错误", "请选择输出文件路径")
            return
        
        try:
            # 处理表头行号（仅当有标题行时需要）
            header_row = 0
            if has_header:
                header_row = int(self.header_row.get()) - 1  # 转换为0-based索引
                if header_row < 0:
                    raise ValueError
            
        except ValueError:
            messagebox.showerror("错误", "表头行号必须是正整数")
            return
        
        try:
            # 创建进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("合并中")
            progress_window.geometry("300x100")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            ttk.Label(progress_window, text="正在合并工作表...").pack(pady=10)
            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=250, mode="determinate")
            progress_bar.pack(pady=10)
            
            self.root.update()
            
            all_data = []
            total_sheets = sum(len(sheets) for sheets in self.selected_sheets.values())
            processed = 0
            
            for file_path in self.file_paths:
                # 获取该文件选中的工作表
                sheets_to_process = self.selected_sheets.get(file_path, [])
                if not sheets_to_process:
                    continue
                    
                try:
                    # 读取Excel文件
                    excel_file = pd.ExcelFile(file_path)
                    
                    for sheet_name in sheets_to_process:
                        # 更新进度
                        processed += 1
                        progress_bar["value"] = (processed / total_sheets) * 100
                        progress_window.update()
                        
                        # 读取工作表数据
                        # 无标题行模式下，header=None表示不将任何行作为表头
                        df = pd.read_excel(
                            excel_file,
                            sheet_name=sheet_name,
                            header=header_row if has_header else None,
                            engine="openpyxl"
                        )
                        
                        # 添加来源信息列
                        if self.include_filename.get():
                            df.insert(0, "来源文件", os.path.basename(file_path))
                        if self.include_sheetname.get():
                            df.insert(0, "来源工作表", sheet_name)
                            
                        all_data.append(df)
                        
                except Exception as e:
                    messagebox.showwarning("警告", f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                    continue
            
            # 合并所有数据
            if not all_data:
                messagebox.showinfo("提示", "没有可合并的数据")
                progress_window.destroy()
                return
                
            merged_df = pd.concat(all_data, ignore_index=True)
            
            # 保存合并结果
            with pd.ExcelWriter(self.output_path, engine="openpyxl") as writer:
                merged_df.to_excel(writer, sheet_name="合并结果", index=False)
                
                # 美化表格
                workbook = writer.book
                worksheet = writer.sheets["合并结果"]
                
                # 设置列宽自适应
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(cell.value)
                        except:
                            pass
                    adjusted_width = (max_length + 2) * 1.2
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # 设置表头样式（仅当有标题行时）
                if has_header:
                    header_font = Font(bold=True)
                    thin_border = Border(
                        left=Side(style='thin'), 
                        right=Side(style='thin'),
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
                    
                    for cell in worksheet[1]:  # 表头行
                        cell.font = header_font
                        cell.border = thin_border
            
            progress_bar["value"] = 100
            progress_window.destroy()
            messagebox.showinfo("成功", f"合并完成！结果已保存至:\n{self.output_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"合并过程中发生错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMerger(root)
    root.mainloop()
