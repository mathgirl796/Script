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
        self.root.geometry("700x550")  # 进一步增加界面高度
        
        # 数据存储变量
        self.file_paths = []
        self.selected_sheets = {}  # 存储每个文件选中的工作表 {file_path: [sheet1, sheet2...]}
        
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
        self.btn_select_sheets_title = ttk.Button(
            frame_lists, 
            text="点此全选/全不选", 
            command=self.toggle_select_all_sheets
        )
        self.btn_select_sheets_title.pack(anchor="w", pady=(0, 2))
        
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
        
        # 合并按钮区域
        frame_buttons = ttk.Frame(self.root)
        frame_buttons.pack(pady=5)
        
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
        
        # 按某列拆分按钮
        btn_split_by_column = ttk.Button(
            frame_buttons, 
            text="按某列拆分", 
            command=self.split_by_column, 
            style="Accent.TButton"
        )
        btn_split_by_column.pack(side="left", padx=10)
        
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
    
    def toggle_select_all_sheets(self):
        """切换全选/全不选当前选中文件的所有工作表"""
        selected = self.file_listbox.curselection()
        if not selected:
            return
            
        # 获取选中的文件路径
        idx = selected[0]
        file_path = self.file_paths[idx]
        
        # 检查当前是否全选
        if not hasattr(self, 'sheet_vars'):
            return
            
        current_selected_count = sum(1 for var in self.sheet_vars.values() if var.get())
        total_count = len(self.sheet_vars)
        
        # 如果当前已全选，则全不选；否则全选
        new_state = current_selected_count < total_count
        
        # 更新所有复选框状态
        for sheet_name, var in self.sheet_vars.items():
            var.set(new_state)
            # 更新选中工作表列表
            if new_state:
                if sheet_name not in self.selected_sheets[file_path]:
                    self.selected_sheets[file_path].append(sheet_name)
            else:
                if sheet_name in self.selected_sheets[file_path]:
                    self.selected_sheets[file_path].remove(sheet_name)
    
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
            
        # 弹出文件选择对话框
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="保存合并结果"
        )
        if not output_file:
            return  # 用户取消了选择
        self.output_path = output_file
        
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
    
    def split_by_column(self):
        """按某列拆分功能"""
        # 检查是否有选中的文件
        if not self.file_paths:
            messagebox.showerror("错误", "请先添加要拆分的Excel文件")
            return
            
        # 检查是否有选中的工作表
        has_selected = False
        for sheets in self.selected_sheets.values():
            if sheets:
                has_selected = True
                break
                
        if not has_selected:
            messagebox.showerror("错误", "请先选择要拆分的工作表")
            return
        
        # 获取第一个选中的文件和工作表
        first_file = None
        first_sheet = None
        for file_path in self.file_paths:
            sheets = self.selected_sheets.get(file_path, [])
            if sheets:
                first_file = file_path
                first_sheet = sheets[0]
                break
        
        if not first_file or not first_sheet:
            messagebox.showerror("错误", "无法找到可用的文件和工作表")
            return
        
        try:
            # 获取表头行号
            header_row = int(self.header_row.get()) - 1
            if header_row < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("错误", "表头行号必须是正整数")
            return
        
        try:
            # 读取第一个工作表获取列标题
            df = pd.read_excel(first_file, sheet_name=first_sheet, header=header_row, engine="openpyxl")
            column_names = df.columns.tolist()
            
            if not column_names:
                messagebox.showerror("错误", "无法读取列标题")
                return
            
            # 创建选择列的对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("选择拆分列")
            dialog.geometry("400x300")
            dialog.transient(self.root)
            dialog.grab_set()
            
            ttk.Label(dialog, text="请选择要按哪一列拆分:").pack(pady=10)
            
            # 列选择下拉框
            selected_column = tk.StringVar()
            column_combo = ttk.Combobox(dialog, textvariable=selected_column, values=column_names, state="readonly")
            column_combo.pack(pady=5, padx=20, fill="x")
            column_combo.current(0)  # 默认选择第一列
            
            # 输出目录选择
            ttk.Label(dialog, text="输出目录:").pack(pady=(20, 5))
            
            output_dir_var = tk.StringVar()
            output_frame = ttk.Frame(dialog)
            output_frame.pack(pady=5, padx=20, fill="x")
            
            output_entry = ttk.Entry(output_frame, textvariable=output_dir_var)
            output_entry.pack(side="left", fill="x", expand=True)
            
            def browse_output_dir():
                directory = filedialog.askdirectory(title="选择输出目录")
                if directory:
                    output_dir_var.set(directory)
            
            ttk.Button(output_frame, text="浏览", command=browse_output_dir).pack(side="right", padx=(5, 0))
            
            # 确认按钮
            def confirm_split():
                if not selected_column.get():
                    messagebox.showerror("错误", "请选择拆分列")
                    return
                if not output_dir_var.get():
                    messagebox.showerror("错误", "请选择输出目录")
                    return
                
                dialog.destroy()
                self.perform_split(selected_column.get(), output_dir_var.get(), header_row)
            
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=20)
            
            ttk.Button(button_frame, text="确定", command=confirm_split).pack(side="left", padx=5)
            ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side="left", padx=5)
            
        except Exception as e:
            messagebox.showerror("错误", f"读取列标题时出错:\n{str(e)}")
    
    def perform_split(self, split_column, output_dir, header_row):
        """执行拆分操作"""
        try:
            # 先合并所有选中的工作表
            all_data = []
            
            # 创建进度窗口
            progress_window = tk.Toplevel(self.root)
            progress_window.title("处理中")
            progress_window.geometry("350x120")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # 创建状态标签
            status_label = ttk.Label(progress_window, text="正在读取工作表...")
            status_label.pack(pady=5)
            
            # 创建详细信息标签
            detail_label = ttk.Label(progress_window, text="")
            detail_label.pack(pady=2)
            
            progress_bar = ttk.Progressbar(progress_window, orient="horizontal", length=300, mode="determinate")
            progress_bar.pack(pady=10)
            
            self.root.update()
            
            total_sheets = sum(len(sheets) for sheets in self.selected_sheets.values())
            processed = 0
            
            for file_path in self.file_paths:
                sheets_to_process = self.selected_sheets.get(file_path, [])
                if not sheets_to_process:
                    continue
                    
                try:
                    excel_file = pd.ExcelFile(file_path)
                    
                    for sheet_name in sheets_to_process:
                        processed += 1
                        progress_value = (processed / total_sheets) * 50  # 合并阶段占50%
                        progress_bar["value"] = progress_value
                        status_label.config(text=f"正在读取工作表... ({processed}/{total_sheets})")
                        detail_label.config(text=f"文件: {os.path.basename(file_path)} - 工作表: {sheet_name}")
                        progress_window.update()
                        
                        df = pd.read_excel(
                            excel_file,
                            sheet_name=sheet_name,
                            header=header_row,
                            engine="openpyxl"
                        )
                        
                        if self.include_filename.get():
                            df.insert(0, "来源文件", os.path.basename(file_path))
                        if self.include_sheetname.get():
                            df.insert(0, "来源工作表", sheet_name)
                            
                        all_data.append(df)
                        
                except Exception as e:
                    messagebox.showwarning("警告", f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
                    continue
            
            if not all_data:
                messagebox.showinfo("提示", "没有可处理的数据")
                progress_window.destroy()
                return
                
            merged_df = pd.concat(all_data, ignore_index=True)
            
            # 检查拆分列是否存在
            if split_column not in merged_df.columns:
                messagebox.showerror("错误", f"拆分列 '{split_column}' 不存在于合并的数据中")
                progress_window.destroy()
                return
            
            # 获取拆分列的所有唯一值
            unique_values = merged_df[split_column].dropna().unique()
            
            if len(unique_values) == 0:
                messagebox.showerror("错误", f"拆分列 '{split_column}' 中没有有效数据")
                progress_window.destroy()
                return
            
            # 按值拆分并保存
            progress_bar["value"] = 50
            progress_window.update()
            status_label.config(text="正在拆分并保存文件...")
            
            for i, value in enumerate(unique_values):
                # 筛选该值对应的数据
                filtered_df = merged_df[merged_df[split_column] == value]
                
                # 更新进度和状态
                progress_value = 50 + (i + 1) / len(unique_values) * 50
                progress_bar["value"] = progress_value
                status_label.config(text=f"正在拆分并保存文件... ({i+1}/{len(unique_values)})")
                detail_label.config(text=f"当前拆分值: {value} ({len(filtered_df)} 行数据)")
                progress_window.update()
                
                # 生成文件名（处理特殊字符）
                safe_filename = str(value).replace('/', '_').replace('\\', '_').replace(':', '_').replace('*', '_').replace('?', '_').replace('"', '_').replace('<', '_').replace('>', '_').replace('|', '_')
                if not safe_filename.strip():
                    safe_filename = f"未知值_{i}"
                
                output_path = os.path.join(output_dir, f"{safe_filename}.xlsx")
                
                # 保存到Excel文件
                with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                    filtered_df.to_excel(writer, sheet_name="数据", index=False)
                    
                    # 美化表格
                    workbook = writer.book
                    worksheet = writer.sheets["数据"]
                    
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
                    
                    # 设置表头样式
                    header_font = Font(bold=True)
                    thin_border = Border(
                        left=Side(style='thin'), 
                        right=Side(style='thin'),
                        top=Side(style='thin'), 
                        bottom=Side(style='thin')
                    )
                    
                    for cell in worksheet[1]:
                        cell.font = header_font
                        cell.border = thin_border
                
                progress_bar["value"] = 50 + (i + 1) / len(unique_values) * 50
                progress_window.update()
            
            progress_bar["value"] = 100
            status_label.config(text="拆分完成！")
            detail_label.config(text=f"共生成 {len(unique_values)} 个文件")
            progress_window.update()
            
            # 短暂延迟让用户看到完成状态
            self.root.after(1000, progress_window.destroy())
            
            messagebox.showinfo("成功", f"拆分完成！\n共生成 {len(unique_values)} 个文件\n保存在: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("错误", f"拆分过程中发生错误:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMerger(root)
    root.mainloop()