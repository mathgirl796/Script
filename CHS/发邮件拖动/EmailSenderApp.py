from tkinter import *
from tkinterdnd2 import *
from tkinter import messagebox, filedialog
import json
import os
import pandas as pd
from FileListbox import FileListbox
from ConfigManager import ConfigManager
from EmailSender import EmailSender
import threading

class EmailSenderApp:
    """邮件发送主应用"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("批量邮件发送系统")
        self.root.geometry("450x800")
        
        # 初始化配置和邮箱对应表
        try:
            self.config_manager = ConfigManager()
        except ValueError as e:
            messagebox.showerror("配置错误", str(e))
            # 提供创建配置文件的选项
            if messagebox.askyesno("创建配置", "是否创建一个空的配置文件？"):
                config_manager = ConfigManager.__new__(ConfigManager)
                success, message = config_manager.create_default_config()
                if success:
                    messagebox.showinfo("成功", message)
                    messagebox.showinfo("提示", "请重新启动应用程序并填写配置")
                else:
                    messagebox.showerror("错误", message)
            self.root.destroy()
            return
        
        self.email_sender = None
        self.institutions_data = []  # 邮箱对应表数据
        self.file_lists = []  # FileListbox实例列表
        self.is_sending = False  # 防止重复发送的标志
        
        # 创建UI
        self._create_config_area()
        self._create_scrollable_frame()
        self._create_status_bar()
        
        # 初始化数据
        self._load_email_config()
        self._load_institution_data()
        
    def _create_config_area(self):
        """创建简化的配置区域"""
        self.config_frame = Frame(self.root, relief=RAISED, bd=1)
        self.config_frame.pack(fill="x", padx=5, pady=5)
        
        # 配置输入框（隐藏在设置窗口中）
        self.config_entries = {}
        
        # 简化的按钮区域
        button_frame = Frame(self.config_frame)
        button_frame.pack(fill="x", padx=5, pady=5)
        
        Button(button_frame, text="⚙️ 设置", command=self._open_settings_window, 
               bg="lightblue", width=10).pack(side="left", padx=5)
        Button(button_frame, text="📧 发送邮件", command=self._send_emails_thread, 
               bg="salmon", fg="white", width=10).pack(side="left", padx=5)
        Button(button_frame, text="🗑️ 清空全部", command=self._clear_all_files, 
               bg="orange", fg="white", width=10).pack(side="left", padx=5)
        
        # 当前Excel文件路径显示
        self.excel_label = Label(self.config_frame, text="当前邮箱对应表: 未选择", fg="blue")
        self.excel_label.pack(fill="x", padx=5, pady=2)
    
    def _open_settings_window(self):
        """打开设置窗口"""
        settings_window = Toplevel(self.root)
        settings_window.title("系统设置")
        settings_window.geometry("500x450")
        settings_window.resizable(False, False)
        
        # 确保窗口显示在最前面
        settings_window.lift()
        settings_window.focus_force()
        
        # 配置输入框
        self.config_entries = {}
        
        # 发件人配置
        sender_frame = LabelFrame(settings_window, text="发件人配置", padx=10, pady=10)
        sender_frame.pack(fill="x", padx=10, pady=5)
        
        Label(sender_frame, text="发件人邮箱:").grid(row=0, column=0, sticky="w", pady=2)
        self.config_entries['sender_email'] = Entry(sender_frame, width=30)
        self.config_entries['sender_email'].grid(row=0, column=1, pady=2)
        
        Label(sender_frame, text="授权码:").grid(row=1, column=0, sticky="w", pady=2)
        self.config_entries['sender_password'] = Entry(sender_frame, width=30)
        self.config_entries['sender_password'].grid(row=1, column=1, pady=2)
        
        Label(sender_frame, text="SMTP服务器:").grid(row=2, column=0, sticky="w", pady=2)
        self.config_entries['smtp_server'] = Entry(sender_frame, width=30)
        self.config_entries['smtp_server'].grid(row=2, column=1, pady=2)
        
        Label(sender_frame, text="端口:").grid(row=3, column=0, sticky="w", pady=2)
        self.config_entries['smtp_port'] = Entry(sender_frame, width=30)
        self.config_entries['smtp_port'].grid(row=3, column=1, pady=2)
        
        Label(sender_frame, text="发件人名称:").grid(row=4, column=0, sticky="w", pady=2)
        self.config_entries['sender_name'] = Entry(sender_frame, width=30)
        self.config_entries['sender_name'].grid(row=4, column=1, pady=2)
        
        # 邮件配置
        email_frame = LabelFrame(settings_window, text="邮件配置", padx=10, pady=10)
        email_frame.pack(fill="x", padx=10, pady=5)
        
        Label(email_frame, text="邮件主题:").grid(row=0, column=0, sticky="w", pady=2)
        self.config_entries['email_subject'] = Entry(email_frame, width=30)
        self.config_entries['email_subject'].grid(row=0, column=1, pady=2)
        
        Label(email_frame, text="邮件正文:").grid(row=1, column=0, sticky="w", pady=2)
        self.config_entries['email_body'] = Entry(email_frame, width=30)
        self.config_entries['email_body'].grid(row=1, column=1, pady=2)
        
        # 文件配置
        file_frame = LabelFrame(settings_window, text="文件配置", padx=10, pady=10)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        Button(file_frame, text="选择邮箱对应表", command=self._select_excel_file, 
               bg="lightblue").pack(pady=5)
        
        # 按钮区域
        button_frame = Frame(settings_window)
        button_frame.pack(fill="x", padx=10, pady=15)
        
        Button(button_frame, text="保存配置", command=lambda: self._save_config_and_close(settings_window), 
               bg="lightgreen", width=12, height=2).pack(side="left", padx=10)
        Button(button_frame, text="取消", command=settings_window.destroy, 
               width=12, height=2).pack(side="left", padx=10)
        
        # 加载现有配置到输入框
        config = self.config_manager.get_config()
        for key in ['sender_email', 'sender_password', 'smtp_server', 
                   'smtp_port', 'sender_name', 'email_subject', 'email_body']:
            if key in self.config_entries:
                self.config_entries[key].delete(0, END)
                self.config_entries[key].insert(0, str(config.get(key, '')))
        
    def _create_scrollable_frame(self):
        """创建可滚动的文件列表区域"""
        scroll_frame = Frame(self.root)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Canvas和Scrollbar
        self.canvas = Canvas(scroll_frame)
        scrollbar = Scrollbar(scroll_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _on_mousewheel(self, event):
        """鼠标滚轮滚动"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def _create_file_lists(self):
        """根据邮箱对应表创建FileListbox"""
        # 清空现有的FileListbox
        for file_list in self.file_lists:
            file_list.destroy()
        self.file_lists.clear()
        
        # 每行2个FileListbox，适应450像素宽度
        items_per_row = 2
        col_width = 220  # 每列宽度，适应450像素窗口
        
        for i, institution in enumerate(self.institutions_data):
            # 计算行列位置
            row = i // items_per_row
            col = i % items_per_row
            
            # 创建FileListbox实例，仅显示第二列的定点名称
            file_list = FileListbox(
                self.scrollable_frame, 
                label_text=f"{institution['name']}",
                width=6  # 适应较小窗口的字符宽度
            )
            
            # 使用place进行精确定位，高度减少到原来的三分之一
            file_list.place(
                x=col * col_width,
                y=row * 83,  # 250的三分之一约等于83
                width=col_width - 4,  # 减去边距
                height=80  # 240的三分之一约等于80
            )
            
            # 保存机构信息到file_list
            file_list.institution_info = institution
            
            self.file_lists.append(file_list)
        
        # 设置scrollable_frame的总尺寸以包含所有FileListbox
        total_rows = (len(self.institutions_data) + items_per_row - 1) // items_per_row
        self.scrollable_frame.config(
            width=items_per_row * col_width,
            height=total_rows * 83
        )
            
    def _create_status_bar(self):
        """创建状态栏"""
        self.status_frame = Frame(self.root, relief=SUNKEN, bd=1)
        self.status_frame.pack(fill="x", side="bottom")
        
        self.status_label = Label(self.status_frame, text="就绪", anchor="w")
        self.status_label.pack(fill="x", padx=5, pady=2)
        
    def _load_email_config(self):
        """加载邮件配置到输入框"""
        config = self.config_manager.get_config()
        for key in ['sender_email', 'sender_password', 'smtp_server', 
                   'smtp_port', 'sender_name', 'email_subject', 'email_body']:
            if key in self.config_entries and hasattr(self.config_entries[key], 'delete'):
                self.config_entries[key].delete(0, END)
                self.config_entries[key].insert(0, str(config.get(key, '')))
        
        # 更新Excel路径显示
        excel_path = config.get('email_excel_path', '')
        if excel_path:
            self.excel_label.config(text=f"当前邮箱对应表: {excel_path}")
            
    def _save_config(self):
        """保存配置"""
        try:
            config_updates = {}
            for key, entry in self.config_entries.items():
                value = entry.get().strip()
                if key == 'smtp_port':
                    config_updates[key] = int(value) if value else 25
                else:
                    config_updates[key] = value
            
            success, message = self.config_manager.update_config(**config_updates)
            if success:
                self._update_status("配置保存成功")
                self._last_save_success = True  # 设置保存成功标志
                self._load_email_config()  # 重新加载配置
                return True
            else:
                self._last_save_success = False
                messagebox.showerror("错误", f"配置保存失败: {message}")
                return False
        except Exception as e:
            self._last_save_success = False
            messagebox.showerror("错误", f"配置保存出错: {str(e)}")
            return False
    
    def _save_config_and_close(self, window):
        """保存配置并关闭设置窗口"""
        if self._save_config():
            window.destroy()
            
    def _select_excel_file(self):
        """选择邮箱对应表文件"""
        file_path = filedialog.askopenfilename(
            title="选择邮箱对应表",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            # 保存到配置
            self.config_manager.update_config(email_excel_path=file_path)
            # 重新加载数据
            self._load_institution_data()
            
    def _load_institution_data(self):
        """加载邮箱对应表数据"""
        excel_path = self.config_manager.get_config('email_excel_path')
        
        if not excel_path or not os.path.exists(excel_path):
            self._update_status("邮箱对应表文件不存在，请重新选择")
            return
            
        try:
            # 读取Excel文件
            df = pd.read_excel(excel_path)
            
            # 清空现有数据
            self.institutions_data.clear()
            
            # 假设Excel格式：第一列机构名称，第二列邮箱，第三列收件人邮箱
            for index, row in df.iterrows():
                if len(row) >= 3:  # 确保至少有3列
                    self.institutions_data.append({
                        'name': str(row.iloc[4]), 
                        'code': str(row.iloc[0]),   # 第一列：机构编号
                        'email': str(row.iloc[2])   # 第三列：邮箱地址
                    })
            
            # 创建FileListbox
            self._create_file_lists()
            
            # 更新Excel路径显示
            self.excel_label.config(text=f"当前邮箱对应表: {excel_path}")
            self._update_status(f"成功加载 {len(self.institutions_data)} 个机构信息")
            
        except Exception as e:
            messagebox.showerror("错误", f"读取邮箱对应表失败: {str(e)}")
            self._update_status("邮箱对应表读取失败")
            
    def _send_emails_thread(self):
        """在线程中发送邮件"""
        if self.is_sending:
            self._update_status("正在发送邮件中，请勿重复点击...")
            return
            
        self.is_sending = True
        threading.Thread(target=self._send_emails, daemon=True).start()
        
    def _send_emails(self):
        """发送邮件"""
        try:
            # 获取邮件配置
            email_config = self.config_manager.get_email_sender_config()
            self.email_sender = EmailSender(**email_config)
            
            # 获取邮件主题和正文（从配置文件读取）
            config = self.config_manager.get_config()
            subject = str(config.get('email_subject', '')).strip()
            body = str(config.get('email_body', '')).strip()
            
            if not subject:
                subject = "支付明细"
            if not body:
                body = "请查收"
                
            total_sent = 0
            total_failed = 0
            
            for file_list in self.file_lists:
                institution = file_list.institution_info
                file_paths = file_list.get_file_paths()
                
                if not file_paths:
                    continue  # 跳过没有文件的
                    
                # 更新状态
                self._update_status(f"正在发送给 {institution['name']} ({institution['email']})")
                
                # 发送邮件
                success, message = self.email_sender.send_single_email(
                    receiver_email=institution['email'],
                    subject=subject,
                    body=body,
                    attachments=file_paths,
                    receiver_name=institution['name']
                )
                
                if success:
                    total_sent += 1
                    self._update_status(f"✓ {institution['name']} 邮件发送成功")
                else:
                    total_failed += 1
                    self._update_status(f"✗ {institution['name']} 邮件发送失败: {message}")
                    
            # 完成总结
            self._update_status(f"邮件发送完成 - 成功: {total_sent}, 失败: {total_failed}")
            messagebox.showinfo("发送完成", 
                             f"邮件发送完成\n成功: {total_sent}\n失败: {total_failed}")
                             
        except Exception as e:
            self._update_status(f"邮件发送出错: {str(e)}")
            messagebox.showerror("错误", f"邮件发送失败: {str(e)}")
        finally:
            self.is_sending = False  # 重置发送标志
            
    def _clear_all_files(self):
        """清空所有文件列表"""
        if not self.file_lists:
            self._update_status("没有文件可清空")
            return
            
        # 确认对话框
        if not messagebox.askyesno("确认清空", 
                                 f"确定要清空所有 {len(self.file_lists)} 个机构的文件吗？\n此操作不可撤销。"):
            return
            
        try:
            total_cleared = 0
            for file_list in self.file_lists:
                if file_list.get_file_paths():  # 如果有文件
                    file_list.file_paths.clear()
                    file_list.listbox.delete(0, END)
                    total_cleared += 1
                    
            self._update_status(f"已清空 {total_cleared} 个机构的文件")
            messagebox.showinfo("清空完成", f"成功清空 {total_cleared} 个机构的文件")
            
        except Exception as e:
            self._update_status(f"清空文件失败: {str(e)}")
            messagebox.showerror("错误", f"清空文件失败: {str(e)}")
    
    def _update_status(self, message):
        """更新状态栏"""
        self.status_label.config(text=message)
        self.root.update_idletasks()


def main():
    root = TkinterDnD.Tk()
    app = EmailSenderApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()