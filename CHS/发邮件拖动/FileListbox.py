from tkinter import *
from tkinterdnd2 import *
import os

class FileListbox(Frame):
    """支持拖拽文件和右键删除的文件列表框组件（包含标签）"""
    
    def __init__(self, master=None, label_text="拖拽文件到此处", width=20, **kwargs):
        super().__init__(master, **kwargs)
        
        # 文件路径数组
        self.file_paths = []
        
        # 创建标签
        self.label = Label(self, text=label_text, fg='blue', font=('Arial', 9))
        self.label.pack(anchor='w', pady=(0, 5))
        
        # 创建列表框
        self.listbox = Listbox(self, selectmode=SINGLE, width=width)
        self.listbox.pack(fill=BOTH, expand=True)
        
        # 绑定拖拽事件
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self._on_drop)
        self.listbox.dnd_bind('<<DragEnter>>', self._on_drag_enter)
        self.listbox.dnd_bind('<<DragLeave>>', self._on_drag_leave)
        
        # 绑定右键菜单
        self.listbox.bind('<Button-3>', self._show_context_menu)
        
        # 创建右键菜单
        self._create_context_menu()
    
    def _on_drop(self, event):
        """处理文件拖入事件"""
        files = self.tk.splitlist(event.data)
        
        for file_path in files:
            # 去除文件路径的前后引号（如果有）
            file_path = file_path.strip('"\'')
            
            if os.path.exists(file_path):
                filename = os.path.basename(file_path)
                if file_path not in self.file_paths:
                    # 添加到文件路径数组
                    self.file_paths.append(file_path)
                    
                    # 在列表中显示文件名
                    self.listbox.insert(END, filename)
                else:
                    print(f"文件已存在：{filename}")
            else:
                print(f"文件不存在：{file_path}")
    
    def _on_drag_enter(self, event):
        """鼠标拖入时的视觉反馈"""
        self.listbox.config(bg='lightgray')
        return 'copy'  # 表示接受拖放
    
    def _on_drag_leave(self, event):
        """鼠标拖出时恢复原样"""
        self.listbox.config(bg='white')
    
    def _create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = Menu(self, tearoff=0)
        self.context_menu.add_command(label="删除", command=self._delete_selected)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="清空全部", command=self._clear_all)
    
    def _show_context_menu(self, event):
        """显示右键菜单"""
        # 确保点击的是有效项
        index = self.listbox.nearest(event.y)
        if index >= 0:
            self.listbox.selection_clear(0, END)
            self.listbox.selection_set(index)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _delete_selected(self):
        """删除选中的文件"""
        selected_indices = self.listbox.curselection()
        
        # 从后往前删除，避免索引变化
        for index in reversed(selected_indices):
            # 从列表中删除
            self.listbox.delete(index)
            # 从文件路径数组中删除
            del self.file_paths[index]
    
    def _clear_all(self):
        """清空所有文件"""
        self.listbox.delete(0, END)
        self.file_paths.clear()
    
    def get_file_paths(self):
        """获取所有文件路径"""
        return self.file_paths.copy()
    
    def add_file(self, file_path):
        """手动添加文件"""
        if os.path.exists(file_path) and file_path not in self.file_paths:
            self.file_paths.append(file_path)
            filename = os.path.basename(file_path)
            self.listbox.insert(END, filename)
            return True
        return False
    
    def remove_file(self, index):
        """删除指定索引的文件"""
        if 0 <= index < len(self.file_paths):
            self.listbox.delete(index)
            del self.file_paths[index]
            return True
        return False


# 测试代码
if __name__ == '__main__':
    # root = TkinterDnD.Tk()
    # root.title("文件拖拽列表测试")
    # root.geometry("400x300")
    
    # # 创建文件列表，自定义标签文本
    # file_list = FileListbox(root, label_text="请拖拽要发送的附件文件到此处")
    # file_list.pack(fill=BOTH, expand=True, padx=10, pady=10)
    
    # # 添加测试按钮
    # def show_files():
    #     print("当前文件路径：")
    #     for i, path in enumerate(file_list.get_file_paths()):
    #         print(f"{i+1}. {path}")
    
    # button = Button(root, text="显示文件路径", command=show_files)
    # button.pack(pady=5)
    
    # root.mainloop()
    pass