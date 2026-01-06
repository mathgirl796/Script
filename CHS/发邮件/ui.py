import FreeSimpleGUI as sg
from ConfigManager import ConfigManager
from EmailSender import EmailSender
import pandas as pd
import os


class EmailSenderGUI:
    def __init__(self):
        self.config_manager = ConfigManager()
        self.window = None
        
    def create_layout(self):
        """创建GUI布局"""
        layout = [
            # 第一行：发件人配置
            [
                sg.Text('发件人邮箱:', size=(12, 1)),
                sg.InputText(key='sender_email', size=(25, 1)),
                sg.Text('授权码:', size=(6, 1)),
                sg.InputText(key='sender_password', size=(25, 1)),
                sg.Text('SMTP服务器:', size=(10, 1)),
                sg.InputText(key='smtp_server', size=(15, 1)),
                sg.Text('端口:', size=(4, 1)),
                sg.InputText(key='smtp_port', size=(5, 1))
            ],
            
            # 第二行：Excel文件路径和保存按钮
            [
                sg.Text('邮箱对应表:', size=(12, 1)),
                sg.InputText(key='email_excel_path', size=(40, 1)),
                sg.FileBrowse('浏览', file_types=(('Excel文件', '*.xlsx'),))
            ],
            
            # 第三行：发件人名称和邮件主题
            [
                sg.Text('发件人名称:', size=(12, 1)),
                sg.InputText(key='sender_name', size=(25, 1)),
                sg.Text('邮件主题:', size=(8, 1)),
                sg.InputText(key='email_subject', size=(35, 1))
            ],
            
            # 第四行：邮件正文
            [
                sg.Text('邮件正文:', size=(12, 1)),
                sg.Multiline(key='email_body', size=(70, 4), default_text='')
            ],
            
            # 第五行：功能按钮
            [
                sg.Button('发送测试邮件', button_color=('white', 'blue'), size=(12, 1)),
                sg.Button('保存配置', button_color=('white', 'green'), size=(10, 1)),
                sg.Button('退出', button_color=('white', 'black'), size=(8, 1))
            ],
            
            # 分隔线
            [sg.HSeparator()],
            
            # 批量发送区域标题
            [sg.Text('批量发送邮件', font=('Arial', 12, 'bold'))],
            
            # 批量发送控制按钮
            [
                sg.Button('添加文件', button_color=('white', 'blue'), size=(10, 1), key='-ADD_FILES-'),
                sg.Button('刷新匹配', button_color=('white', 'green'), size=(10, 1), key='-REFRESH_MATCH-'),
                sg.Button('清空文件', button_color=('white', 'red'), size=(10, 1), key='-CLEAR_FILES-'),
                sg.Button('开始发送', button_color=('white', 'orange'), size=(10, 1), key='-START_SEND-'),
            ],
            
            # 文件列表表格
            [
                sg.Table(
                    values=[],
                    headings=['文件名', '定点名称', '邮箱', '发送状态'],
                    col_widths=[25, 20, 30, 15],
                    auto_size_columns=False,
                    display_row_numbers=True,
                    justification='left',
                    num_rows=15,
                    key='-FILE_TABLE-',
                    enable_events=True,
                    vertical_scroll_only=False
                )
            ]
        ]
        
        return layout
    
    def load_config_to_window(self):
        """加载配置到窗口"""
        config = self.config_manager.get_config()
        for key, value in config.items():
            if key in self.window.AllKeysDict:
                self.window[key].update(value)
    
    def save_config_from_window(self):
        """从窗口保存配置"""
        config = {}
        for key in ['sender_email', 'sender_password', 'smtp_server', 'smtp_port', 
                   'email_excel_path', 'sender_name', 'email_subject', 'email_body']:
            if key in self.window.AllKeysDict:
                value = self.window[key].get()
                if key == 'smtp_port':
                    try:
                        config[key] = int(value)
                    except ValueError:
                        config[key] = 25
                else:
                    config[key] = value
        
        success, message = self.config_manager.update_config(**config)
        return success, message
    
    def send_test_email(self):
        """发送测试邮件"""
        try:
            # 保存当前配置
            success, message = self.save_config_from_window()
            if not success:
                sg.popup_error(f'配置保存失败: {message}')
                return
            
            # 获取邮件发送器配置
            email_config = self.config_manager.get_email_sender_config()
            email_sender = EmailSender(**email_config)
            
            # 发送测试邮件给自己
            receiver_email = self.config_manager.get_config('sender_email')
            receiver_name = '测试发送对象'
            subject = self.config_manager.get_config('email_subject') or '测试邮件'
            body = self.config_manager.get_config('email_body') or '这是一封测试邮件'
            
            success, message = email_sender.send_single_email(
                receiver_email, subject, body, ['测试.xlsx'], 
            )
            
            if success:
                sg.popup(f'测试邮件发送成功!\n{message}', title='成功')
            else:
                sg.popup_error(f'测试邮件发送失败!\n{message}', title='失败')
                
        except Exception as e:
            sg.popup_error(f'发送测试邮件时出错: {str(e)}', title='错误')
    
    def load_email_mapping(self):
        """加载邮件对应表"""
        try:
            excel_path = self.config_manager.get_config('email_excel_path')
            if not excel_path or not os.path.exists(excel_path):
                return None
            
            df = pd.read_excel(excel_path)
            # 假设四列分别是：定点编号、定点名称、邮箱、别名
            if len(df.columns) >= 4:
                df.columns = ['定点编号', '定点名称', '邮箱', '别名'] + list(df.columns[4:])
                return df[['定点编号', '定点名称', '邮箱', '别名']]
            elif len(df.columns) >= 3:
                df.columns = ['定点编号', '定点名称', '邮箱'] + list(df.columns[3:])
                # 如果没有别名列，添加一个空的别名列
                df['别名'] = ''
                return df[['定点编号', '定点名称', '邮箱', '别名']]
            else:
                sg.popup_error('邮件对应表格式错误，需要至少三列数据', title='错误')
                return None
        except Exception as e:
            sg.popup_error(f'读取邮件对应表失败: {str(e)}', title='错误')
            return None
    
    def match_file_to_email(self, filename, email_mapping):
        """根据文件名匹配定点名称和邮箱
        
        匹配优先级：
        1. 首先尝试匹配定点编号
        2. 如果匹配不上，尝试匹配定点名称
        3. 如果还匹配不上，尝试匹配别名（用|分割的字符串）
        """
        if email_mapping is None:
            return '', ''
        
        # 获取文件名（不含扩展名）
        basename = os.path.splitext(os.path.basename(filename))[0]
        
        # 第一级：尝试精确匹配定点编号
        match = email_mapping[email_mapping['定点编号'].astype(str) == basename]
        if not match.empty:
            return match.iloc[0]['定点名称'], match.iloc[0]['邮箱']
        
        # 如果精确匹配失败，尝试模糊匹配定点编号
        for _, row in email_mapping.iterrows():
            if str(row['定点编号']) in basename or basename in str(row['定点编号']):
                return row['定点名称'], row['邮箱']
        
        # 第二级：尝试匹配定点名称
        for _, row in email_mapping.iterrows():
            if not pd.isna(row['定点名称']):
                # 精确匹配
                if str(row['定点名称']) == basename:
                    return row['定点名称'], row['邮箱']
                # 模糊匹配
                if str(row['定点名称']) in basename or basename in str(row['定点名称']):
                    return row['定点名称'], row['邮箱']
        
        # 第三级：尝试匹配别名
        for _, row in email_mapping.iterrows():
            if not pd.isna(row['别名']) and str(row['别名']).strip():
                # 分割别名字符串（用|分割）
                aliases = str(row['别名']).split('|')
                for alias in aliases:
                    alias = alias.strip()
                    if alias:  # 确保别名不为空
                        # 精确匹配
                        if alias == basename:
                            return row['定点名称'], row['邮箱']
                        # 模糊匹配
                        if alias in basename or basename in alias:
                            return row['定点名称'], row['邮箱']
        
        return '未匹配', '未匹配'
    
    def perform_batch_send(self, file_list, parent_window):
        """执行批量发送"""
        try:
            # 获取邮件发送器
            email_config = self.config_manager.get_email_sender_config()
            email_sender = EmailSender(**email_config)
            
            subject = self.config_manager.get_config('email_subject') or '批量发送文件'
            body = self.config_manager.get_config('email_body') or '您好！附件是您的相关文件。'
            
            success_count = 0
            fail_count = 0
            
            # 逐个发送并实时更新表格
            for i, (file_path, filename, name, email, status) in enumerate(file_list):
                # 更新状态为"发送中"
                file_list[i][4] = '发送中...'
                display_list = [[f[1], f[2], f[3], f[4]] for f in file_list]
                parent_window['-FILE_TABLE-'].update(values=display_list)
                parent_window.refresh()  # 强制刷新界面
                
                if email == '未匹配':
                    file_list[i][4] = '跳过-未匹配邮箱'
                    fail_count += 1
                    display_list = [[f[1], f[2], f[3], f[4]] for f in file_list]
                    parent_window['-FILE_TABLE-'].update(values=display_list)
                    parent_window.refresh()
                    continue
                
                try:
                    success, message = email_sender.send_single_email(
                        email, subject, body, [file_path], receiver_name=name
                    )
                    
                    if success:
                        file_list[i][4] = '发送成功'
                        success_count += 1
                    else:
                        file_list[i][4] = f'发送失败'
                        fail_count += 1
                        
                except Exception as e:
                    file_list[i][4] = f'发送异常'
                    fail_count += 1
                
                # 更新表格显示当前发送结果
                display_list = [[f[1], f[2], f[3], f[4]] for f in file_list]
                parent_window['-FILE_TABLE-'].update(values=display_list)
                parent_window.refresh()
            
            # 显示最终统计结果
            sg.popup(
                f'批量发送完成!\n成功: {success_count} 个\n失败: {fail_count} 个\n总计: {len(file_list)} 个文件',
                title='发送完成'
            )
            
        except Exception as e:
            sg.popup_error(f'批量发送过程中出错: {str(e)}', title='错误')
    
    def reset_config(self):
        """重置配置"""
        if sg.popup_yes_no('确定要重置所有配置为默认值吗?', title='确认重置'):
            success, message = self.config_manager.reset_to_default()
            if success:
                self.load_config_to_window()
                sg.popup('配置已重置为默认值', title='成功')
            else:
                sg.popup_error(f'配置重置失败: {message}', title='失败')
    
    def run(self):
        """运行GUI应用"""
        sg.theme('DefaultNoMoreNagging')
        
        layout = self.create_layout()
        self.window = sg.Window('邮件发送系统', layout, resizable=True, finalize=True)
        
        # 加载初始配置
        self.load_config_to_window()
        
        # 初始化批量发送相关变量
        file_list = []
        email_mapping = None
        
        while True:
            event, values = self.window.read()
            
            if event in (sg.WIN_CLOSED, '退出'):
                break
            
            elif event == '保存配置':
                success, message = self.save_config_from_window()
                if success:
                    sg.popup('配置保存成功!', title='成功')
                else:
                    sg.popup_error(f'配置保存失败: {message}', title='失败')
            
            elif event == '加载配置':
                self.load_config_to_window()
                sg.popup('配置加载成功!', title='成功')
            
            elif event == '发送测试邮件':
                self.send_test_email()
            
            elif event == '-ADD_FILES-':
                files = sg.popup_get_file(
                    '选择要发送的文件',
                    multiple_files=True,
                    file_types=(('Excel文件', '*.xlsx'), ('所有文件', '*.*')),
                    no_window=True
                )
                
                if files:
                    if isinstance(files, str):
                        files = files.split(';')
                    
                    # 加载邮件对应表（如果还没有加载）
                    if email_mapping is None:
                        email_mapping = self.load_email_mapping()
                        if email_mapping is None:
                            sg.popup_error('邮件对应表加载失败，请先配置邮箱对应表路径')
                            continue
                    
                    # 添加新文件到列表
                    for file_path in files:
                        if file_path not in [f[0] for f in file_list]:
                            name, email = self.match_file_to_email(file_path, email_mapping)
                            # 存储完整路径，但只显示文件名
                            filename = os.path.basename(file_path)
                            file_list.append([file_path, filename, name, email, '待发送'])

                    # 更新表格（只显示文件名，不显示完整路径）
                    display_list = [[f[1], f[2], f[3], f[4]] for f in file_list]
                    self.window['-FILE_TABLE-'].update(values=display_list)
            
            elif event == '-REFRESH_MATCH-':
                # 重新加载邮件对应表并刷新匹配
                email_mapping = self.load_email_mapping()
                if email_mapping is not None:
                    # 重新匹配所有文件
                    updated_file_list = []
                    for file_path, filename, _, _, status in file_list:
                        name, email = self.match_file_to_email(file_path, email_mapping)
                        updated_file_list.append([file_path, filename, name, email, status])
                    file_list = updated_file_list
                    # 更新表格（只显示文件名）
                    display_list = [[f[1], f[2], f[3], f[4]] for f in file_list]
                    self.window['-FILE_TABLE-'].update(values=display_list)
                    sg.popup('匹配已刷新!', title='成功')
            
            elif event == '-CLEAR_FILES-':
                # 确认清空文件列表
                if file_list:
                    if sg.popup_yes_no('确定要清空所有文件吗?', title='确认清空'):
                        file_list = []
                        self.window['-FILE_TABLE-'].update(values=file_list)
                        sg.popup('文件列表已清空!', title='成功')
                else:
                    sg.popup('文件列表已经是空的!', title='提示')
            
            elif event == '-START_SEND-':
                if not file_list:
                    sg.popup_error('请先添加文件!', title='错误')
                    continue
                
                # 保存当前配置
                success, message = self.save_config_from_window()
                if not success:
                    sg.popup_error(f'配置保存失败: {message}')
                    continue
                
                # 检查是否有未匹配的文件
                unmatched = [f for f in file_list if f[3] == '未匹配']
                if unmatched:
                    if not sg.popup_yes_no(
                        f'有 {len(unmatched)} 个文件未匹配到邮箱，是否继续发送?\n'
                        '未匹配的文件将跳过。',
                        title='确认发送'
                    ):
                        continue
                
                # 开始批量发送
                self.perform_batch_send(file_list, self.window)
            
            elif event == '重置配置':
                self.reset_config()
        
        self.window.close()


if __name__ == '__main__':
    app = EmailSenderGUI()
    app.run()