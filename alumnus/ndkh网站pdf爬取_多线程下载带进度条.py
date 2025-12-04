from requests import get
from bs4 import BeautifulSoup
import os
import re
from PIL import Image
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
from rich.console import Console
from rich.progress import Progress, TaskID, BarColumn, TextColumn, TimeRemainingColumn, SpinnerColumn
from rich.table import Table
from rich.panel import Panel
from rich.layout import Layout
import time
from datetime import datetime
import sys
import platform

# Windows系统兼容性设置
def get_terminal_width():
    """获取终端宽度"""
    try:
        import shutil
        return shutil.get_terminal_size().columns
    except:
        return 120  # 默认宽度

if platform.system() == "Windows":
    # 设置Windows终端模式
    import ctypes
    kernel32 = ctypes.windll.kernel32
    # 启用虚拟终端处理
    kernel32.SetConsoleMode(kernel32.GetStdHandle(-11), 7)

def clean_filename(filename):
    """清理Windows不支持的文件名字符"""
    return re.sub(r'[\\/:*?"<>|]', '_', filename)

def get_survey_links(cookie):
    """获取所有测评的名称和实际链接"""
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
        "Cookie": cookie
    }
    
    try:
        res = get("https://ndkh.hit.edu.cn/user/eva", headers=head)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'lxml')
        
        surveys = []
        start_buttons = soup.find_all(string=re.compile(r'开始测评'))
        
        for button_text in start_buttons:
            parent = button_text.parent
            onclick = parent.get('onclick', '')
            survey_match = re.search(r"openSurvey\('([^']+)'\)", onclick)
            
            if survey_match:
                survey_id = survey_match.group(1)
                actual_url = f"https://ndkh.hit.edu.cn/user/eva_survey?annualSurveyId={survey_id}"
                
                survey_name = "未知测评"
                row = parent.find_parent('tr')
                if row:
                    name_cell = row.find('td')
                    if name_cell:
                        survey_name = name_cell.get_text(strip=True)
                
                surveys.append({
                    "name": survey_name,
                    "url": actual_url,
                    "id": survey_id
                })
        
        return surveys
        
    except Exception as e:
        return []

def download_single_file(args):
    """下载单个文件"""
    name, a, head, output_dir, survey_name = args
    
    try:
        id = str(a).split("'")[1]
        isPlan = 1 if "计划" in str(a) else 0
        url = f"https://ndkh.hit.edu.cn/reportWork_show?id={id}&isPlan={isPlan}"
        res = get(url, headers=head)
        pic_links = ["https://ndkh.hit.edu.cn" + img['data-src'] for img in BeautifulSoup(res.text, 'lxml').find_all('img')]
        pdf_filename = f"{clean_filename(name)}（{clean_filename(a.text.strip())}）.pdf"
        
        images = []
        for image_id, link in enumerate(pic_links):
            img_response = get(link, headers=head)
            img = Image.open(io.BytesIO(img_response.content))
            if img.mode != 'RGB':
                img = img.convert('RGB')
            images.append(img)
        
        if images:
            images[0].save(os.path.join(output_dir, pdf_filename), save_all=True, append_images=images[1:])
        
        return {'success': True, 'filename': pdf_filename}
    except Exception as e:
        return {'success': False, 'error': str(e)}

def get_survey_file_count(cookie, url):
    """获取测评的文件数量"""
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
        "Cookie": cookie
    }
    
    try:
        res = get(url, headers=head)
        soup = BeautifulSoup(res.text, 'lxml')
        items = soup.find_all('td', class_="name")
        
        file_count = 0
        for item in items:
            file_count += len(item.find_all('a'))
        
        return file_count
    except Exception:
        return 0

def fetch_one_survey_with_progress(console, progress, task_id, cookie, url, output_dir, survey_name):
    """使用rich进度条下载单个测评"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
        "Cookie": cookie
    }
    
    res = get(url, headers=head)
    soup = BeautifulSoup(res.text, 'lxml')
    items = soup.find_all('td', class_="name")
    
    download_tasks = []
    for item in items:
        name = item['title'].strip()
        for a in item.find_all('a'):
            download_tasks.append((name, a, head, output_dir, survey_name))
    
    if not download_tasks:
        progress.update(task_id, completed=1, description=f"[yellow]{survey_name[:30]}: 无文件")
        return {'total': 0, 'success': 0, 'errors': 0}
    
    # 更新进度条描述为下载状态
    progress.update(task_id, description=f"[blue]{survey_name[:30]}")
    
    success_count = 0
    error_count = 0
    
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(download_single_file, task) for task in download_tasks]
        
        for future in as_completed(futures):
            result = future.result()
            if result['success']:
                success_count += 1
            else:
                error_count += 1
            
            # 更新进度
            completed = success_count + error_count
            progress.update(task_id, completed=completed, 
                          description=f"[green]{survey_name[:30]}",
                          refresh=True)  # 强制刷新
    
    # 完成状态
    if error_count == 0:
        progress.update(task_id, description=f"[green]{survey_name[:30]}")
    else:
        progress.update(task_id, description=f"[red]{survey_name[:30]}")
    
    return {'total': len(download_tasks), 'success': success_count, 'errors': error_count}

def download_all_surveys_rich(cookie, max_workers=3):
    """使用rich并行下载所有测评"""
    # 为Windows系统优化console设置
    terminal_width = get_terminal_width()
    console = Console(
        legacy_windows=False,  # 使用现代Windows终端API
        force_terminal=True,    # 强制终端模式
        no_color=False,         # 保持颜色支持
        width=terminal_width    # 使用检测到的终端宽度
    )
    
    console.print(Panel.fit("[bold blue]测评文件下载器[/bold blue]", padding=1))
    console.print("正在获取测评列表...")
    
    surveys = get_survey_links(cookie)
    
    if not surveys:
        console.print("[red]未找到任何测评[/red]")
        return
    
    console.print(f"[green]找到 {len(surveys)} 个测评，正在收集文件信息...[/green]")
    
    # 先收集每个测评的文件数量
    survey_file_counts = []
    for survey in surveys:
        file_count = get_survey_file_count(cookie, survey['url'])
        survey_file_counts.append((survey, file_count))
        console.print(f"[cyan]{survey['name'][:40]}: {file_count} 个文件[/cyan]")
    
    total_files = sum(count for _, count in survey_file_counts)
    console.print(f"[green]共 {total_files} 个文件，开始并行下载...[/green]")
    
    base_output_dir = datetime.now().strftime("%Y-%m-%d")
    
    # 使用rich的Progress创建多进度条，针对Windows优化
    with Progress(
        SpinnerColumn(),
        TextColumn("[progress.description]{task.description}"),
        BarColumn(bar_width=None),  # 自动调整进度条宽度
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TextColumn("({task.completed}/{task.total})"),
        console=console,
        expand=True,              # 允许扩展使用全宽度
        refresh_per_second=4,      # 降低刷新频率
        transient=True             # 完成后清除进度条
    ) as progress:
        
        # 为每个测评创建任务，使用已收集的文件数量
        task_ids = {}
        for survey, file_count in survey_file_counts:
            survey_dir = os.path.join(base_output_dir, clean_filename(survey['name']))
            task_id = progress.add_task(f"[yellow]{survey['name'][:30]}", total=file_count)
            task_ids[task_id] = (survey, survey_dir, file_count)
        
        # 并行下载
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_task = {}
            for task_id, (survey, survey_dir, file_count) in task_ids.items():
                future = executor.submit(
                    fetch_one_survey_with_progress,
                    console, progress, task_id,
                    cookie, survey['url'], survey_dir, survey['name']
                )
                future_to_task[future] = (task_id, survey)
            
            # 等待所有完成
            all_results = []
            for future in as_completed(future_to_task):
                task_id, survey = future_to_task[future]
                try:
                    result = future.result()
                    all_results.append((survey, result))
                except Exception as e:
                    console.print(f"[red]{survey['name']}: 下载失败 - {str(e)}[/red]")
                    all_results.append((survey, {'total': 0, 'success': 0, 'errors': 1}))
    
    # 显示最终统计
    total_files = sum(r['total'] for _, r in all_results)
    total_success = sum(r['success'] for _, r in all_results)
    total_errors = sum(r['errors'] for _, r in all_results)
    
    # 创建统计表格
    table = Table(title="下载统计", show_header=True, header_style="bold magenta")
    table.add_column("项目", style="cyan", no_wrap=True)
    table.add_column("数值", justify="right")
    
    table.add_row("总测评数", str(len(surveys)))
    table.add_row("总文件数", str(total_files))
    table.add_row("成功下载", f"[green]{total_success}[/green]")
    table.add_row("下载失败", f"[red]{total_errors}[/red]")
    if total_files > 0:
        table.add_row("成功率", f"{total_success/total_files*100:.1f}%")
    table.add_row("保存位置", base_output_dir)
    
    console.print("\n")
    console.print(table)

if __name__ == "__main__":
    COOKIE="sid=" + input("请输入cookie中的sid：").strip()
    download_all_surveys_rich(COOKIE, max_workers=5)
