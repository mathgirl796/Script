from requests import get
from bs4 import BeautifulSoup
import os
import re
from PIL import Image
import io

def clean_filename(filename):
    """清理Windows不支持的文件名字符"""
    # Windows不支持的字符: \ / : * ? " < > |
    return re.sub(r'[\\/:*?"<>|]', '_', filename)

def get_survey_links(cookie):
    """
    获取所有测评的名称和实际链接
    
    参数:
        cookie (str): 认证cookie
    
    返回:
        list: 包含字典的列表，每个字典包含 name 和 url 键
    """
    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
        "Cookie": cookie
    }
    
    try:
        # 获取页面内容
        res = get("https://ndkh.hit.edu.cn/user/eva", headers=head)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'lxml')
        
        surveys = []
        
        # 查找所有包含"开始测评"文本的元素
        start_buttons = soup.find_all(string=re.compile(r'开始测评'))
        
        for button_text in start_buttons:
            parent = button_text.parent
            
            # 查找onclick事件中的survey ID
            onclick = parent.get('onclick', '')
            survey_match = re.search(r"openSurvey\('([^']+)'\)", onclick)
            
            if survey_match:
                survey_id = survey_match.group(1)
                
                # 构建实际的测评链接
                actual_url = f"https://ndkh.hit.edu.cn/user/eva_survey?annualSurveyId={survey_id}"
                
                # 获取测评名称
                survey_name = "未知测评"
                row = parent.find_parent('tr')
                if row:
                    # 查找测评名称（通常在第一列）
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
        print(f"获取测评链接时出错: {e}")
        return []

def fetch_one_survey(cookie, url, output_dir):

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    os.chdir(output_dir)

    head = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
        "Cookie": cookie
    }
    res = get(url, headers=head)
    
    soup = BeautifulSoup(res.text, 'lxml')

    items = soup.find_all('td', class_="name")

    for item in items:
        name = item['title'].strip()
        print(f"正在读取：{name}")
        for a in item.find_all('a'):
            id = str(a).split("'")[1]
            isPlan = 1 if "计划" in str(a) else 0
            url = f"https://ndkh.hit.edu.cn/reportWork_show?id={id}&isPlan={isPlan}"
            res = get(url, headers=head)
            pic_links = ["https://ndkh.hit.edu.cn" + img['data-src'] for img in BeautifulSoup(res.text, 'lxml').find_all('img')]
            pdf_filename = f"{clean_filename(name)}（{clean_filename(a.text.strip())}）.pdf"
            
            # 下载所有图片并转换为PDF页面
            images = []
            for image_id, link in enumerate(pic_links):
                print(f"下载第{image_id+1}张图片中")
                img_response = get(link, headers=head)
                img = Image.open(io.BytesIO(img_response.content))
                # 转换为RGB模式（PDF需要）
                if img.mode != 'RGB':
                    img = img.convert('RGB')
                images.append(img)
            
            # 保存为PDF文件
            if images:
                print("合并pdf中")
                images[0].save(pdf_filename, save_all=True, append_images=images[1:])
                print(f"已保存PDF: {pdf_filename}")

if __name__ == "__main__":
    COOKIE="sid=" + input("请输入cookie中的sid：").strip()
    for survey in get_survey_links(cookie=COOKIE):
        fetch_one_survey(cookie=COOKIE, url=survey["url"], output_dir=f"output/{survey['name']}")
