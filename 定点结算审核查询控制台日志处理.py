import re
import pandas as pd

def get_brackets_content(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        content = f.read()

    # 匹配 [] 中的内容，不包含括号本身
    pattern = r'\[([^\]]+)\]'
    matches = re.findall(pattern, content)

    return matches

def get_first_number(text):
    match = re.search(r'\d+', text)
    if match:
        return match.group()
    return None

brackets_content = get_brackets_content('all.log')

all_entries = []
for i in range(0, len(brackets_content), 4):
    id = brackets_content[i]
    page_num = get_first_number(brackets_content[i+1])
    info = brackets_content[i+2]+brackets_content[i+3]
    infos = re.split(r'\\[\\tn\s]+',info)
    entry = [id, page_num] + infos[1:]
    all_entries.append(entry)

# 写入Excel
df = pd.DataFrame(all_entries)
df.to_excel('output.xlsx', index=False, header=False)
print(f'成功导出 {len(all_entries)} 条数据到 output.xlsx')
    