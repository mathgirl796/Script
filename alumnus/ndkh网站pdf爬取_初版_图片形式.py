from requests import get
from bs4 import BeautifulSoup
import os

head = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.46",
    "Cookie": "sid=2bf6fbe8-264d-4548-a00c-6c986b6493cd"
}
res = get(
    "https://ndkh.hit.edu.cn/user/eva_survey?inspectorId=&annualSurveyId=NDA4",
    headers=head)
# print(res.text)


soup = BeautifulSoup(res.text, 'lxml')
items = soup.find_all('td', class_="name")
for item in items:
    name = item['title'].strip()
    for a in item.find_all('a'):
        id = str(a).split("'")[1]
        isPlan = 1 if "计划" in str(a) else 0
        url = f"https://ndkh.hit.edu.cn/reportWork_show?id={id}&isPlan={isPlan}"
        res = get(url, headers=head)
        pic_links = ["https://ndkh.hit.edu.cn" + img['data-src'] for img in BeautifulSoup(res.text, 'lxml').find_all('img')]
        # print(pic_links)
        path = f"{name}/{a.text.strip()}"
        if not os.path.exists(path):
            os.makedirs(path)
        for i, link in enumerate(pic_links):
            with open(f"{path}/{i}.jpg", "wb") as f:
                f.write(get(link, headers=head).content)


