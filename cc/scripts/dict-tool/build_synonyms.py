#!/usr/bin/env python3
"""
從教育部《重編國語辭典修訂本》附錄「相似詞索引表」抓全部詞目，存成 synonyms.json
（{詞目: [相似詞, ...]}），給 CNdictionary.py 的近義詞替換功能離線查用。
辭典改版時重跑一次即可：python3 build_synonyms.py
"""
import json, re
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import requests, urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL = "https://dict.revised.moe.edu.tw/appendix.jsp"
OUT = Path(__file__).parent / "synonyms.json"
_SPLIT = re.compile(r"[、，,；;\s]+")
_GLYPH_IMG = re.compile(r"[A-Za-z0-9&?_.]")  # 缺字以圖片（如 &3493_.png）或「?」代替的詞，無法顯示，略過

def fetch_page(session, page):
    r = session.get(URL, params={"ID": "6", "page": page, "la": "0", "powerMode": "0"}, timeout=30)
    r.encoding = "utf-8"
    rows = []
    for tr in BeautifulSoup(r.text, "html.parser").select("table.appendV tr"):
        val, sub = tr.find("td", class_="val"), tr.find("td", class_="sub")
        if val and sub:
            rows.append((val.get_text(strip=True), sub.get_text(" ", strip=True)))
    return rows

def main():
    s = requests.Session()
    s.headers["User-Agent"] = "Mozilla/5.0"
    s.verify = False
    first = s.get(URL, params={"ID": "6", "page": 1, "la": "0", "powerMode": "0"}, timeout=30)
    first.encoding = "utf-8"
    pages = max(int(p) for p in re.findall(r"page=(\d+)", first.text))
    with ThreadPoolExecutor(max_workers=4) as ex:
        results = list(ex.map(lambda p: fetch_page(s, p), range(1, pages + 1)))

    table = {}
    for rows in results:
        for head, sub in rows:
            syns = table.setdefault(head, [])
            for w in _SPLIT.split(sub):
                w = w.strip()
                if w and w != head and w not in syns and not _GLYPH_IMG.search(w):
                    syns.append(w)
    OUT.write_text(json.dumps(table, ensure_ascii=False, indent=0), encoding="utf-8")
    print(f"{pages} 頁，{sum(len(r) for r in results)} 筆，{len(table)} 個詞目 → {OUT}")

if __name__ == "__main__":
    main()
