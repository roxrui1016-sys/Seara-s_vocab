import csv, glob, hashlib, io, json, os, re, zipfile
HEAD = re.compile(r"^(english|英文|英语|word|words|单词|中文|chinese)$", re.I)

def unesc(s):
    for a, b in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'), ("&apos;", "'"), ("&amp;", "&")):
        s = s.replace(a, b)
    return s

def text_of(x):
    x = re.sub(r"<rPh\b.*?</rPh>", "", x, flags=re.S)
    return "".join(unesc(m) for m in re.findall(r"<t\b[^>]*>(.*?)</t>", x, flags=re.S))

def xlsx_rows(path):
    z = zipfile.ZipFile(path)
    names = z.namelist()
    ss = []
    if "xl/sharedStrings.xml" in names:
        x = z.read("xl/sharedStrings.xml").decode("utf8")
        for m in re.finditer(r"<si\b[^>]*?/>|<si\b[^>]*>(.*?)</si>", x, flags=re.S):
            ss.append(text_of(m.group(1) or ""))
    sheets = sorted(n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    xml = z.read(sheets[0]).decode("utf8")
    xml = re.sub(r"<row\b[^>]*/>", "", xml)
    rows = []
    for rm in re.finditer(r"<row\b[^>]*>(.*?)</row>", xml, flags=re.S):
        row = []
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", rm.group(1), flags=re.S):
            at, body = cm.group(1), cm.group(2) or ""
            t = re.search(r'\bt="([^"]+)"', at)
            v = re.search(r"<v>(.*?)</v>", body, flags=re.S)
            if t and t.group(1) == "s":
                val = ss[int(v.group(1))] if v else ""
            elif t and t.group(1) == "inlineStr":
                val = text_of(body)
            else:
                val = unesc(v.group(1)) if v else ""
            row.append(val)
        rows.append(row)
    return rows

def csv_rows(path):
    raw = open(path, "rb").read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("gbk", errors="replace")
    return list(csv.reader(io.StringIO(text)))

def count_words(rows):
    n = 0
    for i, r in enumerate(rows):
        cells = [c.strip() for c in r if c and c.strip()]
        if len(cells) < 2:
            continue
        if i == 0 and HEAD.match(cells[0]):
            continue
        n += 1
    return n

os.makedirs("books", exist_ok=True)
CATS = ("口语", "阅读", "写作")
books = []
paths = glob.glob("books/**/*.xlsx", recursive=True) + glob.glob("books/**/*.csv", recursive=True)
for p in sorted(paths):
    p = p.replace(os.sep, "/")
    name = os.path.splitext(os.path.basename(p))[0]
    if name.startswith("~$") or os.path.basename(p) == "index.json":
        continue
    parts = p.split("/")                      # books / 大类 / 小类 / 文件
    cat = parts[1] if len(parts) >= 3 and parts[1] in CATS else ""
    sub = parts[2] if cat and len(parts) >= 4 else ""
    try:
        rows = xlsx_rows(p) if p.lower().endswith(".xlsx") else csv_rows(p)
        count = count_words(rows)
    except Exception as e:
        print("跳过无法读取的文件:", p, e)
        continue
    ver = hashlib.sha1(open(p, "rb").read()).hexdigest()[:10]
    books.append({"name": name, "file": p, "ver": ver, "count": count, "cat": cat, "sub": sub})
json.dump({"books": books}, open("books/index.json", "w", encoding="utf8"), ensure_ascii=False, indent=1)
print("词本库：", len(books), "本")
for b in books:
    print(" ", b["cat"] or "-", b["sub"] or "-", b["name"], b["count"], "词", b["ver"])
