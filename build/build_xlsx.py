"""まとめたデータからExcelを作る。
使い方: python -I build_xlsx.py まとめたデータのフォルダ 出力xlsxのパス
"""
import json
import math
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding="utf-8")
data_dir, out_path = sys.argv[1], sys.argv[2]
load = lambda n: json.load(open(os.path.join(data_dir, f"{n}.json"), encoding="utf-8"))  # noqa: E731
annual, ytd, breakdown, notes, sources, world = (load(n) for n in ["annual", "ytd2026", "breakdown", "notes", "sources", "world"])
WORLD = "HIV/エイズ（世界）"

FONT = "游ゴシック"
GREEN = "28765B"
f_body = Font(name=FONT, size=10.5)
f_bold = Font(name=FONT, size=10.5, bold=True)
f_title = Font(name=FONT, size=15, bold=True, color=GREEN)
f_sec = Font(name=FONT, size=12, bold=True, color=GREEN)
f_head = Font(name=FONT, size=10.5, bold=True, color="FFFFFF")
f_note = Font(name=FONT, size=9.5, color="5F6B66")
f_link = Font(name=FONT, size=10.5, color="1F5FA8", underline="single")
fill_head = PatternFill("solid", fgColor=GREEN)
fill_soku = PatternFill("solid", fgColor="FFF4D1")
fill_soft = PatternFill("solid", fgColor="EEF5F1")
thin = Side(style="thin", color="C9D3CE")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="center", wrap_text=True)

SHEETS = [("HIV/エイズ", "HIV・エイズ", "HIV/エイズ", "合計", "合計"),
          ("梅毒", "梅毒", "梅毒", "合計", "合計"),
          ("はしか", "はしか", "はしか（麻しん）", "届出数", "届出数"),
          ("結核", "結核", "結核", "新登録患者数", "届出数（発病していない感染者も含む）")]
WHAT = {"HIV/エイズ": "新たに報告されたHIV感染者とエイズ患者の合計（件）",
        "梅毒": "医師が届け出た患者の数（人）",
        "はしか": "医師が届け出た患者の数（人）",
        "結核": "新たに登録された結核患者の数（人）"}
WIDTH = 9  # 文章を横に広げる列数


def numfmt(col, v):
    if isinstance(v, float) or "率" in col or "(%)" in col or "人口10万対" in col:
        return "0.0" if not ("人口10万対" in col and isinstance(v, float) and abs(v * 1000 - round(v * 1000)) < 1e-6 and round(v, 1) != v) else "0.000"
    return "#,##0"


def put(ws, r, c, v, font=f_body, fill=None, fmt=None, align=None, border=True):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if align:
        cell.alignment = align
    if border:
        cell.border = box
    return cell


def para(ws, r, text, font=f_body, cols=WIDTH, per_line=62):
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=cols)
    c = ws.cell(row=r, column=1, value=text)
    c.font = font
    c.alignment = wrap
    ws.row_dimensions[r].height = max(18, 16 * math.ceil(len(text) / per_line))
    return r + 1


def header(ws, r, labels):
    for i, h in enumerate(labels, 1):
        put(ws, r, i, h, f_head, fill_head, align=center)
    ws.row_dimensions[r].height = 30
    return r + 1


def section(ws, r, text):
    ws.cell(row=r, column=1, value=text).font = f_sec
    ws.row_dimensions[r].height = 22
    return r + 1


def notes_of(name, kind, table=None):
    return [n for n in notes if n["病気"] == name and n["種類"] == kind and n["表"] == table]


wb = Workbook()
intro = wb.active
intro.title = "はじめに"
links = {}  # 病気 -> まとめ表から参照するセル

for name, sheet, display, main, ymain in SHEETS:
    ws = wb.create_sheet(sheet)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 40
    for i in range(2, 12):
        ws.column_dimensions[get_column_letter(i)].width = 15
    ws.cell(row=1, column=1, value=display).font = f_title
    ws.row_dimensions[1].height = 26
    r = para(ws, 2, "数え方：" + notes_of(name, "数え方")[0]["文"], f_note)
    r += 1

    # 年ごとの推移
    rows = [a for a in annual if a["病気"] == name]
    metrics = list(dict.fromkeys(a["指標"] for a in rows))
    years = sorted({a["年"] for a in rows})
    r = section(ws, r, "年ごとの推移")
    r = header(ws, r, ["年"] + metrics + [f"前年比（{main}）", "区分", "出典ID"])
    first = r
    main_col = metrics.index(main) + 2
    by = {(a["年"], a["指標"]): a for a in rows}
    for y in years:
        status = by[(y, metrics[0])]["区分"]
        fill = fill_soku if status == "速報" else None
        put(ws, r, 1, y, f_bold, fill, '0"年"', Alignment(horizontal="left"))
        for j, m in enumerate(metrics, 2):
            a = by.get((y, m))
            if name == "HIV/エイズ" and m == "合計":
                # 合計は感染者と患者の和（2016〜2022年は資料に合計がないため）
                put(ws, r, j, f"=B{r}+C{r}", fill=fill, fmt="#,##0")
            else:
                put(ws, r, j, a["値"] if a else None, fill=fill, fmt=numfmt(m, a["値"] if a else 0))
        cl = get_column_letter(main_col)
        put(ws, r, len(metrics) + 2, f'=IF({cl}{r - 1}="","",{cl}{r}/{cl}{r - 1}-1)' if r > first else "—",
            fill=fill, fmt='+0.0%;-0.0%;0.0%', align=Alignment(horizontal="right"))
        put(ws, r, len(metrics) + 3, status, fill=fill, align=Alignment(horizontal="center"))
        put(ws, r, len(metrics) + 4, by[(y, metrics[0])]["出典ID"], fill=fill)
        r += 1
    confirmed = [y for y in years if by[(y, metrics[0])]["区分"] == "確定"]
    row_of = {y: first + i for i, y in enumerate(years)}
    cl = get_column_letter(main_col)
    links[name] = {"latest_year": f"'{sheet}'!A{row_of[confirmed[-1]]}", "latest": f"'{sheet}'!{cl}{row_of[confirmed[-1]]}",
                   "prev": f"'{sheet}'!{cl}{row_of[confirmed[-2]]}", "first_year": f"'{sheet}'!A{row_of[confirmed[0]]}",
                   "first": f"'{sheet}'!{cl}{row_of[confirmed[0]]}",
                   "soku": f"'{sheet}'!{cl}{row_of[years[-1]]}" if years[-1] not in confirmed else None}
    r = para(ws, r, "黄色の行は速報値（あとで変わることがあります）。それ以外は確定値。前年比はこの表で計算したもの。", f_note)
    r += 1

    # 2026年の途中経過
    r = section(ws, r, "2026年の途中経過（速報）")
    r = header(ws, r, ["指標", "期間", "2026年", "前年同期", "前年同期比", "区分", "出典ID"])
    for t in [t for t in ytd if t["病気"] == name]:
        put(ws, r, 1, t["指標"], align=wrap)
        put(ws, r, 2, t["期間"], align=wrap)
        put(ws, r, 3, t["2026年"], fmt="#,##0")
        put(ws, r, 4, t["前年同期"], fmt="#,##0")
        put(ws, r, 5, f"=C{r}/D{r}-1", fmt='+0.0%;-0.0%;0.0%')
        put(ws, r, 6, t["区分"], align=Alignment(horizontal="center"))
        put(ws, r, 7, t["出典ID"])
        ws.row_dimensions[r].height = 50
        if t["指標"] == ymain:
            links[name].update(ytd_label=f"'{sheet}'!B{r}", ytd=f"'{sheet}'!C{r}", ytd_prev=f"'{sheet}'!D{r}")
        r += 1
    detail = next(t for t in ytd if t["病気"] == name)["詳しい期間"]
    r = para(ws, r, "期間のくわしい説明：" + detail, f_note)
    for n in notes_of(name, "表の注", "2026年の途中経過"):
        r = para(ws, r, "注：" + n["文"], f_note)
    r += 1

    # 内訳の表
    rows = [b for b in breakdown if b["病気"] == name]
    for title in dict.fromkeys(b["表"] for b in rows):
        t_rows = [b for b in rows if b["表"] == title]
        cols = list(dict.fromkeys(b["列"] for b in t_rows))
        r = section(ws, r, title)
        r = header(ws, r, ["区分"] + cols + ["出典ID"])
        cell = {(b["区分"], b["列"]): b for b in t_rows}
        for label in dict.fromkeys(b["区分"] for b in t_rows):
            put(ws, r, 1, label, align=wrap)
            for j, c in enumerate(cols, 2):
                v = cell[(label, c)]["値"] if (label, c) in cell else None
                if v is None:
                    put(ws, r, j, "資料になし", f_note, align=Alignment(horizontal="right"))
                else:
                    put(ws, r, j, v, fmt="0" if c == "順位" else numfmt(c, v))
            put(ws, r, len(cols) + 2, t_rows[0]["出典ID"])
            if len(str(label)) > 22:
                ws.row_dimensions[r].height = 30
            r += 1
        for n in notes_of(name, "表の注", title):
            r = para(ws, r, "注：" + n["文"], f_note)
        r += 1

    for kind, head in [("特徴", "最近の特徴"), ("対策", "国の主な対策")]:
        r = section(ws, r, head)
        for n in notes_of(name, kind):
            r = para(ws, r, f"{n['番号']}. {n['文']}（出典 {n['出典ID']}）")
        r += 1
    r = section(ws, r, "この病気の数字を見るときの注意")
    for n in notes_of(name, "見るときの注意"):
        r = para(ws, r, f"・{n['文']}")
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

# HIV/エイズの世界の動向（UNAIDSの推計）
wsw = wb.create_sheet("HIV・世界の動向", wb.sheetnames.index("HIV・エイズ") + 1)
wsw.sheet_view.showGridLines = False
wsw.cell(row=1, column=1, value="HIV/エイズ 世界の動向").font = f_title
wsw.row_dimensions[1].height = 26
r = para(wsw, 2, "数え方：" + notes_of(WORLD, "数え方")[0]["文"], f_note, cols=13, per_line=110)
r += 1
wcols = ["区分", "協力覚書", "感染している人", "感染している人（下限）", "感染している人（上限）", "15〜49歳の感染率（%）", "新たな感染",
         "新たな感染の増減（2010年比・%）", "エイズによる死亡", "死亡の増減（2010年比・%）", "感染を知っている人（%）",
         "治療を受けている人（%）", "ウイルスが抑えられている人（%）"]
for part, rows_w in [(f"世界と地域（{world[0]['年']}年）", [w for w in world if w["区分"] != "国"]),
                     (f"日本と協力覚書を結んだ国（{world[0]['年']}年、感染している人の多い順）",
                      sorted([w for w in world if w["区分"] == "国"], key=lambda w: -(w["感染している人"] if isinstance(w["感染している人"], (int, float)) else -1)))]:
    r = section(wsw, r, part)
    r = header(wsw, r, ["名前"] + wcols + ["出典ID"])
    for w in rows_w:
        put(wsw, r, 1, w["名前"], f_bold, fill_soft)
        for j, c in enumerate(wcols, 2):
            v = w[c]
            if v is None:
                put(wsw, r, j, "—" if c != "協力覚書" else "", f_note, align=Alignment(horizontal="right"))
            elif isinstance(v, str):
                put(wsw, r, j, v, align=Alignment(horizontal="right" if c not in ("区分", "協力覚書") else "left", wrap_text=True))
            else:
                put(wsw, r, j, v, fmt='+0;-0;0' if "増減" in c else ("0.0" if isinstance(v, float) else "#,##0"))
        put(wsw, r, len(wcols) + 2, w["出典ID"])
        r += 1
    r += 1
r = para(wsw, r, "「<100」は100人未満、「—」はUNAIDSに数字がないことを表します。増減（2010年比・%）と割合（%）はUNAIDSの資料に書かれた値です。", f_note, cols=13, per_line=110)
r += 1
for kind, head in [("協力覚書", "日本の協力覚書（特定技能・技能実習）"), ("移住者とHIV", "移住者とHIV（国連の機関の見方）"),
                   ("入国前の健康診断", "日本の入国前の健康診断とIOM")]:
    r = section(wsw, r, head)
    for n in notes_of(WORLD, kind):
        r = para(wsw, r, f"{n['番号']}. {n['文']}" + (f"（出典 {n['出典ID']}）" if n["出典ID"] else ""), cols=13, per_line=110)
    r += 1
r = section(wsw, r, "この数字を見るときの注意")
for n in notes_of(WORLD, "見るときの注意"):
    r = para(wsw, r, f"・{n['文']}", cols=13, per_line=110)
wsw.column_dimensions["A"].width = 26
for i, w in enumerate([8, 22, 13, 12, 12, 12, 12, 13, 12, 13, 12, 12, 13, 10], 2):
    wsw.column_dimensions[get_column_letter(i)].width = w
wsw.freeze_panes = "B4"
wsw.page_setup.orientation = "landscape"
wsw.page_setup.fitToWidth = 1
wsw.page_setup.fitToHeight = 0
wsw.sheet_properties.pageSetUpPr.fitToPage = True

# 4つの比較
sm = wb.create_sheet("4つの比較", 1)
sm.sheet_view.showGridLines = False
sm.cell(row=1, column=1, value="4つの病気の比較").font = f_title
para(sm, 2, "数字は各病気のシートから自動で読み込んでいます。前年比・10年前と比べた倍率・前年同期比はここで計算したものです。", f_note, cols=14)
heads = ["病気", "数えているもの", "最新の確定値の年", "最新の確定値", "その前の年", "前年比",
         "10年前の年", "10年前の数", "10年前と比べて", "2025年（速報値）", "2026年の途中経過の期間",
         "2026年の途中経過", "前年の同じ時期", "前年同期比"]
header(sm, 4, heads)
widths = [18, 30, 12, 13, 13, 11, 11, 12, 12, 13, 30, 13, 13, 11]
for i, w in enumerate(widths, 1):
    sm.column_dimensions[get_column_letter(i)].width = w
r = 5
for name, sheet, display, main, ymain in SHEETS:
    L = links[name]
    put(sm, r, 1, display, f_bold, fill_soft)
    put(sm, r, 2, WHAT[name] if name != "結核" else WHAT[name], align=wrap)
    put(sm, r, 3, f"={L['latest_year']}", fmt='0"年"')
    put(sm, r, 4, f"={L['latest']}", f_bold, fmt="#,##0")
    put(sm, r, 5, f"={L['prev']}", fmt="#,##0")
    put(sm, r, 6, f"=D{r}/E{r}-1", fmt='+0.0%;-0.0%;0.0%')
    put(sm, r, 7, f"={L['first_year']}", fmt='0"年"')
    put(sm, r, 8, f"={L['first']}", fmt="#,##0")
    put(sm, r, 9, f"=D{r}/H{r}", fmt='0.0"倍"')
    put(sm, r, 10, f"={L['soku']}" if L["soku"] else "—", fmt="#,##0", fill=fill_soku if L["soku"] else None,
        align=None if L["soku"] else Alignment(horizontal="center"))
    put(sm, r, 11, f"={L['ytd_label']}", align=wrap)
    put(sm, r, 12, f"={L['ytd']}", fmt="#,##0", fill=fill_soku)
    put(sm, r, 13, f"={L['ytd_prev']}", fmt="#,##0", fill=fill_soku)
    put(sm, r, 14, f"=L{r}/M{r}-1", fmt='+0.0%;-0.0%;0.0%', fill=fill_soku)
    sm.row_dimensions[r].height = 36
    r += 1
r += 1
for t in ["・最新の確定値が出ている年は病気ごとに違います（HIV/エイズと結核は2025年、梅毒とはしかは2024年）。",
          "・「10年前の年」は、最新の確定値の年から数えて10年分の最初の年です。",
          "・2026年の途中経過は速報値で、病気によって期間が違います。結核の途中経過は医師の届出の数で、発病していない感染者も含むため、年報の「新登録患者数」とは比べられません。",
          "・黄色のセルは速報値です。"]:
    r = para(sm, r, t, f_note, cols=14, per_line=120)
sm.freeze_panes = "B5"
sm.page_setup.orientation = "landscape"
sm.sheet_properties.pageSetUpPr.fitToPage = True
sm.page_setup.fitToHeight = 0

# 見るときの注意
cs = wb.create_sheet("見るときの注意")
cs.sheet_view.showGridLines = False
cs.column_dimensions["A"].width = 16
cs.column_dimensions["B"].width = 110
cs.cell(row=1, column=1, value="数字を見るときの注意").font = f_title
r = 3
for who in ["全体", "HIV/エイズ", WORLD, "梅毒", "はしか", "結核"]:
    cs.cell(row=r, column=1, value=who if who != "全体" else "全体にかかわること").font = f_sec
    r += 1
    for n in [n for n in notes if n["病気"] == who and n["種類"] == "見るときの注意"]:
        put(cs, r, 1, n["番号"], align=Alignment(horizontal="center", vertical="top"), border=False)
        c = put(cs, r, 2, n["文"], align=wrap, border=False)
        cs.row_dimensions[r].height = max(18, 16 * math.ceil(len(n["文"]) / 70))
        r += 1
    r += 1

# 出典一覧
ss = wb.create_sheet("出典一覧")
ss.sheet_view.showGridLines = False
ss.cell(row=1, column=1, value="出典一覧（国と国連機関の公式資料のみ）").font = f_title
para(ss, 2, "各シートの「出典ID」はこの表の番号です。URLはクリックで開けます。確認日は2026年10月26日。", f_note, cols=7)
header(ss, 4, ["出典ID", "病気", "資料名", "発行元", "URL", "公表日・更新日", "確認日"])
for i, w in enumerate([10, 11, 60, 30, 55, 22, 12], 1):
    ss.column_dimensions[get_column_letter(i)].width = w
r = 5
for s in sources:
    put(ss, r, 1, s["出典ID"])
    put(ss, r, 2, s["病気"])
    put(ss, r, 3, s["資料名"], align=wrap)
    put(ss, r, 4, s["発行元"], align=wrap)
    c = put(ss, r, 5, s["URL"], f_link, align=wrap)
    c.hyperlink = s["URL"]
    put(ss, r, 6, s["公表日・更新日"], align=wrap)
    put(ss, r, 7, s["確認日"])
    ss.row_dimensions[r].height = max(18, 15 * math.ceil(max(len(s["資料名"]) / 30, len(s["URL"]) / 50)))
    r += 1
ss.freeze_panes = "A5"
ss.auto_filter.ref = f"A4:G{r - 1}"

# はじめに
intro.sheet_view.showGridLines = False
intro.column_dimensions["A"].width = 22
intro.column_dimensions["B"].width = 100
intro.cell(row=1, column=1, value="日本の感染症を知る：HIV/エイズ・梅毒・はしか・結核の状況").font = f_title
rows = [
    ("作成日", "2026年10月26日（この日までに国が公表した資料をもとにしています）"),
    ("内容", "4つの感染症について、最新の人数、過去10年の動き、男女・年代・都道府県、2026年の途中経過、国の対策をまとめました。"),
    ("使った資料", "国の公式資料だけを使いました（厚生労働省、国立健康危機管理研究機構〔JIHS〕）。HIV/エイズの年報は、厚生労働省のページから案内されている「エイズ予防情報ネット」（国の委託先）に置かれているものを使っています。HIV/エイズの「世界の動向」だけは、国連の機関（国連合同エイズ計画〔UNAIDS〕・国際移住機関〔IOM〕）の公式資料を使っています。"),
    ("シートの説明", "「4つの比較」…4つの病気を1つの表で比べます。\n「HIV・エイズ」「梅毒」「はしか」「結核」…病気ごとの詳しい数字です。\n「HIV・世界の動向」…世界全体と、日本と協力覚書を結んだ国のHIVの状況（UNAIDSの推計）、移住者とHIV、入国前の健康診断。\n「見るときの注意」…数字を読むときに気をつけること。\n「出典一覧」…すべての資料名とURL。"),
    ("確定値", "国が1年分の届出を集計し直した数。"),
    ("速報値", "毎週や3か月ごとに出る途中の数。あとから変わることがあります。表では黄色で示しています。"),
    ("届出数", "医師が感染症と診断して保健所に届け出た数。検査を受けていない人などは入らないため、実際に感染している人の数とは同じではありません。"),
    ("罹患率（りかんりつ）", "人口10万人あたり、1年間に新しく病気になった人の数。人口の違う地域や年どうしを比べやすくなります。"),
]
r = 3
for k, v in rows:
    put(intro, r, 1, k, f_bold, fill_soft, align=Alignment(vertical="top"))
    put(intro, r, 2, v, align=wrap)
    intro.row_dimensions[r].height = max(20, 16 * (v.count("\n") + math.ceil(len(v) / 64)))
    r += 1

wb.save(out_path)
print("saved", out_path)
