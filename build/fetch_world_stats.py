"""国別の感染状況（WHOの公開データ）を取ってきて、data/research/world_stats.json を作る。

- 結核：WHO Global Tuberculosis Database（推計値。人口10万人あたりの罹患率と人数）
- はしか：WHO Global Health Observatory（各国がWHOに報告した数）
- HIV：WHO Global Health Observatory（UNAIDSの推計値をWHOが載せたもの）
- 梅毒：WHO Global Health Observatory（各国がWHOに報告した数・割合。報告している国は一部だけ）
- 地域：国連統計部の M49（国・地域の区分）

使い方（感染症フォルダで）: python -I build/fetch_world_stats.py （ダウンロードは build/_dl に置く）
"""
import csv
import html
import json
import os
import re
import sys
import urllib.request

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DL = os.path.join(ROOT, "build", "_dl")
OUT = os.path.join(ROOT, "data", "research", "world_stats.json")
CHECKED = "2026-10-26"
os.makedirs(DL, exist_ok=True)

TB_URL = "https://extranet.who.int/tme/generateCSV.asp?ds=estimates"
GHO = "https://ghoapi.azureedge.net/api/{}?$filter=SpatialDimType%20eq%20%27COUNTRY%27"
M49_URL = "https://unstats.un.org/unsd/methodology/m49/overview/"
IND = {
    "measles": "WHS3_62",            # Measles - number of reported cases
    "plhiv": "HIV_0000000001",       # Estimated number of people (all ages) living with HIV
    "hiv_new": "HIV_0000000026",     # Number of new HIV infections
    "hiv_prev": "MDG_0000000029",    # Prevalence of HIV among adults aged 15 to 49 (%)
    "syph_anc": "PercposANC",        # Antenatal care attendees who were positive for syphilis (%), reported
    "syph_cases": "STI_SYPHILIS_CASES_NUM",  # Reported cases of syphilis (number)
    "cs_rate": "RateCS",             # Congenital syphilis rate per 100 000 live births, reported
}


def get(url, name, refresh):
    path = os.path.join(DL, name)
    if refresh or not os.path.exists(path):
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=120) as r, open(path, "wb") as f:
            f.write(r.read())
    return path


refresh = "--refresh" in sys.argv

# 地域（国連 M49）
m49 = open(get(M49_URL, "m49.html", refresh), encoding="utf-8", errors="replace").read()
tbl = m49[m49.find('id = "downloadTableEN"'):]
tbl = tbl[:tbl.find("</table>")]
REGION = {}
for tr in re.findall(r"<tr>(.*?)</tr>", tbl, re.S)[1:]:
    td = [html.unescape(re.sub(r"<.*?>", "", x)).strip() for x in re.findall(r"<td>(.*?)</td>", tr, re.S)]
    if len(td) >= 12 and td[11]:
        REGION[td[11]] = {"region": td[3], "sub": td[5], "inter": td[7]}

# 地域の日本語名（国連 M49 のサブ地域をまとめたもの）
def region_ja(iso3):
    r = REGION.get(iso3)
    if not r:
        return "その他"
    sub, inter = r["sub"], r["inter"]
    return {
        "Eastern Asia": "東アジア", "South-eastern Asia": "東南アジア", "Southern Asia": "南アジア",
        "Central Asia": "中央アジア", "Western Asia": "西アジア・中東", "Northern Africa": "アフリカ",
        "Sub-Saharan Africa": "アフリカ", "Northern America": "北アメリカ",
        "Latin America and the Caribbean": "中南米・カリブ", "Australia and New Zealand": "オセアニア",
        "Melanesia": "オセアニア", "Micronesia": "オセアニア", "Polynesia": "オセアニア",
    }.get(sub, "ヨーロッパ" if r["region"] == "Europe" else "その他")


# 国名（日本語）
JA = {
    "ABW": "アルバ（オランダ領）", "AFG": "アフガニスタン", "AGO": "アンゴラ", "AIA": "アンギラ（英国領）", "ALB": "アルバニア", "AND": "アンドラ",
    "ARE": "アラブ首長国連邦", "ARG": "アルゼンチン", "ARM": "アルメニア", "ASM": "米領サモア", "ATG": "アンティグア・バーブーダ", "AUS": "オーストラリア",
    "AUT": "オーストリア", "AZE": "アゼルバイジャン", "BDI": "ブルンジ", "BEL": "ベルギー", "BEN": "ベナン", "BFA": "ブルキナファソ", "BGD": "バングラデシュ",
    "BGR": "ブルガリア", "BHR": "バーレーン", "BHS": "バハマ", "BIH": "ボスニア・ヘルツェゴビナ", "BLR": "ベラルーシ", "BLZ": "ベリーズ", "BMU": "バミューダ（英国領）",
    "BOL": "ボリビア", "BRA": "ブラジル", "BRB": "バルバドス", "BRN": "ブルネイ", "BTN": "ブータン", "BWA": "ボツワナ", "CAF": "中央アフリカ", "CAN": "カナダ",
    "CHE": "スイス", "CHL": "チリ", "CHN": "中国", "CIV": "コートジボワール", "CMR": "カメルーン", "COD": "コンゴ民主共和国", "COG": "コンゴ共和国",
    "COK": "クック諸島", "COL": "コロンビア", "COM": "コモロ", "CPV": "カーボベルデ", "CRI": "コスタリカ", "CUB": "キューバ", "CUW": "キュラソー（オランダ領）",
    "CYM": "ケイマン諸島（英国領）", "CYP": "キプロス", "CZE": "チェコ", "DEU": "ドイツ", "DJI": "ジブチ", "DMA": "ドミニカ国", "DNK": "デンマーク",
    "DOM": "ドミニカ共和国", "DZA": "アルジェリア", "ECU": "エクアドル", "EGY": "エジプト", "ERI": "エリトリア", "ESP": "スペイン", "EST": "エストニア",
    "ETH": "エチオピア", "FIN": "フィンランド", "FJI": "フィジー", "FRA": "フランス", "FSM": "ミクロネシア連邦", "GAB": "ガボン", "GBR": "英国",
    "GEO": "ジョージア", "GHA": "ガーナ", "GIN": "ギニア", "GMB": "ガンビア", "GNB": "ギニアビサウ", "GNQ": "赤道ギニア", "GRC": "ギリシャ",
    "GRD": "グレナダ", "GRL": "グリーンランド（デンマーク領）", "GTM": "グアテマラ", "GUM": "グアム（米国領）", "GUY": "ガイアナ", "HKG": "中国（香港）",
    "HND": "ホンジュラス", "HRV": "クロアチア", "HTI": "ハイチ", "HUN": "ハンガリー", "IDN": "インドネシア", "IND": "インド", "IRL": "アイルランド",
    "IRN": "イラン", "IRQ": "イラク", "ISL": "アイスランド", "ISR": "イスラエル", "ITA": "イタリア", "JAM": "ジャマイカ", "JOR": "ヨルダン", "JPN": "日本",
    "KAZ": "カザフスタン", "KEN": "ケニア", "KGZ": "キルギス", "KHM": "カンボジア", "KIR": "キリバス", "KNA": "セントクリストファー・ネービス", "KOR": "韓国",
    "KWT": "クウェート", "LAO": "ラオス", "LBN": "レバノン", "LBR": "リベリア", "LBY": "リビア", "LCA": "セントルシア", "LKA": "スリランカ",
    "LSO": "レソト", "LTU": "リトアニア", "LUX": "ルクセンブルク", "LVA": "ラトビア", "MAC": "中国（マカオ）", "MAR": "モロッコ", "MCO": "モナコ",
    "MDA": "モルドバ", "MDG": "マダガスカル", "MDV": "モルディブ", "MEX": "メキシコ", "MHL": "マーシャル諸島", "MKD": "北マケドニア", "MLI": "マリ",
    "MLT": "マルタ", "MMR": "ミャンマー", "MNE": "モンテネグロ", "MNG": "モンゴル", "MNP": "北マリアナ諸島（米国領）", "MOZ": "モザンビーク",
    "MRT": "モーリタニア", "MSR": "モントセラト（英国領）", "MUS": "モーリシャス", "MWI": "マラウイ", "MYS": "マレーシア", "NAM": "ナミビア",
    "NCL": "ニューカレドニア（フランス領）", "NER": "ニジェール", "NGA": "ナイジェリア", "NIC": "ニカラグア", "NIU": "ニウエ", "NLD": "オランダ",
    "NOR": "ノルウェー", "NPL": "ネパール", "NRU": "ナウル", "NZL": "ニュージーランド", "OMN": "オマーン", "PAK": "パキスタン", "PAN": "パナマ",
    "PER": "ペルー", "PHL": "フィリピン", "PLW": "パラオ", "PNG": "パプアニューギニア", "POL": "ポーランド", "PRI": "プエルトリコ（米国領）",
    "PRK": "北朝鮮", "PRT": "ポルトガル", "PRY": "パラグアイ", "PSE": "パレスチナ", "PYF": "フランス領ポリネシア", "QAT": "カタール", "ROU": "ルーマニア",
    "RUS": "ロシア", "RWA": "ルワンダ", "SAU": "サウジアラビア", "SDN": "スーダン", "SEN": "セネガル", "SGP": "シンガポール", "SLB": "ソロモン諸島",
    "SLE": "シエラレオネ", "SLV": "エルサルバドル", "SMR": "サンマリノ", "SOM": "ソマリア", "SRB": "セルビア", "SSD": "南スーダン",
    "STP": "サントメ・プリンシペ", "SUR": "スリナム", "SVK": "スロバキア", "SVN": "スロベニア", "SWE": "スウェーデン", "SWZ": "エスワティニ",
    "SXM": "シント・マールテン（オランダ領）", "SYC": "セーシェル", "SYR": "シリア", "TCA": "タークス・カイコス諸島（英国領）", "TCD": "チャド",
    "TGO": "トーゴ", "THA": "タイ", "TJK": "タジキスタン", "TKL": "トケラウ", "TKM": "トルクメニスタン", "TLS": "東ティモール", "TON": "トンガ",
    "TTO": "トリニダード・トバゴ", "TUN": "チュニジア", "TUR": "トルコ", "TUV": "ツバル", "TZA": "タンザニア", "UGA": "ウガンダ", "UKR": "ウクライナ",
    "URY": "ウルグアイ", "USA": "アメリカ合衆国", "UZB": "ウズベキスタン", "VCT": "セントビンセント・グレナディーン", "VEN": "ベネズエラ",
    "VGB": "英領バージン諸島", "VNM": "ベトナム", "VUT": "バヌアツ", "WLF": "ウォリス・フツナ（フランス領）", "WSM": "サモア", "YEM": "イエメン",
    "ZAF": "南アフリカ", "ZMB": "ザンビア", "ZWE": "ジンバブエ", "XKX": "コソボ",
}

# 結核（WHOの推計。最新年と、その前の年）
tb_rows = list(csv.DictReader(open(get(TB_URL, "tb_estimates.csv", refresh), encoding="utf-8")))
TB_Y = max(int(r["year"]) for r in tb_rows)
TB = {}
for r in tb_rows:
    y = int(r["year"])
    if y in (TB_Y, TB_Y - 1):
        TB.setdefault(r["iso3"], {})[y] = r
EN = {r["iso3"]: r["country"] for r in tb_rows if int(r["year"]) == TB_Y}


def n(v):
    try:
        return float(v) if "." in str(v) else int(v)
    except (TypeError, ValueError):
        return None


# WHO Global Health Observatory
def gho(code):
    d = json.load(open(get(GHO.format(code), f"gho_{code}.json", refresh), encoding="utf-8"))["value"]
    out = {}
    for x in d:
        if x.get("Dim1") not in (None, "SEX_BTSX"):  # 男女計だけ
            continue
        out.setdefault(x["SpatialDim"], {})[int(x["TimeDim"])] = x
    return out


G = {k: gho(v) for k, v in IND.items()}


def latest(code, iso3, year=None, min_year=None):
    """その国のいちばん新しい年（year を指定したらその年）の値。値のない年は飛ばす。min_year より古い年は使わない"""
    ys = G[code].get(iso3, {})
    cand = [year] if year else sorted((y for y in ys if not min_year or y >= min_year), reverse=True)
    for y in cand:
        x = ys.get(y)
        if not x:
            continue
        v = x.get("NumericValue")
        txt = (x.get("Value") or "").strip()
        if v is None and not txt:
            continue
        # 「<100」のような値は、GHOの数値欄では100になっているため、文字のまま使う（UNAIDSの表し方と同じ）
        m = re.match(r"^(<\s*[\d.]+)", txt.replace(" ", " "))
        if m:
            v = m.group(1).replace(" ", "")
        return {"year": y, "value": v, "text": txt, "low": x.get("Low"), "high": x.get("High")}
    return None


ME_Y = max(max(v) for v in G["measles"].values())
SYPH_MIN = 2020  # 梅毒は国によって報告の年がばらばら。古すぎる報告（2019年以前）は使わない
HIV_Y = max(max(v) for v in G["plhiv"].values())

rows = []
for iso3 in sorted(set(EN) | {"XKX"} & set(G["measles"])):
    t = TB.get(iso3, {})
    t1, t0 = t.get(TB_Y), t.get(TB_Y - 1)
    row = {"iso3": iso3, "name": JA.get(iso3, EN.get(iso3, iso3)), "en": EN.get(iso3, iso3), "region": region_ja(iso3)}
    if t1 and n(t1["e_inc_100k"]) is not None:
        row["tb"] = {
            "year": TB_Y, "rate": n(t1["e_inc_100k"]), "rate_lo": n(t1["e_inc_100k_lo"]), "rate_hi": n(t1["e_inc_100k_hi"]),
            "num": n(t1["e_inc_num"]), "num_lo": n(t1["e_inc_num_lo"]), "num_hi": n(t1["e_inc_num_hi"]), "pop": n(t1["e_pop_num"]),
            "prev_year": TB_Y - 1 if t0 else None, "prev_rate": n(t0["e_inc_100k"]) if t0 else None, "prev_num": n(t0["e_inc_num"]) if t0 else None,
        }
    m1, m0 = latest("measles", iso3, ME_Y), latest("measles", iso3, ME_Y - 1)
    if m1 or m0:
        pop = row.get("tb", {}).get("pop")
        row["measles"] = {
            "year": ME_Y, "cases": m1["value"] if m1 else None,
            "prev_year": ME_Y - 1, "prev_cases": m0["value"] if m0 else None,
            # 参考：人口100万人あたり（報告数 ÷ WHOの結核推計で使われた人口。このサイトで計算）
            "per_m": round(m1["value"] / pop * 1e6, 1) if m1 and m1["value"] is not None and pop else None, "pop_year": TB_Y if pop else None,
        }
    h1 = latest("plhiv", iso3, HIV_Y)
    if h1:
        nw, pv = latest("hiv_new", iso3, HIV_Y), latest("hiv_prev", iso3, HIV_Y)
        nw0 = latest("hiv_new", iso3, HIV_Y - 1)
        row["hiv"] = {
            "year": HIV_Y, "plhiv": h1["value"], "plhiv_text": h1["text"],
            "new": nw["value"] if nw else None, "new_text": nw["text"] if nw else None,
            "prev_new": nw0["value"] if nw0 else None,
            "prev15_49": pv["value"] if pv else None, "prev15_49_text": pv["text"] if pv else None,
        }
    s = {}
    for k in ("syph_anc", "syph_cases", "cs_rate"):
        v = latest(k, iso3, min_year=SYPH_MIN)
        if v and v["value"] is not None:
            s[k] = {"year": v["year"], "value": v["value"]}
    if s:
        row["syphilis"] = s
    rows.append(row)

missing = [r["iso3"] for r in rows if r["iso3"] not in JA]
if missing:
    print("日本語名がない国:", missing)

out = {
    "checked": CHECKED,
    "years": {"tb": TB_Y, "measles": ME_Y, "hiv": HIV_Y, "syphilis_from": SYPH_MIN},
    "rows": rows,
    "sources": [
        {"id": "W1", "title": f"WHO Global Tuberculosis Database：結核の推計（{TB_Y}年まで。Global Tuberculosis Report 2025 のデータ）", "publisher": "世界保健機関（WHO）",
         "url": "https://www.who.int/teams/global-programme-on-tuberculosis-and-lung-health/data", "published": "2025-11（Global Tuberculosis Report 2025）", "data": TB_URL},
        {"id": "W2", "title": "Global Health Observatory：Measles - number of reported cases（はしかの報告数）", "publisher": "世界保健機関（WHO）",
         "url": "https://www.who.int/data/gho/data/indicators/indicator-details/GHO/measles---number-of-reported-cases", "published": f"{ME_Y}年分まで掲載（{CHECKED}確認）", "data": GHO.format(IND["measles"])},
        {"id": "W3", "title": "Global Health Observatory：HIV（UNAIDSの推計。HIVとともに生きる人・新たな感染・15〜49歳の感染率）", "publisher": "世界保健機関（WHO）／国連合同エイズ計画（UNAIDS）",
         "url": "https://www.who.int/data/gho/data/themes/hiv-aids", "published": f"{HIV_Y}年分まで掲載（{CHECKED}確認）", "data": GHO.format(IND["plhiv"])},
        {"id": "W4", "title": "Global Health Observatory：梅毒（妊婦健診で梅毒が陽性だった割合・報告された梅毒の数・先天梅毒の率。各国の報告）", "publisher": "世界保健機関（WHO）",
         "url": "https://www.who.int/data/gho/data/themes/topics/data-on-syphilis", "published": f"{CHECKED}確認", "data": GHO.format(IND["syph_anc"])},
        {"id": "W5", "title": "Standard country or area codes for statistical use（M49。国・地域の区分）", "publisher": "国際連合統計部（UNSD）",
         "url": M49_URL, "published": "記載なし", "data": M49_URL},
    ],
}
json.dump(out, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"{len(rows)}か国・地域 → {OUT}")
print("年:", out["years"])
for r in rows:
    if r["iso3"] in ("JPN", "PHL", "CHN", "VNM", "NPL", "IDN", "MMR", "LAO"):
        print(r["name"], json.dumps({k: r.get(k) for k in ("tb", "measles", "hiv", "syphilis")}, ensure_ascii=False))
