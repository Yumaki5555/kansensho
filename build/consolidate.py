"""4つの病気の調査JSONを、Excelとページ共通の形（縦長の表）にまとめる。
使い方: python -I consolidate.py データのフォルダ 出力フォルダ
"""
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
src_dir, out_dir = sys.argv[1], sys.argv[2]
os.makedirs(out_dir, exist_ok=True)

DISEASES = [
    ("hiv", "HIV/エイズ", "HIV"),
    ("syphilis", "梅毒", "梅毒"),
    ("measles", "はしか", "はしか"),
    ("tb", "結核", "結核"),
]
# ページでは表示しない表（Excelには載せる）
OFF_PAGE = {
    "2026年第1・第2四半期（速報）の四半期別報告数と前年同期",
    "高齢の新登録患者の割合（2025年）",
    "結核による死亡数・男女別（2025年、人口動態統計 確定数）",
}
YTD_LABEL = {
    "HIV/エイズ": "2026年1〜6月ごろ（第1・第2四半期）",
    "梅毒": "2026年1月〜10月11日（第41週まで）",
    "はしか": "2026年1月〜10月11日（第41週まで）",
    "結核": "2026年1月〜10月11日（第41週まで）",
}
TB_YTD_METRIC = "結核の届出数（感染症発生動向調査・累積）"
TB_YTD_NAME = "届出数（発病していない感染者も含む）"

CAUTIONS = {
    "全体": [
        "ここにある数は、医師が保健所に届け出た数です（結核は、保健所に新しく登録された患者の数）。検査を受けていない人や、感染に気づいていない人は含まれないため、実際に感染している人の数とは同じではありません。",
        "「確定値」は、国が1年分の届出を集計し直した数です。「速報値」は、毎週や3か月ごとに出る途中の数で、あとから届出が追加されて変わることがあります。たとえば梅毒の2024年は、速報値14,663例が確定値では14,829例になりました。",
        "最新の確定値が出ている年は病気ごとに違います。HIV/エイズと結核は2025年、梅毒とはしかは2024年までです（2025年は速報値）。",
        "「2026年の途中経過」は、病気によって期間が違います。HIV/エイズは1〜6月ごろ（第1・第2四半期）、梅毒・はしか・結核は10月11日（第41週）までで、それぞれ前の年の同じ時期と比べています。",
        "出典は国の公式資料（厚生労働省、国立健康危機管理研究機構）だけを使いました。ただしHIV/エイズの年報は、厚生労働省のページから案内されている「エイズ予防情報ネット」（国の委託先）に置かれているものを使っています。HIV/エイズの「世界の動向」だけは、国連の機関（UNAIDS・IOM）の公式資料を使っています。「公費で受けられる支え」のうち結核の医療費助成の一部は、東京都の公式ページで確かめました。",
        "前年比や「10年前と比べて何倍か」などは、資料の数字からこのまとめで計算したものです。",
    ],
    "HIV/エイズ": [
        "年は「報告された年」で、診断された年ではありません。",
        "2016〜2022年の合計は、HIV感染者とエイズ患者の数を足したものです（2023年以降は資料に合計が書かれています）。",
        "2024年は、速報値（合計1,000件）が確定値では994件に直されました。",
        "感染経路の「その他」は、2019年に届出用紙が変わって中身が変わったため、それより前の年とは比べにくくなっています。",
        "2025年の検査件数は、郵送検査を含まない106,756件を使いました。2025年9月から郵送検査も数えるようになったため、過去の年と比べるときは郵送を除く必要があります。",
        "厚生労働省の「HIVとエイズ」のページは、2026年10月26日時点でまだ2024年の数字のままでした。",
    ],
    "梅毒": [
        "2025年の確定値は、国の予定では2027年2月ごろに公表されます。このため2025年は速報値です。",
        "2025年（速報）の男女別は、厚生労働省のページのグラフに書かれた数字を読み取ったものです。",
        "性別がわからない人が1〜2人いる年があるため、男女を足しても合計と合わない年があります（2022年・2024年・2025年）。",
        "都道府県の人口10万人あたりの数は、資料に書かれている東京都と大阪府だけ載せています。",
    ],
    "はしか": [
        "2025年の確定値は、2026年10月26日時点でまだ公表されていません。2025年は速報値です。",
        "2024年のワクチン接種歴別は、確定値の集計表に区分がないため、国の専門誌（IASR、2025年5月8日時点）の数字を使いました。",
        "都道府県の人口あたりの数は、確定値の集計表にないため載せていません。",
        "2回接種した人の届出もありますが、若い世代はもともと2回接種した人が多いため、人数だけでワクチンの効果は比べられません。",
    ],
    "結核": [
        "男女別の新登録患者数は、厚生労働省の年報に表がないため載せていません。",
        "「2026年の途中経過」は医師の届出の数（速報）で、発病していないが治療が必要な感染者（潜在性結核感染症）なども含みます。年報の「新登録患者数」とは数え方が違うため、直接は比べられません。",
        "80歳以上・70歳以上の割合は、年代別の数字を足して計算したものです。",
        "年報のPDFには公表日の記載がないため、出典にはファイルの更新日（2026年10月2日）を書いています。",
        "結核による死亡数は、2017年から死因の分け方が変わっていて、その影響を含みます。",
    ],
}

annual, ytd, breakdown, notes, sources = [], [], [], [], []
overview_keys = []


def sid(prefix, s):
    return f"{prefix}-{s[1:]}" if s else None


for key, name, prefix in DISEASES:
    d = json.load(open(os.path.join(src_dir, f"{key}.json"), encoding="utf-8"))
    S = lambda s: sid(prefix, s)  # noqa: E731

    # 年ごとの推移
    for r in d["annual"]["rows"]:
        status = "確定" if r["status"].startswith("確定") else "速報"
        for col in d["annual"]["columns"]:
            v = r["values"].get(col)
            if v is None:
                continue
            annual.append({"病気": name, "年": r["year"], "指標": col, "値": v,
                           "区分": status, "出典ID": S(r["source_id"])})

    # 2026年の途中経過
    y = d.get("ytd_2026")
    if y:
        prev = y.get("prev_year_same_period") or {}
        for col, v in y["values"].items():
            metric = TB_YTD_NAME if col == TB_YTD_METRIC else col
            ytd.append({"病気": name, "指標": metric, "期間": YTD_LABEL[name],
                        "詳しい期間": y["as_of"], "2026年": v, "前年同期": prev.get(col),
                        "区分": "速報", "出典ID": S(y["source_id"])})
        if y.get("note"):
            notes.append({"病気": name, "種類": "表の注", "表": "2026年の途中経過", "文": y["note"], "出典ID": S(y["source_id"])})
        if y.get("prev_year_note"):
            notes.append({"病気": name, "種類": "表の注", "表": "2026年の途中経過", "文": y["prev_year_note"], "出典ID": S(y["source_id"])})

    tables = []
    sx = d.get("sex_latest")
    if sx and any(v is not None for r in sx["rows"] for v in r["values"].values()):
        tables.append({"title": f"男女別（{sx['year']}年）", "columns": sx["columns"], "rows": sx["rows"],
                       "source_id": sx["source_id"], "note": sx.get("note"), "kind": "男女別"})
    ag = d["age_latest"]
    tables.append({"title": f"年代別（{ag['year']}年）", "columns": ag["columns"], "rows": ag["rows"],
                   "source_id": ag["source_id"], "note": ag.get("note"), "kind": "年代別"})
    pf = d["prefecture_top5"]
    if key == "hiv":
        cols = ["順位", "人数（感染者＋患者）", "HIV感染者", "エイズ患者", "HIV感染者（人口10万対）", "エイズ患者（人口10万対）"]
        prow = [{"label": r["pref"], "values": {"順位": r["rank"], "人数（感染者＋患者）": r["count"], "HIV感染者": r["hiv"],
                                                 "エイズ患者": r["aids"], "HIV感染者（人口10万対）": r["rate_hiv_per_100k"],
                                                 "エイズ患者（人口10万対）": r["rate_aids_per_100k"]}} for r in pf["rows"]]
    elif key == "tb":
        cols = ["順位", "罹患率（人口10万対）", "人数"]
        prow = [{"label": r["pref"], "values": {"順位": r["rank"], "罹患率（人口10万対）": r["rate_per_100k"], "人数": r["count"]}}
                for r in pf["rows"]]
    else:
        has_rate = any(r["rate_per_100k"] is not None for r in pf["rows"])
        cols = ["順位", "人数"] + (["人口10万対"] if has_rate else [])
        prow = [{"label": r["pref"], "values": {"順位": r["rank"], "人数": r["count"],
                                                 **({"人口10万対": r["rate_per_100k"]} if has_rate else {})}} for r in pf["rows"]]
    ptitle = f"都道府県別 上位5（{pf['year']}年、{'罹患率' if key == 'tb' else '人数'}の多い順）"
    tables.append({"title": ptitle, "columns": cols, "rows": prow, "source_id": pf["source_id"], "note": pf.get("note"), "kind": "都道府県"})

    if key == "hiv":
        a = d["aids_share_latest"]
        tables.append({"title": f"新規報告に占めるエイズ患者の割合（{a['year']}年）",
                       "columns": [f"{a['year']}年(%)", f"{a['year'] - 1}年(%)"],
                       "rows": [{"label": b["area"], "values": {f"{a['year']}年(%)": b["percent"], f"{a['year'] - 1}年(%)": b["prev_year_percent"]}}
                                for b in a["by_area"]],
                       "source_id": a["source_id"], "note": "エイズを発症してから初めてHIV感染がわかった人の割合。資料に書かれた値。", "kind": "その他"})
        c = d["cumulative_latest"]
        tables.append({"title": f"1985年からの累積報告数（{c['as_of']}）", "columns": list(c["values"].keys()),
                       "rows": [{"label": "累積", "values": c["values"]}], "source_id": c["source_id"], "note": c.get("note"), "kind": "その他"})
    for e in d["extra_tables"]:
        tables.append({**e, "kind": "その他"})

    for ti, t in enumerate(tables):
        on_page = t["title"] not in OFF_PAGE
        for ri, r in enumerate(t["rows"]):
            vals = [r["values"].get(c) for c in t["columns"]]
            if all(v is None for v in vals):
                continue
            for c in t["columns"]:
                breakdown.append({"病気": name, "表": t["title"], "種類": t["kind"], "区分": r["label"], "列": c,
                                  "値": r["values"].get(c), "出典ID": S(t["source_id"]), "ページ": on_page})
        if t.get("note"):
            notes.append({"病気": name, "種類": "表の注", "表": t["title"], "文": t["note"], "出典ID": S(t["source_id"])})

    notes.append({"病気": name, "種類": "数え方", "表": None, "文": d["metric_note"], "出典ID": None})
    for f in d["features"]:
        notes.append({"病気": name, "種類": "特徴", "表": None, "文": f["text"], "出典ID": S(f["source_id"])})
    for m in d["measures"]:
        notes.append({"病気": name, "種類": "対策", "表": None, "文": m["text"], "出典ID": S(m["source_id"])})
    for s in d["sources"]:
        sources.append({"出典ID": S(s["id"]), "病気": name, "資料名": s["title"], "発行元": s["publisher"],
                        "URL": s["url"], "公表日・更新日": s.get("published") or "記載なし", "確認日": s.get("accessed") or "2026-10-26"})

# HIV/エイズの世界の動向（UNAIDSの推計と、IOM・日本の協力覚書の資料）
WORLD = "HIV/エイズ（世界）"
world = []
wd = json.load(open(os.path.join(src_dir, "world_hiv.json"), encoding="utf-8"))
WCOLS = [("plhiv", "感染している人"), ("plhiv_low", "感染している人（下限）"), ("plhiv_high", "感染している人（上限）"),
         ("prev", "15〜49歳の感染率（%）"), ("new", "新たな感染"), ("deaths", "エイズによる死亡"),
         ("chg_new", "新たな感染の増減（2010年比・%）"), ("chg_deaths", "死亡の増減（2010年比・%）"),
         ("know", "感染を知っている人（%）"), ("art", "治療を受けている人（%）"), ("vls", "ウイルスが抑えられている人（%）")]
for r in wd["rows"]:
    world.append({"区分": r["kind"], "名前": r["name"], "英語名": r["en"], "協力覚書": r.get("moc"), "年": wd["data_year"],
                  **{label: r.get(k) for k, label in WCOLS}, "出典ID": sid("世界", r["source_id"])})
notes.append({"病気": WORLD, "種類": "数え方", "表": None, "文": wd["metric_note"], "出典ID": None})
for kind, items in wd["notes"].items():
    for n in items:
        n = n if isinstance(n, dict) else {"text": n, "source_id": None}
        notes.append({"病気": WORLD, "種類": kind, "表": None, "文": n["text"], "出典ID": sid("世界", n.get("source_id"))})
for s in wd["sources"]:
    sources.append({"出典ID": sid("世界", s["id"]), "病気": WORLD, "資料名": s["title"], "発行元": s["publisher"],
                    "URL": s["url"], "公表日・更新日": s.get("published") or "記載なし", "確認日": s.get("accessed") or "2026-10-26"})

# 公費で受けられる支え（検査・予防・治療）。出典の「S1」などは公費の資料、「HIV-14」などは病気ごとの資料
KOUHI = "公費の制度"
kouhi = []
kd = json.load(open(os.path.join(src_dir, "kouhi.json"), encoding="utf-8"))
K = lambda s: sid("公費", s) if s and re.fullmatch(r"S\d+", s) else s  # noqa: E731
for r in kd["summary"]:
    kouhi.append({"病気": r["disease"], "場面": r["scene"], "目安": r["level"], "内容": r["text"], "出典ID": K(r["source_id"])})
notes.append({"病気": KOUHI, "種類": "説明", "表": None, "文": kd["intro"], "出典ID": None})
for n in kd["common"]:
    notes.append({"病気": KOUHI, "種類": "しくみ", "表": None, "文": n["text"], "出典ID": K(n["source_id"])})
SCENE_ORDER = ["【検査】", "【予防】", "【治療】"]
for who, items in kd["items"].items():
    # 検査 → 予防 → 治療 の順に並べる（同じ場面の中は書いた順）
    for n in sorted(items, key=lambda n: next(i for i, p in enumerate(SCENE_ORDER) if n["text"].startswith(p))):
        notes.append({"病気": who, "種類": "公費", "表": None, "文": n["text"], "出典ID": K(n["source_id"])})
for t in kd["tables"]:
    for r in t["rows"]:
        for c in t["columns"]:
            breakdown.append({"病気": t["disease"], "表": t["title"], "種類": "公費", "区分": r["label"], "列": c,
                              "値": r["values"].get(c), "出典ID": K(t["source_id"]), "ページ": True})
    if t.get("note"):
        notes.append({"病気": t["disease"], "種類": "表の注", "表": t["title"], "文": t["note"], "出典ID": K(t["source_id"])})
for t in kd["cautions"]:
    notes.append({"病気": KOUHI, "種類": "見るときの注意", "表": None, "文": t, "出典ID": None})
for s in kd["sources"]:
    sources.append({"出典ID": K(s["id"]), "病気": KOUHI, "資料名": s["title"], "発行元": s["publisher"],
                    "URL": s["url"], "公表日・更新日": s.get("published") or "記載なし", "確認日": s.get("accessed") or "2026-10-26"})
skeys = [(k["病気"], k["場面"]) for k in kouhi]
assert len(skeys) == len(set(skeys)) and set(k["場面"] for k in kouhi) <= set(kd["scenes"])

for who, items in CAUTIONS.items():
    for t in items:
        notes.append({"病気": who, "種類": "見るときの注意", "表": None, "文": t, "出典ID": None})

# 文章の中の仮の出典番号（S12など）を、出典一覧の番号（結核-12など）に置きかえる
PREFIX = {name: prefix for _, name, prefix in DISEASES}
PREFIX[WORLD] = "世界"
for n in notes:
    if n["病気"] in PREFIX:
        n["文"] = re.sub(r"(?<![A-Za-z0-9])S(\d+)(?![0-9])", lambda m: f"{PREFIX[n['病気']]}-{m.group(1)}", n["文"])

# 番号を振る（同じ病気・種類の中での順番）
seen = {}
for n in notes:
    k = (n["病気"], n["種類"], n["表"])
    seen[k] = seen.get(k, 0) + 1
    n["番号"] = seen[k]

# ページ用の内訳は、同じ行が2つないことを確かめる
keys = [(b["病気"], b["表"], b["区分"], b["列"]) for b in breakdown]
assert len(keys) == len(set(keys)), "内訳の行が重複しています"
akeys = [(a["病気"], a["年"], a["指標"]) for a in annual]
assert len(akeys) == len(set(akeys))
wkeys = [w["名前"] for w in world]
assert len(wkeys) == len(set(wkeys))
used = {x["出典ID"] for x in annual + ytd + breakdown + notes + world + kouhi if x["出典ID"]}
known = {s["出典ID"] for s in sources}
assert used <= known, used - known

for fname, rows in [("annual", annual), ("ytd2026", ytd), ("breakdown", breakdown), ("notes", notes), ("sources", sources), ("world", world), ("kouhi", kouhi)]:
    with open(os.path.join(out_dir, f"{fname}.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=0)
    print(fname, len(rows), "rows", os.path.getsize(os.path.join(out_dir, f"{fname}.json")), "bytes")
