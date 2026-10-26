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
        "出典は国の公式資料（厚生労働省、国立健康危機管理研究機構）だけを使いました。ただしHIV/エイズの年報は、厚生労働省のページから案内されている「エイズ予防情報ネット」（国の委託先）に置かれているものを使っています。HIV/エイズの「世界の動向」だけは、国連の機関（UNAIDS・IOM）の公式資料を使っています。「公費で受けられる支援」のうち結核の医療費助成の一部は、東京都の公式ページで確かめました。",
        "前年比や「収録開始年（基準年）と比べて何倍か」などは、資料の数字からこのまとめで計算したものです。基準年はHIV/エイズ・結核が2016年、梅毒・はしかが2015年で、最新の確定値の年との間は9年です（「10年前」ではありません）。",
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
        "性別がわからない人（性別不明）が、2022年に1人、2024年に1人、2025年（速報）に2人います（2022年・2024年は国立健康危機管理研究機構の2024年の報告、2025年は厚生労働省のグラフに書かれた数）。このため、男女を足しても合計と合いません。ほかの年の性別不明0人は、合計と男女の和が同じことから、このサイトで計算したものです。",
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

# 公費で受けられる支援（検査・予防・治療）。出典の「S1」などは公費の資料、「HIV-14」などは病気ごとの資料
KOUHI = "公費の制度"
kouhi = []
kd = json.load(open(os.path.join(src_dir, "kouhi.json"), encoding="utf-8"))
K = lambda s: sid("公費", s) if s and re.fullmatch(r"S\d+", s) else s  # noqa: E731
LEVELS = [lv["name"] for lv in kd["levels"]]
for r in kd["summary"]:
    assert r["level"] in LEVELS, r
    kouhi.append({"病気": r["disease"], "場面": r["scene"], "目安": r["level"], "内容": r["text"], "出典ID": K(r["source_id"])})
notes.append({"病気": KOUHI, "種類": "説明", "表": None, "文": kd["intro"], "出典ID": None})
# 一覧の「目安」の見方（表＝目安の名前）
for lv in kd["levels"]:
    notes.append({"病気": KOUHI, "種類": "目安", "表": lv["name"], "文": lv["text"], "出典ID": None, "印": lv["mark"]})
for n in kd["common"]:
    notes.append({"病気": KOUHI, "種類": "しくみ", "表": None, "文": n["text"], "出典ID": K(n["source_id"])})
SCENE_ORDER = ["【検査】", "【予防】", "【治療】"]
SCENE_ITEM = {"【検査】": "検査・健診", "【予防】": "予防", "【治療】": "治療"}
for who, items in kd["items"].items():
    # 検査 → 予防 → 治療 の順に並べる（同じ場面の中は書いた順）。「項目」はページの病気別の詳しい説明の見出し
    for n in sorted(items, key=lambda n: next(i for i, p in enumerate(SCENE_ORDER) if n["text"].startswith(p))):
        item = n.get("section") or SCENE_ITEM[n["text"][:4]]
        notes.append({"病気": who, "種類": "公費", "表": None, "文": n["text"], "出典ID": K(n["source_id"]), "項目": item})
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

# ===== 国別の比較（WHOの公開データ）・協力覚書・入国前の健康確認 =====
COUNTRY = "国別の比較"
ENTRY = "入国前対策"
COST = "医療費"
INFO = "病気の解説"
OTHERS = "その他の感染症"


def load(name):
    p = os.path.join(src_dir, name)
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


def add_sources(items, prefix, who):
    for s in items or []:
        sources.append({"出典ID": ids(prefix, [s["id"]]), "病気": who, "資料名": s["title"], "発行元": s["publisher"],
                        "URL": s["url"], "公表日・更新日": s.get("published") or "記載なし", "確認日": s.get("checked") or s.get("accessed") or "2026-10-26"})


def ids(prefix, lst):
    """["制S1", "制S2"] → "制度-1・制度-2"（出典の仮の番号を、出典一覧の番号にする）"""
    out = []
    for s in lst or []:
        m = re.fullmatch(r"\D*?(\d+)", str(s))
        if m:
            out.append(f"{prefix}-{m.group(1)}")
    return "・".join(dict.fromkeys(out)) or None


seido = load("seido.json")
ws = load("world_stats.json")
# 協力覚書（seido.json があればそちら、なければ世界の動向の表）
moc_of = {}
if seido:
    for c in seido["moc"]["countries"]:
        moc_of[c["iso3"]] = "・".join(s["scheme"] for s in sorted(c["schemes"], key=lambda s: ["特定技能", "技能実習", "育成就労"].index(s["scheme"])))
else:
    en2iso = {r["en"]: r["iso3"] for r in ws["rows"]}
    for r in world:
        if r["区分"] == "国" and r["協力覚書"]:
            moc_of[en2iso[r["英語名"]]] = r["協力覚書"]
jpets_of = {c["iso3"]: c for c in (seido["preentry"]["jpets_countries"] if seido else [])}
countries = []
for r in ws["rows"]:
    t, m, hv, sy = r.get("tb") or {}, r.get("measles") or {}, r.get("hiv") or {}, r.get("syphilis") or {}
    jp = jpets_of.get(r["iso3"])
    countries.append({
        "ISO3": r["iso3"], "国": r["name"], "英語名": r["en"], "地域": r["region"],
        "協力覚書": moc_of.get(r["iso3"]),
        "入国前結核検査": jp["status"] if jp else None, "入国前結核検査_開始": jp.get("start") if jp else None,
        "関係の深い国": bool(moc_of.get(r["iso3"]) or jp or r["iso3"] == "JPN"),
        "結核_年": t.get("year"), "結核_罹患率": t.get("rate"), "結核_罹患率_下限": t.get("rate_lo"), "結核_罹患率_上限": t.get("rate_hi"),
        "結核_人数": t.get("num"), "結核_人数_下限": t.get("num_lo"), "結核_人数_上限": t.get("num_hi"),
        "結核_前年の罹患率": t.get("prev_rate"), "結核_前年の人数": t.get("prev_num"),
        "はしか_年": m.get("year"), "はしか_報告数": m.get("cases"), "はしか_前年の報告数": m.get("prev_cases"), "はしか_100万人あたり": m.get("per_m"), "はしか_人口の年": m.get("pop_year") if m.get("per_m") is not None else None,
        "HIV_年": hv.get("year"), "HIV_感染している人": hv.get("plhiv"), "HIV_感染している人_幅": hv.get("plhiv_text"),
        "HIV_新たな感染": hv.get("new"), "HIV_新たな感染_幅": hv.get("new_text"), "HIV_前年の新たな感染": hv.get("prev_new"),
        "HIV_15〜49歳の感染率": hv.get("prev15_49"),
        "梅毒_妊婦の陽性率": (sy.get("syph_anc") or {}).get("value"), "梅毒_妊婦の陽性率_年": (sy.get("syph_anc") or {}).get("year"),
        "梅毒_報告数": (sy.get("syph_cases") or {}).get("value"), "梅毒_報告数_年": (sy.get("syph_cases") or {}).get("year"),
        "梅毒_先天梅毒の率": (sy.get("cs_rate") or {}).get("value"), "梅毒_先天梅毒の率_年": (sy.get("cs_rate") or {}).get("year"),
        "出典ID": "WHO-1・WHO-2・WHO-3・WHO-4・WHO-5",
    })
    countries[-1] = {k: v for k, v in countries[-1].items() if v is not None}  # 値のない項目は書かない（ファイルを小さくする）
assert set(moc_of) <= {c["ISO3"] for c in countries}, set(moc_of) - {c["ISO3"] for c in countries}
add_sources(ws["sources"], "WHO", COUNTRY)
Y_ = ws["years"]
notes.append({"病気": COUNTRY, "種類": "数え方", "表": None, "出典ID": None, "文":
              f"国別の数字は、世界保健機関（WHO）が公開しているデータです。結核は1年間に新たに結核になった人の推計（{Y_['tb']}年。人口10万人あたりの罹患率と人数）、"
              f"はしかは各国がWHOに報告した患者の数（{Y_['measles']}年）、HIVは国連合同エイズ計画（UNAIDS）の推計（{Y_['hiv']}年。WHOのデータベースに載っているもの）、"
              f"梅毒は各国がWHOに報告した数字（妊婦健診で梅毒が陽性だった人の割合など。{Y_['syphilis_from']}年以降でいちばん新しい年）です。"})
for t in [
    "ここにある数字は、その国に住む人全体（一般人口）の数字です。その国から日本へ入国する人や、日本に住むその国籍の人の感染の割合ではありません。",
    "国籍や出身地域から、個人の感染や危険性を判断することはできません。",
    "結核とHIVは推計値（調査や報告をもとに計算した、おおよその数）で、幅があります。はしかと梅毒は各国の報告の数で、検査や届出のしくみの違いによって、実際より少なく出る国があります。国どうしを単純に比べることはできません。",
    f"日本の結核は、WHOの推計（{Y_['tb']}年、人口10万人あたり"
    f"{next(c['結核_罹患率'] for c in countries if c['ISO3'] == 'JPN')}）と、国の年報の数字（新しく登録された患者の罹患率）が違います。WHOの推計は、届け出られなかった人の分も見込んだ数のためです。国内の数字は、結核のページの「日本国内の推移」をご覧ください。",
    f"はしかの報告数は、流行のある年とない年で大きく変わります。{Y_['measles']}年の数は各国がWHOに報告したもので、あとから直されることがあります。人口100万人あたりの数は、{Y_['measles']}年の報告数を、WHOが結核の推計で使った{Y_['tb']}年の人口で割って、このサイトで計算した目安です（{Y_['measles']}年の人口で割った正式な率ではないため、暫定の値として見てください）。",
    f"梅毒は、国どうしを比べられる統計が少ないため、WHOに報告された「妊婦健診で梅毒が陽性だった人の割合」を中心に載せています。報告がない国は「WHOに報告なし」です（{Y_['syphilis_from'] - 1}年以前の古い報告は使っていません）。国によって報告の年が違います。",
    "HIVは、UNAIDSが国の推計を公表していない国（中国・日本など）は空欄です。「<100」は100人未満という意味です。",
]:
    notes.append({"病気": COUNTRY, "種類": "見るときの注意", "表": None, "文": t, "出典ID": None})

# 協力覚書・入国前の健康確認（seido.json）
entry = []
if seido:
    for c in seido["moc"]["countries"]:
        for s in c["schemes"]:
            entry.append({"区分": "協力覚書", "国": c["name"], "ISO3": c["iso3"], "制度": s["scheme"], "名前": s.get("title"),
                          "日付": s.get("date"), "日付の種類": s.get("date_kind"), "内容": s.get("purpose"), "資料URL": s.get("doc_url"),
                          "出典ID": ids("制度", s.get("source_ids"))})
        entry.append({"区分": "覚書の健康の記載", "国": c["name"], "ISO3": c["iso3"], "内容": c.get("health_clause"), "出典ID": None})
    for n in seido["moc"].get("notes", []):
        notes.append({"病気": ENTRY, "種類": "協力覚書", "表": None, "文": n["text"], "出典ID": ids("制度", n.get("source_ids"))})
    P = seido["preentry"]
    for s in P["systems"]:
        entry.append({"区分": "制度", "ID": s.get("id"), "名前": s["name"], "病気": "・".join(s.get("diseases") or []), "内容": s.get("what"),
                      "検査の内容": "／".join(s.get("content") or []), "対象": s.get("targets"), "対象外": "／".join(s.get("exempt") or []),
                      "開始": s.get("start"), "有効期間": s.get("validity"), "根拠": s.get("basis"), "状態": s.get("status"), "出典ID": ids("制度", s.get("source_ids"))})
    for c in P["jpets_countries"]:
        entry.append({"区分": "JPETSの国", "国": c["name"], "ISO3": c["iso3"], "状態": c["status"], "開始": c.get("start"), "内容": c.get("detail"),
                      "出典ID": ids("制度", c.get("source_ids"))})
    for x in P["matrix"]:
        entry.append({"区分": "病気と制度", "病気": x["disease"], "名前": x["item"], "状態": x["status"], "内容": x.get("text"), "出典ID": ids("制度", x.get("source_ids"))})
    for x in P.get("schemes", []):
        entry.append({"区分": "在留資格", "名前": x["scheme"], "JPETS": x.get("jpets"), "内容": x.get("health_check"), "状態": x.get("status"), "出典ID": ids("制度", x.get("source_ids"))})
    for x in P.get("gaps", []):
        entry.append({"区分": "確認事項", "名前": x["kind"], "内容": x["text"], "出典ID": ids("制度", x.get("source_ids"))})
    add_sources(seido["sources"], "制度", ENTRY)

# 医療費（costs.json）
costs = []
cd = load("costs.json")
if cd:
    for f in cd.get("financing", []):
        costs.append({"区分": "財源", "病気": "共通", "項目": "公的医療保険の財源", "名前": f.get("text"), "値": json.dumps(f.get("values"), ensure_ascii=False) if f.get("values") else None,
                      "単位": f.get("unit"), "年": f.get("year"), "出典ID": ids("費用", f.get("source_ids"))})
    for f in cd.get("insurance_rules", []):
        costs.append({"区分": "保険のきまり", "病気": "共通", "項目": "保険のきまり", "名前": f.get("text"), "値": json.dumps(f.get("table"), ensure_ascii=False) if f.get("table") else None,
                      "年": f.get("valid"), "出典ID": ids("費用", f.get("source_ids"))})
    for it in cd["items"]:
        # 1人当たりの値かどうかを報告書から確かめられない研究推計は載せない（生涯・長期の費用は条件を示せるものだけ）
        if "1人当たり" in (it.get("note") or "") and "確かめられない" in (it.get("note") or ""):
            continue
        costs.append({"区分": "費用", "病気": it["disease"], "項目": it["item"], "名前": it.get("label") or it["item"], "値": it.get("value"), "単位": it.get("unit"),
                      "年": it.get("year"), "種類": it.get("kind"), "対象": it.get("target"), "範囲": it.get("scope"), "負担": it.get("payer"),
                      "計算式": it.get("calc"), "注": it.get("note"), "状態": it.get("status"), "探した資料": it.get("searched"), "出典ID": ids("費用", it.get("source_ids"))})
    calc = cd.get("calculator") or {}
    for r in calc.get("rules", []):
        costs.append({"区分": "試算のきまり", "病気": "共通", "項目": r["case"], "名前": r["case"], "値": json.dumps(r.get("steps"), ensure_ascii=False),
                      "注": "　".join(t for t in [r.get("conditions"), f"例：{r['example']}" if r.get("example") else None] if t), "出典ID": ids("費用", r.get("source_ids"))})
    if calc.get("why"):
        notes.append({"病気": COST, "種類": "試算について", "表": None, "文": calc["why"], "出典ID": None})
    if cd.get("comparable"):
        notes.append({"病気": COST, "種類": "比べ方", "表": None, "文": cd["comparable"], "出典ID": None})
    add_sources(cd["sources"], "費用", COST)

# 病気の解説（disease_info.json）
info = []
di = load("disease_info.json")
if di:
    for who, d in di["diseases"].items():
        for sec in ["overview", "routes", "not_routes", "world", "test", "prevent", "treat"]:
            for n in d.get(sec) or []:
                info.append({"病気": who, "節": sec, "見出し": None, "文": n["text"], "出典ID": ids("解説", n.get("source_ids"))})
        for x in d.get("world_numbers") or []:
            info.append({"病気": who, "節": "world_numbers", "見出し": x["label"], "値": x.get("value"), "単位": x.get("unit"), "年": x.get("year"),
                         "種類": x.get("kind"), "出典ID": ids("解説", x.get("source_ids"))})
        for sp in d.get("special") or []:
            for n in sp.get("items") or []:
                info.append({"病気": who, "節": "special", "見出し": sp["heading"], "文": n["text"], "出典ID": ids("解説", n.get("source_ids"))})
            for x in sp.get("numbers") or []:
                info.append({"病気": who, "節": "special_numbers", "見出し": sp["heading"], "文": x["label"], "値": x.get("value"), "単位": x.get("unit"),
                             "年": x.get("year"), "種類": x.get("kind"), "出典ID": ids("解説", x.get("source_ids"))})
    add_sources(di["sources"], "解説", INFO)

# その他の感染症と、さらに詳しく調べたい方への資料（others_guide.json）
others, guide = [], []
og = load("others_guide.json")
if og:
    for o in og["others"]:
        others.append({"病名": o["name"], "特徴": o.get("feature"), "感染経路": o.get("routes"), "予防・対策": o.get("prevent"), "出典ID": ids("他", o.get("source_ids"))})
    for g in og["guide"]:
        guide.append({"分類": g["category"], "資料名": g["title"], "発行機関": g["publisher"], "対象期間": g.get("period"), "公表日": g.get("published"),
                      "調べられる内容": g.get("content"), "URL": g["url"], "言語": g.get("lang")})
    add_sources(og["sources"], "他", OTHERS)

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
used = {i for x in annual + ytd + breakdown + notes + world + kouhi + countries + entry + costs + info + others if x.get("出典ID") for i in x["出典ID"].split("・")}
known = {s["出典ID"] for s in sources}
assert used <= known, used - known
assert len(known) == len(sources), "出典IDが重複しています"

for fname, rows in [("annual", annual), ("ytd2026", ytd), ("breakdown", breakdown), ("notes", notes), ("sources", sources), ("world", world), ("kouhi", kouhi),
                    ("countries", countries), ("entry", entry), ("costs", costs), ("info", info), ("others", others), ("guide", guide)]:
    with open(os.path.join(out_dir, f"{fname}.json"), "w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=0)
    print(fname, len(rows), "rows", os.path.getsize(os.path.join(out_dir, f"{fname}.json")), "bytes")
