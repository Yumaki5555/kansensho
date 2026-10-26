"""GitHub Pages 用のページ（docs フォルダ）を作る。
claude.ai のダッシュボードと同じページ本体（page_fragment.html）を、
単体で動くように「見た目の設定」と「データを渡す小さな仕組み」で包む。

使い方（感染症フォルダで）: python -I build/build_site.py
"""
import hashlib
import json
import os
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCS = os.path.join(ROOT, "docs")
SITE_URL = "https://yumaki5555.github.io/kansensho/"
XLSX_SRC = os.path.join(ROOT, "日本の感染症の状況_2026年10月.xlsx")
XLSX_PUB = "kansensho_2026-10.xlsx"
TITLE = "日本の感染症を知る"
DESC = "HIV/エイズ・梅毒・はしか・結核は、増えている？減っている？ 国の公式資料をもとに、10年の動きや2026年の最新状況をグラフと表で比べられます。HIV/エイズは、日本と協力覚書を結んだ国など世界の動向も。"

os.makedirs(DOCS, exist_ok=True)
data = {n: json.load(open(os.path.join(ROOT, "data", "out", f"{n}.json"), encoding="utf-8"))
        for n in ["annual", "ytd2026", "breakdown", "notes", "sources", "world"]}
data_js = "window.KJ_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n"
open(os.path.join(DOCS, "data.js"), "w", encoding="utf-8").write(data_js)
ver = hashlib.sha1(data_js.encode("utf-8")).hexdigest()[:8]
shutil.copyfile(XLSX_SRC, os.path.join(DOCS, XLSX_PUB))
fragment = open(os.path.join(ROOT, "build", "page_fragment.html"), encoding="utf-8").read()

head = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>{TITLE}｜HIV/エイズ・梅毒・はしか・結核の今</title>
<meta name="description" content="{DESC}">
<meta property="og:type" content="website">
<meta property="og:site_name" content="日本を知る。">
<meta property="og:title" content="{TITLE}｜HIV/エイズ・梅毒・はしか・結核の今">
<meta property="og:description" content="{DESC}">
<meta property="og:url" content="{SITE_URL}">
<meta property="og:image" content="{SITE_URL}ogp.png">
<meta name="twitter:card" content="summary_large_image">
<link rel="icon" href="favicon.svg" type="image/svg+xml">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700&family=Noto+Serif+JP:wght@400;600&display=swap">
<style>
/* 見た目の設定：ページ本体が使う色・文字・すき間の名前をここで決める（明るい表示と暗い表示） */
:root {{
  color-scheme: light;
  --font-anthropic-sans: "Noto Sans JP", "Hiragino Sans", "Yu Gothic UI", Meiryo, sans-serif;
  --font-anthropic-serif: "Noto Serif JP", "Hiragino Mincho ProN", "Yu Mincho", serif;
  --color-bg: #f6f7f5; --color-panel: #ffffff; --color-fg: #1c2321; --color-fg-muted: #59645f;
  --color-border-line: #dde3df; --color-ok: #23704a; --color-bad: #b13a28; --color-warn: #8f5f00;
  --series-1: #2f6cb3; --series-2: #c2532d; --series-3: #a87a12; --series-4: #3b8a56;
  --cds-chart-grid: #e9ecea; --cds-chart-axis: #b4bdb8; --cds-chart-reference: #6f7a75; --cds-chart-muted: #c8cfcb;
  --cds-surface-popover: #ffffff; --cds-border: #d3dad6; --cds-shadow-popover: 0 6px 18px rgba(20, 30, 25, .14); --cds-text-secondary: #59645f;
  --cds-font-size-caption: 0.8125rem; --cds-font-size-body: 0.9375rem; --cds-font-size-heading: 1.0625rem; --cds-font-size-title: 1.5rem;
  --cds-font-weight-medium: 600;
  --cds-gap-xs: 4px; --cds-gap-sm: 8px; --cds-gap-md: 16px; --cds-gap-lg: 24px; --cds-gap-xl: 48px;
  --cds-pad-xs: 4px; --cds-pad-sm: 8px; --cds-pad-md: 12px; --cds-pad-lg: 18px;
  --cds-radius: 10px; --cds-dur-slow: .3s;
}}
@media (prefers-color-scheme: dark) {{
  :root {{
    color-scheme: dark;
    --color-bg: #111614; --color-panel: #19201d; --color-fg: #e5ebe8; --color-fg-muted: #a1aca7;
    --color-border-line: #2b3430; --color-ok: #6cc391; --color-bad: #f08c74; --color-warn: #e0b04f;
    --series-1: #70a6e8; --series-2: #ee8d68; --series-3: #ddb34f; --series-4: #75c48f;
    --cds-chart-grid: #252d2a; --cds-chart-axis: #4a5551; --cds-chart-reference: #95a09b; --cds-chart-muted: #3d4743;
    --cds-surface-popover: #212926; --cds-border: #38423e; --cds-shadow-popover: 0 6px 18px rgba(0, 0, 0, .45); --cds-text-secondary: #a1aca7;
  }}
}}
html {{ -webkit-text-size-adjust: 100%; }}
body {{ margin: 0; background: var(--color-bg); color: var(--color-fg); font-family: var(--font-anthropic-sans); font-size: var(--cds-font-size-body); }}
.site-bar {{ border-bottom: 1px solid var(--color-border-line); background: var(--color-panel); }}
.site-bar-in, #dash-root, .site-foot-in {{ max-width: 74rem; margin-inline: auto; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }}
.site-bar-in {{ display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 8px 16px; padding-block: 10px; font-size: var(--cds-font-size-caption); }}
.site-bar a, .site-foot a {{ color: var(--color-fg); text-underline-offset: 3px; }}
.site-home {{ font-weight: 700; text-decoration: none; }}
.site-home span {{ color: var(--color-fg-muted); font-weight: 400; margin-left: .4em; }}
#dash-root {{ padding-block: 32px 48px; }}
.site-foot {{ border-top: 1px solid var(--color-border-line); }}
.site-foot-in {{ padding-block: 20px 32px; color: var(--color-fg-muted); font-size: var(--cds-font-size-caption); line-height: 1.8; }}
.site-foot p {{ margin: 0; }}
.dash-skeleton {{ color: transparent; background: var(--cds-chart-grid); border-radius: 4px; }}
a:focus-visible {{ outline: 2px solid var(--color-fg); outline-offset: 2px; }}
</style>
<script src="https://cdnjs.cloudflare.com/ajax/libs/d3/7.9.0/d3.min.js"></script>
<script src="data.js?v={ver}"></script>
<script>
// ページ本体（claude.ai 版と共通）にデータを渡す小さな仕組み
(function () {{
  var D = window.KJ_DATA, calcs = {{}}, cbs = [];
  function run() {{ cbs.forEach(function (f) {{ f(); }}); }}
  window.dash = {{
    colors: ["var(--series-1)", "var(--series-2)", "var(--series-3)", "var(--series-4)"],
    calc: function (id, o) {{ calcs[id] = o; }},
    data: function (id) {{
      if (D && D[id]) return {{ status: "ok", data: D[id] }};
      var c = calcs[id];
      if (D && c) return {{ status: "ok", data: c.fn.apply(null, c.inputs.map(function (i) {{ return D[i]; }})) }};
      return {{ status: "error", data: [] }};
    }},
    onData: function (f) {{ cbs.push(f); f(); }},
    setLink: function (id) {{ try {{ history.replaceState(null, "", "#" + id); }} catch (e) {{}} }}
  }};
  var lastW = window.innerWidth, timer = null;
  window.addEventListener("resize", function () {{
    if (window.innerWidth === lastW) return;
    lastW = window.innerWidth;
    clearTimeout(timer);
    timer = setTimeout(run, 150);
  }});
  // 文字の形が読み込まれたら、グラフの文字幅を測り直して描き直す
  if (document.fonts && document.fonts.ready) document.fonts.ready.then(run);
}})();
</script>
</head>
<body>
<header class="site-bar"><div class="site-bar-in">
  <a class="site-home" href="https://yumaki5555.github.io/">日本を知る。<span>トップへ戻る</span></a>
  <a href="{XLSX_PUB}" download>Excelの表をダウンロード（.xlsx）</a>
</div></header>
<main id="dash-root">
"""
tail = f"""
</main>
<footer class="site-foot"><div class="site-foot-in">
  <p>出典：厚生労働省、国立健康危機管理研究機構（JIHS）などの公式資料。HIV/エイズの世界の動向は国連合同エイズ計画（UNAIDS）・国際移住機関（IOM）の公式資料。くわしくは「出典一覧」をご覧ください。</p>
  <p>同じ内容のExcelの表：<a href="{XLSX_PUB}" download>{XLSX_PUB}</a>（4つの比較、病気ごとの詳細、HIV/エイズの世界の動向、見るときの注意、出典一覧のシート入り）</p>
  <p><a href="https://yumaki5555.github.io/">「日本を知る。」のトップへ</a></p>
</div></footer>
</body>
</html>
"""
open(os.path.join(DOCS, "index.html"), "w", encoding="utf-8").write(head + fragment + tail)
open(os.path.join(DOCS, "favicon.svg"), "w", encoding="utf-8").write(
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="7" fill="#2f6cb3"/>'
    '<path d="M5 17h5l3-7 5 13 3-6h6" fill="none" stroke="#fff" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg>\n')
for f in sorted(os.listdir(DOCS)):
    print(f, os.path.getsize(os.path.join(DOCS, f)))
