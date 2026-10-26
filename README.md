# 日本の感染症を知る

HIV/エイズ・梅毒・はしか・結核の状況を、国の公式資料（厚生労働省・国立健康危機管理研究機構など）だけでまとめたページです。

- 公開先：https://yumaki5555.github.io/kansensho/ （リポジトリ：Yumaki5555/kansensho、main ブランチの `docs` フォルダを公開）
- 入口サイト「日本を知る。」の「調べる・比較する」に掲載
- 作成日：2026年10月26日（この日までに国が公表した資料をもとにしています）

## フォルダの役割

| 場所 | 役割 |
|---|---|
| `docs/` | **公開されるページ**（index.html、data.js、Excel、サムネイル、アイコン） |
| `data/research/` | 病気ごとに調べた結果（数字・出典URL・注意点）。ここが元データ |
| `data/out/` | 元データを、Excelとページ共通の形にまとめたもの |
| `build/page_fragment.html` | ページ本体（グラフ・表・文章の作り）。claude.ai 版のページと共通 |
| `build/consolidate.py` | `data/research` → `data/out` にまとめる |
| `build/build_xlsx.py` | `data/out` → Excel を作る |
| `build/excel_recalc.ps1` | Excel で数式を計算し、エラーがないか確かめて保存する |
| `build/build_site.py` | `data/out` とページ本体 → `docs/` を作る |
| `ogp/ogp.html` | サムネイル（docs/ogp.png）の元 |

## 数字を新しくするとき

1. `data/research/` の該当する病気のファイルを、国の新しい資料で直す（出典URL・公表日も）
2. 感染症フォルダで次を順に実行する

```
python -I build/consolidate.py data/research data/out
python -I build/build_xlsx.py data/out 日本の感染症の状況_2026年10月.xlsx
powershell -File build/excel_recalc.ps1 -Path "（Excelの場所）"
python -I build/build_site.py
```

3. `docs/` の変更を GitHub に送る（数分で公開ページに反映されます）

## サムネイルを撮り直すとき

```
"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --headless=new --hide-scrollbars --window-size=1200,630 --screenshot=docs\ogp.png ogp\ogp.html
```
