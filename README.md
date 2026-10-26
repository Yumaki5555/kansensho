# 日本の感染症を知る

HIV/エイズ・梅毒・はしか・結核の状況を、国の公式資料（厚生労働省・国立健康危機管理研究機構など）だけでまとめたページです。
HIV/エイズには「世界の動向」（UNAIDSの推計、日本と協力覚書を結んだ国、移住者とHIV、入国前の健康診断とIOM）もあります。こちらは国連の機関（UNAIDS・IOM）と出入国在留管理庁・外国人技能実習機構の公式資料を使っています。

- 公開先：https://yumaki5555.github.io/kansensho/ （リポジトリ：Yumaki5555/kansensho、main ブランチの `docs` フォルダを公開）
- 入口サイト「日本を知る。」の「調べる・比較する」に掲載
- 作成日：2026年10月26日（この日までに国が公表した資料をもとにしています）

## フォルダの役割

| 場所 | 役割 |
|---|---|
| `docs/` | **公開されるページ**（index.html、data.js、Excel、サムネイル、アイコン） |
| `data/research/` | 病気ごとに調べた結果（数字・出典URL・注意点）。ここが元データ。`world_hiv.json` はHIV/エイズの世界の動向、`kouhi.json` は公費で受けられる支援（検査・予防・治療。比較表の目安は「公費制度あり／条件付きで公費制度あり／対象となる公費制度を確認できず」の3つ） |
| `data/out/` | 元データを、Excelとページ共通の形にまとめたもの |
| `build/page_fragment.html` | ページ本体（グラフ・表・文章の作り）。claude.ai 版のページと共通の作り（`dash` という小さな仕組みでデータを受け取る） |
| `build/consolidate.py` | `data/research` → `data/out` にまとめる |
| `build/build_xlsx.py` | `data/out` → Excel を作る |
| `build/excel_recalc.ps1` | Excel で数式を計算し、エラーがないか確かめて保存する |
| `build/build_site.py` | `data/out` とページ本体 → `docs/` を作る |
| `ogp/ogp.html` | サムネイル（docs/ogp.png）の元 |

## ページの作り（2026年10月26日の改修後）

- 上の追従メニューは5つ：国内の感染症（4つの病気の比較・HIV/エイズ・梅毒・はしか・結核）／HIV世界の動向／公費で受けられる支援／数字の見方・注意事項／出典・資料。`#hiv` のように # の後ろでページが切りかわる（ブラウザの「戻る」も使える）
- どのページでも下に「本サイトのご利用にあたって」（6項目。③偏見・差別の防止と⑤冷静な理解と対策は最初から開いている）。フッターから「出典」「サイトの利用にあたって」「免責事項」へ移動できる
- 検索（右上、または「/」キー）：見出し・説明・表・国名・出典から探して、その場所へ移動する
- グラフの「画像で保存（PNG）」：1920×1080（16:9）で、タイトル・対象年・単位・出典を入れて保存する
- 掲載内容の日付（2026年10月26日時点）は `build/build_site.py` の `AS_OF` と、`page_fragment.html` のトップ・フッターの文にある。数字を新しくしたら一緒に直す

## 数字を新しくするとき

1. `data/research/` の該当する病気のファイルを、国の新しい資料で直す（出典URL・公表日も）
   - 世界の動向（`world_hiv.json`）：UNAIDSの推計は毎年7月ごろに新しくなる。AIDSinfo（https://aidsinfo.unaids.org/）の「Factsheets」で世界・地域・各国を選んで数字を読み取る。協力覚書の国は出入国在留管理庁（特定技能・育成就労）と外国人技能実習機構（技能実習）の一覧で、入国前結核スクリーニングは厚生労働省のページで確かめる
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
