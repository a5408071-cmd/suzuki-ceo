# 資料作成のルール

## 納品ファイル
- 資料を作り直した（再ビルドした）ときは、**PDFだけ**を送ること。PPTXは送らない。
- PPTXが必要なときは鈴木から明示的に依頼があったときだけ送る。
- 完成したPDFは必ず `/home/user/suzuki-ceo/downloads/` にコピーしておく。ここが鈴木の取り出し口。プロジェクトのディレクトリに置いたままにしない。
- 作り直すたびに downloads/ の同名ファイルを上書きする（古い版を残さない）。

## ビルド手順
各ディレクトリで:
1. `node build.js`
2. `python3.13 ../fix-pptx.py <file>.pptx`  ← 必須。やらないとPowerPointが修復ダイアログを出す
3. PDF変換:
   `PYTHONPATH=/mnt/skills/public/pptx/scripts python3.13 -c "from office.soffice import run_soffice; run_soffice(['--headless','--convert-to','pdf','--outdir','.','x.pptx'])"`
4. PNGに書き出して目視チェック（重なり・はみ出しを毎回確認する）

## 注意
- `python3` は 3.11。pymupdf / openpyxl は **python3.13** にしか入っていない。
- メール本文は太字（markdown）を使わない。プレーンテキストで書く。
- 直営/FC比率、店舗ごとの売上・利益、どの店舗が直営かFCかが分かる情報は資料に出さない。
- 印西牧の原・新守山は別会社。社名は出さず、杯数だけを使う。
