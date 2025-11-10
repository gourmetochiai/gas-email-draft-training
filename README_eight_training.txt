# Eight連携トレーニング用 ダミーデータ一式

## ファイル
- eight_dummy_contacts.csv … 連絡先ダミーデータ（Name, Company, Title, Email, Tag, MetAt, Note）
- config_templates.csv … A〜Eカテゴリ別の件名・本文テンプレ（{Name}, {Company}, {MeetingLink}を差し込み）
- gas_create_drafts.gs … Gmail下書き作成の最小GASサンプル

## 想定ルール（分類 A〜E）
- A: Titleに「代表」「部長」または Tagが「ホット」
- B: Titleに「課長」「マネージャ」または Tagが「既存」
- C: Titleに「担当」または Tagが「要フォロー」
- D: Tagが「展示会」
- E: その他

上から順に評価（条件に当てはまった時点でカテゴリ確定）。

## 手順（概要）
1. Googleスプレッドシートに `eight_dummy_contacts.csv` をインポート（シート名: data）
2. `config_templates.csv` を別シート（config）にインポート
3. Apps Scriptで `gas_create_drafts.gs` のコードを貼り付けて保存
4. メニュー「下書き作成」 → 未処理行の宛先・件名・本文を下書き生成
5. 生成結果とログを確認（logシート）

## 差し込み変数
- {Name}, {Company}, {MeetingLink}

MeetingLink はスクリプト内の定数で指定してください。
