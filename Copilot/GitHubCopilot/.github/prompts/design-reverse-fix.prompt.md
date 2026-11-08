---
mode: agent
description: ソース突合（Excelマクロ）で出た抜け・誤りだけを取込ファイルに追記・修正する
tools: ['codebase', 'search', 'editFiles']
---
# 突合結果の反映（抜けだけ追記）

- 対象ファイル: ${input:file:取込ファイル（例 design-import/SPN40190）}
- 起点クラス: ${input:target:Serviceクラス名（例 SPN40190）}
- 突合結果: ${input:issues:Excelの「突合結果」シートの依頼文を貼り付け}

## 手順
1. 突合結果に挙がった項目だけを対応する。挙がっていない既存ブロックは変更しない。
   - メソッド抜け … そのメソッドの PROC ブロックを追加（処理Noは呼出順に合わせて P1xx、未使用 public は P9xx）。
     呼出元の PROC に「～する（P1xx参照）」の行が無ければ追加する。
   - SQL抜け(処理) … 呼出元メソッドの PROC の呼出位置に、説明T行＋SQL行を追加する。
   - SQL抜け(アクセス) … そのSQLの ACCESS ブロックを追加する（@SQLID 必須）。
   - 編集定義抜け … 参照されている @編集名 の EDIT ブロックを追加する。
   - まとめ書き … 1行1SQL・1ブロック1SQL に分割する。
   - SQL名称 … Javadoc/コメントから日本語名を付ける。
   - メソッド不明 / SQL不明 … ソースを確認して名前を正す。
2. 書き方は [取込フォーマット](../instructions/design-import-format.instructions.md) に従う。
3. 終わったら、対応した件数と、対応できなかった項目（理由付き）をチャットに報告する。
