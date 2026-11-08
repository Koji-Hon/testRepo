---
mode: agent
description: フォルダ内の複数プログラムをまとめて設計書取込ファイルにする
tools: ['codebase', 'search', 'editFiles']
---
# 設計書リバース（一括）

- 対象フォルダ: ${input:folder:Serviceクラスのあるフォルダ（例 src/main/java/jp/co/xxx/tabs/service）}
- 対象フォルダ内の Service クラスを列挙し、1クラス＝1プログラムとして順に処理する。
- 各プログラムについて [design-reverse](./design-reverse.prompt.md) と同じ手順・同じ出力ルールで
  `design-import/<処理ID>` を作成する（既にあるものはスキップ）。
- コンテキストが大きくなりすぎないよう、1プログラム終わるごとに読み込んだ内容を引きずらないこと。
- 最後に一覧（処理ID / 作成・スキップ / 要確認件数）をチャットに表で報告する。
