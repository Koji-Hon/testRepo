# リポジトリ共通の指示（Copilot）

- このリポジトリは COBOL → Java 移行プロジェクトである。
- `design-import/` 配下は Excelマクロで設計書に展開する取込ファイル。編集するときは
  `.github/instructions/design-import-format.instructions.md` の書式を必ず守る。
- 取込ファイルを作るときは、プロンプト `/design-reverse`（1本）または `/design-reverse-batch`（一括）を使う。
- 処理定義だけが粗い・省略されている場合は `/design-reverse-proc` で PROC だけ作り直す。
- Excelの「取込ファイル_ソース突合」で出た抜けは `/design-reverse-fix` に依頼文を貼って追記する。
- 処理定義は1メソッド＝1ブロック。入口 P100、呼び先 P101, P102 …、入口から未使用の public は P901 …。
- 取込ファイルを書くときは要約・省略をしない。長くなる場合はブロック単位でファイルに追記して書き切る。
- ソースから判断できない内容は推測で埋めず「要確認」と書く。
