# Open Excel Readonly

Excelアプリに関連付けられたファイルを開く時のデフォルトの動作を編集可能または読取専用に変更するツール


## 読み取り専用とは

- 開いたExcelファイルを上書き保存できなくなるモード

- タスクバー上に「読み取り専用」と表示されていれば読み取り専用モード

- 保護モードと違って編集自体は可能で、別名を付ければ編集内容を保存することができる



## 使い方

- Excel起動モード設定.batを右クリックして「管理者として実行」を選択して起動する
  - ファイルの関連付けを変更するために管理者権限が必要なため

- コンソール上の表示にしたがって変更したいモードを選択する
  - 0 : 編集可能
  - 1 : 読取専用

- 設定が完了すると終了メッセージが表示され自動でウィンドウが閉じる



## Tips

### デフォルトの動作を読取専用にした場合の編集方法

- <kbd>Shift</kbd>を押しながらファイルを右クリックし、コンテキストメニューの中から「編集」を選択するとExcelを編集可能な状態で開くことができる
  - ファイルを選択している状態で<kbd>Shift</kbd>+<kbd>F10</kbd>を押すことでも同様のコンテキストメニューを表示できる


### 起動設定の初期化

- 「設定 > アプリ > アプリと機能」で表示されるアプリケーション一覧から「Microsoft 365 Apps」を選択して「変更」を実行し、クイック修復を行うとExcelの起動設定を初期設定に戻すことができる

- 同様の原理で、Excelを再インストールした時やアップデートしたときも設定が初期設定に戻る

