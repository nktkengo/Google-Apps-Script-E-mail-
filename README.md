# 各自に異なる添付ファイルを送付する。
Googel Apps Script
## 注意
`Browser.msgBox`は非推奨メソッドの為、`ui.alert`を使うべきだそうだ。
## 作成経緯

コロナ禍により、remote授業となった。<br>
これに伴い、周りのprofessor方の資料配布の手間が増え、グチグチ言っておられたのでGASを用いて作ってみた。<br>

## 機能
* spread上のデータを基にメールを送信する。
* 各自に異なる添付ファイルを送信できる。<br>
  -pdf、document、spreadsheet、pngは確認済み。<br>
  注意点；添付ファイルはすべてpdf化して送信される。また、scriptの送信はできない。
* 空欄チェック機能がある。<br>
  -空欄のままで送信することも可能である
* 返信を拒否することも可能。


* 300件程を一気に送信する用途であるため、７件まではチェック機能を付けていない。

### 参考文献
- https://tonari-it.com/gas-mail-magazine/
