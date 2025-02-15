# dotnet_npoi_text_function_investigate

## 概要
* NPOI で数式の TEXT 関数を評価した際の挙動を調査する
* 特に知りたいのは、"yyyy-mm-dd" 形式の文字列を "yyyy/mm/dd" でフォーマットした際にエラーになるパターンがあるか

## 結果
* エラーになるパターンがある
* "yyyy-mm-dd" 形式の文字列を直接渡した場合と、セルに格納して参照させた場合で結果が異なり、セル参照ではエラーとなる
  * `TEXT("2012-12-31","yyyy/mm/dd")` → 2012-12-31
  * `TEXT(A1,"yyyy/mm/dd")` → #VALUE! ※A1 に "2012-12-31" を格納
* DATEVALUE 関数で明示的に変換すると正常に動作する
  * `TEXT(DATEVALUE(A1),"yyyy/mm/dd")` → 2012/12/31

## 所感
* TEXT 関数の処理が NPOI と Excel で異なっている
* セルを渡された際、Excel は当該セルが日付に変換可能かまで考慮するが NPOI は考慮していない
* また、同じ "2012-12-31" でも直接渡した場合とセル経由で挙動が異なることから、これは不具合と思われる

※POI の TEXT 関数は不具合が多いらしいのでその一つ？

## 詳細
```
dotnet new console
dotnet add package NPOI --version 2.7.1
```

https://www.nuget.org/packages/NPOI/

```
$ dotnet run
#1 日付型の値を直接渡した場合 → 変換され yyyy/mm/dd 形式で出力
String: 2024/09/27
#2 数値型の値を直接渡した場合 → 変換され yyyy/mm/dd 形式で出力
String: 1900/05/02
#3 文字列型(日付)の値を直接渡した場合 → 変換されず元の値を出力
String: 2012-12-31
#4 文字列型(数字)の値を直接渡した場合 → 変換されず元の値を出力
String: 123
#5 文字列型(文字)の値を直接渡した場合 → 変換されず元の値を出力
String: abc
#6 日付型の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力
String: 2024/09/27
#7 数値型の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力
String: 1900/05/02
#8 文字列型(日付)の値をセル経由で渡した場合(ハイフン) → #VALUE! エラーが発生
Error: 15: #VALUE!
#9 文字列型(数字)の値をセル経由で渡した場合 → 変換され yyyy/mm/dd 形式で出力
String: 1900/05/02
#10 文字列型(文字)の値をセル経由で渡した場合 → #VALUE! エラーが発生
Error: 15: #VALUE!
#11 文字列型(日付)の値をセル経由で渡した場合(スラッシュ) → #VALUE! エラーが発生
Error: 15: #VALUE!
#12 文字列型(日付)の値をセル経由で渡した場合(スラッシュ+日付書式) → #VALUE! エラーが発生
Error: 15: #VALUE!
#13 文字列型(日付)の値をセル経由で渡した場合(ハイフン+日付書式) → #VALUE! エラーが発生
Error: 15: #VALUE!
#14 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(ハイフン) → 変換され yyyy/mm/dd 形式で出力
String: 2012/12/31
#15 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(スラッシュ) → 変換され yyyy/mm/dd 形式で出力
String: 2012/12/31
#16 文字列型(日付)の値をセル経由且つ日付に変換して渡した場合(スラッシュ+日付書式) → 変換され yyyy/mm/dd 形式で出力
String: 2012/12/31
#17 文字列型(日付)の値をセル経由で渡した場合(ハイフン+日付書式) → 変換され yyyy/mm/dd 形式で出力
String: 2012/12/31
```