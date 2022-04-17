### ファイル作成/編集 ###
# 空のファイル作成
New-Item file.txt

# コマンドの結果をファイルに出力 (日本語OK)
Get-ChildItem C:\Users\opkd0\Desktop | Out-File desktop.txt


### CSV ### 
# コマンドの結果をcsvファイルに出力 (日本語OK: Unicode -> SJIS)
Get-ChildItem C:\Users\opkd0\Desktop | Export-Csv desktop.csv -NoType -Encoding Default

# csvファイルを読み込む (日本語OK: SJIS -> Unicode)
$csv = Import-Csv .\desktop.csv -Encoding Default

# カラム名を指定してcsvファイルに出力 (日本語OK: Unicode -> SJIS)
$csv | Select-Object Name, LastWriteTime | Export-Csv desktop_filenames.csv -NoType -Encoding Default

# ヘッダー情報がないcsvファイルを読み込む
$csv = Get-Content .\sample.csv | ConvertFrom-Csv -Header Name,Date


### 編集 ###
# 文字列を置換して上書き保存
(Get-Content '.\sample3.csv').Replace('"','').Replace('-','') | Set-Content .\sample3.csv

# 先頭行を削除して上書き保存
$content = Get-Content .\sample3.csv
$content[0] = $null # 先頭行 = 0
$content | Set-Content .\sample3.csv

# 重複を削除して上書き保存
Get-Content .\sample4.txt | Sort-Object | Get-Unique | Set-Content .\sample4.txt


### ファイル情報取得 ###
# 行数取得
(Get-Content '.\sample3.csv').Length

# grep (正規表現適用 & 大文字小文字を区別する)
$pattern = "^aM.*?:"
(Select-String $pattern .\sample3.csv -CaseSensitive -Encoding Default).Line # -> String

# grep -v (正規表現適用 & 大文字小文字を区別する)
$pattern = "^aM.*?:"
(Select-String $pattern .\sample3.csv -NotMatch -CaseSensitive -Encoding Default).Line # -> String

# 正規表現にマッチした部分だけを抽出 [sed s/(^aM.*?)/\1/gと同じ]
$pattern = "^aM.*?:"
(Select-String "^aM.*?:" .\sample3.csv -CaseSensitive -Encoding Default).Matches.Value


### 応用編 ###
# パターンにマッチした部分だけを置換して上書き保存
$pattern = "[0-9]{8}"
$before = (Select-String $pattern .\sample.sql).Matches.Value | Get-Unique
$after = '20220401'
(Get-Content .\sample.sql).Replace($before, $after) | Set-Content .\sample.sql