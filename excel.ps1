### 準備 ###
# Excelオブジェクトを取得 (Excelを起動)
$excel = New-Object -ComObject Excel.Application

# Excelを可視化
$excel.Visible = $True

# 既存のExcelファイルを開く (パスワードなし) #フルパスで指定しないとエラーになるので注意
$fullname = (Get-ChildItem test.xlsx).FullName
$workbook = $excel.Workbooks.Open($fullname)  # -> Workbook

# 既存のExcelファイルを開く (パスワードあり) #フルパスで指定しないとエラーになるので注意
$fullname = (Get-ChildItem test_pass.xlsx).FullName
$password = "password"
$workbook = $excel.Workbooks.Open($fullname, [type]::Missing, [type]::Missing, [type]::Missing, $password) # -> Workbook

# Sheetを取得
$worksheet = $workbook.Worksheets(1) # -> WorkSheet


### 書き込み ###
# 指定範囲のデータを削除
$worksheet.Range("C5:E8").Clear()

# クリップボードのデータをセルに張り付け
$range = $worksheet.Range("C5")
$worksheet.Paste($range)

# 上書き保存
$workbook.Save()

### 後処理 ###
# Excelを閉じる
$excel.Quit()

# プロセス開放
$excel = $null 