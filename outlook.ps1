# OutlookAPIを呼び出す (= Outlookオブジェクトを取得)
# 参考: https://docs.microsoft.com/en-us/archive/msdn-magazine/2013/march/powershell-managing-an-outlook-mailbox-with-powershell
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -comobject Outlook.Application
$namespace = $outlook.GetNameSpace("MAPI")

# ルートフォルダに接続 ([注意!]ここでは最初に登録されたメールアカウントのみを対象とする)
$root = $namespace.Folders.Item(1) # -> Folder

# 受信トレイに接続
$inbox = $root.Folders('受信トレイ') # -> Folder

# タイトルでフィルター
$regex = ".*【1日1フレーズ！生英語】.*"
$result = $inbox.items | Where-Object {$_.Subject -match $regex} # -> Item

# 本文でフィルター
$regex = ".*type IT.*"
$result = $inbox.items | Where-Object {$_.Body -match $regex} # -> Item

# 差出人でフィルター (宛先: To, )
$regex = ".*@hapaeikaiwa.com"
$result = $inbox.items | Where-Object {$_.SenderEmailAddress -match $regex} # -> Item

# 宛先でフィルター
$regex = ".*@gmail.com"
$result = $inbox.items | Where-Object {$_.To -match $regex} # -> Item

# 日付でフィルター
$from = "2021/07/19 06:26:43"
$to = "2021/08/01 00:00:00"
$result = $inbox.items | Where-Object {$_.ReceivedTime -gt $from} | Where-Object {$_.ReceivedTime -lt $to}

# カラムを抽出
$result = $result | Select-Object ReceivedTime, Subject, SenderEmailAddress, To, CC, Body

# 結果をCsvに出力
$result | Export-Csv -NoType C:\temp\ps.csv -Encoding Default

# Csvを開く
Invoke-Item c:\Temp\ps.csv