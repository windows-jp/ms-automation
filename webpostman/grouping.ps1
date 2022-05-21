$work_dir = "C:\Users\opkd0\開発\ms-automation\webpostman\"
$csv_file = "sample.csv"

$send_column = "送信時刻"
$title_column = "件名"
$sender_id = "送信者ID" 
$recipient_email = "受信メールアドレス"

$out_dir = "out"
$temp_dir = "temp"
$duplicate_dir = "duplicate"
$out_file = "out.txt"

Set-Location $work_dir

foreach ($directory in $out_dir, $temp_dir, $duplicate_dir) {
    if (Test-Path $directory) {
        Remove-Item -Recurse $directory
    }
    New-Item -Path $work_dir -Name $directory -ItemType "directory"
}

$group_identifier = Import-Csv $csv_file -Encoding default | Select-Object $send_column, $title_column -uniq

foreach ($gi in $group_identifier) {Import-Csv $csv_file -Encoding default | Where-Object {($_.$send_column -eq $gi.$send_column) -and ($_.$title_column -eq $gi.$title_column)} | Select-Object $sender_id, $recipient_email | Export-Csv -LiteralPath ($temp_dir + '\' +  $gi.$send_column + '_' + $gi.$title_column + '.csv').replace(' ', '_').replace('/','').replace(':','') -NoTypeInformation -Encoding default}

foreach ($temp_csv in (Get-ChildItem $temp_dir)) {
    if (Test-Path $temp_csv.FullName) {
        foreach ($compared_csv in (Get-ChildItem $temp_dir -Exclude $temp_csv.Name)) {
                $difference_is_detected = Compare-Object (Get-Content -LiteralPath $temp_csv.FullName) (Get-Content -LiteralPath $compared_csv.FullName);
                if (-Not $difference_is_detected) {
                    Move-Item ($compared_csv.FullName) -Destination ($duplicate_dir)
                }
        }
    }
}

foreach ($uniq_csv in (Get-ChildItem $temp_dir)){Get-Content -LiteralPath $uniq_csv.FullName | Add-Content ($work_dir + $out_dir + '/' + $out_file) }
