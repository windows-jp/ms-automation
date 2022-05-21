$work_dir = "C:\Users\opkd0\開発\ms-automation\webpostman\"
$csv_file = "sample.csv"
$email_list = "email.csv"

# csv_file's columns
$send_column = "送信時刻"
$title_column = "件名"
$sender_id = "送信者ID" 
$recipient_email = "受信メールアドレス"

#email_list's columns
$id = "アカウントID"
$email = "メールアドレス"

$out_dir = "out"
$temp_dir = "temp"
$duplicate_dir = "duplicate"
$converted_dir = "converted"
$out_file = $work_dir + $out_dir + '/' + "out.txt"


Set-Location $work_dir

foreach ($directory in $out_dir, $temp_dir, $duplicate_dir, $converted_dir) {
    if (Test-Path $directory) {
        Remove-Item -Recurse $directory
    }
    New-Item -Path $work_dir -Name $directory -ItemType "directory" | Out-Null
}

$group_identifier = Import-Csv $csv_file -Encoding default | Select-Object $send_column, $title_column -uniq

foreach ($gi in $group_identifier) {Import-Csv $csv_file -Encoding default | Where-Object {($_.$send_column -eq $gi.$send_column) -and ($_.$title_column -eq $gi.$title_column)} | Select-Object $sender_id, $recipient_email | Export-Csv -LiteralPath ($temp_dir + '\' +  $gi.$send_column + '_' + $gi.$title_column + '.csv').replace(' ', '_').replace('/','').replace(':','') -NoTypeInformation -Encoding default}

foreach ($temp_csv in (Get-ChildItem $temp_dir)) {
    
    # convert id to email
    $id_group_string = (Import-Csv -LiteralPath $temp_csv.FullName -Encoding default | Select-Object $sender_id -Unique).$sender_id
    $webpostman_id, $webpostman_group = $id_group_string -split '/'
    $webpostman_email = (Import-Csv $email_list -Encoding default | Where-Object {$_.$id -eq $webpostman_id}).$email

    # remove header & sort & uniq & save
    $lines = (Get-Content -LiteralPath $temp_csv.FullName).Replace($id_group_string,$webpostman_email)
    $converted_file_name = $converted_dir + '\' + ($temp_csv.Name).Replace($temp_csv.Extension,'.txt')
    $converted_content = $lines[1, ($lines.Length - 1)] -Split ',' | Sort-Object -Unique
    $converted_content | Out-File -LiteralPath $converted_file_name
}

# remove duplicated files
foreach ($temp_txt in (Get-ChildItem $converted_dir)) {
    if (Test-Path $temp_txt.FullName) {
        foreach ($compared_txt in (Get-ChildItem $converted_dir -Exclude $temp_txt.Name)) {
                $difference_is_detected = Compare-Object (Get-Content -LiteralPath $temp_txt.FullName) (Get-Content -LiteralPath $compared_txt.FullName);
                if (-Not $difference_is_detected) {
                    Move-Item ($compared_txt.FullName) -Destination ($duplicate_dir)
                }
        }
    }
}

foreach ($uniq_txt in (Get-ChildItem $converted_dir)){
    "-----------------" | Add-Content $out_file
    Get-Content -LiteralPath $uniq_txt.FullName | Add-Content $out_file

}
