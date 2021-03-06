$work_dir = "C:\Users\opkd0\開発\ms-automation\webpostman\"
$csv_file = "JOIN.dat"
$email_list = "email.csv"

# csv_file's columns
$send_column = "送信日時"
$title_column = "件名"
$sender_id = "送信者アカウントID" 
$recipient_email = "受信者メールアドレス"

#email_list's columns
$id = "アカウントID"
$email = "メールアドレス"

$out_dir = "out"
$temp_dir = "temp"
$duplicate_dir = "duplicate"
$converted_dir = "converted"
$log_dir = "log"

$out_file = $work_dir + $out_dir + '/' + "out.txt"
$log_file = $log_dir + '/' + "log.txt"


Set-Location $work_dir

foreach ($directory in $out_dir, $temp_dir, $duplicate_dir, $converted_dir, $log_dir) {
    if (Test-Path $directory) {
        Remove-Item -Recurse $directory
    }
    New-Item -Path $work_dir -Name $directory -ItemType "directory" | Out-Null
}

$group_identifier = Import-Csv $csv_file -Encoding default | Select-Object $send_column, $title_column -uniq

foreach ($gi in $group_identifier) {Import-Csv $csv_file -Encoding default | Where-Object {($_.$send_column -eq $gi.$send_column) -and ($_.$title_column -eq $gi.$title_column)} | Select-Object $sender_id, $recipient_email | Export-Csv -LiteralPath ($temp_dir + '\' +  $gi.$send_column.Replace('-','').Replace(':','').Replace(' ','') + '_' + $gi.$title_column.Replace('/','').Replace(':','').Replace('?','').Replace('[','').Replace(']','') + '.csv') -NoTypeInformation -Encoding default}

foreach ($temp_csv in (Get-ChildItem $temp_dir)) {

    $converted_file_name = $converted_dir + '\' + ($temp_csv.Name).Replace($temp_csv.Extension,'.txt')

    foreach ($data in Import-Csv -LiteralPath $temp_csv.FullName -Encoding Default){
        if ($data.$sender_id -match '.*/.*') {
            $webpostman_id, $webpostman_group = $data.$sender_id -split '/'
            $data.$sender_id = (Import-Csv $email_list -Encoding default | Where-Object {$_.$id -eq $webpostman_id}).$email
        }
        if ($data.$recipient_email -match '.*/.*') {
            $webpostman_id, $webpostman_group = $data.$recipient_email -split '/'
            $data.$recipient_email = (Import-Csv $email_list -Encoding default | Where-Object {$_.$id -eq $webpostman_id}).$email
        }

        if ($data.$recipient_email -match '<.*>') {
            $start_index = $Matches.Values.IndexOf('<')
            $chara_count = $Matches.Values.IndexOf('>') - $start_index
            $data.$recipient_email = $Matches.Values.Substring($start_index + 1, $chara_count - 1)
        }

        $data.$sender_id | Add-Content -LiteralPath $converted_file_name -Encoding Default
        $data.$recipient_email | Add-Content -LiteralPath $converted_file_name -Encoding Default

        Get-Content -LiteralPath $converted_file_name -Encoding Default | Sort-Object -Unique | Set-Content -Encoding Default -LiteralPath $converted_file_name

    }
}

# remove duplicated files
foreach ($temp_txt in (Get-ChildItem $converted_dir)) {
    if (Test-Path $temp_txt.FullName) {
        foreach ($compared_txt in (Get-ChildItem $converted_dir -Exclude $temp_txt.Name)) {
                $difference_is_detected = Compare-Object (Get-Content -LiteralPath $temp_txt.FullName) (Get-Content -LiteralPath $compared_txt.FullName);
                if (-Not $difference_is_detected) {
                    Move-Item ($compared_txt.FullName) -Destination ($duplicate_dir)
                    Write-Output ((Get-Date -UFormat "%Y-%m-%d %H:%M:%S") + ' [Duplicated] ' + $temp_txt.Name  + ' <=> ' + $compared_txt.Name) | Add-Content $log_file -Encoding Default
                }
        }
    }
}

foreach ($uniq_txt in (Get-ChildItem $converted_dir)){
    "-----------------" | Add-Content $out_file
    Get-Content -LiteralPath $uniq_txt.FullName | Add-Content $out_file
}