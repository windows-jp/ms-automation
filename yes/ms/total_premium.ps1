foreach ($file in Get-ChildItem C:\Users\opkd0\Desktop\MS -Recurse -Filter "MSFJ*.txt"){
    $null = $file.Name -match '.*_(?<date>\d{8})_.*'
    $date_str = [string]::Format("{0}/{1}/{2}", $matches.date.Substring(0,4), $matches.date.Substring(4,2), $matches.date.Substring(6,2))

    $content = Get-Content $file.FullName

    $total_premium = 0

    foreach ($line in $content[1..($content.Length - 2)]){
        $premium = [int]$line.Substring(728,4)
        $total_premium += $premium
    }

    Write-Host "${date_str},${total_premium}"
    
}