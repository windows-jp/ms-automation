@echo off
setlocal enabledelayedexpansion

set config=.\ip_setting.conf
set interface=""
set address=""
set mask=""
set gateway=""
set dns_01=""
set dns_02=""
set /a line_number=1

for /f "tokens=1-5" %%a in (!config!) do (
    if %%a==インターフェイスの構成 (
        set interface=%%b
        set interface=!interface:~1,-1!
    )
    if %%a==IP set address=%%c
    if %%a==サブネット (
        set mask=%%e
        set mask=!mask:~0,-1!
    )
    if %%b==ゲートウェイ: set gateway=%%c
    if %%b==DNS (
        set dns_01=%%d
        set /a dns_02_line_num=line_num+1
    )
    if !line_num!==!dns_02_line_num! set dns_02=%%a
    
    set /a line_num=line_num+1
)

netsh interface ipv4 set address !interface! static !address! !mask! !gateway! 
netsh interface ipv4 set dnsservers !interface! static !dns_01! validate=no
netsh interface ipv4 add dnsservers !interface! !dns_02! validate=no

if !errorlevel!==0 (
    netsh interface ipv4 show config !interface!
)

pause