@echo off
setlocal enabledelayedexpansion

set config=.\ip_setting.conf

for /f "usebackq tokens=1-4" %%a in (`netsh interface show interface ^| findstr "Ú‘±"`) do (
    set interface=%%d
)

if not exist !config! (
    echo create config
    netsh interface ipv4 show config !interface! > !config!
)

netsh interface ipv4 set address !interface! dhcp
netsh interface ipv4 set dnsservers !interface! dhcp

if !errorlevel!==0 (
    netsh interface ipv4 show config !interface!
)

pause