@echo off
cd /d %~dp0

if defined ProgramFiles(x86) (
    REM We are in a x64 system
    C:\Windows\SysWoW64\wscript fone.vbs
) else (
    REM We are in a x86 system
    wscript fone.vbs
)