IF EXIST C:\WINNT\WINNT.BMP GOTO WINNT

GOTO DONE

:WINNT
rem net use E: \\Bluegill\Intranet
copy E:\Intranet\HelpDesk\chat\chatsock.dll C:\winnt\system32
copy E:\Intranet\HelpDesk\chat\mschat.ocx C:\winnt\system32 
rem net use E: /d
regsvr32 mschat.ocx
regsvr32 chatsock.dll

:DONE