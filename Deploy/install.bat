@echo off


:MAIN

cls
echo This will Setup IRC Dominator on your system
echo Press Ctrl+C to abort, or
pause

echo.
echo Copying files...
copy nslock15vb5.ocx %WinDir%\system
copy Splitter.ocx %WinDir%\system
copy RICHTX32.OCX %WinDir%\system
copy mswinsck.ocx %WinDir%\system

echo.
echo Registering files...
cd %WinDir%\system
regsvr32/s %WinDir%\system\nslock15vb5.ocx
regsvr32/s %WinDir%\system\Splitter.ocx 
regsvr32/s %WinDir%\system\RICHTX32.OCX 
regsvr32/s %WinDir%\system\mswinsck.ocx

echo.
echo IRC Dominator was just setup on your system!
echo.

goto END

:ERRO


goto END

:END
