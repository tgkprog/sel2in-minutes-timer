g:
cd\
set rootMI=G:\prog\vb\minsTmr\g\h\miti4
cd "%rootMi%\packBinOnly\bin\"
del /s/q "%rootMi%\packBinOnly\bin\*.*"
rmdir /s/q "%rootMi%\packBinOnly\bin\"
mkdir "%rootMi%\packBinOnly\bin\res"
cd  %rootMi%\packHelpers\binOnly

copy /y %rootMi%\src\res\help*.* %rootMi%\bin\res
rem /h if you want hidden folders too - we don't
xcopy /c /y/e/v/r/k/q %rootMi%\bin\*
rem del /q/f/s %rootMi%\packBinOnly\bin\res\.svn\*
rem for /r  "%rootMi%\packBinOnly\bin\res\.svn" %a  in (.)  do rmdir %a

rem not req as hidden folder will not be copied but is good to clean up
rmdir /s/q "%rootMi%\packBinOnly\bin\res\.svn"
rmdir /s/q "%rootMi%\packBinOnly\bin\.svn"

del /f ..\MinsTimer.zip
rem -pass=20 for password
"C:\Program Files\7-Zip\7z.exe" a -tzip -mx9    ..\MinsTimer.zip @%rootMi%\packBinOnly\pack-list
rem will run sed now %rootMi%\Package\single_support\"%rootMi%\Package\single_support\Minutes_Timer_Installer.sed
pause
cd %rootMi%\Package\single_support\
"%rootMi%\Package\single_support\Minutes_Timer_Installer.sed