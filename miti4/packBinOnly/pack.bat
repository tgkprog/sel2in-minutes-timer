d:
cd\
rmdir /s/q "D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\"
mkdir "D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\"
cd  D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\

rem /h if you want hidden folders too - we don't
xcopy /c /y/e/v/r/k/q D:\prog\vb\minuteAlarm\g\miti4\bin\*
rem del /q/f/s D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\res\.svn\*
rem for /r  "D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\res\.svn" %a  in (.)  do rmdir %a

rem not req as hidden folder will not be copied but is good to clean up
rmdir /s/q "D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\res\.svn"
rmdir /s/q "D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\bin\.svn"

del /f ..\MinsTimer.zip
D:\prgFiles\7zip\7z.exe a -tzip -mx9    ..\MinsTimer.zip @D:\prog\vb\minuteAlarm\g\miti4\packBinOnly\pack-list
rem will run sed now D:\prog\vb\minuteAlarm\g\miti4\Package\single_support\"D:\prog\vb\minuteAlarm\g\miti4\Package\single_support\Minutes_Timer_Installer.sed
pause

cd D:\prog\vb\minuteAlarm\g\miti4\Package\Support
copy  /y D:\prog\vb\minuteAlarm\g\miti4\bin\*
copy  /y D:\prog\vb\minuteAlarm\g\miti4\bin\res\*
call D:\prog\vb\minuteAlarm\g\miti4\Package\Support\MinutesTimer_Vb6_no_prompt.BAT
cd D:\prog\vb\minuteAlarm\g\miti4\Package\single_support\
"D:\prog\vb\minuteAlarm\g\miti4\Package\single_support\Minutes_Timer_Installer.sed