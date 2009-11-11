set tmpCopy=%1
if "%tmpCopy%" == "" set tmpCopy=a.wav
echo tmp is %tmpCopy%
copy %tmpCopy% "%tmpCopy%_1.wav"
copy %tmpCopy% "%tmpCopy%_2.wav"
copy %tmpCopy% "%tmpCopy%_3.wav"
copy %tmpCopy% "%tmpCopy%_4.wav"
pause