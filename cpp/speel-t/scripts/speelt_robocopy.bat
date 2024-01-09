@ROBOCOPY %*

@IF %ERRORLEVEL% EQU 16 echo Robocopy Result: ***FATAL ERROR*** & goto fatal
@IF %ERRORLEVEL% EQU 15 echo Robocopy Result: OKCOPY + FAIL + MISMATCHES + XTRA & goto fatal
@IF %ERRORLEVEL% EQU 14 echo Robocopy Result: FAIL + MISMATCHES + XTRA & goto fatal
@IF %ERRORLEVEL% EQU 13 echo Robocopy Result: OKCOPY + FAIL + MISMATCHES & goto fatal
@IF %ERRORLEVEL% EQU 12 echo Robocopy Result: FAIL + MISMATCHES& goto fatal
@IF %ERRORLEVEL% EQU 11 echo Robocopy Result: OKCOPY + FAIL + XTRA & goto fatal
@IF %ERRORLEVEL% EQU 10 echo Robocopy Result: FAIL + XTRA & goto fatal
@IF %ERRORLEVEL% EQU 9 echo Robocopy Result: OKCOPY + FAIL & goto fatal
@IF %ERRORLEVEL% EQU 8 echo Robocopy Result: FAIL & goto fatal
@IF %ERRORLEVEL% EQU 7 echo Robocopy Result: OKCOPY + MISMATCHES + XTRA & goto nonfatal
@IF %ERRORLEVEL% EQU 6 echo Robocopy Result: MISMATCHES + XTRA & goto nonfatal
@IF %ERRORLEVEL% EQU 5 echo Robocopy Result: OKCOPY + MISMATCHES & goto nonfatal
@IF %ERRORLEVEL% EQU 4 echo Robocopy Result: MISMATCHES & goto nonfatal
@IF %ERRORLEVEL% EQU 3 echo Robocopy Result: OKCOPY + XTRA & goto nonfatal
@IF %ERRORLEVEL% EQU 2 echo Robocopy Result: XTRA & goto nonfatal
@IF %ERRORLEVEL% EQU 1 echo Robocopy Result: OKCOPY & goto nonfatal
@IF %ERRORLEVEL% EQU 0 echo Robocopy Result: No Change & goto nonfatal

:fatal
@EXIT /B %ERRORLEVEL%

:nonfatal
@EXIT /B 0