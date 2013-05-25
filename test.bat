@echo off
set ERRORLEVEL=0

echo wget.vbs unit tests:

echo TEST #1: Download a file
cscript wget.vbs "http://httpbin.org/robots.txt" >NUL
call :basic_file_tests "1" "robots.txt"
if %ERRORLEVEL% NEQ 0 goto :eof
echo TEST #1: PASS
del /F /Q "robots.txt" 1>NUL 2>&1

echo TEST #2: Download a file to a specific filename
cscript wget.vbs "http://httpbin.org/user-agent" "user-agent.json" >NUL
call :basic_file_tests "2" "user-agent.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "2" "user-agent.json" "user-agent"
if %ERRORLEVEL% NEQ 0 goto :eof
echo TEST #2: PASS
del /F /Q "user-agent.json" 1>NUL 2>&1

echo TEST #3: Support response headers
cscript wget.vbs "http://httpbin.org/response-headers?wget=vbs" "response-headers.json" >NUL
call :basic_file_tests "3" "response-headers.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "3" "response-headers.json" "wget"
if %ERRORLEVEL% NEQ 0 goto :eof
echo TEST #3: PASS
del /F /Q "response-headers.json" 1>NUL 2>&1

echo TEST #4: Support no-clobber
cscript wget.vbs "http://httpbin.org/ip" "httpbin.json" >NUL
call :basic_file_tests "4" "httpbin.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "4" "httpbin.json" "origin"
if %ERRORLEVEL% NEQ 0 goto :eof
cscript wget.vbs "http://httpbin.org/user-agent" "httpbin.json" /NC >NUL
call :basic_file_tests "4" "httpbin.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "4" "httpbin.json" "user-agent" "assert-not"
if %ERRORLEVEL% NEQ 0 goto :eof
echo TEST #4: PASS
del /F /Q "httpbin.json" 1>NUL 2>&1

echo TEST #5: Support supression of prompting
cscript wget.vbs "http://httpbin.org/ip" "httpbin.json" >NUL
call :basic_file_tests "5" "httpbin.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "5" "httpbin.json" "origin"
if %ERRORLEVEL% NEQ 0 goto :eof
cscript wget.vbs "http://httpbin.org/user-agent" "httpbin.json" /Y >NUL
call :basic_file_tests "5" "httpbin.json"
if %ERRORLEVEL% NEQ 0 goto :eof
call :text_file_tests "5" "httpbin.json" "user-agent"
if %ERRORLEVEL% NEQ 0 goto :eof
echo TEST #5: PASS
del /F /Q "httpbin.json" 1>NUL 2>&1

echo ALL TESTS PASS
goto :eof

:basic_file_tests
if not exist "%~2" (
    echo TEST #%1: FAIL: File "%~2" not created
    set ERRORLEVEL=1
    goto :eof
)

for /f "usebackq tokens=4 delims= " %%f in (`dir ^| find "%~2"`) do (
    if %%f LEQ 0 (
        echo TEST #%1: FAIL: File "%~2" is zero-length
        set ERRORLEVEL=1
        goto :eof
    )
)
goto :eof

:text_file_tests
set OUTPUT=
for /f "usebackq tokens=1 delims=: " %%f in (`type "%~2" ^| find "%~3"`) do (
    set OUTPUT=%%~f
)

if "%~4" == "assert-not" (
    if not "%OUTPUT%" == "" (
        echo TEST #%~1: FAIL: File "%~2" is not supposed to contain "%~3"
        echo                  Found: "%OUTPUT%"
        set ERRORLEVEL=1
        goto :eof
    )
) else (
    if "%OUTPUT%" == "" (
        echo TEST #%~1: FAIL: File "%~2" does not contain "%~3"
        set ERRORLEVEL=1
        goto :eof
    )
)
goto :eof

:eof
