@echo off
setlocal EnableExtensions
setlocal EnableDelayedExpansion

::: Usage: %~nx0 [-f] [-k] [-a <accountid>] [-d]
:::   -f: forces the retrieval of transactions (10 maximum), otherwise
:::       retrieves only the transactions newer than the previous retrieval date
:::   -a <accountid>: retrieves only the specified account. By default, all
:::       account known by boobank are retrieved.
:::   -k: keeps the produced intermediate .ofx file (in Downloads). By default, the files 
:::       are deleted.
:::   -c: clear cache. Forces the creation of intermediate files. 
:::       are deleted.
:::   -d: debug option. Uses intermediate files, producing them if needed. Implies
:::       -k option.
:::   -n: no import. Files are not imported in MSMoney. To be used for debugging 
:::       in conjunction with -d.

set localfile=%~nx0
set localdir=%~dp0
set datesFile=%~dpn0.last

set force=0
set keep=0
set clearcache=0
set debug=0
set noimport=0

:parse
IF "%~1"=="" ( GOTO endparse
) else if "%~1"=="-c" ( set clearcache=1
) else if "%~1"=="/c" ( set clearcache=1
) else if "%~1"=="-C" ( set clearcache=1
) else if "%~1"=="/C" ( set clearcache=1
) else if "%~1"=="-n" ( set noimport=1
) else if "%~1"=="/n" ( set noimport=1
) else if "%~1"=="-N" ( set noimport=1
) else if "%~1"=="/N" ( set noimport=1
) else if "%~1"=="-d" ( set debug=1& set keep=1
) else if "%~1"=="/d" ( set debug=1& set keep=1
) else if "%~1"=="-D" ( set debug=1& set keep=1
) else if "%~1"=="/D" ( set debug=1& set keep=1
) else if "%~1"=="-f" ( set force=1 
) else if "%~1"=="-F" ( set force=1
) else if "%~1"=="/f" ( set force=1
) else if "%~1"=="/F" ( set force=1
) else if "%~1"=="-k" ( set keep=1
) else if "%~1"=="-K" ( set keep=1
) else if "%~1"=="/k" ( set keep=1
) else if "%~1"=="/K" ( set keep=1
) else if "%~1"=="/a" ( set theAccount=%2& SHIFT
) else if "%~1"=="/A" ( set theAccount=%2& SHIFT
) else if "%~1"=="-a" ( set theAccount=%2& SHIFT
) else if "%~1"=="-A" ( set theAccount=%2& SHIFT
) else (
    : Display help
    for /F "tokens=*" %%A in ('findstr "^:::" "%localdir%%localfile%"') do (
        set line=%%A
        set line=!line:%%~nx0=%localfile%!
        echo.!line:~4!
    )
    exit /B
)
SHIFT
GOTO parse
:endparse

if %debug%==1 echo WARNING: intermediate files will not be deleted^^!

rem Find MSMoney path
rem
rem trouve la parenthese dans: (par d�faut)   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
rem                                       |<----- %%a ----
for /f "tokens=1,* delims=)" %%a in ('reg query HKEY_CLASSES_ROOT\money\Shell\Open\Command /ve') do (
    set l=%%b
    rem trouve la fin dans:   REG_SZ    D:\Program Files\Microsoft Money 2005\MNYCoreFiles\MSMoney.exe -url:%1
    rem                                |<----- %%b ----
    for /f "tokens=1,* delims= " %%a in ('echo.!l!') do (
        set moneyPath=%%~dpb
    )
)

rem last download dates file
if not exist %datesFile% (
    echo Id              Minimum date            Last download date >%datesFile%
    echo ----------------------------------------------------------->>%datesFile%
)

if "%theAccount%"=="" (
    set doGetList=0
    if %clearcache%==1 set doGetList=1
    if %debug%==0 set doGetList=1
    if %debug%==1 (
        if not exist %temp%\accounts.list.from_boobank set doGetList=1
    )
    if !doGetList!==1 (
        echo Getting the list of bank accounts from boobank...
        call :requestPwd
        python C:\Python27\Scripts\boobank list > %temp%\accounts.list
        if %debug%==1 copy %temp%\accounts.list %temp%\accounts.list.from_boobank >NUL
    ) else (
         echo Using previously retrieved list of bank accounts...   
         copy %temp%\accounts.list.from_boobank %temp%\accounts.list >NUL    
    )
    type %temp%\accounts.list
) else (
    echo %theAccount% Unknown > %temp%\accounts.list
)

for /f "tokens=1,*" %%a in ('type %temp%\accounts.list ^| find "@"') do (
    set rest=%%b
    set account=%%a
    set label=!rest:~0,26!
    : remove trailing spaces
    for /l %%a in (1,1,31) do if "!label:~-1!"==" " set label=!label:~0,-1!

    rem find last download date
    set lastdate=
    set mindate=
    for /F "tokens=2,3" %%a in ('findstr /B /C:!account! %datesFile%') do (
        set lastdate=%%b
        set mindate=%%a
        if "!mindate!"=="null" set mindate=
    )
    if %force%==1 set lastdate=

    set today=!DATE:~6,4!-!DATE:~3,2!-!DATE:~0,2!

    : diff is < 0 if lastdate strictly before today
    if "!lastdate!"=="" (
        set diff=-1
    ) else (
        set /A diff=!lastdate:-=! - !today:-=!
    )

    echo -------------------
    if "!label!"=="Unkown" (
        set theLabel=
    ) else set theLabel=^(!label!^)
    echo Account !account! !theLabel!
    set ofxfile=C:\Users\bruno\Downloads\!label!_!account!.ofx
    set ofxfile=!ofxfile:/=_!
    if "!diff:~0,1!"=="-" (

        set backend=!account:*@=!
        set backendHandler=%localdir%!backend!.bat

        set inTransaction=0
        (
            set doRetrieve=0
            if %debug%==0 set doRetrieve=1
            if %debug%==1 if not exist !ofxfile!.from_boobank set doRetrieve=1
            if %clearcache%==1 set doRetrieve=1
            if !doRetrieve!==1 (
                if "!lastdate!"=="" (
                    set prevdate=!mindate! 
                ) else set prevdate=!lastdate!
                call :requestPwd
                set retrieveCmd=python c:\Python27\scripts\boobank history !account! !prevdate! -f ofx
                if %debug%==1 set retrieveCmd=!retrieveCmd! ^> "!ofxfile!.from_boobank" ^& type "!ofxfile!.from_boobank"
            ) else (
                set retrieveCmd=type "!ofxfile!.from_boobank"
            )
            for /F "tokens=*" %%A in ('!retrieveCmd!') do (
                set line=%%A

                if "!line!"=="<STMTTRN>" set inTransaction=1
                if !inTransaction!==0 (
                    : MSMoney does not support some account types
                    if "!line!"=="<ACCTTYPE>CARD" set line=^<ACCTTYPE^>CHECKING
                    if "!line!"=="<ACCTTYPE>LOAN" set line=^<ACCTTYPE^>CHECKING
                    if "!line!"=="<ACCTTYPE>MARKET" set line=^<ACCTTYPE^>CHECKING

                    echo.!line!
                )
                if "!line!"=="</STMTTRN>" (
                    : go through specific backend process if any
                    if exist !backendHandler! (
                        : apply the transformations, in the form
                        : field_NAME=...
                        : field_MEMO=...
                        : fields=...

                        : escape the ) which cause the backendHandler to fail
                        for %%F in (!fields!) do call set field_%%F=%%field_%%F:^)=^^^)%%

                        for /F "tokens=1,* delims==" %%s in ('call !backendHandler!') do (
                            : apply the changes
                            call set %%s=%%t
                        )
                    )
                    echo.^<STMTTRN^>
                    for %%F in (!fields!) do (
                        set field=%%F
                        call set value=%%field_!field!%%
                        if !field!==NAME (
                            if "!value!"=="" (
                                : MSMoney does not support empty NAME fields
                                set value=^</NAME^>
                            ) else (
                                : MSMoney does not support NAME fields longer than 64
                                set value=!value:~0,64!
                            )
                        )
                        echo.^<!field!^>!value!
                    )
                    echo.^</STMTTRN^>
                    set inTransaction=0
                )
                if !inTransaction!==1 (
                    if "!line!"=="<STMTTRN>" (
                        for %%F in (!fields!) do call set field_%%F=
                        set fields=
                    ) else (
                        for /F "tokens=1,* delims=>" %%A in ('echo."!line!"') do (
                            set field=%%~A
                            set rest=%%~B
                            if not "!rest!"=="" set rest=!rest:~0,-1!
                            set field=!field:~1!
                            set fields=!fields! !field!
                            call set field_!field!=!rest!
                        )
                    )
                )
            ) 
        ) > "!ofxfile!"

        : check if the file contains transactions
        set nbTransactions=0
        for /F "tokens=*" %%N in ('type "!ofxfile!" ^| find /C "<STMTTRN>"') do set nbTransactions=%%N
        if "!nbTransactions!"=="0" (
            echo No transaction found.
            set status=0
        ) else (
            if !noimport!==0 (
                <nul set /p =Importing !ofxfile! into Money (!nbTransactions! transaction(s^)^)...
                "%moneyPath%mnyimprt.exe" !ofxfile!
                set status=!errorlevel!
                echo.
            ) else (
                echo !nbTransactions! transaction(s^).
                set status=0
            )
        )

        if !status!==0 (
            : mark current date
            findstr /V /B /C:!account! %datesFile% > %TEMP%list.new
            if "!mindate!"=="" set mindate=null
            echo !account!       !mindate!             !today! >>%TEMP%list.new
            copy /Y %TEMP%list.new %datesFile% >NUL
            del %TEMP%list.new
        )

        if %keep%==0 (
            del "!ofxfile!"
        ) else (
            echo WARNING: File "!ofxfile!" is left undeleted.
        )
        if exist "!ofxfile!.from_boobank" (
            if %debug%==0 (
                del "!ofxfile!.from_boobank"
            ) else (
                echo WARNING: File "!ofxfile!.from_boobank" is left undeleted.
            )
        )
    ) else (
        echo Last download date is !lastdate!, no need to download again...
    )
)

del %temp%\accounts.list

if exist %temp%\accounts.list.from_boobank (
    if %debug%==0 (
        del %temp%\accounts.list.from_boobank
    ) else (
        echo WARNING: File "%temp%\accounts.list.from_boobank" is left undeleted.
    )
)

GOTO:EOF

:requestPwd
if not "%password%"=="" goto:EOF

set "psCommand=powershell -Command "$pword = read-host 'Enter Password' -AsSecureString ; ^
    $BSTR=[System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pword); ^
        [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)""
for /f "usebackq delims=" %%p in (`%psCommand%`) do set password=%%p

goto:EOF