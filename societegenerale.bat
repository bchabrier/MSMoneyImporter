@echo OFF
setlocal EnableExtensions
setlocal EnableDelayedExpansion

: Transforms a transaction
: 
: Available transaction fields are stored in the 'fields' variable, e.g. field=NAME, 
: field=MEMO. Each field value is stored in variable field_<field>, e.g. field_NAME, 
: field_MEMO.
: 
: New fields can be added in 'fields', with their values defined accordingly.
: 
: Transformed or added fields must be echo'ed with their value, in order to be taken 
: into account, e.g:
: echo field_NAME=the new value of NAME
: 
if not "!field_MEMO!"=="" if not "!field_MEMO:CHEQUE=!"=="!field_MEMO!" (
    rem MEMO contains "CHEQUE"

    rem check if NAME contains a number
    call :isNumeric !field_NAME!
    if !ERRORLEVEL!==0 (
       rem NAME contains a number

        rem replace NAME by CHECKNUM
        echo field_MEMO=
        echo field_CHECKNUM=!field_NAME!
        echo fields=!fields:NAME=CHECKNUM!
        echo field_NAME=
    )
)
exit /B

:isNumeric
: - checks if parameter is numeric
: - return code set to 1 if not numeric
set i=%*
set /A n=1%i% 2>NUL
if "%n%"=="1%i%" exit /B 0
exit /B 1
