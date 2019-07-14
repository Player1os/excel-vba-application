@echo off

:: Set the SQL*PLUS script parameter environment variables.
set SQL_PLUS_USERNAME=DIHASSANEIN
set SQL_PLUS_DATABASE=EWH9
set SQL_PLUS_PASSWORD_FILE_PATH=%~dp0decoded.txt

:: Decode and write the password.
cscript /NoLogo %~dp0decoder.vbs <%~dp0password.txt >%SQL_PLUS_PASSWORD_FILE_PATH%

:: Store the curret working directory.
set TEMP_CWD=%CD%

:: Relocate to the project directory.
cd "%1"
