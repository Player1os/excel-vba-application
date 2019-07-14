@echo off

:: Relocate to the original working directory and unset the temp cwd environment variable.
cd "%TEMP_CWD%"
set TEMP_CWD=

:: Delete the decoded password file.
del "%SQL_PLUS_PASSWORD_FILE_PATH%"

:: Unset the SQL*PLUS script parameter environment variables.
set SQL_PLUS_USERNAME=
set SQL_PLUS_DATABASE=
set SQL_PLUS_PASSWORD_FILE_PATH=
