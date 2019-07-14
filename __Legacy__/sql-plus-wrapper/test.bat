@echo off
setlocal enabledelayedexpansion
setlocal enableextensions

set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=a a\^
   %COMPUTERNAME%\bb\cc

echo %APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%

set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=

:: type projects\plsql-execution-monitor\verify.bat >&2
