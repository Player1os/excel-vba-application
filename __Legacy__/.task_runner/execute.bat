@echo off
setlocal enabledelayedexpansion
setlocal enableextensions

:: Set environment variables.
call "%~dp0.env.set.bat"

:: Define runtime parameters.
set APP_TASK_RUNNER_TASK_NAME=%1
set APP_TASK_RUNNER_SCRIPT_FILE_PATH=%2

:: Set the iteration counter.
set APP_TASK_RUNNER_ITERATION_COUNTER=0

:: Store the start timestamp.
call "%~dp0timestamp.bat"
set APP_TASK_RUNNER_START_TIMESTAMP=%APP_TIMESTAMP%
set APP_TIMESTAMP=

:: Determine the current task log directory.
set APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH=%APP_TASK_RUNNER_LOG_DIRECTORY_PATH%\%APP_TASK_RUNNER_TASK_NAME%
set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=%APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH%\%APP_TASK_RUNNER_START_TIMESTAMP%

:: Check whether the target script is available.
:check_availability
if not exist %APP_TASK_RUNNER_SCRIPT_FILE_PATH% (
	:: Wait for the specified amount of time.
	ping 127.0.0.1 -n %APP_TASK_RUNNER_WAIT_SECONDS% >nul

	:: Increment the iteration counter.
	set /a APP_TASK_RUNNER_ITERATION_COUNTER=%APP_TASK_RUNNER_ITERATION_COUNTER%+1

	:: Check whether the iteration count was reached.
	if not %APP_TASK_RUNNER_ITERATION_COUNTER% lss %APP_TASK_RUNNER_ITERATION_COUNT% (
		:: Reset the return code, current log directory path and end timestamp.
		set APP_TASK_RUNNER_RETURN_CODE=N/A
		set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=N/A
		set APP_TASK_RUNNER_END_TIMESTAMP=N/A

		:: Report the error to the user.
		cscript /NoLogo "%~dp0\send_error_mail.vbs" "The '%APP_TASK_RUNNER_TASK_NAME%' task script file was not found to be available."

		:: Jump to the termination section.
		goto :terminate
	)

	:: Jump back to the iteration condition.
	goto :check_availability
)

:: Create the current task log directory.
mkdir "%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%"

:: Execute the script in the submitted file path.
call "%APP_TASK_RUNNER_SCRIPT_FILE_PATH%" ^
	>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\out.log" ^
	2>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\err.log"

:: Store the return code.
set APP_TASK_RUNNER_RETURN_CODE=%ERRORLEVEL%

:: Store the end timestamp.
call "%~dp0timestamp.bat"
set APP_TASK_RUNNER_END_TIMESTAMP=%APP_TIMESTAMP%
set APP_TIMESTAMP=

:: Output additional information about the execution of the script.
echo User name: %USERNAME% >"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
echo Machine name: %COMPUTERNAME% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
echo Script file path: %APP_TASK_RUNNER_SCRIPT_FILE_PATH% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
echo Start timestamp: %APP_TASK_RUNNER_START_TIMESTAMP% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
echo End timestamp: %APP_TASK_RUNNER_END_TIMESTAMP% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"
echo Return code: %APP_TASK_RUNNER_RETURN_CODE% >>"%APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH%\info.log"

:: Remove any extraenous log files, that exceed the specified limit.
cscript /NoLogo "%~dp0\remove_extra_logs.vbs"

:: Check whether an error had occurred.
if %APP_TASK_RUNNER_RETURN_CODE% neq 0 (
	:: Report the error to the user.
	cscript /NoLogo "%~dp0\send_error_mail.vbs" "The '%APP_TASK_RUNNER_TASK_NAME%' task script has returned a non-zero status code."

	:: Jump to the termination section.
	goto :terminate
)

:terminate

:: Clear runtime parameters.
set APP_TASK_RUNNER_RETURN_CODE=
set APP_TASK_RUNNER_END_TIMESTAMP=
set APP_TASK_RUNNER_CURRENT_LOG_DIRECTORY_PATH=
set APP_TASK_RUNNER_TASK_LOG_DIRECTORY_PATH=
set APP_TASK_RUNNER_START_TIMESTAMP=
set APP_TASK_RUNNER_ITERATION_COUNTER=
set APP_TASK_RUNNER_SCRIPT_FILE_PATH=

:: Reset environment variables.
call "%~dp0.env.reset.bat"
