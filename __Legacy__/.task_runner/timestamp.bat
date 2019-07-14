:: Store the current time and date.
set APP_TIME=%TIME%
set APP_DATE=%DATE%

:: Determine the current year.
set APP_YEAR=%APP_DATE:~-4%

:: Determine the current month.
set APP_MONTH=%APP_DATE:~-7,2%
if "%APP_MONTH:~0,1%" == " " (
	set APP_MONTH=0%APP_MONTH:~1,1%
)

:: Determine the current day.
set APP_DAY=%APP_DATE:~-10,2%
if "%APP_DAY:~0,1%" == " " (
	set APP_DAY=0%APP_DAY:~1,1%
)

:: Determine the current hour.
set APP_HOUR=%APP_TIME:~0,2%
if "%APP_HOUR:~0,1%" == " " (
	set APP_HOUR=0%APP_HOUR:~1,1%
)

:: Determine the current minute.
set APP_MINUTE=%APP_TIME:~3,2%
if "%APP_MINUTE:~0,1%" == " " (
	set APP_MINUTE=0%APP_MINUTE:~1,1%
)

:: Determine the current second.
set APP_SECOND=%APP_TIME:~6,2%
if "%APP_SECOND:~0,1%" == " " (
	set APP_SECOND=0%APP_SECOND:~1,1%
)

:: Combine the collected parts into a timestamp.
set APP_TIMESTAMP=%APP_YEAR%%APP_MONTH%%APP_DAY%_%APP_HOUR%%APP_MINUTE%%APP_SECOND%

:: Clear the collected parts.
set APP_SECOND=
set APP_MINUTE=
set APP_HOUR=
set APP_DAY=
set APP_MONTH=
set APP_YEAR=

:: Clear the stored time and date.
set APP_DATE=
set APP_TIME=
