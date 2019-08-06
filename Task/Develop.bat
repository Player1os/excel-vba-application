@echo off
setlocal enabledelayedexpansion
setlocal enableextensions

:: Set the working directory to the project's directory.
cd %~dp0..\..

:: Load the corresponding vbscript with the base script.
call .\Script\_Base.vbs .\Script\Develop.vbs
