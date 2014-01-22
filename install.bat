@echo off

setlocal

if exist install.bat goto ok
echo install.bat must be run from its folder
goto end

: ok

set OLDGOPATH=%GOPATH%
set GOPATH=%~dp0

gofmt -tabs=false -tabwidth=4 -w src

go install -v doc2pdf

:end
echo finished