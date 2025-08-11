@echo off
set CONTAINER_NAME=obsidian-onlyoffice

echo WARNING: This will permanently remove the OnlyOffice container.
set /p "choice=Are you sure? (y/n): "
if /i not "%choice%"=="y" goto :eof

echo Stopping and removing container '%CONTAINER_NAME%'...
docker stop %CONTAINER_NAME% > nul
docker rm %CONTAINER_NAME% > nul
echo Container removed.
echo.
echo --- Operation Complete ---
pause
