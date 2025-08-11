@echo off
set CONTAINER_NAME=obsidian-onlyoffice
set JWT_SECRET=your-secret-key-please-change
set HOST_PORT=8080

echo --- OnlyOffice Docker Manager ---
echo.
echo Checking for container '%CONTAINER_NAME%'...

docker ps -a --filter "name=%CONTAINER_NAME%" --format "{{.Names}}" | findstr /r "^%CONTAINER_NAME%$" > nul
if %errorlevel% == 0 (
    echo Container found.
    docker ps --filter "name=%CONTAINER_NAME%" --format "{{.Names}}" | findstr /r "^%CONTAINER_NAME%$" > nul
    if %errorlevel% == 0 (
        echo Container is already running.
    ) else (
        echo Container is stopped. Starting...
        docker start %CONTAINER_NAME%
    )
) else (
    echo Container not found. Creating and starting a new one.
    echo IMPORTANT: Make sure the JWT_SECRET in this script matches the one in your Obsidian plugin settings.
    echo Current secret: %JWT_SECRET%
    echo.
    
    docker run -d --name %CONTAINER_NAME% -p %HOST_PORT%:80 --restart always --add-host=host.docker.internal:host-gateway ^
    -e JWT_ENABLED=true ^
    -e JWT_SECRET=%JWT_SECRET% ^
    -e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_ORIGIN="app://obsidian.md" ^
    -e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_METHODS="GET, POST, OPTIONS, PUT, DELETE" ^
    -e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_HEADERS="Content-Type, Authorization" ^
    -e ONLYOFFICE_NGINX_ACCESS_CONTROL_ALLOW_CREDENTIALS="true" ^
    onlyoffice/documentserver
)

echo.
echo --- Operation Complete ---
pause
