@echo off
echo Preparing GitHub deployment for ILB Excel Add-in...

REM Build the project
echo Building project...
npm run build
if %ERRORLEVEL% neq 0 (
    echo Build failed!
    pause
    exit /b 1
)

REM Copy GitHub manifest to dist
echo Copying GitHub manifest...
copy manifest-github.xml dist\manifest.xml
if %ERRORLEVEL% neq 0 (
    echo Failed to copy manifest!
    pause
    exit /b 1
)

REM Create GitHub deployment package
echo Creating GitHub deployment package...
if exist github-package rmdir /s /q github-package
mkdir github-package
xcopy dist\* github-package\ /s /e /y

echo.
echo ================================
echo GitHub Deployment Package Ready!
echo ================================
echo.
echo Next steps:
echo 1. Create a GitHub repository
echo 2. Push your code to the repository
echo 3. Enable GitHub Pages in repository settings
echo 4. Point GitHub Pages to the main branch
echo.
echo Your add-in will be available at:
echo https://YOUR-USERNAME.github.io/REPOSITORY-NAME/manifest.xml
echo.
echo Files prepared in: github-package\
echo.

pause