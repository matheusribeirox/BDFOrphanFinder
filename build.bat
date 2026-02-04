@echo off
echo ============================================
echo   Compilando BDFOrphanFinder
echo ============================================
echo.

:: Verifica se o dotnet esta instalado
where dotnet >nul 2>&1
if %errorlevel% neq 0 (
    echo ERRO: .NET SDK nao encontrado!
    echo Baixe em: https://dotnet.microsoft.com/download
    pause
    exit /b 1
)

echo Restaurando pacotes...
dotnet restore BDFOrphanFinder.csproj

echo.
echo Escolha o tipo de build:
echo   1. Self-contained (executavel unico ~150MB, nao precisa de .NET instalado)
echo   2. Framework-dependent (menor ~1MB, requer .NET 6 instalado no destino)
echo.
set /p opcao=Digite a opcao (1 ou 2):

if "%opcao%"=="2" (
    echo.
    echo Compilando versao Framework-dependent...
    dotnet publish BDFOrphanFinder.csproj -c Release -o .\publish-small --no-self-contained
    echo.
    echo ============================================
    echo   Compilacao concluida!
    echo   Executavel: .\publish-small\BDFOrphanFinder.exe
    echo   NOTA: Requer .NET 6 Desktop Runtime instalado no destino
    echo ============================================
) else (
    echo.
    echo Compilando versao Self-contained...
    dotnet publish BDFOrphanFinder.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o .\publish
    echo.
    echo ============================================
    echo   Compilacao concluida!
    echo   Executavel: .\publish\BDFOrphanFinder.exe
    echo   (Executavel unico, nao requer .NET instalado)
    echo ============================================
)

echo.
pause
