@echo off
REM pushd em caminho UNC cria letra de drive temporaria automaticamente
pushd "%~dp0"

echo ================================================
echo  DFC Dashboard - Atualizacao de Dados
echo ================================================
echo.
echo Gerando dados da planilha...
echo.

powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0gerar_dados.ps1"

if errorlevel 1 (
    echo.
    echo Falha ao gerar os dados. Verifique as mensagens acima.
    popd
    pause
    exit /b 1
)

echo.
echo Abrindo dashboard no navegador...
start "" "%CD%\Fluxo de caixa.html"

echo.
echo Pronto!
popd
timeout /t 2 /nobreak >nul
