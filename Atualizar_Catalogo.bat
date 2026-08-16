@echo off
chcp 65001 > nul
title Atualizar Catalogo Brasil Botoes

echo.
echo ====================================================
echo   Brasil Botoes ^| Atualizando Catalogo
echo ====================================================
echo.

cd /d "%~dp0"

echo [1/4] Gerando dados do catalogo geral...
python "%~dp0gerar_json.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao gerar o JSON do catalogo geral.
    echo  Verifique se a planilha esta fechada e tente novamente.
    pause
    exit /b 1
)

echo.
echo [2/4] Gerando catalogos por representante...
python "%~dp0gerar_catalogos_representantes.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao gerar os catalogos de representantes.
    echo  Verifique se a planilha de Clientes esta fechada e tente novamente.
    pause
    exit /b 1
)

echo.
echo [3/4] Enviando para o GitHub...
"C:\Program Files\Git\bin\git.exe" add -A
"C:\Program Files\Git\bin\git.exe" commit -m "Atualizacao %date%"
if %errorlevel% neq 0 (
    echo  Nenhuma alteracao detectada ou erro no commit.
)
"C:\Program Files\Git\bin\git.exe" push origin main
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao enviar para o GitHub.
    echo  Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo.
echo [4/4] Catalogo atualizado com sucesso!
echo.
echo  Catalogo geral: https://brasilbotoes.github.io/catalogo
echo  Catalogos de representantes: https://brasilbotoes.github.io/catalogo/representantes/^<slug^>
echo.
echo ====================================================
timeout /t 5
