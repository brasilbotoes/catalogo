@echo off
chcp 65001 > nul
title Atualizar Catalogo Brasil Botoes

echo.
echo ====================================================
echo   Brasil Botoes ^| Atualizando Catalogo
echo ====================================================
echo.

cd /d "%~dp0"

echo [1/3] Gerando dados do catalogo...
python "%~dp0gerar_json.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao gerar o JSON.
    echo  Verifique se a planilha esta fechada e tente novamente.
    pause
    exit /b 1
)

echo.
echo [2/3] Enviando para o GitHub...
"C:\Program Files\Git\bin\git.exe" add produtos.json imagens/
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
echo [3/3] Catalogo atualizado com sucesso!
echo.
echo  Acesse: https://brasilbotoes.github.io/catalogo
echo.
echo ====================================================
timeout /t 5
