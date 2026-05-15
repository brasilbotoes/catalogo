@echo off
chcp 65001 > nul
title Atualizar Catálogo Brasil Botões

echo.
echo ================================================
echo   Brasil Botões — Atualizando Catálogo
echo ================================================
echo.

:: Vai para a pasta do script
cd /d "%~dp0"

:: Gera o JSON atualizado
echo [1/3] Gerando dados do catálogo...
python gerar_json.py
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao gerar o JSON. Verifique a planilha.
    pause
    exit /b 1
)

:: Envia para o GitHub
echo [2/3] Enviando para o GitHub...
git add produtos.json imagens/
git commit -m "Atualização do catálogo %date% %time%"
if %errorlevel% neq 0 (
    echo  Nenhuma alteração detectada ou erro no commit.
)

git push origin main
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao enviar para o GitHub.
    echo  Verifique sua conexão com a internet.
    pause
    exit /b 1
)

echo.
echo [3/3] Catálogo atualizado com sucesso!
echo.
echo  Acesse: https://brasilbotoes.github.io/catalogo
echo.
echo ================================================
timeout /t 5
