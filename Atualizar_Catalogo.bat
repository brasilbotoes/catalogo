@echo off
chcp 65001 > nul
title Atualizar Catálogo Brasil Botões

echo.
echo ================================================
echo   Brasil Botões — Atualizando Catálogo
echo ================================================
echo.

:: Pasta do catálogo local
cd /d "C:\catalogo_bb"

:: Caminhos
set PLANILHA_ORIGEM=C:\Users\Usuario\BRASIL BOTOES LTDA\Público - Documentos\catalogo\apoio\DADOS_DE_PRODUTOS_E_ESTOQUE.xlsm
set PLANILHA_DESTINO=C:\catalogo_bb\apoio\DADOS_DE_PRODUTOS_E_ESTOQUE.xlsm

echo [1/4] Copiando planilha do Sharepoint...
copy /Y "%PLANILHA_ORIGEM%" "%PLANILHA_DESTINO%" > nul
if %errorlevel% neq 0 (
    echo.
    echo  ERRO: Planilha nao encontrada no Sharepoint.
    echo  Verifique se o arquivo esta sincronizado.
    pause
    exit /b 1
)
echo  Planilha copiada.

echo.
echo [2/4] Gerando dados do catalogo...
python gerar_json.py
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao gerar o JSON.
    pause
    exit /b 1
)

echo.
echo [3/4] Enviando para o GitHub...
git add produtos.json
git commit -m "Atualizacao de estoque %date%"
git push origin main
if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao enviar para o GitHub.
    echo  Verifique sua conexao com a internet.
    pause
    exit /b 1
)

echo.
echo [4/4] Concluido!
echo.
echo  Catalogo atualizado em:
echo  https://brasilbotoes.github.io/catalogo
echo.
echo ================================================
timeout /t 5
