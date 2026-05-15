@echo off
chcp 65001 > nul
title Publicar Catálogo Brasil Botões no GitHub

echo.
echo ================================================
echo   Brasil Botões — Publicação Inicial no GitHub
echo ================================================
echo.

:: Vai para a pasta do script
cd /d "%~dp0"

echo [1/5] Verificando Git instalado...
git --version > nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Git não encontrado. Baixe em: https://git-scm.com/download/win
    echo  Instale e execute este script novamente.
    pause
    exit /b 1
)
echo  Git OK.

echo.
echo [2/5] Configurando repositório local...
git init
git checkout -b main

echo.
echo [3/5] Adicionando arquivos...
git add index.html produtos.json logo_simbolo.jpeg gerar_json.py Atualizar_Catalogo.bat imagens/
git commit -m "Publicação inicial do Catálogo Brasil Botões"

echo.
echo [4/5] Conectando ao GitHub...
git remote add origin https://github.com/brasilbotoes/catalogo.git
git remote -v

echo.
echo [5/5] Enviando para o GitHub...
git push -u origin main

if %errorlevel% neq 0 (
    echo.
    echo  ERRO ao enviar. Isso pode acontecer se o repositório
    echo  ainda não foi criado no GitHub.
    echo.
    echo  Acesse github.com, clique em New Repository,
    echo  nome: catalogo, deixe vazio e clique Create.
    echo  Depois execute este script novamente.
    pause
    exit /b 1
)

echo.
echo ================================================
echo  Publicado com sucesso!
echo.
echo  Agora ative o GitHub Pages:
echo  1. Acesse github.com/brasilbotoes/catalogo
echo  2. Settings ^> Pages
echo  3. Source: Deploy from branch
echo  4. Branch: main / (root)
echo  5. Save
echo.
echo  Link do catálogo:
echo  https://brasilbotoes.github.io/catalogo
echo ================================================
echo.
pause
