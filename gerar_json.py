"""
gerar_json.py — Brasil Botões
Lê a planilha de estoque e gera o produtos.json atualizado.
Execute via Atualizar_Catalogo.bat
"""

import os, re, json, sys
from datetime import datetime

# ══════════════════════════════════════════════════════════════
#  CONFIGURAÇÃO
# ══════════════════════════════════════════════════════════════

PASTA_BASE   = os.path.dirname(os.path.abspath(__file__))
PLANILHA     = os.path.join(PASTA_BASE, "apoio", "DADOS_DE_PRODUTOS_E_ESTOQUE.xlsm")
TABELA_CORES = os.path.join(PASTA_BASE, "apoio", "tabela_cores_final.xlsx")
SAIDA_JSON   = os.path.join(PASTA_BASE, "produtos.json")

FURACAO_MAP = {
    'QUATRO FUROS':'4F','DOIS FUROS':'2F','UM FURO':'1F',
    'TRES FUROS':'3F','SEM FURO':'SF','OUTROS (PE)':'Pé',
    '5 FUROS':'5F','SEIS FUROS':'6F','9 FUROS':'9F',
}

# ══════════════════════════════════════════════════════════════

def log(msg):
    print(f"  {msg}")

def verificar_dependencias():
    try:
        import openpyxl
    except ImportError:
        log("Instalando dependência openpyxl...")
        os.system("pip install openpyxl --break-system-packages -q")

def carregar_bd_prod(wb):
    ws = wb['BD_PROD']
    bd = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0] and isinstance(r[0], int):
            bd[r[0]] = {
                'descricao':  str(r[1]).strip() if r[1] else '',
                'categoria':  str(r[3]).strip() if r[3] else '',
                'referencia': str(r[4]).strip() if r[4] else '',
                'cor':        str(r[5]).strip() if r[5] else '',
                'tamanho':    str(r[6]).strip() if r[6] else '',
                'furacao':    str(r[7]).strip() if r[7] else '',
            }
    log(f"BD_PROD: {len(bd)} registros")
    return bd

def carregar_cores():
    import openpyxl
    if not os.path.exists(TABELA_CORES):
        log("AVISO: tabela_cores_final.xlsx não encontrada — cores sem classificação")
        return {}
    wb = openpyxl.load_workbook(TABELA_CORES, read_only=True)
    ws = wb.active
    cores = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[0]:
            cores[str(r[0]).strip().upper()] = {
                'cor_nome':    str(r[1]).strip() if r[1] else '',
                'cor_familia': str(r[2]).strip() if r[2] else '',
                'cor_hex':     str(r[3]).strip() if r[3] else '',
            }
    log(f"Cores: {len(cores)} classificadas")
    return cores

def gerar_json():
    import openpyxl

    print()
    print("=" * 50)
    print("  Brasil Botões — Atualização do Catálogo")
    print("=" * 50)
    print(f"  {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    print()

    # Verifica arquivos
    if not os.path.exists(PLANILHA):
        print(f"  ERRO: Planilha não encontrada em:\n  {PLANILHA}")
        return False

    log("Carregando planilha...")
    wb = openpyxl.load_workbook(PLANILHA, read_only=True, keep_vba=True)

    bd    = carregar_bd_prod(wb)
    cores = carregar_cores()

    # Lê aba PRODUTOS
    ws_prod = wb['PRODUTOS']
    todas   = list(ws_prod.iter_rows(values_only=True))
    cab_idx = next(i for i, r in enumerate(todas) if r[0] == 'REDUZIDO')

    vistos   = set()
    produtos = []

    for r in todas[cab_idx + 1:]:
        reduzido = r[0]
        if not reduzido or not isinstance(reduzido, int):
            continue
        qtde  = r[11]
        saldo = int(qtde) if isinstance(qtde, (int, float)) else 0
        if saldo <= 0:
            continue
        info = bd.get(reduzido, {})
        if not info.get('descricao'):
            continue
        chave = str(reduzido)
        if chave in vistos:
            continue
        vistos.add(chave)

        cor_raw  = info.get('cor', '')
        cor_norm = re.sub(r'\s+', '', cor_raw).strip().upper()
        cor_info = cores.get(cor_norm, {})
        fur_raw  = info.get('furacao', '').strip()
        furacao  = FURACAO_MAP.get(fur_raw, fur_raw)
        ref      = info.get('referencia', '').strip().lstrip('0') or info.get('referencia', '').strip()

        produtos.append({
            'reduzido':    reduzido,
            'descricao':   info.get('descricao', ''),
            'referencia':  ref,
            'cor_codigo':  cor_norm,
            'cor_nome':    cor_info.get('cor_nome', ''),
            'cor_familia': cor_info.get('cor_familia', ''),
            'cor_hex':     cor_info.get('cor_hex', ''),
            'tamanho':     info.get('tamanho', '').strip(),
            'furacao':     furacao,
            'acabamento':  info.get('categoria', '').strip(),
            'saldo':       saldo,
            'imagem_base': f"{ref}_{cor_norm}_{furacao}".replace(' ', ''),
        })

    # Salva JSON
    with open(SAIDA_JSON, 'w', encoding='utf-8') as f:
        json.dump(produtos, f, ensure_ascii=False, indent=2)

    print()
    log(f"Produtos em estoque: {len(produtos)}")
    log(f"JSON salvo: produtos.json")
    print()
    return True

if __name__ == "__main__":
    verificar_dependencias()
    ok = gerar_json()
    if not ok:
        input("  Pressione Enter para fechar...")
        sys.exit(1)
