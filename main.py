import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import time
import sys

# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
EXCEL_FILE = "professores.xlsx"
URL = "https://agendadocente.pucpr.br/associacao-coordenador"
HEADLESS = False  # True para rodar sem abrir o navegador

# ─────────────────────────────────────────
# CARREGA A PLANILHA
# ─────────────────────────────────────────
def carregar_dados(arquivo: str) -> pd.DataFrame:
    df = pd.read_excel(arquivo)
    df.columns = df.columns.str.strip().str.lower()
    colunas_necessarias = {"professor", "coordenador", "campus"}
    if not colunas_necessarias.issubset(df.columns):
        raise ValueError(
            f"A planilha deve conter as colunas: {colunas_necessarias}\n"
            f"Colunas encontradas: {set(df.columns)}"
        )
    df = df.dropna(subset=["professor", "coordenador"])
    df["professor"] = df["professor"].str.strip().str.upper()
    df["coordenador"] = df["coordenador"].str.strip().str.upper()
    df["campus"] = df["campus"].str.strip()
    return df

# ─────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ─────────────────────────────────────────
def aguardar_e_clicar(page, selector, timeout=10000):
    page.wait_for_selector(selector, timeout=timeout)
    page.click(selector)

def fechar_modal_sucesso(page):
    """Clica em OK no modal 'Salvo com sucesso.' se ele aparecer."""
    try:
        page.wait_for_selector("text=Salvo com sucesso.", timeout=8000)
        page.click("button:has-text('OK')")
        page.wait_for_selector("text=Salvo com sucesso.", state="hidden", timeout=5000)
        print("  ✅ Salvo com sucesso.")
    except PlaywrightTimeoutError:
        print("  ⚠️  Modal de sucesso não apareceu.")

def limpar_campo_professor(page):
    """Clica no X para limpar o campo de professor."""
    try:
        page.wait_for_selector("button[aria-label='Clear'], .vs__clear", timeout=5000)
        page.click("button[aria-label='Clear'], .vs__clear")
        page.wait_for_selector("input[placeholder='Professor']", timeout=5000)
    except PlaywrightTimeoutError:
        pass  # Campo já pode estar vazio

def selecionar_professor(page, nome_professor: str) -> bool:
    """Digita o nome no campo e seleciona a opção no dropdown."""
    print(f"  🔍 Buscando professor: {nome_professor}")
    try:
        campo = page.wait_for_selector("input[placeholder='Professor']", timeout=8000)
        campo.click()
        campo.fill(nome_professor)
        time.sleep(1)  # aguarda o dropdown carregar

        opcao = page.wait_for_selector(
            f".vs__dropdown-menu li:has-text('{nome_professor}')", timeout=8000
        )
        opcao.click()
        page.wait_for_selector("text=2. Selecione o coordenador", timeout=8000)
        print(f"  ✔ Professor selecionado: {nome_professor}")
        return True
    except PlaywrightTimeoutError:
        print(f"  ❌ Professor não encontrado: {nome_professor}")
        return False

def coordenador_ja_atribuido(page, nome_coordenador: str) -> bool:
    """Verifica se o coordenador já está na coluna 'Coordenador atribuído'."""
    try:
        page.wait_for_selector(".atribuido, [class*='atribuido']", timeout=3000)
    except PlaywrightTimeoutError:
        pass
    conteudo = page.inner_text("body")
    return nome_coordenador in conteudo and "Coordenador atribuído" in conteudo

def buscar_e_selecionar_coordenador(page, nome_coordenador: str) -> bool:
    """Pesquisa o coordenador na lista disponível e marca o checkbox."""
    print(f"  🔍 Buscando coordenador: {nome_coordenador}")
    try:
        # Campo de pesquisa na lista "Coordenador disponível"
        campos_pesquisa = page.query_selector_all("input[placeholder='Pesquise aqui']")
        if not campos_pesquisa:
            print("  ❌ Campo 'Pesquise aqui' não encontrado.")
            return False

        campo_disponivel = campos_pesquisa[0]
        campo_disponivel.click()
        campo_disponivel.fill(nome_coordenador)
        time.sleep(1)

        # Localiza o checkbox do coordenador filtrado
        linha = page.wait_for_selector(
            f"text={nome_coordenador}", timeout=8000
        )
        # Sobe até o checkbox da linha
        checkbox = page.query_selector(
            f"input[type='checkbox'] ~ * >> text={nome_coordenador}"
        )
        if checkbox is None:
            # Tenta localizar via linha da tabela
            linha.evaluate("el => el.closest('tr, li, .item, label')?.querySelector('input[type=checkbox]')?.click()")
        else:
            linha.click()

        time.sleep(0.5)

        # Verifica se algum checkbox ficou marcado
        checkboxes = page.query_selector_all("input[type='checkbox']:checked")
        if not checkboxes:
            # Fallback: clica direto no elemento de texto para marcar
            page.click(f"text={nome_coordenador}")

        print(f"  ✔ Coordenador selecionado: {nome_coordenador}")
        return True
    except PlaywrightTimeoutError:
        print(f"  ❌ Coordenador não encontrado: {nome_coordenador}")
        return False

def clicar_adicionar(page) -> bool:
    try:
        btn = page.wait_for_selector("button:has-text('Adicionar'):not([disabled])", timeout=5000)
        btn.click()
        time.sleep(1)
        print("  ✔ Clicou em Adicionar.")
        return True
    except PlaywrightTimeoutError:
        print("  ⚠️  Botão Adicionar não encontrado ou desabilitado.")
        return False

def clicar_salvar(page) -> bool:
    try:
        btn = page.wait_for_selector("button:has-text('Salvar'):not([disabled])", timeout=5000)
        btn.click()
        return True
    except PlaywrightTimeoutError:
        print("  ⚠️  Botão Salvar não encontrado.")
        return False

# ─────────────────────────────────────────
# FLUXO PRINCIPAL
# ─────────────────────────────────────────
def processar_linha(page, professor: str, coordenador: str, campus: str):
    print(f"\n{'─'*60}")
    print(f"📋 Professor : {professor}")
    print(f"👤 Coordenador: {coordenador}")
    print(f"🏫 Campus    : {campus}")
    print(f"{'─'*60}")

    if not selecionar_professor(page, professor):
        return

    # Verifica se coordenador já está atribuído
    atribuidos = page.inner_text("body")
    if coordenador in atribuidos:
        print(f"  ℹ️  Coordenador '{coordenador}' já está atribuído. Pulando...")
        limpar_campo_professor(page)
        return

    if not buscar_e_selecionar_coordenador(page, coordenador):
        limpar_campo_professor(page)
        return

    if not clicar_adicionar(page):
        limpar_campo_professor(page)
        return

    if not clicar_salvar(page):
        limpar_campo_professor(page)
        return

    fechar_modal_sucesso(page)
    limpar_campo_professor(page)

def main():
    print("📂 Carregando planilha...")
    try:
        df = carregar_dados(EXCEL_FILE)
    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        sys.exit(1)

    print(f"✅ {len(df)} registros encontrados.\n")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()

        print(f"🌐 Acessando {URL}...")
        page.goto(URL, wait_until="networkidle")

        # Aguarda autenticação automática (SSO)
        print("⏳ Aguardando autenticação automática...")
        page.wait_for_selector("text=Associação de Coordenador", timeout=30000)
        print("✅ Autenticado e na página correta.\n")

        erros = []
        for i, row in df.iterrows():
            try:
                processar_linha(
                    page,
                    professor=row["professor"],
                    coordenador=row["coordenador"],
                    campus=row["campus"],
                )
            except Exception as e:
                print(f"  ❌ Erro inesperado na linha {i+2}: {e}")
                erros.append({"linha": i+2, "professor": row["professor"], "erro": str(e)})
                # Tenta limpar o campo para continuar
                try:
                    limpar_campo_professor(page)
                except Exception:
                    pass

        browser.close()

    print(f"\n{'═'*60}")
    print(f"✅ Processamento concluído!")
    if erros:
        print(f"⚠️  {len(erros)} erro(s) encontrado(s):")
        for e in erros:
            print(f"   Linha {e['linha']} | {e['professor']} → {e['erro']}")
    else:
        print("🎉 Todos os registros foram processados sem erros.")

if __name__ == "__main__":
    main()