import sys
import time

import pandas as pd
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright

# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
EXCEL_FILE = "professores.xlsx"
URL = "https://agendadocente.pucpr.br/associacao-coordenador"
HEADLESS = False

# Códigos de status por linha processada
STATUS_OK = "OK"
STATUS_PULADO = "PULADO"
STATUS_ERRO_PROFESSOR = "ERRO_PROFESSOR"
STATUS_ERRO_ANALISTA = "ERRO_ANALISTA"
STATUS_ERRO_ADICIONAR = "ERRO_ADICIONAR"
STATUS_ERRO_SALVAR = "ERRO_SALVAR"
STATUS_ERRO_INESPERADO = "ERRO_INESPERADO"

# ─────────────────────────────────────────
# CARREGA A PLANILHA
# ─────────────────────────────────────────
def carregar_dados(arquivo: str) -> pd.DataFrame:
    df = pd.read_excel(arquivo)
    df.columns = df.columns.str.strip().str.lower()

    colunas_necessarias = {"professor", "analista"}
    if not colunas_necessarias.issubset(df.columns):
        raise ValueError(
            f"A planilha deve conter as colunas: {colunas_necessarias}\n"
            f"Colunas encontradas: {set(df.columns)}"
        )

    df = df.dropna(subset=["professor", "analista"])
    df["professor"] = df["professor"].str.strip().str.upper()
    df["analista"] = df["analista"].str.strip().str.upper()
    if "campus" in df.columns:
        df["campus"] = df["campus"].str.strip()
    if "escola" in df.columns:
        df["escola"] = df["escola"].str.strip()

    return df.reset_index(drop=True)


# ─────────────────────────────────────────
# FUNÇÕES AUXILIARES
# ─────────────────────────────────────────
def fechar_modal_sucesso(page) -> None:
    try:
        page.wait_for_selector("text=Salvo com sucesso.", timeout=8000)
        page.click("button:has-text('OK')")
        page.wait_for_selector("text=Salvo com sucesso.", state="hidden", timeout=5000)
        print("  ✅ Salvo com sucesso.")
    except PlaywrightTimeoutError:
        print("  ⚠️  Modal de sucesso não apareceu.")


def limpar_campo_professor(page) -> None:
    try:
        page.wait_for_selector("button[aria-label='Clear'], .vs__clear", timeout=5000)
        page.click("button[aria-label='Clear'], .vs__clear")
        page.wait_for_selector("input[placeholder='Professor']", timeout=5000)
    except PlaywrightTimeoutError:
        pass


def selecionar_professor(page, nome_professor: str) -> bool:
    print(f"  🔍 Buscando professor: {nome_professor}")
    try:
        campo = page.wait_for_selector("input[placeholder='Professor']", timeout=8000)
        campo.click()
        campo.fill(nome_professor)
        time.sleep(1.5)

        opcao = page.wait_for_selector(
            f".vs__dropdown-menu li:has-text('{nome_professor}')", timeout=8000
        )
        opcao.click()
        page.wait_for_selector("text=2. Selecione o coordenador", timeout=10000)
        print(f"  ✔ Professor selecionado.")
        return True
    except PlaywrightTimeoutError:
        print(f"  ❌ Professor não encontrado no dropdown.")
        return False


def analista_ja_atribuida(page, nome_analista: str) -> bool:
    """
    Verifica se a analista já consta na tabela 'Coordenador atribuído' (lado direito).
    Usa JavaScript para checar apenas o texto que vem APÓS o cabeçalho da seção,
    evitando falso-positivo com a tabela 'Coordenador disponível' (lado esquerdo).
    """
    try:
        resultado = page.evaluate(
            """(nome) => {
                const texto = document.body.innerText;
                const idx = texto.indexOf('Coordenador atribuído');
                if (idx === -1) return false;
                return texto.slice(idx).includes(nome);
            }""",
            nome_analista,
        )
        return bool(resultado)
    except Exception:
        return False


def buscar_e_selecionar_analista(page, nome_analista: str) -> bool:
    """
    Pesquisa a analista na lista 'Coordenador disponível' e marca seu checkbox.
    Tenta primeiro via linhas de tabela (tr), depois via itens de lista (li/label).
    """
    print(f"  🔍 Buscando analista: {nome_analista}")
    try:
        campos_pesquisa = page.query_selector_all("input[placeholder='Pesquise aqui']")
        if not campos_pesquisa:
            print("  ❌ Campo 'Pesquise aqui' não encontrado.")
            return False

        campo_disponivel = campos_pesquisa[0]
        campo_disponivel.click()
        campo_disponivel.fill(nome_analista)
        time.sleep(1)

        page.wait_for_selector(f"text={nome_analista}", timeout=8000)

        # Estratégia 1: linhas de tabela
        for linha in page.query_selector_all("tr"):
            try:
                if nome_analista in linha.inner_text():
                    checkbox = linha.query_selector("input[type='checkbox']")
                    if checkbox and not checkbox.is_checked():
                        checkbox.click()
                        time.sleep(0.3)
                        print(f"  ✔ Analista selecionada.")
                        return True
            except Exception:
                continue

        # Estratégia 2: itens de lista ou labels
        for item in page.query_selector_all("li, label"):
            try:
                if nome_analista in item.inner_text():
                    checkbox = item.query_selector("input[type='checkbox']")
                    if checkbox:
                        if not checkbox.is_checked():
                            checkbox.click()
                    else:
                        item.click()
                    time.sleep(0.3)
                    print(f"  ✔ Analista selecionada (fallback).")
                    return True
            except Exception:
                continue

        print(f"  ❌ Analista não encontrada na lista disponível.")
        return False

    except PlaywrightTimeoutError:
        print(f"  ❌ Analista não apareceu após pesquisa.")
        return False


def clicar_adicionar(page) -> bool:
    try:
        btn = page.wait_for_selector(
            "button:has-text('Adicionar'):not([disabled])", timeout=5000
        )
        btn.click()
        time.sleep(1)
        print("  ✔ Clicou em Adicionar.")
        return True
    except PlaywrightTimeoutError:
        print("  ⚠️  Botão Adicionar não encontrado ou desabilitado.")
        return False


def clicar_salvar(page) -> bool:
    try:
        btn = page.wait_for_selector(
            "button:has-text('Salvar'):not([disabled])", timeout=5000
        )
        btn.click()
        return True
    except PlaywrightTimeoutError:
        print("  ⚠️  Botão Salvar não encontrado.")
        return False


# ─────────────────────────────────────────
# FLUXO POR LINHA
# ─────────────────────────────────────────
def processar_linha(page, row: dict) -> str:
    professor = row["professor"]
    analista = row["analista"]
    campus = row.get("campus", "")

    print(f"\n{'─' * 60}")
    print(f"📋 Professor : {professor}")
    print(f"👤 Analista  : {analista}")
    if campus:
        print(f"🏫 Campus    : {campus}")
    print(f"{'─' * 60}")

    if not selecionar_professor(page, professor):
        return STATUS_ERRO_PROFESSOR

    if analista_ja_atribuida(page, analista):
        print(f"  ℹ️  Analista já atribuída. Pulando...")
        limpar_campo_professor(page)
        return STATUS_PULADO

    if not buscar_e_selecionar_analista(page, analista):
        limpar_campo_professor(page)
        return STATUS_ERRO_ANALISTA

    if not clicar_adicionar(page):
        limpar_campo_professor(page)
        return STATUS_ERRO_ADICIONAR

    if not clicar_salvar(page):
        limpar_campo_professor(page)
        return STATUS_ERRO_SALVAR

    fechar_modal_sucesso(page)
    limpar_campo_professor(page)
    return STATUS_OK


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────
def main():
    print("📂 Carregando planilha...")
    try:
        df = carregar_dados(EXCEL_FILE)
    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        sys.exit(1)

    print(f"✅ {len(df)} registros encontrados.\n")

    resultados = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        context = browser.new_context()
        page = context.new_page()

        print(f"🌐 Acessando {URL}...")
        page.goto(URL, wait_until="networkidle")

        print("⏳ Aguardando autenticação SSO...")
        page.wait_for_selector("text=Associação de Coordenador", timeout=60000)
        print("✅ Autenticado e na página correta.\n")

        for i, row in df.iterrows():
            row_dict = row.to_dict()
            try:
                status = processar_linha(page, row_dict)
            except Exception as e:
                print(f"  ❌ Erro inesperado na linha {i + 2}: {e}")
                status = STATUS_ERRO_INESPERADO
                try:
                    limpar_campo_professor(page)
                except Exception:
                    pass

            resultados.append(
                {
                    "linha": i + 2,
                    "professor": row_dict.get("professor", ""),
                    "analista": row_dict.get("analista", ""),
                    "campus": row_dict.get("campus", ""),
                    "status": status,
                }
            )

        browser.close()

    # ─── RELATÓRIO FINAL ───
    print(f"\n{'═' * 60}")
    print("📊 RELATÓRIO FINAL")
    print(f"{'═' * 60}")

    contagem: dict[str, int] = {}
    for r in resultados:
        contagem[r["status"]] = contagem.get(r["status"], 0) + 1

    icones = {
        STATUS_OK: "✅",
        STATUS_PULADO: "ℹ️ ",
    }
    for status, qtd in sorted(contagem.items()):
        icone = icones.get(status, "❌")
        print(f"  {icone} {status}: {qtd}")

    erros = [
        r for r in resultados if r["status"] not in (STATUS_OK, STATUS_PULADO)
    ]
    if erros:
        print(f"\n⚠️  Registros com erro ({len(erros)}):")
        print(f"  {'Linha':>5}  {'Professor':<32}  {'Analista':<32}  Status")
        print(f"  {'─'*5}  {'─'*32}  {'─'*32}  {'─'*20}")
        for r in erros:
            print(
                f"  {r['linha']:>5}  {r['professor'][:32]:<32}  "
                f"{r['analista'][:32]:<32}  {r['status']}"
            )

    print(f"\nTotal: {len(resultados)} | OK: {contagem.get(STATUS_OK, 0)} | "
          f"Pulados: {contagem.get(STATUS_PULADO, 0)} | "
          f"Erros: {sum(v for k, v in contagem.items() if k not in (STATUS_OK, STATUS_PULADO))}")
    print(f"{'═' * 60}")


if __name__ == "__main__":
    main()
