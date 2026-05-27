# RPA — Inclusão de Analista por Professor

Automação web para vincular analistas a professores no sistema **Agenda Docente** da PUCPR,
acessando a página de *Associação de Coordenador* e processando os registros de uma planilha Excel.

---

## Visão geral

O sistema Agenda Docente exige que cada professor tenha analistas associadas manualmente,
uma a uma, por uma interface web. Este script automatiza esse processo lendo uma planilha
com os pares professor ↔ analista e executando o fluxo completo no navegador.

---

## Pré-requisitos

- Python 3.10+
- Pacotes Python:

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

- Acesso à rede PUCPR (VPN ou rede interna) com sessão SSO ativa no navegador
- Arquivo `professores.xlsx` na mesma pasta do script

---

## Estrutura do projeto

```
rpa-inclusao-analista-prof/
├── main.py            # Script principal da automação
├── professores.xlsx   # Planilha de entrada (base consolidada)
└── README.md
```

---

## Planilha de entrada (`professores.xlsx`)

| Coluna | Descrição | Exemplo |
|---|---|---|
| `PROFESSOR` | Nome completo em maiúsculas | `HELIO JOSE BAMBERG` |
| `CAMPUS` | Nome do campus | `Curitiba` |
| `ESCOLA` | Nome da escola/unidade | `MEDICINA` |
| `ANALISTA` | Nome completo em maiúsculas | `LUCIANA GUIDARINI DOS SANTOS` |
| `CHAVE_RELACIONAMENTO` | Chave única do par professor-analista | `CAMPUS: LONDRINA` |

> As colunas `CAMPUS`, `ESCOLA` e `CHAVE_RELACIONAMENTO` são carregadas mas não utilizadas
> na automação atual — estão disponíveis para filtragens ou validações futuras.

---

## Como executar

```bash
python main.py
```

O script abre o navegador Chromium, aguarda a autenticação SSO automática e começa a
processar os registros da planilha. Não é necessário nenhuma interação manual após iniciar.

### Configurações no topo do `main.py`

| Variável | Padrão | Descrição |
|---|---|---|
| `EXCEL_FILE` | `professores.xlsx` | Nome do arquivo Excel de entrada |
| `URL` | `https://agendadocente.pucpr.br/...` | URL do sistema |
| `HEADLESS` | `False` | `True` para rodar sem abrir o navegador |

---

## Fluxo da automação

Para cada linha da planilha, o script executa:

```
1. Selecionar professor no campo de busca (Vue Select)
      ↓
2. Verificar se a analista já está atribuída ao professor
   → Se sim: pula o registro (PULADO)
      ↓
3. Pesquisar a analista na tabela "Coordenador disponível"
   → Se não encontrada: registra erro (ERRO_ANALISTA)
      ↓
4. Marcar o checkbox da analista
      ↓
5. Clicar em "Adicionar"
      ↓
6. Clicar em "Salvar"
      ↓
7. Confirmar modal "Salvo com sucesso."
      ↓
8. Limpar o campo de professor → próximo registro
```

---

## Regras de negócio

- Se a analista **já estiver atribuída** ao professor → registro pulado (`PULADO`)
- Se o professor **não for encontrado** no dropdown → erro registrado, continua (`ERRO_PROFESSOR`)
- Se a analista **não for encontrada** na lista disponível → erro registrado, continua (`ERRO_ANALISTA`)
- Erros inesperados são capturados individualmente — o script **não para**, continua o próximo registro

---

## Relatório final

Ao término, o script exibe no terminal:

```
════════════════════════════════════════════════════════════
📊 RELATÓRIO FINAL
════════════════════════════════════════════════════════════
  ✅ OK: 6100
  ℹ️  PULADO: 1900
  ❌ ERRO_PROFESSOR: 45
  ❌ ERRO_ANALISTA: 83

⚠️  Registros com erro (128):
  Linha  Professor                          Analista                           Status
  ─────  ────────────────────────────────   ────────────────────────────────   ────────────────────
    102  HELIO JOSE BAMBERG                 ANALISTA INEXISTENTE                ERRO_ANALISTA
    ...

Total: 8628 | OK: 6100 | Pulados: 1900 | Erros: 128
════════════════════════════════════════════════════════════
```

### Códigos de status

| Status | Significado |
|---|---|
| `OK` | Analista associada com sucesso |
| `PULADO` | Analista já estava atribuída ao professor |
| `ERRO_PROFESSOR` | Professor não encontrado no dropdown |
| `ERRO_ANALISTA` | Analista não encontrada na lista disponível |
| `ERRO_ADICIONAR` | Botão "Adicionar" não encontrado ou desabilitado |
| `ERRO_SALVAR` | Botão "Salvar" não encontrado |
| `ERRO_INESPERADO` | Exceção não tratada durante o processamento |

---

## Estimativa de tempo de execução

Com 8.628 registros e ~15 segundos por linha em média (happy path):

| Cenário | Tempo estimado |
|---|---|
| Otimista (~10s/linha) | ~24 horas |
| Realista (~15s/linha) | ~36 horas |
| Pessimista (~20s/linha) | ~48 horas |

> **Atenção:** a sessão SSO expira em aproximadamente 8–12 horas. Para execuções longas,
> recomenda-se rodar em blocos menores (ex: 500 linhas por vez) controlando o offset de início.

---

## Detalhes técnicos

### Sistema alvo
- SPA (Single Page Application) em Vue.js
- Autenticação via SSO — sem login manual no script
- Campo de professor: componente **Vue Select** (seletores `.vs__*`)

### Seletores principais utilizados

| Elemento | Seletor |
|---|---|
| Campo professor | `input[placeholder='Professor']` |
| Dropdown de opções | `.vs__dropdown-menu li:has-text(...)` |
| Botão limpar professor | `button[aria-label='Clear'], .vs__clear` |
| Campo de busca analista | `input[placeholder='Pesquise aqui']` |
| Checkbox da analista | `input[type='checkbox']` dentro do `tr` |
| Botão Adicionar | `button:has-text('Adicionar'):not([disabled])` |
| Botão Salvar | `button:has-text('Salvar'):not([disabled])` |
| Modal de confirmação | `text=Salvo com sucesso.` |

### Detecção de "já atribuída"
A verificação usa JavaScript para checar o texto da página **somente após** o marcador
"Coordenador atribuído", evitando falso-positivo com a tabela da esquerda ("disponível"):

```javascript
const idx = texto.indexOf('Coordenador atribuído');
return texto.slice(idx).includes(nome);
```

---

## Possíveis melhorias futuras

- Suporte a **offset de início** para retomar a execução em blocos
- Detecção de **expiração de sessão SSO** com pausa e aviso
- **Agrupamento por professor** para evitar re-seleção em sequências do mesmo nome
- Exportação do relatório final para **Excel ou CSV**
- Modo `--dry-run` para validar a planilha sem executar alterações no sistema
