# ID01_2_ConcEst_TratarDados — Documentação do Projeto

> Documentação gerada a partir da leitura do código-fonte, do `Config.xlsx`, dos scripts SQL, do banco SQLite de template e do histórico Git do pacote `ID01_2_ConcEst_TratarDados.zip`. Nenhum arquivo do projeto foi alterado.

---

## 1. Visão geral

| Item | Valor |
|---|---|
| Nome | `ID01_2_ConcEst_TratarDados` |
| Descrição | Faz a conciliação dos estacionamentos |
| Cliente | Partage (`Details_Project.json`) |
| Ferramenta | Python RPA (framework no estilo REFramework) |
| Sistemas de interação | Excel, SQLite, sites (`Details_Project.json`) |
| Tipo de fila | SQLite |
| Versão (`VERSION`) | 1.0.0 |
| Equipe (README) | GP / RO / DEV: Felipe Mello |
| Repositório | `github.com/maurovmorais/ID01_2_ConcEst_TratarDados` (branch `main`) |
| Python | 3.13 (pelos `.pyc` `cpython-313`) |

**O que o robô faz:** é a **segunda etapa** da conciliação de estacionamentos. Ele consome os relatórios que a etapa anterior (`ID01_1_ConcEst_Download`) baixou de cada adquirente, normaliza tudo em um banco SQLite, cruza com o DEPARA de regras de negócio (empresas, formas de pagamento, taxas, prazos de compensação) e **monta a fila de trabalho** (`tbl_Fila_Proc_Performer`) que a etapa seguinte (`ID01_3_ConcEst_LanctoRPS`, lançamento no portal) vai consumir.

**Adquirentes tratados:** ConectCar, Cielo, Greenpass, SemParar, Veloe e Bradesco.

---

## 2. Posição no fluxo de conciliação

```
ID01_1_ConcEst_Download            ID01_2_ConcEst_TratarDados (este projeto)        ID01_3_ConcEst_LanctoRPS
─────────────────────────          ───────────────────────────────────────          ────────────────────────
Baixa relatórios dos      ──────►  Lê relatórios + DEPARA, trata dados,     ──────► Lê tbl_Fila_Proc_Performer
adquirentes                         grava no SQLite e cria a fila                    e lança no portal
(Arquivos_Baixados)                 do Performer                                     (banco: CaminhoBancoLanctoRPS)
```

Os pontos de contato são pastas e bancos configurados no `Config.xlsx` (seção 7).

---

## 3. Arquitetura

O projeto segue um framework com ciclo **Inicialização → Loop de itens → Finalização**.

```
__main__.py
   └─ bot.Bot.main() / Bot.action()
        ├─ Initialization.execute()   ← prepara ambiente e CHAMA InitAllApplications (onde está a ETL)
        ├─ LoopStation.execute()      ← processa itens da fila local (tbl_Fila_Processamento)
        │     └─ Process.execute()    ← cria a fila do Performer
        └─ EndProcess.execute()       ← fecha apps, relatórios, e-mails, finaliza task
```

### 3.1 Estrutura de pastas

```
ID01_2_ConcEst_TratarDados/
├── README.md · VERSION · requirements.txt · setup.py · MANIFEST.in
├── build.bat / build.sh            → python setup.py sdist
├── Details_Project.json            → metadados do projeto
└── ID01_2_ConcEst_TratarDados/     → pacote Python
    ├── __main__.py · bot.py
    ├── classes/
    │   ├── framework/   Initialization, InitAllApplications, InitAllSettings, LoopStation,
    │   │                GetTransaction, Process, EndProcess, CloseAllApplications, KillAllProcesses
    │   ├── sqlite/      config, sqlite (GerenciadorSQLite), manipular_tabelas,
    │   │                gravador_tbl_aux_dados, criar_fila_performer
    │   ├── excel/       excel (LeitorExcel), leitura_depara, leitura_relat_adquirentes,
    │   │                ExcelManagerOpenpyxl, ExcelManagerXlwings
    │   ├── queue/       QueueManager, QueueManagerPerformer
    │   ├── dados_execucao/  DadosExecucao (tbl_dados_execucao / tbl_dados_itens_fila)
    │   ├── relatorios/  Relatorios (analítico/sintético, CSV RAAS)
    │   ├── email/       send (SMTP, Outlook, Outlook API) · receive
    │   ├── databases_manager/  MariaDB, SQLite, SQLServer
    │   ├── utils/       Log, LogMethod, ExecutionControl, CredentialManager, BackupSqlite,
    │   │                GenericReusable, RobotStream, ScreenRecorder, Exceptions
    │   └── chrome/google/Homepage.py
    └── resources/
        ├── config/Config.xlsx
        ├── script_dados_estruturados/  dados_estruturados.sql (+ _old.sql)
        ├── scripts/analitico_sintetico/ SQLs dos relatórios
        ├── sqlite/banco_dados.db       banco template (vazio)
        ├── templates/  e-mails e relatórios .xlsx
        ├── logs/ · relatorios/ · videos/ · exception/ · robot_stream/
```

> As classes de framework (e-mail, relatórios, gravação de tela, RobotStream, gerenciadores de banco, Selenium em `InitAllSettings`) vêm do template. O trabalho específico deste projeto está em `classes/sqlite`, `classes/excel/leitura_*`, `resources/script_dados_estruturados` e nos pontos de extensão `InitAllApplications.add_to_queue` e `Process.execute`.

---

## 4. Fluxo de execução detalhado

### 4.1 Inicialização (`Initialization.execute`)
1. Registra info do sistema, uso da máquina e monitor no log.
2. Opcionalmente: inicia RobotStream, faz backup do SQLite, grava a tela, envia e-mail inicial (conforme `Config.xlsx`).
3. Registra a execução em `tbl_dados_execucao`.
4. Encerra `excel.exe` e `winword.exe` (`KillAllProcesses`).
5. Chama `InitAllApplications.execute(first_run=True)`, que dispara a **ETL** (4.2).

### 4.2 ETL — `InitAllApplications.add_to_queue`
Executa, nesta ordem:

| # | Passo | Código | Efeito |
|---|---|---|---|
| 1 | Ler DEPARA e popular tabelas | `processar_depara` | Lê 7 abas do `DEPARA_Estacionamento.xlsx` e **recria** (DROP + CREATE + INSERT) as tabelas de regra |
| 2 | Abandonar fila remanescente | `QueueManagerPerformer.abandon_queue` | Itens `NEW` em `tbl_Fila_Proc_Performer` viram `ABANDONED` |
| 3 | Ler relatórios dos adquirentes | `LeitorRelatoriosAdquirentes.processar_todos` | Normaliza os arquivos em um único DataFrame |
| 4 | Gravar `tbl_aux_dados` | `GravadorTblAuxDados.gravar_tbl_aux_dados` | `DELETE FROM` + carga (transação com rollback) |
| 5 | Popular `tbl_dados_estruturados` | `atualizar_dados_estruturados` | `DELETE FROM` + `INSERT … SELECT` do `dados_estruturados.sql` |
| 6 | Criar item único da fila local | `QueueManager.insert_new_queue_item` | 1 item em `tbl_Fila_Processamento`; referência = data de ontem (`dd/mm/aaaa`); info adicional = lista de adquirentes |

### 4.3 Loop (`LoopStation` → `Process.execute`)
- `GetTransaction` pega o item `NEW` da fila local e marca `ON QUEUE` → `RUNNING`.
- Até `MaxRetryNumber` (3) tentativas por item; erros de negócio (`BusinessRuleException`) não repetem; erros de sistema reiniciam as aplicações e tentam de novo.
- `Process.execute()` chama `criar_fila_performer(CaminhoBancoSqlite)`, que **insere os registros em `tbl_Fila_Proc_Performer`** (seção 6).
- O loop para após `MaxConsecutiveSystemExceptions` (3) falhas de sistema seguidas.

### 4.4 Finalização (`EndProcess.execute`)
Fecha aplicações e `chromedriver.exe`, finaliza gravação de tela, atualiza `tbl_dados_execucao`, gera relatórios analítico e sintético, envia e-mail final e RAAS (se habilitados) e chama `ExecutionControl.finish_task`.

---

## 5. Dados

### 5.1 Entradas
| Entrada | Origem (Config.xlsx) | Observação |
|---|---|---|
| `DEPARA_Estacionamento.xlsx` | `depara` | 7 abas: Empresas, FormaDePagamento, CondPagamt, Adquirentes, TaxaAdquirentes, TaxaCielo, Dinheiro |
| Relatórios dos adquirentes (`.csv/.xlsx/.xls`) | `arquivos_baixados` (`...\ID01_1_ConcEst_Download\Arquivos_Baixados`) | Arquivo localizado pelo nome (ex.: `cielo`, `semparar`), ignorando maiúsculas e símbolos; um adquirente pode ter vários arquivos |

### 5.2 Tabelas do SQLite (`CaminhoBancoSqlite`)

**Tabelas de regra (recriadas a cada execução a partir do DEPARA):**

| Tabela | Aba de origem | Colunas principais |
|---|---|---|
| `tbl_empresas` | Empresas | ID_UNICO, Razao Social, Nomenclatura SoftCase, CNPJ, SIGLA, e uma coluna por adquirente (Cielo, SemParar, Greenpass, ConectCar, Veloe, Bradesco), Observação |
| `tbl_FormaDePagamento` | FormaDePagamento | Forma de Pagamento, Tipo, Bandeira, Valida, Não atuar, Forma de Pagamento Principal, Origem |
| `tbl_CondPagamt` | CondPagamt | Tipo de Lançamento, Dias Comp. |
| `tbl_Adquirentes` | Adquirentes | Lista de Adquirentes, Extração via, Observação |
| `tbl_TaxaAdquirentes` | TaxaAdquirentes | ID_UNICO, Sigla Empresa, CONECTCAR, VELOE, GREENPASS, SEMPARAR (REAL) |
| `tbl_TaxaCielo` | TaxaCielo | TAXAS CIELO, Bandeira, Taxa (TEXT: mistura número e texto, p.ex. PIX) |
| `tbl_Dinheiro` | Dinheiro | Dia da Semana, Dias Comp. Dinheiro |

**Tabelas de dados:**

| Tabela | Alimentada por | Conteúdo |
|---|---|---|
| `tbl_aux_dados` | `GravadorTblAuxDados` | Dados normalizados dos adquirentes: `adquirente, empresa, data_processamento, forma_pagto, Bandeira, valor_lancamento, valor_taxa, data_atualizacao` |
| `tbl_dados_estruturados` | `dados_estruturados.sql` | `tbl_aux_dados` cruzada com empresas e taxas (ver 5.4) |
| `tbl_Fila_Processamento` | `QueueManager` | Fila local do framework (1 item por execução) |
| `tbl_Fila_Proc_Performer` | `criar_fila_performer` | Fila de saída, consumida pelo Performer |
| `tbl_dados_execucao`, `tbl_dados_itens_fila` | `DadosExecucao` | Telemetria da execução (base dos relatórios analítico/sintético) |

> ⚠️ O `banco_dados.db` que vai dentro do pacote é um **template**: contém apenas `tbl_Fila_Processamento`, `tbl_dados_execucao` e `tbl_dados_itens_fila`, todas vazias. As tabelas `tbl_aux_dados`, `tbl_dados_estruturados` e `tbl_Fila_Proc_Performer` **não são criadas por nenhum código deste projeto** e precisam existir no banco real (`C:\Armazenamento\...\banco_dados.db`).

### 5.3 Normalização dos relatórios (`leitura_relat_adquirentes.py`)

Cada adquirente é descrito em `MAPEAMENTO_ADQUIRENTES` (colunas do relatório → colunas de `tbl_aux_dados`):

| Adquirente | Empresa | Data | Forma pagto | Bandeira | Valor lançamento | Valor taxa |
|---|---|---|---|---|---|---|
| ConectCar | Conveniado | Data do Processamento | — | — | Valor Cobrado | Tarifa de interconexão |
| Cielo | CPF/CNPJ do estabelecimento (só dígitos) | Data da venda | Forma de pagamento | Bandeira | Valor bruto | Taxa/tarifa |
| Greenpass | CONVENIADO | DATA ESTADIA | — | — | VALOR TOTAL | — |
| SemParar | credenciado | data_periodo | — | — | valor_total | — |
| Veloe | Estabelecimento | Data Saída | Tipo | — | Valor Transação | Valor MDR |
| Bradesco | inscricao_empresa | data_lancamento | historico | — | valor_assinado | — |

Comportamentos relevantes:
- **CSV:** testa codificações (`utf-8`, `utf-8-sig`, `cp1252`, `latin1`) e separadores (`,` `;` `\t` `|`) até o cabeçalho esperado bater.
- **Cabeçalho:** varre até 60 linhas procurando a linha cujas duas primeiras colunas são as de identificação do adquirente (relatórios têm preâmbulo). Nomes duplicados recebem sufixo `__duplicada_N`.
- **Valores:** converte formato brasileiro (`R$ 1.234,56` → `1234.56`).
- **Datas:** converte para ISO; intervalos `dd/mm/aaaa - dd/mm/aaaa` usam a primeira data (SemParar).
- **Rodapés:** linhas sem data válida são descartadas (com aviso no log).
- **Tolerância a falhas:** adquirente sem arquivo é ignorado com `WARNING`; falha de parse (`ParserError/ValueError/KeyError`) registra `ERROR` e segue para o próximo adquirente.

### 5.4 Regras de negócio do `dados_estruturados.sql`

- **Vínculo da empresa:** pega o texto de `aux.empresa` (se contiver `" - "`, usa o trecho após o primeiro `" - "`) e compara com a coluna do adquirente em `tbl_empresas` (`Cielo`, `Veloe`, `SemParar`, `Greenpass`, `ConectCar`, `Bradesco`).
- **`forma_pagto` normalizada:**
  - `TAG` para Veloe com `CREDITO` e para ConectCar/Greenpass/SemParar/Veloe sem forma de pagamento;
  - `DEBITO` para `Débito à vista` e `Débito pré-pago`;
  - `PIX` para Bradesco `CRED PIX QR CODE DINAMIC`.
- **`Dia_Comp` (dias de compensação):** Crédito à vista = 31; Crédito pré-pago = 2; Crédito conversor de moedas = 5; demais = `NULL` (dinheiro é tratado na fila, por dia da semana).
- **`taxa_adquirente`:** Cielo vem de `tbl_TaxaCielo` (por tipo e bandeira); os demais vêm de `tbl_TaxaAdquirentes` pela sigla da empresa. Resultado em percentual com vírgula decimal.
- **Exclusão:** Veloe com `DEBITO` é descartado.
- **PIX Bradesco:** quando existem, para a mesma empresa/sigla/data, `CRED PIX QR CODE DINAMIC` e `DEVOLUCOES PIX`, os dois são somados em **uma linha `PIX`**.
- `valor_taxa` é gravado em valor absoluto.

---

## 6. Fila do Performer (`criar_fila_performer.py`)

1. Busca combinações distintas `(adquirente, forma_pagto)` em `tbl_dados_estruturados`.
2. Para cada combinação, executa `QUERY_TOTAL_POR_FORMA`, que agrupa por Nomenclatura SoftCase, bandeira, taxa, forma e data, e calcula:
   - `total` = soma de `valor_lancamento` (itens com total zero são descartados);
   - `valor_taxa`: para **Greenpass e SemParar** = `total × taxa_adquirente / 100`; para os demais = soma de `valor_taxa`;
   - `Dia_Comp`: para `DINHEIRO`, vem de `tbl_Dinheiro` pelo dia da semana; senão, o `Dia_Comp` da linha;
   - `forma_pagamento`: `TAG` → nome do adquirente; demais → `tbl_FormaDePagamento` (somente `Forma de Pagamento Principal = true`), casando por Origem (extraída dos parênteses da Nomenclatura SoftCase), bandeira e tipo; `DINHEIRO` sem bandeira usa a forma principal de dinheiro da Origem;
   - `data_ref` = `aaaa-mm-dd - DIA_DA_SEMANA`.
3. Gera **um registro de fila por item**, com `referencia = "Adquirente/forma_pagto"`, `status = 'NEW'` e `info_adicionais` em JSON (lista com 1 item) contendo: `adquirente, softcase, bandeira, dias_comp, valor_taxa, taxa_adquirente, forma_pagto, forma_pagamento, valor, data_ref` (números com vírgula decimal).
4. Insere tudo em `tbl_Fila_Proc_Performer`.

Existe também `montar_registros_fila` (um registro por adquirente com todos os itens), hoje **não utilizado** (chamada comentada).

---

## 7. Configuração (`resources/config/Config.xlsx`)

Abas: `Settings`, `Constants`, `Credentials` (vazia), `Assets` (vazia). Valores relevantes lidos:

| Chave | Valor atual | Uso |
|---|---|---|
| `CaminhoBancoSqlite` | `C:\Armazenamento\ID01_2_ConcEst_TratarDados\Banco Dados\banco_dados.db` | Banco principal |
| `CaminhoBancoLanctoRPS` | `C:\Armazenamento\ID01_3_ConcEst_LanctoRPS\Banco Dados\banco_dados.db` | Banco da etapa 3 |
| `depara` | `C:\Armazenamento\ID01_2_ConcEst_TratarDados\DePara\DEPARA_Estacionamento.xlsx` | Regras de negócio |
| `arquivos_baixados` | `C:\Armazenamento\ID01_1_ConcEst_Download\Arquivos_Baixados` | Relatórios dos adquirentes |
| `FilaProcessamento` | *(vazio → padrão `tbl_Fila_Processamento`)* | Fila local |
| `FilaProcessamentoPerformer` | `tbl_Fila_Proc_Performer` | Fila de saída |
| `CaminhoPastaRelatorios`, `CaminhoPastaLogs`, `CaminhoExceptionScreenshots`, `CaminhoSalvarVideo`, `CaminhoAnexo` | pastas em `C:\Armazenamento\ID01_2_ConcEst_TratarDados\...` | Saídas |
| `MaxRetryNumber` / `MaxConsecutiveSystemExceptions` / `MaxTimeoutRequests` | 3 / 3 / 30 | Robustez |
| `EmailInicial/CadaErro/ErroInicializacao/Final`, `RelatorioRAAS`, `GravarTela`, `IniciarRobotStream`, `BackupSqlite`, `CapturarScreenshot`, `AtivarLogs` | todos `NÃO` | Recursos opcionais desligados |
| `NomeCliente` | `IAFastLab` | Identificação nos relatórios |

---

## 8. Execução e instalação

**Requisitos:** Windows (usa `taskkill`, caminhos com `\` e Excel/Word), Python 3.13, acesso às pastas de `C:\Armazenamento\`.

```bash
# dentro da pasta do projeto, com venv ativo
pip install -r requirements.txt

# execução local
python -m ID01_2_ConcEst_TratarDados

# empacotar (gera sdist)
build.bat        # Windows
./build.sh       # Linux/macOS
```

`__main__.py` chama `Bot.action()` quando recebe `--execution` com 5+ argumentos; caso contrário chama `Bot.main()`.

**Dependências (`requirements.txt`):** selenium 4.28.1, webdriver-manager 4.0.2, pandas 2.3.1, openpyxl 3.1.5, pyodbc 5.2.0, opencv-python 4.12.0.88, pyautogui 0.9.54, msal 1.32.3, beautifulsoup4 4.13.4, screeninfo 0.8.1, xlwings 0.33.15, setuptools 80.9.0, mysql-connector-python 9.3.0, psutil 6.1.1, numpy 2.2.1.

**Logs:** `resources/logs/execucao_AAAAMMDD.log` (também há o parâmetro `CaminhoPastaLogs`). O log de 08/10/2026 mostra o DEPARA lido com 16 empresas, 108 formas de pagamento, 3 condições, 6 adquirentes, 9 taxas, 14 taxas Cielo e 7 linhas de dinheiro, e termina com `End Process Finished`.

---

## 9. Tratamento de erros e observabilidade

- Exceções próprias: `BusinessRuleException` (sem retry) e `TerminateException` (encerra o item como sucesso).
- Cada item registra tentativas, screenshot (se habilitado) e status em `tbl_dados_itens_fila`.
- `ExecutionControl` é local: **não há orquestrador integrado** (`is_interrupted()` sempre retorna `False`; `finish_task` e `send_error` só registram em log). Há um ponto de extensão documentado no código.
- `EndProcess` gera relatório analítico e sintético a partir de `tbl_dados_execucao`/`tbl_dados_itens_fila`.

---

## 10. Histórico Git (branch `main`)

| Data | Commit |
|---|---|
| 13/09/2026 | first commit |
| 17/09/2026 | adiciona leitura de arquivos e gravação tabela · corrigi lógica do process |
| 19/09/2026 | ajusta leitura relat adquirentes cielo |
| 25/09/2026 | adiciona adquirente bradesco |
| 28/09/2026 | adiciona tratamento dados estruturados |
| 29/09/2026 | adiciona lógica de criação da fila Performer · corrigi select que cria a fila |
| 30/09/2026 | corrigi forma_pagto bradesco |
| 01/10/2026 | adiciona dias_comp na fila |
| 02/10/2026 | ajusta dias_comp para dinheiro |
| 05/10/2026 | corrigi data_ref (agora pega do arquivo adquirente) |
| 06/10/2026 | ajusta data_ref |

Observação: `InitAllApplications.py` e `dados_estruturados.sql` têm data de modificação em 08/10/2026, posterior ao último commit (06/10). Pode haver alterações ainda não commitadas.

---

## 11. Pontos de atenção e pendências identificadas

1. **Docstring desatualizada:** `criar_fila_performer` diz que a inserção está comentada, mas o `executemany` está **ativo**. A fila é gravada de fato.
2. **Tabelas fora do código:** `tbl_aux_dados`, `tbl_dados_estruturados` e `tbl_Fila_Proc_Performer` não têm `CREATE TABLE` no projeto nem no `.db` de template. Recomenda-se documentar/versionar o DDL.
3. **Sem deduplicação da fila:** cada execução insere novos registros `NEW` e abandona os anteriores; reexecutar no mesmo dia gera nova leva (a anterior fica `ABANDONED`).
4. **Variável sem uso efetivo:** `data_processamento` (ontem) só serve de referência do item local; as datas dos registros vêm dos arquivos dos adquirentes.
5. **Caminhos fixos:** `Config.xlsx` e `InitAllSettings` usam `\` e `C:\Armazenamento`; não roda em Linux sem ajuste.
6. **`dados_estruturados_old.sql`** permanece em `resources` (versão anterior do script).
7. **README com placeholders:** seção de releases ainda com texto-modelo; `MANIFEST.in` referencia `README.rst`, mas o arquivo é `README.md`.
8. **Classes herdadas do template** (BotCity Maestro removido, ExcelManagerXlwings/Openpyxl, bancos MariaDB/SQL Server, e-mail) não são usadas pelo fluxo atual; avaliar remoção para reduzir superfície.
9. **Sem testes automatizados** no pacote (não há pasta `tests/`).
10. **Credenciais:** abas `Credentials` e `Assets` do `Config.xlsx` estão vazias e nenhuma credencial aparece no código lido; manter assim (usar `CredentialManager`/cofre se surgir necessidade).

---

## 12. Glossário

| Termo | Significado |
|---|---|
| DEPARA | Planilha de regras de negócio (empresas, formas de pagamento, taxas, prazos) |
| Adquirente | Fonte do relatório: ConectCar, Cielo, Greenpass, SemParar, Veloe, Bradesco |
| Performer | Robô que consome a fila gerada aqui (etapa `ID01_3_ConcEst_LanctoRPS`) |
| SoftCase / RPS | Sistema/portal de destino dos lançamentos |
| Dia_Comp | Dias de compensação do recebimento |
| TAG | Forma de pagamento por tag de pedágio/estacionamento (ConectCar, Greenpass, SemParar, Veloe) |
| RAAS | Relatório de execução enviado por e-mail/CSV (desligado na configuração atual) |
