# Conciliador Bancário — Advbox

Automação da **conciliação bancária mensal** de escritórios de advocacia que usam o
**Advbox**. Lê os **extratos dos bancos** (OFX / PDF) + a **planilha exportada do Financeiro
do Advbox**, cruza os dois e gera uma planilha de conciliação com três listas de ação:

- **Dar baixa** — o movimento já existe no Advbox (é só marcar o pagamento).
- **Criar lançamento** — não existe: já classificado e desdobrado (honorários + repasse).
- **Revisar** — casos ambíguos que precisam de decisão humana.

Opcionalmente, envia os lançamentos aprovados direto para o Advbox via API (com modo
`dry-run` por padrão, para simular antes de postar de verdade).

> Base reutilizável — mantida pela **Método Adv Digital**. As contas, categorias e regras
> de classificação vêm com exemplos genéricos: ajuste-as para o Advbox de cada escritório
> (ver a seção **Personalização**).

## Como rodar

### Modo "pasta" (usuário final, sem código)
1. Coloque na mesma pasta: os extratos (`.ofx`/`.pdf`) e o export do Advbox (`.xlsx`).
2. Nomeie cada extrato com o **nome da conta** (o nome do arquivo vira a conta na planilha).
   Ex.: `BANCO DO BRASIL.ofx`, `ASAAS.ofx`, Caixa em `.pdf`.
3. Dê dois cliques em `EXECUTAR_CONCILIACAO.bat` (Windows) — ou rode `python index.py .`.
   Gera `conciliacao_AAAA-MM-DD.xlsx` e abre.

Veja o passo a passo detalhado para o usuário não técnico em [`LEIA-ME.txt`](LEIA-ME.txt).

### Interface web (Flask)
```bash
pip install -r requirements.txt
python app.py            # http://localhost:5000
```
Suba os extratos + o export do Advbox pela tela, revise/edite os itens e baixe o Excel
ou envie para o Advbox.

### Docker
```bash
docker compose up -d     # publica atrás do Caddy (ver compose.yaml / Caddyfile)
```

## Configuração da API (opcional)
Copie `.env.example` para `.env` e preencha:
```
ADVBOX_TOKEN=...         # Advbox > Conta e assinatura > Nossos produtos > API ADVBOX
ADVBOX_DRY_RUN=true      # true = simula; mude para false para enviar de verdade
```
Nunca commite o `.env` nem o token.

## Personalização (por escritório)
- `app.py` → `CONTAS` e `MAPA_CONTA`: nomes das contas do Advbox e o mapeamento
  palavra-chave-do-arquivo → conta.
- `app.py` → `CENTROS_CUSTO`, `SETORES`, `MAPA_CENTRO_CUSTO`: estrutura de centros de custo.
- `index.py` → `_RULES`: as ~77 regras (regex do texto do extrato → categoria Advbox) e o
  `PERC_HONORARIOS_PADRAO` (percentual de honorários no desdobramento de alvará/RPV).

## Estrutura
- `index.py` — o motor de conciliação (parsers OFX/PDF, rulebook, matching, geração do Excel).
- `app.py` — interface web (Flask).
- `advbox_client.py` — cliente da API do Advbox (criar/dar baixa, com dry-run).
- `templates/` — telas web (upload, resultado, configuração).
- `EXECUTAR_CONCILIACAO.bat` — lançador de duplo-clique (Windows).
- `HANDOFF.md` — notas técnicas e backlog.
