# Conciliação Bancária — Advbox · Handoff

## O que estamos construindo
Uma automação da **conciliação bancária mensal** de um escritório de advocacia que usa o
**Advbox** (sistema jurídico/financeiro). Objetivo: parar de conferir o extrato linha por
linha à mão. O programa lê os **extratos dos bancos** + a **planilha exportada do Financeiro
do Advbox**, cruza os dois e gera uma **planilha de conciliação** com três listas de ação:
o que dar baixa (já existe no sistema), o que criar (não existe) e o que revisar (ambíguo).

Regras de negócio já embutidas (do "passo a passo" do escritório):
- Classificação por categoria a partir do texto do extrato (~77 regras).
- **Desdobramento** de alvará / RPV / precatório / levantamento judicial em 3 lançamentos:
  valor identificado + **30% honorários** + **repasse ao cliente**.
- Casos de risco (Pix genérico, migração de conta, % variável) vão para revisão humana.

## Estado atual (validado em dados reais de abril/2026)
Rodado contra os 4 bancos do cliente, cruzando com o export real do Advbox:

| Conta | Movimentos | Dar baixa | Criar | Revisar |
|---|---|---|---|---|
| BB (OFX) | 181 | 159 | 9 | 13 |
| Asaas (OFX) | 73 | 72 | 0 | 1 |
| Itaú (OFX) | 35 | 30 | 3 | 2 |
| Caixa (PDF) | 5 | 4 | 3 | 0 |
| **Total** | **294** | **265** | **15** | **16** |

Trabalho humano cai de 294 linhas para ~31. O Asaas quase se concilia sozinho por causa da
integração nativa Asaas↔Advbox. A Caixa só fornece PDF (sem OFX) e é lida direto do PDF.

## Como rodar
Requisitos: Python 3.11+, e `poppler` (para o `pdftotext`, usado no leitor da Caixa).
```bash
pip install pandas openpyxl
# Windows: instalar poppler e por no PATH  |  macOS: brew install poppler  |  Linux: apt install poppler-utils
```
Modo "pasta" (o usuário final usa assim, via EXECUTAR_CONCILIACAO.bat):
1. Coloque na mesma pasta: os extratos e o export do Advbox (.xlsx).
2. Nomeie cada extrato com o NOME DA CONTA — o nome do arquivo vira a conta na planilha.
   Ex.: `BANCO DO BRASIL.ofx`, `ASAAS.ofx`, Caixa em `.pdf`.
3. `python index.py .`  → gera `conciliacao_AAAA-MM-DD.xlsx` e abre.

Uso como biblioteca:
```python
import index as C
sistema = C.carregar_advbox_export("Advbox-export.xlsx")
ext  = C.carregar_ofx("extrato_bb.ofx", "BANCO DO BRASIL")
ext += C.carregar_caixa_pdf("caixa.pdf", "CAIXA ECONOMICA")
C.gerar_planilha(ext, sistema, "conciliacao.xlsx")
```

## Estrutura do código (`index.py`)
- `_RULES` / `classificar()` — rulebook regex (texto do extrato → categoria Advbox).
- `desdobrar()` — gera os 3 lançamentos de alvará/RPV (30% honorários + repasse).
- `carregar_ofx()` — parser OFX (BB, Itaú, Asaas). Ignora linhas de saldo e valor zero.
- `carregar_caixa_pdf()` — parser do PDF da Caixa (usa `pdftotext -layout`).
- `carregar_advbox_export()` — lê o Excel exportado do Financeiro do Advbox.
- `conciliar()` — matching por valor + janela de datas (±5 dias) + similaridade de texto.
- `gerar_planilha()` — escreve o .xlsx (abas Resumo / Dar baixa / Criar / Revisar).
- `executar_pasta()` — modo "clica e roda": varre a pasta e processa tudo.

## API do Advbox — próximo grande passo
Documentação oficial confirmada (docs em `api.softwareadvbox.com.br/docs`, base
`https://app.advbox.com.br/api/v1`, auth `Authorization: Bearer <token>`). Endpoints úteis:
- `GET /settings` — **chamar primeiro**; traz os IDs de conta, categoria e centro de custo
  (a API usa IDs numéricos, não os nomes de texto). Montar um mapa nome→ID.
- `POST /transactions` — criar lançamento (receita/despesa) = a aba "Criar".
- `PUT /transactions/{id}` — atualizar/dar baixa (marcar pagamento) = a aba "Dar baixa".

Observações: acesso pode exigir liberação de "parceiro/integrador" pelo Advbox e plano
compatível (verificar no comercial). Havia um `AdvboxClient` na 1ª versão do projeto (modo
dry-run) que serve de ponto de partida para reescrever com estes endpoints reais.

## Backlog priorizado
1. **Camada de API (com aprovação):** reintroduzir `AdvboxClient` usando os endpoints acima.
   - `GET /settings` → cache nome→ID (conta, categoria, centro de custo).
   - Substituir a aba "Dar baixa"/"Criar" por chamadas reais, **mas** atrás de um gate de
     aprovação: mostrar "vou dar baixa em N e criar M, confirma?" antes de postar.
   - Manter `dry_run=True` como padrão. Nunca postar automaticamente a aba "Revisar".
2. **Checagem de saldo:** ler o saldo final do OFX (`<LEDGERBAL><BALAMT>`) e da Caixa (última
   linha do PDF) e comparar com o saldo pós-conciliação → sinal de "fechou / não fechou".
3. **Reduzir a fila "Revisar":** mais regras para Itaú/Asaas; tratar transferências entre
   contas como pares espelhados (débito numa conta = crédito noutra).
4. **Portabilidade do leitor da Caixa:** trocar `pdftotext` por `pdfplumber` (pip) para não
   depender de poppler no Windows.
5. **Empacotar para o usuário final:** validar o `.bat`; avaliar `.exe` (pyinstaller) ou a
   versão em Google Colab (o escritório é 100% Google Workspace, sem servidor).
6. **Alternativa sem código:** a Advbox tem integração com **n8n** — avaliar montar a
   criação/baixa via n8n como opção de manutenção mais simples para o cliente.

## Segurança
- O cliente compartilhou um token da API em texto aberto durante a conversa. **Gerar um token
  novo e descartar o antigo** (Advbox → Conta e assinatura → Nossos produtos → API ADVBOX).
- Guardar o token em variável de ambiente / `.env` (fora do git). Nunca commitar token.

## Arquivos do projeto
- `index.py` — o motor.
- `EXECUTAR_CONCILIACAO.bat` — lançador de duplo-clique (Windows).
- `LEIA-ME.txt` — instruções para o usuário final (não técnico).
- `conciliacao_abril_4contas.xlsx` — saída de exemplo (abril real, 4 contas).


