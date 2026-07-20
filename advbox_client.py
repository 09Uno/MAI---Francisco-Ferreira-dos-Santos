# -*- coding: utf-8 -*-
"""
Client da API do Advbox (v1)
Docs: api.softwareadvbox.com.br/docs
Base: https://app.advbox.com.br/api/v1
Auth: Authorization: Bearer <token>
"""

import os
import json
import logging
from datetime import datetime
from typing import Optional

import requests

logger = logging.getLogger(__name__)

BASE_URL = "https://app.advbox.com.br/api/v1"


class AdvboxAPIError(Exception):
    """Erro da API do Advbox."""
    def __init__(self, status_code, message, response_body=None):
        self.status_code = status_code
        self.response_body = response_body
        super().__init__(f"Advbox API {status_code}: {message}")


class AdvboxClient:
    """
    Client para a API do Advbox.

    Uso:
        client = AdvboxClient(token="...", dry_run=True)
        client.carregar_settings()  # mapeia nomes → IDs
        client.criar_lancamento(...)
        client.dar_baixa(id, data_pagamento)
    """

    def __init__(self, token: Optional[str] = None, dry_run: bool = True):
        self.token = token or os.environ.get("ADVBOX_TOKEN", "")
        self.dry_run = dry_run

        self.contas = {}        # nome da conta (banco) → ID
        self.categorias = {}    # nome da categoria → ID
        self.centros_custo = {} # nome do centro de custo → ID
        self.departamentos = {} # nome do departamento/setor → ID
        self.usuarios = []      # lista de {id, name}
        self._default_user_id = None
        self._settings_raw = None

    @property
    def configurado(self) -> bool:
        return bool(self.token)

    @property
    def settings_carregados(self) -> bool:
        return bool(self.contas)

    def _headers(self):
        return {
            "Authorization": f"Bearer {self.token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "ConciliadorBancario/1.0",
        }

    def _request(self, method, endpoint, data=None):
        """Faz uma requisição à API. Em dry_run, loga e retorna sem postar."""
        url = f"{BASE_URL}{endpoint}"

        if self.dry_run and method in ("POST", "PUT", "DELETE"):
            logger.info(f"[DRY RUN] {method} {url}")
            if data:
                logger.info(f"[DRY RUN] Payload: {json.dumps(data, ensure_ascii=False, indent=2)}")
            return {
                "_dry_run": True,
                "_method": method,
                "_url": url,
                "_payload": data,
            }

        try:
            resp = requests.request(method, url, headers=self._headers(),
                                    json=data, timeout=30)
        except requests.RequestException as e:
            raise AdvboxAPIError(0, f"Erro de conexão: {e}")

        if resp.status_code >= 400:
            try:
                body = resp.json()
            except Exception:
                body = resp.text
            if resp.status_code == 403:
                raise AdvboxAPIError(403,
                    "Acesso negado. Verifique se o token da API está correto e completo.",
                    body)
            if resp.status_code == 401:
                raise AdvboxAPIError(401,
                    "Token inválido ou expirado. Gere um novo token no Advbox.",
                    body)
            raise AdvboxAPIError(resp.status_code, str(body), body)

        if resp.status_code == 204:
            return {}
        return resp.json()

    # ----------------------------------------------------------------
    # GET /settings — carrega mapas nome→ID
    # ----------------------------------------------------------------
    def carregar_settings(self):
        data = self._request("GET", "/settings")
        self._settings_raw = data

        # Usuários (top-level)
        for u in data.get("users", []):
            uid = u.get("id")
            nome = u.get("name", "")
            if uid:
                self.usuarios.append({"id": uid, "name": nome})
        if self.usuarios:
            self._default_user_id = self.usuarios[0]["id"]

        # Dados financeiros (aninhados em "financial")
        fin = data.get("financial", {})

        # Contas bancárias: financial.banks[].{id, name}
        for banco in fin.get("banks", []):
            nome = banco.get("name", "")
            bid = banco.get("id")
            if nome and bid:
                self.contas[nome.strip().upper()] = bid

        # Categorias: financial.categories[].{id, category, type}
        for cat in fin.get("categories", []):
            nome = cat.get("category", "")
            cid = cat.get("id")
            if nome and cid:
                self.categorias[nome.strip().upper()] = cid

        # Centros de custo: financial.cost_centers[].{id, cost_center}
        for cc in fin.get("cost_centers", []):
            nome = cc.get("cost_center", "")
            cid = cc.get("id")
            if nome and cid:
                self.centros_custo[nome.strip().upper()] = cid

        # Departamentos/Setores: financial.departments[].{id, department}
        for dep in fin.get("departments", []):
            nome = dep.get("department", "")
            did = dep.get("id")
            if nome and did:
                self.departamentos[nome.strip().upper()] = did

        logger.info(f"Settings carregados: {len(self.contas)} contas, "
                     f"{len(self.categorias)} categorias, "
                     f"{len(self.centros_custo)} centros de custo, "
                     f"{len(self.departamentos)} departamentos, "
                     f"{len(self.usuarios)} usuários")
        return {
            "contas": len(self.contas),
            "categorias": len(self.categorias),
            "centros_custo": len(self.centros_custo),
        }

    def _resolver_id(self, mapa, nome, tipo_label):
        """Busca o ID correspondente a um nome. Faz match case-insensitive."""
        if not nome:
            return None
        key = nome.strip().upper()
        cid = mapa.get(key)
        if cid is None:
            for k, v in mapa.items():
                if key in k or k in key:
                    return v
            logger.warning(f"{tipo_label} não encontrado nos settings: '{nome}'")
        return cid

    # ----------------------------------------------------------------
    # POST /transactions — criar lançamento
    # ----------------------------------------------------------------
    def criar_lancamento(self, tipo, categoria, descricao, valor,
                         data_vencimento, data_pagamento=None,
                         conta=None, centro_custo=None, setor=None,
                         **_kwargs):
        conta_id = self._resolver_id(self.contas, conta, "Conta")
        categoria_id = self._resolver_id(self.categorias, categoria, "Categoria")
        cc_id = self._resolver_id(self.centros_custo, centro_custo, "Centro de custo")
        setor_id = self._resolver_id(self.departamentos, setor, "Setor")

        entry_type = "income" if tipo.upper() in ("RECEITA", "INCOME") else "expense"

        data_iso = _formatar_data(data_vencimento)

        payload = {
            "entry_type": entry_type,
            "description": descricao or "",
            "amount": _formatar_valor(valor),
            "date_due": data_iso,
            "competence": _competencia(data_iso),
        }

        if self._default_user_id:
            payload["users_id"] = self._default_user_id
        if conta_id:
            payload["debit_account"] = conta_id
        if categoria_id:
            payload["categories_id"] = categoria_id
        if cc_id:
            payload["cost_centers_id"] = cc_id
        if setor_id:
            payload["sectors_id"] = setor_id
        if data_pagamento:
            payload["date_payment"] = _formatar_data(data_pagamento)

        return self._request("POST", "/transactions", payload)

    # ----------------------------------------------------------------
    # PUT /transactions/{id} — dar baixa (marcar pagamento)
    # ----------------------------------------------------------------
    def dar_baixa(self, transaction_id, data_pagamento, valor=None):
        payload = {
            "date_payment": _formatar_data(data_pagamento),
        }
        if valor is not None:
            payload["amount"] = _formatar_valor(valor)

        return self._request("PUT", f"/transactions/{transaction_id}", payload)

    # ----------------------------------------------------------------
    # DELETE /transactions/{id} — excluir lançamento
    # ----------------------------------------------------------------
    def excluir_lancamento(self, transaction_id):
        return self._request("DELETE", f"/transactions/{transaction_id}")

    # ----------------------------------------------------------------
    # Operações em lote (com relatório)
    # ----------------------------------------------------------------
    def executar_conciliacao(self, itens):
        resultados = {
            "sucesso": [],
            "erros": [],
            "ignorados": [],
        }

        for item in itens:
            acao = item.get("acao")

            if acao == "revisar":
                resultados["ignorados"].append({
                    "id": item.get("id"),
                    "descricao": item.get("descricao"),
                    "motivo": "Itens em revisão não são enviados automaticamente",
                })
                continue

            try:
                if acao == "baixa":
                    advbox_id = item.get("advbox_id")
                    if not advbox_id:
                        resultados["ignorados"].append({
                            "id": item.get("id"),
                            "descricao": item.get("descricao"),
                            "motivo": "Sem ID do lançamento no Advbox",
                        })
                        continue
                    resp = self.dar_baixa(
                        transaction_id=advbox_id,
                        data_pagamento=item.get("data"),
                        valor=item.get("valor"),
                    )
                    resultados["sucesso"].append({
                        "id": item.get("id"),
                        "acao": "baixa",
                        "descricao": item.get("descricao"),
                        "resposta": resp,
                    })

                elif acao == "criar":
                    if not item.get("categoria"):
                        resultados["erros"].append({
                            "id": item.get("id"),
                            "descricao": item.get("descricao"),
                            "erro": "Categoria não definida",
                        })
                        continue
                    resp = self.criar_lancamento(
                        tipo=item.get("tipo", "RECEITA"),
                        categoria=item.get("categoria"),
                        descricao=item.get("descricao", ""),
                        valor=item.get("valor", 0),
                        data_vencimento=item.get("data"),
                        data_pagamento=item.get("data"),
                        conta=item.get("conta"),
                        centro_custo=item.get("centro_custo"),
                        setor=item.get("setor"),
                    )
                    resultados["sucesso"].append({
                        "id": item.get("id"),
                        "acao": "criar",
                        "descricao": item.get("descricao"),
                        "resposta": resp,
                    })

            except AdvboxAPIError as e:
                resultados["erros"].append({
                    "id": item.get("id"),
                    "descricao": item.get("descricao"),
                    "erro": str(e),
                })
            except Exception as e:
                resultados["erros"].append({
                    "id": item.get("id"),
                    "descricao": item.get("descricao"),
                    "erro": f"Erro inesperado: {e}",
                })

        return resultados


# ----------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------

def _formatar_data(data_str):
    """Converte dd/mm/yyyy para yyyy-mm-dd (formato da API)."""
    if not data_str:
        return None
    if isinstance(data_str, datetime):
        return data_str.strftime("%Y-%m-%d")
    if "-" in str(data_str) and len(str(data_str)) >= 10:
        return str(data_str)[:10]
    try:
        parts = str(data_str).split("/")
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    except (IndexError, ValueError):
        return str(data_str)


def _formatar_valor(valor):
    """Converte float para formato brasileiro: 1500.50 → '1.500,50'"""
    s = f"{abs(float(valor)):,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"-{s}" if float(valor) < 0 else s


def _competencia(data_iso):
    """Extrai competência MM/YYYY de uma data ISO yyyy-mm-dd."""
    if not data_iso or len(str(data_iso)) < 7:
        return None
    parts = str(data_iso).split("-")
    try:
        return f"{parts[1]}/{parts[0]}"
    except (IndexError, ValueError):
        return None
