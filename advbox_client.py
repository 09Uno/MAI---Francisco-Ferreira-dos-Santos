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

        # Mapas nome→ID carregados do GET /settings
        self.contas = {}        # nome da conta → ID
        self.categorias = {}    # nome da categoria → ID
        self.centros_custo = {} # nome do centro de custo → ID
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
            raise AdvboxAPIError(resp.status_code, str(body), body)

        if resp.status_code == 204:
            return {}
        return resp.json()

    # ----------------------------------------------------------------
    # GET /settings — carrega mapas nome→ID
    # ----------------------------------------------------------------
    def carregar_settings(self):
        """
        Chama GET /settings e monta mapas de nome→ID para
        contas, categorias e centros de custo.
        """
        data = self._request("GET", "/settings")
        self._settings_raw = data

        # Mapear contas
        for conta in data.get("accounts", data.get("contas", [])):
            nome = conta.get("name", conta.get("nome", ""))
            cid = conta.get("id")
            if nome and cid:
                self.contas[nome.strip().upper()] = cid

        # Mapear categorias
        for cat in data.get("categories", data.get("categorias", [])):
            nome = cat.get("name", cat.get("nome", ""))
            cid = cat.get("id")
            if nome and cid:
                self.categorias[nome.strip().upper()] = cid

        # Mapear centros de custo
        for cc in data.get("cost_centers", data.get("centros_custo", [])):
            nome = cc.get("name", cc.get("nome", ""))
            cid = cc.get("id")
            if nome and cid:
                self.centros_custo[nome.strip().upper()] = cid

        logger.info(f"Settings carregados: {len(self.contas)} contas, "
                     f"{len(self.categorias)} categorias, "
                     f"{len(self.centros_custo)} centros de custo")
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
            # Tenta match parcial
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
                         conta=None, pessoa=None, processo=None,
                         registro_interno=False, centro_custo=None):
        """
        Cria um lançamento (receita ou despesa) no Advbox.
        Retorna a resposta da API (ou o payload em dry_run).
        """
        conta_id = self._resolver_id(self.contas, conta, "Conta")
        categoria_id = self._resolver_id(self.categorias, categoria, "Categoria")
        cc_id = self._resolver_id(self.centros_custo, centro_custo, "Centro de custo")

        payload = {
            "type": tipo.lower(),  # "receita" ou "despesa"
            "description": descricao,
            "value": round(valor, 2),
            "due_date": _formatar_data(data_vencimento),
            "internal_record": registro_interno,
        }

        if conta_id:
            payload["account_id"] = conta_id
        if categoria_id:
            payload["category_id"] = categoria_id
        if cc_id:
            payload["cost_center_id"] = cc_id
        if data_pagamento:
            payload["payment_date"] = _formatar_data(data_pagamento)
        if pessoa:
            payload["person_name"] = pessoa
        if processo:
            payload["lawsuit_number"] = processo

        return self._request("POST", "/transactions", payload)

    # ----------------------------------------------------------------
    # PUT /transactions/{id} — dar baixa (marcar pagamento)
    # ----------------------------------------------------------------
    def dar_baixa(self, transaction_id, data_pagamento, valor=None):
        """
        Marca um lançamento existente como pago.
        """
        payload = {
            "payment_date": _formatar_data(data_pagamento),
        }
        if valor is not None:
            payload["paid_value"] = round(valor, 2)

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
        """
        Executa a conciliação em lote:
        - Itens com acao='baixa' → dar_baixa (se tiver ID do Advbox)
        - Itens com acao='criar' → criar_lancamento
        - Itens com acao='revisar' → IGNORADOS (nunca postar automaticamente)

        Retorna dict com resultado de cada operação.
        """
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
                        pessoa=item.get("pessoa"),
                        registro_interno=item.get("registro_interno", False),
                        centro_custo=item.get("centro_custo"),
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


def _formatar_data(data_str):
    """Converte dd/mm/yyyy para yyyy-mm-dd (formato da API)."""
    if not data_str:
        return None
    if isinstance(data_str, datetime):
        return data_str.strftime("%Y-%m-%d")
    # Se já está no formato ISO
    if "-" in str(data_str) and len(str(data_str)) >= 10:
        return str(data_str)[:10]
    # dd/mm/yyyy → yyyy-mm-dd
    try:
        parts = str(data_str).split("/")
        return f"{parts[2]}-{parts[1]}-{parts[0]}"
    except (IndexError, ValueError):
        return str(data_str)
