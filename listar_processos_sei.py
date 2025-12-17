"""
Script standalone para listar processos do SEI (MG) e exportar metadados para Excel.

Objetivo: permitir que um colega rode a listagem sem depender do pacote `sei_client/`.

Requisitos:
  - `uv sync`
  - Variáveis no `.env` (ou ambiente):
      SEI_USER, SEI_PASS, SEI_ORGAO, SEI_UNIDADE (obrigatórias)
      SEI_DEBUG=1 (opcional)
      SEI_SAVE_DEBUG_HTML=1 (opcional)
      SEI_DATA_DIR=data (opcional)

Execução:
  uv run listar_processos_sei.py
  uv run listar_processos_sei.py --saida ./saida/processos.xlsx
"""

from __future__ import annotations

import argparse
import logging
import math
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Literal, Optional, Set
from urllib.parse import parse_qs, urljoin, urlparse

import requests
from bs4 import BeautifulSoup, Tag
from dotenv import load_dotenv
from openpyxl import Workbook
from requests.adapters import HTTPAdapter

try:
    from urllib3.util.retry import Retry
except Exception:  # pragma: no cover - fallback defensivo
    Retry = None  # type: ignore[assignment,misc]


# Carrega automaticamente um `.env` na raiz do projeto (se existir).
# Isso permite rodar o script sem precisar exportar variáveis no terminal.
load_dotenv()

log = logging.getLogger("listar-processos-sei")

DEFAULT_BASE_URL = "https://www.sei.mg.gov.br"
DEFAULT_LOGIN_PATH = "/sip/login.php?sigla_orgao_sistema=GOVMG&sigla_sistema=SEI&infra_url=L3NlaS8="

DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

RE_ONCLICK_REDIRECT = re.compile(r"window\.location\.href='(?P<url>[^']+)'")
RE_PROCESSO = re.compile(r"\b\d{4}\.\s?\d{2}\.\s?\d{7}\s*/\s*\d{4}\s*[-–—-]\s*\d{2}\b", re.I)
RE_TOOLTIP = re.compile(r"infraTooltipMostrar\('([^']*)',\s*'([^']*)'\)", re.I)


class SEIError(RuntimeError):
    """Erro base para falhas relacionadas ao SEI neste script."""


class SEIConfigError(SEIError):
    """Configuração ausente/inválida (ex.: variáveis obrigatórias não definidas)."""


class SEILoginError(SEIError):
    """Falhas de autenticação (ex.: credenciais inválidas, conta bloqueada)."""


class SEIProcessoError(SEIError):
    """Falhas ao acessar/paginar/listar processos (rede, parsing, paginação indisponível)."""


def _str_to_bool(value: Optional[str]) -> Optional[bool]:
    """Converte strings comuns para booleano, retornando None quando indefinido."""
    if value is None:
        return None
    value_norm = value.strip().lower()
    truthy = {"1", "true", "t", "yes", "y", "sim"}
    falsy = {"0", "false", "f", "no", "n", "nao", "não"}
    if value_norm in truthy:
        return True
    if value_norm in falsy:
        return False
    return None


@dataclass(frozen=True)
class Settings:
    """Configuração calculada a partir do ambiente (.env/variáveis de ambiente)."""

    orgao_value: str
    unidade_value: str
    base_url: str = DEFAULT_BASE_URL
    login_path: str = DEFAULT_LOGIN_PATH
    data_dir: Path = field(default_factory=lambda: Path(os.environ.get("SEI_DATA_DIR", "data")))
    save_debug_html: bool = field(default_factory=lambda: _str_to_bool(os.environ.get("SEI_SAVE_DEBUG_HTML")) is True)
    debug_enabled: bool = field(default_factory=lambda: _str_to_bool(os.environ.get("SEI_DEBUG")) is True)

    @property
    def login_url(self) -> str:
        """URL completa de login (base + login_path)."""
        return f"{self.base_url}{self.login_path}"

    @property
    def unidade_alvo(self) -> str:
        """Nome da unidade que deve ficar ativa após o login."""
        return self.unidade_value.strip()


def load_settings() -> Settings:
    """Lê variáveis obrigatórias do ambiente e constrói `Settings`."""
    orgao_value = os.environ.get("SEI_ORGAO")
    unidade_value = os.environ.get("SEI_UNIDADE")

    if not orgao_value or not orgao_value.strip():
        raise SEIConfigError("Variável de ambiente SEI_ORGAO é obrigatória (ex: SEI_ORGAO=28).")
    if not unidade_value or not unidade_value.strip():
        raise SEIConfigError(
            "Variável de ambiente SEI_UNIDADE é obrigatória (ex: SEI_UNIDADE=SEPLAG/AUTOMATIZAMG)."
        )
    return Settings(orgao_value=orgao_value.strip(), unidade_value=unidade_value.strip())


def configure_logging(settings: Settings) -> None:
    """Configura logging (INFO por padrão; DEBUG quando `SEI_DEBUG=1`)."""
    level = logging.DEBUG if settings.debug_enabled else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%H:%M:%S",
    )


def absolute_to_sei(settings: Settings, href: str) -> str:
    """Converte `href` relativo do SEI em URL absoluta."""
    if href.startswith("http"):
        return href
    return urljoin(f"{settings.base_url}/sei/", href.lstrip("/"))


def save_html(settings: Settings, path: Path, html: str) -> None:
    """Salva HTML para debug quando `SEI_SAVE_DEBUG_HTML=1`."""
    if not settings.save_debug_html:
        return
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(html, encoding="iso-8859-1")
        log.debug("HTML salvo: %s (%s chars)", path, len(html))
    except Exception as exc:  # pragma: no cover
        log.warning("Erro ao salvar HTML %s: %s", path, exc)


def create_session(settings: Settings) -> requests.Session:
    """
    Cria sessão HTTP com cabeçalhos, cookie do órgão e retries.

    - O cookie `SIP_U_GOVMG_SEI` influencia o órgão selecionado no SEI.
    - Retries ajudam com instabilidades (429/5xx), sem mascarar erros de login.
    """
    session = requests.Session()
    session.headers.update(DEFAULT_HEADERS)
    if settings.orgao_value:
        session.cookies.set("SIP_U_GOVMG_SEI", settings.orgao_value, domain="sei.mg.gov.br")

    if Retry is not None:
        retry = Retry(
            total=5,
            connect=5,
            read=5,
            backoff_factor=0.5,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset({"GET", "POST"}),
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
    else:
        log.debug("urllib3 Retry indisponível; seguindo sem retries automáticos.")
    return session


def serializar_inputs(form: Tag) -> Dict[str, str]:
    """Serializa `<input>` (incluindo radio/checkbox) para um dict `name -> value`."""
    data: Dict[str, str] = {}
    for inp in form.find_all("input"):
        if not isinstance(inp, Tag):
            continue
        name = inp.get("name")
        if not name:
            continue
        itype = (inp.get("type") or "").lower()
        val = inp.get("value", "")
        if itype in {"radio", "checkbox"}:
            if inp.has_attr("checked"):
                data[name] = val
        else:
            data[name] = val
    return data


def serializar_selects(form: Tag) -> Dict[str, str]:
    """Serializa `<select>` para `name -> selected value`."""
    data: Dict[str, str] = {}
    for sel in form.find_all("select"):
        if not isinstance(sel, Tag):
            continue
        name = sel.get("name")
        if not name:
            continue
        opt = sel.find("option", selected=True) or sel.find("option")
        if opt and isinstance(opt, Tag):
            data[name] = opt.get("value", "")
        else:
            data[name] = ""
    return data


def serializar_textareas(form: Tag) -> Dict[str, str]:
    """Serializa `<textarea>` para `name -> texto`."""
    data: Dict[str, str] = {}
    for ta in form.find_all("textarea"):
        if not isinstance(ta, Tag):
            continue
        name = ta.get("name")
        if not name:
            continue
        data[name] = (ta.text or "").strip()
    return data


def processar_radios_nao_marcados(form: Tag, data: Dict[str, str]) -> Dict[str, str]:
    """
    Garante que grupos de radio tenham valor mesmo sem `checked`.

    O SEI às vezes espera que o campo do radio seja enviado; este fallback usa o primeiro valor.
    """
    radios_by_name: Dict[str, List[Tag]] = {}
    for radio in form.find_all("input", {"type": "radio"}):
        if not isinstance(radio, Tag):
            continue
        name = radio.get("name")
        if not name:
            continue
        radios_by_name.setdefault(name, []).append(radio)
    for name, radios in radios_by_name.items():
        if name not in data and radios:
            data[name] = radios[0].get("value", "")
    return data


def serializar_formulario(form: Tag) -> Dict[str, str]:
    """Serializa inputs/selects/textareas em um payload de POST."""
    data: Dict[str, str] = {}
    data.update(serializar_inputs(form))
    data.update(serializar_selects(form))
    data.update(serializar_textareas(form))
    return processar_radios_nao_marcados(form, data)


def login_sei(session: requests.Session, settings: Settings, user: str, pwd: str) -> str:
    """
    Realiza login no SEI e retorna o HTML pós-login.

    Observação: o SEI tipicamente responde em `iso-8859-1`, então o script força essa codificação.
    """
    if not user or not pwd:
        raise SEILoginError("Usuário e senha devem ser fornecidos (SEI_USER/SEI_PASS).")

    try:
        log.info("Abrindo página de login…")
        response = session.get(settings.login_url, timeout=30, headers=DEFAULT_HEADERS)
        response.raise_for_status()
        response.encoding = "iso-8859-1"

        session.cookies.set("SIP_U_GOVMG_SEI", settings.orgao_value, domain="sei.mg.gov.br")

        data = {
            "txtUsuario": user,
            "pwdSenha": pwd,
            "selOrgao": settings.orgao_value,
            "hdnAcao": "2",
            "Acessar": "Acessar",
        }

        log.info("Enviando POST de login…")
        response = session.post(settings.login_url, data=data, timeout=30, headers=DEFAULT_HEADERS, allow_redirects=True)
        response.raise_for_status()
        response.encoding = "iso-8859-1"

        save_html(settings, settings.data_dir / "debug" / "login.html", response.text)

        ok = ("Sair" in response.text) or ("Controle de Processos" in response.text)
        if not ok:
            lowered = response.text.lower()
            if "usuário ou senha" in lowered or "inval" in lowered:
                raise SEILoginError("Credenciais inválidas.")
            if "bloqueado" in lowered or "bloqueio" in lowered:
                raise SEILoginError("Conta bloqueada.")
            raise SEILoginError("Login não confirmado - verifique credenciais.")

        log.info("Autenticado com sucesso.")
        return response.text
    except requests.RequestException as exc:
        raise SEILoginError(f"Erro de rede durante login: {exc}") from exc


def descobrir_url_controle_do_html(settings: Settings, html: str) -> Optional[str]:
    """Tenta localizar no HTML pós-login o link para 'Controle de Processos'."""
    try:
        soup = BeautifulSoup(html, "lxml")
        for tag in soup.find_all("a", href=True):
            href = tag["href"]
            if "acao=procedimento_controlar" in href:
                return absolute_to_sei(settings, href)
        return None
    except Exception as exc:  # pragma: no cover
        log.warning("Erro ao descobrir URL de controle: %s", exc)
        return None


def abrir_controle(session: requests.Session, settings: Settings, html_pos_login: str) -> tuple[str, str]:
    """Abre a tela de Controle de Processos e retorna `(html, url)`."""
    try:
        url = descobrir_url_controle_do_html(settings, html_pos_login) or f"{settings.base_url}/sei/controlador.php?acao=procedimento_controlar"
        log.info("Acessando controle de processos: %s", url)
        response = session.get(url, timeout=30, headers=DEFAULT_HEADERS)
        response.raise_for_status()
        response.encoding = "iso-8859-1"
        save_html(settings, settings.data_dir / "debug" / "controle_pagina_1.html", response.text)
        return response.text, url
    except requests.RequestException as exc:
        raise SEIProcessoError(f"Erro ao acessar controle de processos: {exc}") from exc


def obter_unidade_atual(settings: Settings, html_controle: str) -> tuple[Optional[str], Optional[str]]:
    """
    Extrai a unidade atual e a URL de troca de unidade.

    O SEI expõe a troca de unidade via um `onclick` em `#lnkInfraUnidade`.
    """
    try:
        soup = BeautifulSoup(html_controle, "lxml")
        anchor = soup.select_one("#lnkInfraUnidade")
        if not anchor or not isinstance(anchor, Tag):
            log.debug("Elemento #lnkInfraUnidade não encontrado no HTML do controle.")
            return None, None

        nome_unidade = anchor.get_text(" ", strip=True) or None
        onclick = anchor.get("onclick") or ""
        url_troca: Optional[str] = None

        match = RE_ONCLICK_REDIRECT.search(onclick)
        if match:
            url_relativa = match.group("url")
            url_troca = absolute_to_sei(settings, url_relativa)
        else:
            log.debug("Não foi possível identificar URL de troca da unidade no atributo onclick.")

        return nome_unidade, url_troca
    except Exception as exc:  # pragma: no cover
        log.warning("Falha ao determinar unidade SEI atual: %s", exc)
        return None, None


def carregar_pagina_selecao_unidades(session: requests.Session, settings: Settings, url_troca: str) -> str:
    """Carrega a página que lista as unidades disponíveis para o usuário."""
    try:
        log.info("Carregando página de seleção de unidades: %s", url_troca)
        response = session.get(url_troca, timeout=30, headers=DEFAULT_HEADERS)
        response.raise_for_status()
        response.encoding = "iso-8859-1"
        save_html(settings, settings.data_dir / "debug" / "selecao_unidades.html", response.text)
        return response.text
    except requests.RequestException as exc:
        raise SEIProcessoError(f"Erro ao carregar página de seleção de unidades: {exc}") from exc


def selecionar_unidade_sei(
    session: requests.Session,
    settings: Settings,
    html_selecao: str,
    unidade_desejada: str,
    url_troca_origem: str,
) -> tuple[bool, Optional[str]]:
    """
    Seleciona a unidade desejada na tela de unidades e retorna (sucesso, html_resultado).

    Implementação baseada em scraping do formulário HTML do SEI.
    """
    try:
        soup = BeautifulSoup(html_selecao, "lxml")
        tabela = soup.select_one("table[id^='infraTable'], table.infraTable")
        if not tabela:
            tabelas = soup.find_all("table")
            for tab in tabelas:
                caption = tab.find("caption")
                if caption and "unidade" in caption.get_text(" ", strip=True).lower():
                    tabela = tab
                    break

        if not tabela:
            log.warning("Tabela de unidades não encontrada na página de seleção.")
            save_html(settings, settings.data_dir / "debug" / "selecao_unidades_debug.html", html_selecao)
            return False, None

        linhas = tabela.select("tbody tr") or tabela.select("tr")
        linhas = [linha for linha in linhas if isinstance(linha, Tag) and linha.select("th") == []]
        unidade_desejada_normalizada = re.sub(r"\s+", " ", unidade_desejada.strip().upper()).strip()

        for linha in linhas:
            if not isinstance(linha, Tag):
                continue
            celulas = linha.select("td")
            if len(celulas) < 2:
                continue
            texto_unidade = celulas[1].get_text(" ", strip=True)
            texto_limpo = re.sub(r"\s+", " ", texto_unidade.strip().upper()).strip()
            if texto_limpo != unidade_desejada_normalizada:
                continue

            radio = linha.select_one('input[type="radio"][name="chkInfraItem"]')
            if not radio or not isinstance(radio, Tag):
                log.warning("Radio button não encontrado para a unidade %s", unidade_desejada)
                continue
            valor_unidade = radio.get("value")
            if not valor_unidade:
                log.warning("Valor do radio button não encontrado para a unidade %s", unidade_desejada)
                continue

            form = soup.select_one("form#frmInfraSelecaoUnidade, form")
            if not form or not isinstance(form, Tag):
                log.warning("Formulário não encontrado na página de seleção.")
                return False, None

            data = serializar_formulario(form)
            data["selInfraUnidades"] = valor_unidade
            data["chkInfraItem"] = valor_unidade

            action = form.get("action", "")
            url_action = absolute_to_sei(settings, action) if action else url_troca_origem

            headers = dict(DEFAULT_HEADERS)
            headers["Referer"] = url_troca_origem
            headers["Content-Type"] = "application/x-www-form-urlencoded"

            log.info("Selecionando unidade SEI: %s (ID: %s)", unidade_desejada, valor_unidade)
            response = session.post(url_action, data=data, headers=headers, timeout=30, allow_redirects=True)
            response.raise_for_status()
            response.encoding = "iso-8859-1"

            save_html(settings, settings.data_dir / "debug" / "troca_unidade_resultado.html", response.text)

            if "Controle de Processos" in response.text or "procedimento_controlar" in response.text:
                log.info("Unidade SEI alterada com sucesso para: %s", unidade_desejada)
                return True, response.text
            log.warning("Resposta da troca de unidade não parece ter sido bem-sucedida.")
            return False, response.text

        log.warning("Unidade SEI '%s' não encontrada na lista de unidades disponíveis.", unidade_desejada)
        return False, None
    except requests.RequestException as exc:
        raise SEIProcessoError(f"Erro de rede ao selecionar unidade SEI: {exc}") from exc
    except Exception as exc:  # pragma: no cover
        log.error("Erro inesperado ao selecionar unidade SEI: %s", exc, exc_info=True)
        return False, None


@dataclass
class Processo:
    """Modelo de processo (metadados que aparecem no Controle de Processos)."""
    numero_processo: str
    id_procedimento: str
    url: str
    visualizado: bool
    categoria: Literal["Recebidos", "Gerados"]
    titulo: Optional[str] = None
    tipo_especificidade: Optional[str] = None
    responsavel_nome: Optional[str] = None
    responsavel_cpf: Optional[str] = None
    marcadores: List[str] = field(default_factory=list)
    tem_documentos_novos: bool = False
    tem_anotacoes: bool = False
    hash: str = ""
    documentos: List[Dict[str, Any]] = field(default_factory=list)
    eh_sigiloso: bool = False
    assinantes: List[str] = field(default_factory=list)
    metadados: Dict[str, Any] = field(default_factory=dict)


@dataclass
class PaginationInfo:
    """Metadados de paginação inferidos da tela de Controle de Processos."""
    total_registros: int
    pagina_atual: int
    total_paginas: int
    itens_por_pagina: int


def _get_attr_str(tag: Optional[Tag], attr: str, default: str = "") -> str:
    """Obtém um atributo de uma Tag do BeautifulSoup garantindo retorno em string."""
    if not tag:
        return default
    value = tag.get(attr, default)
    if isinstance(value, list):
        return value[0] if value else default
    return str(value) if value else default


def canonizar_processo(txt: str) -> str:
    """Normaliza o número do processo (remove espaços inconsistentes e NBSP)."""
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"\.\s+", ".", txt)
    txt = re.sub(r"\s*/\s*", "/", txt)
    txt = re.sub(r"\s*-\s*", "-", txt)
    return txt.strip()


def extrair_id_procedimento_da_url(url: str) -> str:
    """Extrai `id_procedimento` dos parâmetros da URL do processo."""
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        return params.get("id_procedimento", [""])[0]
    except Exception:
        return ""


def extrair_hash_da_url(url: str) -> str:
    """Extrai `infra_hash` dos parâmetros da URL do processo."""
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query)
        return params.get("infra_hash", [""])[0]
    except Exception:
        return ""


def parse_tooltip(onmouseover: Optional[str]) -> tuple[Optional[str], Optional[str]]:
    """Extrai (titulo, tipo/especificidade) do tooltip do SEI."""
    if not onmouseover:
        return None, None
    match = RE_TOOLTIP.search(onmouseover)
    if match:
        titulo = match.group(1).strip() if match.group(1) else None
        tipo = match.group(2).strip() if match.group(2) else None
        return titulo, tipo
    return None, None


def extrair_processo_da_linha(settings: Settings, linha: Tag, categoria: Literal["Recebidos", "Gerados"]) -> Optional[Processo]:
    """Converte uma `<tr>` do SEI em um `Processo`."""
    try:
        link_processo = linha.select_one('a[href*="acao=procedimento_trabalhar"]')
        if not link_processo or not isinstance(link_processo, Tag):
            return None

        txt = link_processo.get_text(" ", strip=True)
        title_attr = _get_attr_str(link_processo, "title")
        href_attr = _get_attr_str(link_processo, "href")
        match = RE_PROCESSO.search(txt) or RE_PROCESSO.search(title_attr) or RE_PROCESSO.search(href_attr)
        if not match:
            return None

        numero_processo = canonizar_processo(match.group(0))
        if not href_attr:
            return None
        url = absolute_to_sei(settings, href_attr)

        classes = link_processo.get("class", [])
        if isinstance(classes, str):
            classes = [classes]
        visualizado = "processoVisualizado" in classes

        id_procedimento = extrair_id_procedimento_da_url(url)
        hash_proc = extrair_hash_da_url(url)

        onmouseover = _get_attr_str(link_processo, "onmouseover")
        titulo, tipo_especificidade = parse_tooltip(onmouseover if onmouseover else None)

        responsavel_nome = None
        responsavel_cpf = None
        link_responsavel = linha.select_one('a[href*="acao=procedimento_atribuicao_listar"]')
        if link_responsavel and isinstance(link_responsavel, Tag):
            title_resp = _get_attr_str(link_responsavel, "title")
            responsavel_nome = title_resp.replace("Atribuído para ", "") if title_resp else None
            responsavel_cpf = link_responsavel.get_text(strip=True)

        marcadores: List[str] = []
        for img in linha.select("img.imagemStatus"):
            if not isinstance(img, Tag):
                continue
            parent_link = img.find_parent("a")
            if parent_link and isinstance(parent_link, Tag):
                onmouseover_attr = _get_attr_str(parent_link, "onmouseover")
                if onmouseover_attr:
                    tooltip_match = re.search(r"infraTooltipMostrar\('([^']*)'", onmouseover_attr)
                    if tooltip_match:
                        marcadores.append(tooltip_match.group(1).strip())

        tem_documentos_novos = bool(linha.select_one('img[src*="exclamacao.svg"]'))
        tem_anotacoes = bool(linha.select_one('img[src*="anotacao"]'))

        return Processo(
            numero_processo=numero_processo,
            id_procedimento=id_procedimento,
            url=url,
            visualizado=visualizado,
            categoria=categoria,
            titulo=titulo,
            tipo_especificidade=tipo_especificidade,
            responsavel_nome=responsavel_nome,
            responsavel_cpf=responsavel_cpf,
            marcadores=marcadores,
            tem_documentos_novos=tem_documentos_novos,
            tem_anotacoes=tem_anotacoes,
            hash=hash_proc,
        )
    except Exception as exc:  # pragma: no cover
        log.debug("Erro ao extrair processo da linha: %s", exc)
        return None


def extrair_processos(settings: Settings, html_controle: str) -> List[Processo]:
    """Extrai processos (Recebidos e Gerados) do HTML da página de controle."""
    try:
        soup = BeautifulSoup(html_controle, "lxml")
        processos: List[Processo] = []
        processos_ids: Set[str] = set()

        tabela_recebidos = soup.select_one("#tblProcessosRecebidos")
        if tabela_recebidos:
            for linha in tabela_recebidos.select("tr[id^='P']"):
                proc = extrair_processo_da_linha(settings, linha, "Recebidos")
                if proc and proc.id_procedimento and proc.id_procedimento not in processos_ids:
                    processos.append(proc)
                    processos_ids.add(proc.id_procedimento)

        tabela_gerados = soup.select_one("#tblProcessosGerados")
        if tabela_gerados:
            for linha in tabela_gerados.select("tr[id^='P']"):
                proc = extrair_processo_da_linha(settings, linha, "Gerados")
                if proc and proc.id_procedimento and proc.id_procedimento not in processos_ids:
                    processos.append(proc)
                    processos_ids.add(proc.id_procedimento)

        return processos
    except Exception as exc:
        raise SEIProcessoError(f"Erro ao extrair processos: {exc}") from exc


def _parse_caption_info(texto: str) -> tuple[int, int]:
    """Lê total de registros/itens por página a partir do texto do caption da tabela."""
    total_registros = 0
    itens_por_pagina = 0

    m_total = re.search(r"(\d+)\s+registros", texto)
    if m_total:
        total_registros = int(m_total.group(1))

    m_intervalo = re.search(r"-\s*(\d+)\s*a\s*(\d+)", texto)
    if m_intervalo:
        inicio = int(m_intervalo.group(1))
        fim = int(m_intervalo.group(2))
        itens_por_pagina = max(0, fim - inicio + 1)

    if itens_por_pagina == 0 and total_registros:
        itens_por_pagina = total_registros

    return total_registros, itens_por_pagina


def obter_paginacao_info(html_controle: str) -> Dict[str, PaginationInfo]:
    """Lê os campos hidden/caption para inferir paginação de Recebidos/Gerados."""
    soup = BeautifulSoup(html_controle, "lxml")
    info: Dict[str, PaginationInfo] = {}

    for grupo in ("Recebidos", "Gerados"):
        tabela = soup.select_one(f"#tblProcessos{grupo}")
        total_registros = 0
        itens_por_pagina = 0

        if tabela:
            caption = tabela.find("caption")
            if caption:
                total_registros, itens_por_pagina = _parse_caption_info(caption.get_text(" ", strip=True))
            linhas = tabela.select("tr[id^='P']")
            if itens_por_pagina <= 0 and linhas:
                itens_por_pagina = len(linhas)
            if total_registros <= 0 and linhas:
                total_registros = len(linhas)

        hidden_nro = soup.select_one(f"#hdn{grupo}NroItens")
        valor_nro = _get_attr_str(hidden_nro, "value") if hidden_nro else None
        if valor_nro:
            try:
                nro_itens = int(valor_nro)
                if itens_por_pagina <= 0:
                    itens_por_pagina = nro_itens
            except ValueError:
                pass

        hidden_itens = soup.select_one(f"#hdn{grupo}Itens")
        valor_itens = _get_attr_str(hidden_itens, "value") if hidden_itens else None
        if total_registros <= 0 and valor_itens:
            total_registros = len([item for item in valor_itens.split(",") if item])

        hidden_pagina = soup.select_one(f"#hdn{grupo}PaginaAtual")
        valor_pagina = _get_attr_str(hidden_pagina, "value") if hidden_pagina else None
        pagina_atual = 0
        if valor_pagina:
            try:
                pagina_atual = int(valor_pagina)
            except ValueError:
                pagina_atual = 0

        if itens_por_pagina <= 0:
            itens_por_pagina = max(1, total_registros if total_registros else 1)

        total_paginas = max(1, math.ceil(total_registros / itens_por_pagina)) if itens_por_pagina else 1
        info[grupo] = PaginationInfo(
            total_registros=total_registros,
            pagina_atual=pagina_atual,
            total_paginas=total_paginas,
            itens_por_pagina=itens_por_pagina,
        )

    return info


def submeter_paginacao(
    session: requests.Session,
    settings: Settings,
    html_atual: str,
    grupo: Literal["Recebidos", "Gerados"],
    pagina_destino: int,
    controle_url: str,
) -> str:
    """
    Submete o formulário `frmProcedimentoControlar` para trocar a página de um grupo.

    `pagina_destino` é 0-based (o SEI costuma trabalhar com índices numéricos internos).
    """
    soup = BeautifulSoup(html_atual, "lxml")
    form = soup.select_one("#frmProcedimentoControlar")
    if not form:
        raise SEIProcessoError("Formulário de controle não encontrado para paginação.")

    data = serializar_formulario(form)
    alvo = str(pagina_destino)

    select_superior = f"sel{grupo}PaginacaoSuperior"
    select_inferior = f"sel{grupo}PaginacaoInferior"
    hidden_pagina = f"hdn{grupo}PaginaAtual"

    if select_superior in data:
        data[select_superior] = alvo
    if select_inferior in data:
        data[select_inferior] = alvo
    if hidden_pagina in data:
        data[hidden_pagina] = alvo
    else:
        raise SEIProcessoError(f"Paginação indisponível para {grupo}.")

    action = _get_attr_str(form, "action")
    url_action = absolute_to_sei(settings, action)
    headers = dict(DEFAULT_HEADERS)
    headers.setdefault("Referer", controle_url)

    resposta = session.post(url_action, data=data, headers=headers, timeout=60)
    resposta.raise_for_status()
    resposta.encoding = "iso-8859-1"

    save_html(settings, settings.data_dir / "debug" / f"controle_{grupo.lower()}_{pagina_destino + 1}.html", resposta.text)
    return resposta.text


def _adicionar_processos(destino: List[Processo], novos: Iterable[Processo]) -> None:
    """Adiciona processos sem duplicar (chave por `id_procedimento`/`numero_processo`)."""
    vistos: Set[str] = {proc.id_procedimento or proc.numero_processo for proc in destino}
    for processo in novos:
        chave = processo.id_procedimento or processo.numero_processo
        if chave and chave not in vistos:
            destino.append(processo)
            vistos.add(chave)


def coletar_processos_com_paginacao(
    session: requests.Session,
    settings: Settings,
    html_inicial: str,
    controle_url: str,
) -> List[Processo]:
    """
    Coleta todos os processos navegando pelas páginas de Recebidos e Gerados.

    O SEI tem paginação separada por grupo, então o script pagina cada um e acumula.
    """
    processos: List[Processo] = []

    info_inicial = obter_paginacao_info(html_inicial)
    processos_iniciais = extrair_processos(settings, html_inicial)
    _adicionar_processos(processos, processos_iniciais)

    log.info(
        "Total inicial de processos: %s (%s Recebidos, %s Gerados)",
        len(processos),
        sum(1 for p in processos if p.categoria == "Recebidos"),
        sum(1 for p in processos if p.categoria == "Gerados"),
    )

    info_recebidos = info_inicial.get("Recebidos")
    if info_recebidos and info_recebidos.total_paginas > 1:
        html_receb = html_inicial
        for pagina in range(info_recebidos.pagina_atual + 1, info_recebidos.total_paginas):
            log.info("Carregando página %s/%s de Recebidos (acumulado: %s)", pagina + 1, info_recebidos.total_paginas, len(processos))
            html_receb = submeter_paginacao(session, settings, html_receb, "Recebidos", pagina, controle_url)
            _adicionar_processos(processos, extrair_processos(settings, html_receb))

    info_gerados = info_inicial.get("Gerados")
    if info_gerados and info_gerados.total_paginas > 1:
        html_ger = html_inicial
        for pagina in range(info_gerados.pagina_atual + 1, info_gerados.total_paginas):
            log.info("Carregando página %s/%s de Gerados (acumulado: %s)", pagina + 1, info_gerados.total_paginas, len(processos))
            html_ger = submeter_paginacao(session, settings, html_ger, "Gerados", pagina, controle_url)
            _adicionar_processos(processos, extrair_processos(settings, html_ger))

    log.info(
        "Total final de processos: %s (%s Recebidos, %s Gerados)",
        len(processos),
        sum(1 for p in processos if p.categoria == "Recebidos"),
        sum(1 for p in processos if p.categoria == "Gerados"),
    )
    return processos


def exportar_processos_para_excel(processos: List[Processo], caminho: str) -> Path:
    """Gera um `.xlsx` com 1 linha por processo e retorna o caminho final gravado."""
    if not processos:
        raise SEIProcessoError("Nenhum processo para exportar.")

    path = Path(caminho).expanduser()
    if path.is_dir():
        path = path / "processos.xlsx"
    elif path.suffix.lower() != ".xlsx":
        path = path.with_suffix(".xlsx")

    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    if ws is None:
        ws = wb.create_sheet("Processos")
    else:
        ws.title = "Processos"

    cabecalho = [
        "numero_processo",
        "categoria",
        "visualizado",
        "titulo",
        "tipo_especificidade",
        "responsavel_nome",
        "responsavel_cpf",
        "marcadores",
        "tem_documentos_novos",
        "tem_anotacoes",
        "id_procedimento",
        "hash",
        "url",
    ]
    ws.append(cabecalho)

    for proc in processos:
        ws.append(
            [
                proc.numero_processo,
                proc.categoria,
                "Sim" if proc.visualizado else "Não",
                proc.titulo or "",
                proc.tipo_especificidade or "",
                proc.responsavel_nome or "",
                proc.responsavel_cpf or "",
                "; ".join(proc.marcadores),
                "Sim" if proc.tem_documentos_novos else "Não",
                "Sim" if proc.tem_anotacoes else "Não",
                proc.id_procedimento,
                proc.hash,
                proc.url,
            ]
        )

    wb.save(path)
    return path


def executar_listagem(saida_xlsx: str) -> Path:
    """
    Orquestra o fluxo principal:

    1) Login
    2) Abre Controle de Processos
    3) Troca unidade (se necessário)
    4) Pagina e coleta todos os processos (Recebidos + Gerados)
    5) Exporta para Excel
    """
    settings = load_settings()
    configure_logging(settings)

    user = os.environ.get("SEI_USER")
    password = os.environ.get("SEI_PASS")
    if not user or not password:
        raise SEIConfigError("Defina SEI_USER e SEI_PASS no .env para autenticação.")

    session = create_session(settings)
    try:
        html_login = login_sei(session, settings, user, password)
        html_controle, controle_url = abrir_controle(session, settings, html_login)

        unidade_atual, trocar_url = obter_unidade_atual(settings, html_controle)
        if unidade_atual:
            log.info("Unidade SEI atual: %s", unidade_atual)

        if settings.unidade_alvo.strip().upper() != (unidade_atual or "").strip().upper():
            log.info(
                "Unidade SEI atual (%s) difere da desejada (%s). Iniciando troca...",
                unidade_atual or "desconhecida",
                settings.unidade_alvo,
            )
            if not trocar_url:
                log.warning("URL de troca de unidade não disponível. Continuando com a unidade atual.")
            else:
                html_selecao = carregar_pagina_selecao_unidades(session, settings, trocar_url)
                sucesso, html_resultado = selecionar_unidade_sei(
                    session, settings, html_selecao, settings.unidade_alvo, trocar_url
                )
                if sucesso and html_resultado:
                    # Recarrega o controle para estado consistente.
                    html_controle, controle_url = abrir_controle(session, settings, html_resultado)
                    nova_unidade, _ = obter_unidade_atual(settings, html_controle)
                    if nova_unidade:
                        log.info("Unidade SEI após troca: %s", nova_unidade)
                else:
                    log.warning("Falha ao trocar unidade SEI. Seguindo com unidade atual.")

        processos = coletar_processos_com_paginacao(session, settings, html_controle, controle_url)
        destino = exportar_processos_para_excel(processos, saida_xlsx)
        log.info("Planilha Excel gerada: %s", destino)
        return destino
    finally:
        session.close()


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    """Parser de argumentos CLI (mantém o script simples, sem filtros de processos)."""
    parser = argparse.ArgumentParser(
        description="Lista todos os processos (Recebidos e Gerados) do SEI e exporta metadados para .xlsx.",
    )
    parser.add_argument(
        "--saida",
        default="./saida/processos.xlsx",
        help="Caminho do arquivo .xlsx de saída (default: ./saida/processos.xlsx).",
    )
    return parser.parse_args(argv)


def main(argv: Optional[List[str]] = None) -> int:
    """Entry point. Retorna código de saída para facilitar automação/CI local."""
    args = parse_args(argv)
    try:
        executar_listagem(args.saida)
        return 0
    except SEIError as exc:
        log.error("%s", exc)
        return 10
    except KeyboardInterrupt:
        log.error("Interrompido pelo usuário.")
        return 130
    except Exception as exc:  # pragma: no cover
        log.exception("Erro inesperado: %s", exc)
        return 99


if __name__ == "__main__":
    sys.exit(main())
