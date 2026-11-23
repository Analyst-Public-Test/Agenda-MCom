#!/usr/bin/env python3
"""Agenda MCom (SERAD): e-Agendas (CGU) + Portal da Transparência (viagens) -> dados/agenda-mcom.json

Autoridades: Secretário e 3 diretores da SERAD (config/autoridades.json).

Fontes
  e-Agendas : página pública de cada autoridade (a agenda inteira vem embutida no HTML, sem captcha).
  Portal    : zip anual de viagens (GET direto) -> {ano}_Viagem.csv.
  Plano B   : arquivos em dados/entrada/ (xlsx/csv do e-Agendas; zip/csv de viagens em dados/entrada/viagens/).

Regras
  - Só entram anos >= ano_inicio_automatico (2026). Anos anteriores ficam fixos no HTML.
  - Se uma fonte falhar, o que já está no JSON é mantido (nada some por causa de uma falha de rede).
  - Classificação já gravada no JSON não é sobrescrita; config/ajustes.json vence sempre.
  - Compromisso com o mesmo ID (ou mesma data/hora/assunto) em mais de uma agenda vira um registro só, com todas as pessoas.
  - Estados e países novos entram sozinhos no mapa: o JSON leva a tabela de países (código, nome e posição) e o HTML a usa (lugares.py).
  - O JSON só é regravado quando muda o conteúdo OU o estado das fontes (assim o bloco de controle nunca fica desatualizado).
"""
import argparse, csv, datetime as dt, difflib, hashlib, html, io, json, re, sys, tempfile, time, unicodedata, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
import lugares as LG  # noqa: E402

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/157.0.0.0 Safari/537.36 Edg/157.0.0.0")
TIPOS_OK = {"Reunião", "Evento", "Audiência", "Audiência pública"}


def N(s):
    if s is None:
        return ""
    return unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode()


def limpo(s):
    if s is None:
        return ""
    s = re.sub(r"\s+", " ", html.unescape(str(s))).strip()
    return "" if s.upper() == "ND" else s


def nrm(s):
    return re.sub(r"[^a-z0-9 ]", " ", N(s).lower()).strip()


def sim(a, b):
    return difflib.SequenceMatcher(None, nrm(a), nrm(b)).ratio()


def log(*a):
    print(*a, flush=True)


def ler_json(path, padrao=None):
    return json.loads(Path(path).read_text(encoding="utf-8")) if Path(path).exists() else (padrao if padrao is not None else {})


# ----------------------------------------------------------------------------------------------
# e-Agendas
# ----------------------------------------------------------------------------------------------
def extrair_eventos(texto):
    m = re.search(r'ng-init="events=(\[.*?\])"\s*>', texto, re.S)
    if not m:
        raise ValueError("bloco ng-init=\"events=[...]\" não encontrado na página")
    return json.loads(html.unescape(m.group(1)))


def baixar_pagina(cfg, aut, tentativas=3):
    import requests
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "pt-BR,pt;q=0.9"})
    ultimo = None
    for t in range(tentativas):
        try:
            r0 = s.get(cfg["url_eagendas"], timeout=60)
            r0.raise_for_status()
            tok = re.search(r'name="_token"\s+value="([^"]+)"', r0.text)
            params = {
                "filtro_orgaos_ativos": "on", "filtro_orgao": cfg["orgao_id"],
                "filtro_cargos_ativos": "on", "filtro_cargo": aut["cargo"],
                "filtro_apos_ativos": "on", "filtro_servidor": aut["servidor_id"],
                "cargo_confianca_id": "", "is_cargo_vago": "false",
            }
            if tok:
                params = {"_token": tok.group(1), **params}
            r = s.get(cfg["url_eagendas"], params=params, timeout=120)
            r.raise_for_status()
            return extrair_eventos(r.text)
        except Exception as e:  # noqa: BLE001
            ultimo = e
            time.sleep(5 * (t + 1))
    raise RuntimeError(f"e-Agendas {aut['nome']}: {ultimo}")


def texto_detalhe(h):
    h = re.sub(r"<br\s*/?>", "\n", h or "", flags=re.I)
    h = re.sub(r"<[^>]+>", "", h)
    return html.unescape(h).strip()


def reg_de_evento(e, chave):
    """evento do e-Agendas -> registro comum"""
    if "compromisso_id" not in e or e.get("tipo") not in TIPOS_OK:
        return None
    ini = e["start"]
    titulo = re.sub(r"^(Audiência pública|Audiência|Reunião|Evento)\s+-\s+", "", e.get("title", ""))
    det = e.get("detalhamento_exibicao") or {}
    local = limpo(e.get("local"))
    return {
        "id": str(e["compromisso_id"]), "pes": {chave}, "iso": ini[:10], "hora": ini[11:16],
        "tipo": "Audiência" if e["tipo"].startswith("Audiência") else e["tipo"],
        "assunto": limpo(titulo), "local": local, "virtual": local.lower() == "virtual",
        "participantes": texto_detalhe(e.get("detalhe")),
        "objetivos": "; ".join(det.get("objetivos") or []), "info": limpo(det.get("principal")),
        "mod": e.get("modificado_em", ""), "fonte": "pagina",
    }


def ler_tabela(path):
    if path.suffix.lower() == ".xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb:
            it = ws.iter_rows(values_only=True)
            cab = next(it, None)
            if not cab:
                continue
            for r in it:
                yield {str(c): ("" if v is None else str(v)) for c, v in zip(cab, r)}
    else:
        raw = path.read_bytes()
        try:
            txt = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            txt = raw.decode("latin-1")
        sep = ";" if txt.split("\n", 1)[0].count(";") > txt.split("\n", 1)[0].count(",") else ","
        yield from csv.DictReader(io.StringIO(txt), delimiter=sep)


def reg_de_linha(l, por_nome):
    if l.get("Tipo de registro") not in TIPOS_OK:
        return None
    chave = por_nome.get(N(l.get("Nome", "")).upper().strip())
    if not chave or not l.get("ID do registro"):
        return None
    try:
        iso = dt.datetime.strptime(l["Data de início"], "%d-%m-%Y").strftime("%Y-%m-%d")
    except ValueError:
        return None
    hora = limpo(l.get("Hora de início"))
    local = limpo(l.get("Local"))
    forma = limpo(l.get("Forma de realização"))
    return {
        "id": str(l["ID do registro"]).strip(), "pes": {chave}, "iso": iso, "hora": hora,
        "tipo": "Audiência" if l["Tipo de registro"].startswith("Audiência") else l["Tipo de registro"],
        "assunto": limpo(l.get("Assunto do Compromisso")), "local": local,
        "virtual": forma.lower() == "virtual" or local.lower() == "virtual",
        "participantes": l.get("Participantes", "") or "", "objetivos": limpo(l.get("Objetivos")),
        "info": limpo(l.get("Informações complementares")), "mod": l.get("Última modificação", ""),
        "fonte": "entrada",
    }


def unir_ids(regs):
    """une IDs diferentes do mesmo compromisso (mesma data/hora/assunto)"""
    grupos = sorted(regs.values(), key=lambda r: (r["iso"], r["hora"], r["id"]))
    usados, finais = set(), []
    for i, g in enumerate(grupos):
        if i in usados:
            continue
        g = dict(g, ids=[g["id"]], pes=set(g["pes"]))
        for j in range(i + 1, len(grupos)):
            h = grupos[j]
            if h["iso"] != g["iso"]:
                break
            if j in usados or h["hora"] != g["hora"] or not g["assunto"]:
                continue
            if sim(g["assunto"], h["assunto"]) >= 0.9:
                g["ids"].append(h["id"]); g["pes"] |= h["pes"]; usados.add(j)
        finais.append(g)
    return finais


# ----------------------------------------------------------------------------------------------
# Portal da Transparência (viagens) - compartilhado pelos painéis
# ----------------------------------------------------------------------------------------------
CIDADE = {"Amsterdam": "Amsterdã", "Santo Domingo de Guzmán": "Santo Domingo"}
PAIS = {"Estados Unidos da América": "EUA"}


def lista_destinos(dest):
    """'São Paulo/SP, Campo Grande/MS' -> ['São Paulo/SP', 'Campo Grande/MS']"""
    return [p.strip() for p in re.split(r",\s*(?=[^,/]+/)", dest.strip()) if p.strip()]


def _lugar_de(destino):
    """'São Paulo/SP' -> 'SP:São Paulo' ; 'Tóquio/Japão' -> 'EX:Tóquio (Japão)'"""
    cidade, _, fim = destino.strip().rpartition("/")
    if re.fullmatch(r"[A-Z]{2}", fim):
        return f"{fim}:{cidade}"
    return f"EX:{CIDADE.get(cidade, cidade)} ({PAIS.get(fim, fim)})"


def lugar_destino(dest):
    """lugar do PRIMEIRO destino"""
    ds = lista_destinos(dest)
    return _lugar_de(ds[0]) if ds else ""


def lugares_destino(dest):
    """todos os destinos da viagem (sem repetir), no mesmo formato de lugar_destino"""
    out = []
    for p in lista_destinos(dest):
        l = _lugar_de(p)
        if l not in out:
            out.append(l)
    return out


def baixar_zip_viagens(url_modelo, ano, destino):
    import requests
    url = url_modelo.format(ano=ano)
    r = requests.get(url, headers={"User-Agent": UA}, stream=True, timeout=600, allow_redirects=True)
    if r.status_code == 404:
        return None, "sem arquivo para o ano"
    r.raise_for_status()
    with open(destino, "wb") as f:
        for ch in r.iter_content(1 << 20):
            f.write(ch)
    with open(destino, "rb") as f:
        if f.read(2) != b"PK":
            raise RuntimeError("o download não é um zip (possível bloqueio do site)")
    return r.url.rsplit("/", 1)[-1], "ok"


def viagens_do_csv(fh, ano, por_nome):
    """linhas: [ini, fim, chave, lugar(1º destino), destinos, PCDP, processo, motivo, lugares[]]"""
    rd = csv.reader(io.TextIOWrapper(fh, encoding="latin-1", newline=""), delimiter=";")
    cab = [N(c).lower().strip() for c in next(rd)]
    ix = {k: cab.index(k) for k in [
        "identificador do processo de viagem", "numero da proposta (pcdp)", "situacao", "nome orgao solicitante",
        "nome", "periodo - data de inicio", "periodo - data de fim", "destinos", "motivo"]}
    out = []
    for r in rd:
        if len(r) < len(cab):
            continue
        chave = por_nome.get(N(r[ix["nome"]]).upper().strip())
        if not chave or "ministerio das comunicacoes" not in N(r[ix["nome orgao solicitante"]]).lower():
            continue
        if r[ix["situacao"]].strip() != "Realizada":
            continue
        ini = dt.datetime.strptime(r[ix["periodo - data de inicio"]], "%d/%m/%Y").strftime("%Y-%m-%d")
        fim = dt.datetime.strptime(r[ix["periodo - data de fim"]], "%d/%m/%Y").strftime("%Y-%m-%d")
        if ini[:4] != str(ano):
            continue
        dest = limpo(r[ix["destinos"]])
        out.append([ini, fim, chave, lugar_destino(dest), dest, r[ix["numero da proposta (pcdp)"]].strip(),
                    r[ix["identificador do processo de viagem"]].lstrip("0"), limpo(r[ix["motivo"]]),
                    lugares_destino(dest)])
    return out


def ler_viagens_arquivo(path, ano, por_nome):
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            nome = next(n for n in z.namelist() if n.endswith("_Viagem.csv"))
            with z.open(nome) as fh:
                return viagens_do_csv(fh, ano, por_nome)
    with open(path, "rb") as fh:
        return viagens_do_csv(fh, ano, por_nome)


_VCACHE = {}
_TMP = []


def viagens_do_ano(url_modelo, ano, por_nome, dir_plano_b, sem_rede):
    """Baixa e lê o zip do ano UMA vez (cache) para todos os painéis. Retorna (linhas|None, info)."""
    if ano in _VCACHE:
        return _VCACHE[ano]
    rows, info = None, ""
    try:
        if sem_rede:
            raise RuntimeError("modo --sem-rede")
        if not _TMP:
            _TMP.append(tempfile.TemporaryDirectory())
        z = Path(_TMP[0].name) / f"{ano}.zip"
        nome, info = baixar_zip_viagens(url_modelo, ano, z)
        if nome:
            rows = ler_viagens_arquivo(z, ano, por_nome)
            info = nome
            z.unlink()
    except Exception as e:  # noqa: BLE001
        info = f"falhou: {str(e)[:150]}"
    if rows is None:
        for p in sorted(dir_plano_b.glob("*")):
            if p.suffix.lower() in (".zip", ".csv") and str(ano) in p.name:
                rows = ler_viagens_arquivo(p, ano, por_nome)
                info = f"arquivo local {p.name}"
    _VCACHE[ano] = (rows, info)
    return rows, info


# ----------------------------------------------------------------------------------------------
# Lugar (SERAD)
# ----------------------------------------------------------------------------------------------
DF_FORTE = (r"minist[eé]rio|mist[eé]rio|esplanada|bras[ií]lia|planalto|\bDF\b|anatel|ag[eê]ncia nacional|senado|"
            r"c[aâ]mara dos|congresso nacional|pal[aá]cio|distrito federal|\bTCU\b|\bCGU\b|\bSTF\b|itamaraty|\bEBC\b|"
            r"\bSCS\b|\bSAUS\b|\bSHIS\b|\bSGAN\b|\bSBN\b|\bSBS\b|\bSEPN\b|\bSAIN\b|\bCLN\b|\bCLS\b|\bW3\b|asa (sul|norte)|"
            r"\bGM\b|gabinete|\bMCom\b|\bUnB\b|FACEB|Freitas Nobre|Radiobr|TV C[aâ]mara|IPHAN|Capital Financial|\bB Hotel\b|"
            r"Complexo Brasil 21")
CIDADES = [
    (r"Expo Center Norte|Anhembi|Celso Furtado|\bUSP\b|C[aâ]mara Municipal de S[aã]o Paulo|S[aã]o Paulo Expo", "SP:São Paulo"),
    (r"Atl[aâ]ntica|Copacabana|Barra da Tijuca|Cidade do Rio", "RJ:Rio de Janeiro"),
    (r"Recife|Luzeiros|ASSERPE|Mar (&|&amp;|e) Vista", "PE:Recife"), (r"Foz do Igua", "PR:Foz do Iguaçu"),
    (r"Florian[oó]polis", "SC:Florianópolis"), (r"Las Vegas", "EX:Las Vegas (EUA)"),
    (r"Amsterd|Tuschinski|\bRAI\b", "EX:Amsterdã (Países Baixos)"), (r"Montevid", "EX:Montevidéu (Uruguai)"),
    (r"Santo Domingo", "EX:Santo Domingo (República Dominicana)"), (r"Manaus", "AM:Manaus"),
    (r"Savassi|Belo Horizonte", "MG:Belo Horizonte"), (r"Bel[eé]m", "PA:Belém"), (r"Fortaleza", "CE:Fortaleza"),
    (r"Porto Alegre", "RS:Porto Alegre"), (r"Salvador", "BA:Salvador"), (r"Curitiba", "PR:Curitiba"),
]


def _escolhe_destino(reg, trecho):
    """viagem com mais de um destino: escolhe o que aparece no local/assunto; sem pista, o primeiro"""
    lugares = trecho[8] if len(trecho) > 8 and trecho[8] else [trecho[3]]
    if len(lugares) == 1:
        return lugares[0]
    txt = LG.N(reg["local"] + " " + reg["assunto"]).lower()
    for l in lugares:
        cid = re.sub(r"^[A-Z]{2}:|^EX:|\s*\(.*\)$", "", l)
        if cid and LG.N(cid).lower() in txt:
            return l
    cu = LG.cidade_uf(reg["local"]) or LG.cidade_estado_extenso(reg["local"])
    if cu:
        for l in lugares:
            if l.startswith(cu[0] + ":"):
                return l
    return lugares[0]


def decidir_lugar(reg, vgs, pendencias):
    if reg["virtual"]:
        return "VI:Virtual", "V"
    if re.search(DF_FORTE, reg["local"], re.I):
        return "DF:Brasília", "P"
    for v in vgs:
        if v[2] in reg["pes"] and v[0] <= reg["iso"] <= v[1]:
            return _escolhe_destino(reg, v), "P"
    for rx, l in CIDADES:
        if re.search(rx, reg["local"] + " " + reg["assunto"], re.I):
            return l, "P"
    # estados e países novos: reconhece pelo texto do local (e, na falta dele, do assunto)
    ex = LG.achar_exterior(reg["local"]) or LG.achar_exterior(reg["assunto"])
    if ex:
        return ex, "P"
    cu = LG.cidade_uf(reg["local"]) or LG.cidade_uf(reg["assunto"])
    if cu and cu[0] in LG.UFS and cu[0] != "DF":
        return f"{cu[0]}:{LG.NOMES.get(LG.N(cu[1]).lower(), cu[1])}", "P"
    c = LG.achar_cidade(reg["local"]) or LG.achar_cidade(reg["assunto"])
    if c and c[0] != "DF":
        return f"{c[0]}:{c[1]}", "P"
    ce = LG.cidade_estado_extenso(reg["local"]) or LG.cidade_estado_extenso(reg["assunto"])
    if ce:
        return f"{ce[0]}:{ce[1]}", "P"
    pendencias.append({"id": reg["ids"][0], "data": reg["iso"], "local": reg["local"], "assunto": reg["assunto"][:100]})
    return "DF:Brasília", "P"


def paises_do_mapa(lugares):
    """lugares 'EX:Cidade (País)' -> (paises {iso: {n, c, xy}}, exp {lugar: iso}, sem_pais [lugares])"""
    exp, paises, cidades, sem_pais = {}, {}, {}, []
    for l in sorted(set(lugares)):
        m = re.match(r"EX:(.*?)(?: \((.*)\))?$", l)
        cid, pais = m.group(1), (m.group(2) or m.group(1))
        if LG.pais_ignorar(pais):
            continue
        info = LG.pais_info(pais)
        if not info:
            sem_pais.append(l)
            continue
        iso, nome, x, y = info
        exp[l] = iso
        paises.setdefault(iso, {"n": nome, "c": "", "xy": [x, y]})
        if cid not in cidades.setdefault(iso, []):
            cidades[iso].append(cid)
    for iso, cs in cidades.items():
        paises[iso]["c"] = ", ".join(cs)
    return paises, exp, sem_pais


def gravar_se_mudou(saida, estado, corpo, sinal, meta, extra_log=""):
    h = hashlib.sha256(json.dumps([corpo, sinal], ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:16]
    if estado.get("hash") == h:
        log("Sem mudanças no conteúdo; JSON não foi alterado.")
        return
    out = {"atualizado_em": dt.datetime.now(dt.timezone(dt.timedelta(hours=-3))).strftime("%Y-%m-%dT%H:%M:%S-03:00"),
           "hash": h, **meta, **corpo}
    saida.parent.mkdir(parents=True, exist_ok=True)
    saida.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    log(f"Gravado {saida.name} {extra_log}")


def sinal_fontes(fontes, fonte_viag):
    """assinatura do estado das fontes: sem textos de erro (que variam a cada execução)"""
    return {
        "e": {k: [v["ok"], v.get("compromissos")] for k, v in fontes.items()},
        "v": {y: ["falhou" if "falhou" in x["arquivo"] else x["arquivo"], x["viagens"]] for y, x in fonte_viag.items()},
    }


# ----------------------------------------------------------------------------------------------
# Painel SERAD
# ----------------------------------------------------------------------------------------------
def rodar_serad(a):
    from classificador import classificar

    cfg = json.loads((ROOT / "config" / "autoridades.json").read_text(encoding="utf-8"))
    rot = json.loads((ROOT / "config" / "rotulos.json").read_text(encoding="utf-8"))
    ajustes = {k: v for k, v in ler_json(ROOT / "config" / "ajustes.json").items() if not k.startswith("_")}
    auts = cfg["autoridades"]
    chaves = {x["chave"] for x in auts}
    por_nome = {N(x["nome"]).upper(): x["chave"] for x in auts}
    ano0 = cfg["ano_inicio_automatico"]
    anos = list(range(ano0, dt.date.today().year + 1))
    saida = Path(a.saida) if a.saida else ROOT / "dados" / "agenda-mcom.json"
    estado = ler_json(saida)
    antigo_rows = estado.get("rows", [])
    antigo_vg = estado.get("vg", [])
    ja_class = {} if a.reclassificar else {i: (r[6], r[7]) for r in antigo_rows for i in (r[12] if len(r) > 12 else [])}

    # ---- 1) e-Agendas
    fontes, regs, falhou = {}, {}, set()
    locais = dict(x.split("=", 1) for x in a.pagina_local)
    for aut in auts:
        k = aut["chave"]
        try:
            if k in locais:
                ev = extrair_eventos(Path(locais[k]).read_text(encoding="utf-8", errors="replace"))
            elif a.sem_rede:
                raise RuntimeError("modo --sem-rede")
            else:
                ev = baixar_pagina(cfg, aut)
            n = 0
            for e in ev:
                r = reg_de_evento(e, k)
                if r and r["iso"][:4] >= str(ano0):
                    if r["id"] in regs:
                        regs[r["id"]]["pes"].add(k)
                    else:
                        regs[r["id"]] = r
                    n += 1
            if n == 0 and any(k in r[9] for r in antigo_rows):
                raise RuntimeError("página sem nenhum compromisso (possível falha do site); mantido o que já estava no JSON")
            fontes[k] = {"fonte": "e-Agendas (página)", "ok": True, "compromissos": n}
            log(f"[SERAD][e-Agendas] {aut['nome']}: {n} compromissos desde {ano0}")
        except Exception as e:  # noqa: BLE001
            falhou.add(k)
            fontes[k] = {"fonte": "e-Agendas (página)", "ok": False, "erro": str(e)[:200]}
            log(f"[SERAD][e-Agendas] FALHOU {aut['nome']}: {e}")

    # plano B: arquivos soltos em dados/entrada (completam o que a página não trouxe)
    for p in sorted((ROOT / "dados" / "entrada").glob("*")):
        if p.suffix.lower() not in (".xlsx", ".csv"):
            continue
        n = 0
        for l in ler_tabela(p):
            r = reg_de_linha(l, por_nome)
            if not r or r["iso"][:4] < str(ano0):
                continue
            if r["id"] in regs:
                regs[r["id"]]["pes"] |= r["pes"]
            else:
                regs[r["id"]] = r
            n += 1
        log(f"[SERAD][entrada] {p.name}: {n} registros")

    # ---- 2) viagens
    vg, fonte_viag = [], {}
    for ano in anos:
        rows_ano, info = viagens_do_ano(cfg["url_viagens"], ano, por_nome, ROOT / "dados" / "entrada" / "viagens", a.sem_rede)
        if rows_ano is not None:
            rows_ano = [v for v in rows_ano if v[2] in chaves]
        else:
            rows_ano = [(v if len(v) > 8 else v + [[v[3]]]) for v in antigo_vg if v[0][:4] == str(ano)]
            if rows_ano:
                info += " (mantido o que já estava no JSON)"
        vg += rows_ano
        fonte_viag[str(ano)] = {"arquivo": info, "viagens": len(rows_ano)}
        log(f"[SERAD][Portal] {ano}: {len(rows_ano)} viagens ({info})")
    vg.sort(key=lambda v: (v[0], v[2]))

    # ---- 3) mantém o que já existia para autoridades cuja página falhou
    for r in antigo_rows:
        ids = r[12] if len(r) > 12 else []
        if not (falhou & set(r[9])) or r[0][:4] < str(ano0) or not ids:
            continue
        achou = [i for i in ids if i in regs]
        if achou:  # compromisso compartilhado: a pessoa cuja fonte falhou continua nele
            regs[achou[0]]["pes"] |= falhou & set(r[9])
        else:
            regs[ids[0]] = {"id": ids[0], "pes": set(r[9]), "iso": r[0], "hora": r[1], "tipo": r[2], "assunto": r[3],
                            "local": r[4], "virtual": r[5] == "V", "participantes": "", "objetivos": "", "info": "",
                            "_manter": r, "fonte": "estado"}

    # ---- 4) une IDs diferentes do mesmo compromisso
    finais = unir_ids(regs)

    # ---- 5) linhas no formato do painel
    pendencias, rows = [], []
    for g in finais:
        if "_manter" in g:
            r = list(g["_manter"]); r[9] = "".join(c["chave"] for c in auts if c["chave"] in g["pes"])
            rows.append(r); continue
        lugar, pv = decidir_lugar(g, vg, pendencias)
        previa = next((ja_class[i] for i in g["ids"] if i in ja_class), None)
        cats, tem = previa if previa else classificar(g)
        for i in g["ids"]:
            if i in ajustes:
                cats = ajustes[i].get("cats", cats); tem = ajustes[i].get("tem", tem)
                lugar = ajustes[i].get("lugar", lugar)
        assunto = g["assunto"] or g["info"] or g["objetivos"] or "Compromisso sem assunto informado"
        pes = "".join(c["chave"] for c in auts if c["chave"] in g["pes"])
        rows.append([g["iso"], g["hora"], g["tipo"], assunto, g["local"] or ("Virtual" if pv == "V" else ""), pv,
                     cats, tem, lugar, pes, "EA", None, g["ids"]])
    rows.sort(key=lambda r: (r[0], r[1]), reverse=True)

    # ---- 6) saída
    lug_ex = [r[8] for r in rows if r[8].startswith("EX:")] + [l for v in vg for l in v[8] if l.startswith("EX:")]
    paises, exp, sem_pais = paises_do_mapa(lug_ex)
    corpo = {"cats": rot["cats"], "tem": rot["tem"], "paises": paises, "exp": exp, "rows": rows, "vg": vg}
    ult = {c["chave"]: max([r[0] for r in rows if c["chave"] in r[9]] or [""]) for c in auts}
    meta = {"fontes_eagendas": fontes, "fontes_viagens": fonte_viag, "ultimo_compromisso": ult,
            "pendencias_de_lugar": pendencias, "pendencias_de_pais": sem_pais}
    gravar_se_mudou(saida, estado, corpo, sinal_fontes(fontes, fonte_viag), meta, f"({len(rows)} compromissos, {len(vg)} viagens)")
    if pendencias:
        log(f"[SERAD] Atenção: {len(pendencias)} compromisso(s) presencial(is) sem lugar identificado (ficaram em Brasília); veja 'pendencias_de_lugar'.")
    if sem_pais:
        log(f"[SERAD] Atenção: destino(s) sem país no mapa: {sem_pais}; veja 'pendencias_de_pais'.")
    if len(falhou) == len(auts) and not regs:
        log("[SERAD] ERRO: nenhuma fonte de agenda respondeu.")
        return False
    return True


# ----------------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--saida", default=None, help="JSON de saída (teste)")
    ap.add_argument("--sem-rede", action="store_true", help="usa só dados/entrada (teste local)")
    ap.add_argument("--pagina-local", action="append", default=[], metavar="CHAVE=arquivo.html",
                    help="lê a página salva de uma autoridade em vez de baixar (teste)")
    ap.add_argument("--reclassificar", action="store_true", help="ignora a classificação já gravada")
    a = ap.parse_args()
    if not rodar_serad(a):
        sys.exit(1)


if __name__ == "__main__":
    main()
