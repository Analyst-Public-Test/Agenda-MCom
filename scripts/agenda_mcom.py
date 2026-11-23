#!/usr/bin/env python3
"""Agenda MCom: e-Agendas (CGU) + Portal da Transparência (viagens) -> dados/agenda-mcom.json

Fontes
  e-Agendas : página pública de cada autoridade (a agenda inteira vem embutida no HTML, sem captcha).
  Portal    : zip anual de viagens (GET direto) -> {ano}_Viagem.csv.
  Plano B   : arquivos em dados/entrada/ (xlsx/csv do e-Agendas; zip/csv de viagens em dados/entrada/viagens/).

Regras
  - Só entram anos >= ano_inicio_automatico (2026). Anos anteriores ficam fixos no HTML.
  - Se uma fonte falhar, o que já está no JSON é mantido (nada some por causa de uma falha de rede).
  - Classificação já gravada no JSON não é sobrescrita; config/ajustes.json vence sempre.
  - Compromisso com o mesmo ID (ou mesma data/hora/assunto) em mais de uma agenda vira um registro só, com todas as pessoas.
  - O JSON só é regravado quando muda o conteúdo OU o estado das fontes (assim o bloco de controle nunca fica desatualizado).
"""
import argparse, csv, datetime as dt, difflib, hashlib, html, io, json, re, sys, time, unicodedata, zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent))
from classificador import classificar  # noqa: E402

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


# ----------------------------------------------------------------------------------------------
# Portal da Transparência (viagens)
# ----------------------------------------------------------------------------------------------
CIDADE = {"Amsterdam": "Amsterdã", "Santo Domingo de Guzmán": "Santo Domingo"}
PAIS = {"Estados Unidos da América": "EUA"}


def lugar_destino(dest):
    cidade, _, fim = dest.strip().rpartition("/")
    if re.fullmatch(r"[A-Z]{2}", fim):
        return f"{fim}:{cidade}"
    return f"EX:{CIDADE.get(cidade, cidade)} ({PAIS.get(fim, fim)})"


def baixar_zip_viagens(cfg, ano, destino):
    import requests
    url = cfg["url_viagens"].format(ano=ano)
    r = requests.get(url, headers={"User-Agent": UA}, stream=True, timeout=600, allow_redirects=True)
    if r.status_code == 404:
        return None, "sem arquivo para o ano"
    r.raise_for_status()
    with open(destino, "wb") as f:
        for ch in r.iter_content(1 << 20):
            f.write(ch)
    if destino.read_bytes()[:2] != b"PK":
        raise RuntimeError("o download não é um zip (possível bloqueio do site)")
    return r.url.rsplit("/", 1)[-1], "ok"


def viagens_do_csv(fh, ano, por_nome):
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
                    r[ix["identificador do processo de viagem"]].lstrip("0"), limpo(r[ix["motivo"]])])
    return out


def ler_viagens_arquivo(path, ano, por_nome):
    if path.suffix.lower() == ".zip":
        with zipfile.ZipFile(path) as z:
            nome = next(n for n in z.namelist() if n.endswith("_Viagem.csv"))
            with z.open(nome) as fh:
                return viagens_do_csv(fh, ano, por_nome)
    with open(path, "rb") as fh:
        return viagens_do_csv(fh, ano, por_nome)


# ----------------------------------------------------------------------------------------------
# Lugar
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


def decidir_lugar(reg, vgs, pendencias):
    if reg["virtual"\]:
        return "VI:Virtual", "V"
    if re.search(DF_FORTE, reg["local"], re.I):
        return "DF:Brasília", "P"
    for v in vgs:
        if v[2] in reg["pes"] and v[0] <= reg["iso"] <= v[1\]:
            return v[3], "P"
    for rx, l in CIDADES:
        if re.search(rx, reg["local"] + " " + reg["assunto"], re.I):
            return l, "P"
    pendencias.append({"id": reg["ids"][0], "data": reg["iso"], "local": reg["local"], "assunto": reg["assunto"][:100]})
    return "DF:Brasília", "P"


# ----------------------------------------------------------------------------------------------
# Principal
# ----------------------------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--saida", default=str(ROOT / "dados" / "agenda-mcom.json"))
    ap.add_argument("--sem-rede", action="store_true", help="usa só dados/entrada (teste local)")
    ap.add_argument("--pagina-local", action="append", default=[], metavar="CHAVE=arquivo.html",
                    help="lê a página salva de uma autoridade em vez de baixar (teste)")
    ap.add_argument("--reclassificar", action="store_true", help="ignora a classificação já gravada")
    a = ap.parse_args()

    cfg = json.loads((ROOT / "config" / "autoridades.json").read_text(encoding="utf-8"))
    rot = json.loads((ROOT / "config" / "rotulos.json").read_text(encoding="utf-8"))
    ajustes = {k: v for k, v in json.loads((ROOT / "config" / "ajustes.json").read_text(encoding="utf-8")).items()
               if not k.startswith("_")}
    auts = cfg["autoridades"]
    por_nome = {N(x["nome"]).upper(): x["chave"] for x in auts}
    ano0 = cfg["ano_inicio_automatico"]
    hoje = dt.date.today()
    anos = list(range(ano0, hoje.year + 1))
    saida = Path(a.saida)
    estado = json.loads(saida.read_text(encoding="utf-8")) if saida.exists() else {}
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
            log(f"[e-Agendas] {aut['nome']}: {n} compromissos desde {ano0}")
        except Exception as e:  # noqa: BLE001
            falhou.add(k)
            fontes[k] = {"fonte": "e-Agendas (página)", "ok": False, "erro": str(e)[:200]}
            log(f"[e-Agendas] FALHOU {aut['nome']}: {e}")

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
        log(f"[entrada] {p.name}: {n} registros")

    # ---- 2) viagens
    vg, fonte_viag = [], {}
    import tempfile
    for ano in anos:
        rows_ano, info = None, ""
        try:
            if a.sem_rede:
                raise RuntimeError("modo --sem-rede")
            with tempfile.TemporaryDirectory() as td:
                z = Path(td) / f"{ano}.zip"
                nome, info = baixar_zip_viagens(cfg, ano, z)
                if nome:
                    rows_ano = ler_viagens_arquivo(z, ano, por_nome)
                    info = nome
        except Exception as e:  # noqa: BLE001
            info = f"falhou: {str(e)[:150]}"
        if rows_ano is None:
            for p in sorted((ROOT / "dados" / "entrada" / "viagens").glob("*")):
                if p.suffix.lower() in (".zip", ".csv") and str(ano) in p.name:
                    rows_ano = ler_viagens_arquivo(p, ano, por_nome)
                    info = f"arquivo local {p.name}"
        if rows_ano is None:
            rows_ano = [v for v in antigo_vg if v[0][:4] == str(ano)]
            if rows_ano:
                info += " (mantido o que já estava no JSON)"
        vg += rows_ano
        fonte_viag[str(ano)] = {"arquivo": info, "viagens": len(rows_ano)}
        log(f"[Portal] {ano}: {len(rows_ano)} viagens ({info})")
    vg.sort(key=lambda v: (v[0], v[2]))

    # ---- 3) mantém o que já existia para autoridades cuja página falhou
    for r in antigo_rows:
        ids = r[12] if len(r) > 12 else []
        if not (falhou & set(r[9])) or r[0][:4] < str(ano0or not ids:
            continue
        achou = [i for i in ids if i in regs]
        if achou:  # compromisso compartilhado: a pessoa cuja fonte falhou continua nele
            regs[achou[0]]["pes"] |= falhou & set(r[9])
        else:
            regs[ids[0]] = {"id": ids[0], "pes": set(r[9]), "iso": r[0], "hora": r[1], "tipo": r[2], "assunto": r[3],
                            "local": r[4], "virtual": r[5] == "V", "participantes": "", "objetivos": "", "info": "",
                            "_manter": r, "fonte": "estado"}

    # ---- 4) une IDs diferentes do mesmo compromisso (mesma data/hora/assunto)
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

    # ---- 6) saída (regrava se mudou o conteúdo OU o estado das fontes)
    corpo = {"cats": rot["cats"], "tem": rot["tem"], "rows": rows, "vg": vg}
    ult = {c["chave"]: max([r[0] for r in rows if c["chave"] in r[9]] or [""]) for c in auts}
    meta = {"fontes_eagendas": fontes, "fontes_viagens": fonte_viag, "ultimo_compromisso": ult,
            "pendencias_de_lugar": pendencias}
    # assinatura do estado das fontes: sem textos de erro (que variam a cada execução)
    sinal = {
        "e": {k: [v["ok"], v.get("compromissos")] for k, v in fontes.items()},
        "v": {y: ["falhou" if "falhou" in x["arquivo"] else x["arquivo"], x["viagens"]] for y, x in fonte_viag.items()},
    }
    h = hashlib.sha256(json.dumps([corpo, sinal], ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:16]
    if estado.get("hash") == h:
        log("Sem mudanças no conteúdo; JSON não foi alterado.")
    else:
        out = {"atualizado_em": dt.datetime.now(dt.timezone(dt.timedelta(hours=-3))).strftime("%Y-%m-%dT%H:%M:%S-03:00"),
               "hash": h, **meta, **corpo}
        saida.parent.mkdir(parents=True, exist_ok=True)
        saida.write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
        log(f"Gravado {saida} ({len(rows)} compromissos, {len(vg)} viagens)")
    if pendencias:
        log(f"Atenção: {len(pendencias)} compromisso(s) presencial(is) sem lugar identificado (ficaram em Brasília); veja 'pendencias_de_lugar'.")
    if len(falhou) == len(auts) and not regs:
        log("ERRO: nenhuma fonte de agenda respondeu.")
        sys.exit(1)


if __name__ == "__main__":
    main()
