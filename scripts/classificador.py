"""Regras de classificação de interlocutores e temas (calibradas contra as 222 reuniões já classificadas do Wilson: temas 95%, interlocutores 77%).
Os índices seguem config/rotulos.json. Classificação já gravada no JSON nunca é sobrescrita; ajustes manuais ficam em config/ajustes.json."""
import re,unicodedata
def N(s):
    if s is None or s!=s: return ''
    return unicodedata.normalize('NFD',str(s)).encode('ascii','ignore').decode()
def split_part(p):
    p=N(p); p=re.sub(r'\(CPF:[^)]*\)','',p)
    parts=re.split(r'agentes privados participantes:?',p,flags=re.I)
    pub=parts[0]; priv=' '.join(parts[1:])
    pub=re.sub(r'agentes publicos participantes:?','',pub,flags=re.I)
    toks=re.split(r'\n|\|',pub)
    keep=[t for t in toks if t.strip() and not re.search(r'ministerio das comunicacoes',t,re.I)]
    return ' | '.join(keep),priv
R={}
R['min']=r'ministro de estado|ministro das comunicacoes|frederico de siqueira|gabinete do ministro|reuniao com o ministro|com o ministro'
R[1]=r'anatel|agencia nacional de telecomunicacoes'
R[2]=r'\bUIT\b|\bITU\b|uniao internacional|embaixad|governo japon|japao|\bMIC\b|\bBID\b|banco mundial|banco interamericano|\bONU\b|\bOEA\b|\bIIC\b|\bAIR\b|associacao interamericana|\bABU\b|\bEBU\b|UNESCO|assuntos internos|uniao europeia|comissao europeia|\bNHK\b|\bDVB\b|\bATSC\b|\bBBC\b|\bCEPAL\b|mercosul|\bCAF\b|\bOCDE\b|delegacao|ministerio (da|de|do) .{0,40}(china|coreia|japao|portugal|franca|italia|alemanha|india|estados unidos|eua|uruguai|argentina|paraguai|chile|colombia|mexico|peru|equador|bolivia|venezuela|cuba|angola|mocambique|cabo verde)|international|internacional'
R[3]=r'\bABERT\b|\bABRATEL\b|\bSET\b|\bABRASAT\b|\bSINDISAT\b|\bACAERT\b|\bAERP\b|\bABRACO\b|sindicato|associacao (?!administradora)|federacao|\bFENAJ\b|\bAMARC\b|\bABCCOM\b|\bABRANET\b|\bTELCOMP\b|\bABTA\b|\bNEOTV\b|\bSINDITELEBRASIL\b|\bABRAFIX\b|\bAERT\b|\bARPUB\b|\bABEPEC\b|\bAPTEL\b|\bFNDC\b|intervozes|confederacao|\bFIESP\b|\bCNI\b|\bCNC\b|\bABPI\b|\bABRAPEC\b|\bABEMD\b|\bABINEE\b|\bABRADEE\b|\bFEBRATEL\b|\bAESP\b|\bFENACOM\b|\bABRAFI\b|\bABDTIC\b|\bABRINT\b|conselho (nacional )?de comunicacao|\bSBEG\b|\bFITERT\b|\bFENATRACOOP\b|\bFENEC\b'
R[4]=r'\bglobo\b|\bsbt\b|\brecord\b|\bband\b|bandeirantes|redetv|rede tv|\bcnt\b|rede vida|tv aparecida|\brbs\b|\brpc\b|\beptv\b|\bgcom\b|emissoras? de (radio|tv|televisao)|radiodifusor|geradora|retransmissora|\bjovem pan\b|\bsistema [a-z]+ de comunicacao'
R[5]=r'qualcomm|samsung|\bLG\b|elsys|mileto|gatesair|gates air|rohde|schwarz|st engineering|seja digital|\bEAD\b|\bclaro\b|\bvivo\b|\btelefonica\b|\btim\b|huawei|ericsson|nokia|\bsky\b|operadora|fabricante|\bzte\b|intelbras|positivo|\bsony\b|panasonic|philips|motorola|satelit|brisanet|mediatek|broadcom|dolby|google|\bmeta\b|apple|amazon|microsoft|hughes|viasat|starlink|intelsat|eutelsat|cisco|siemens|thomson|harmonic|digicap|blonder|hitachi|\bDTH\b|associacao administradora do processo|denki|electric|linear|\bCPQD\b|\bRNP\b|brasilsat|allegro|tropico|video technology|technolog|tecnolog|telecom|engenharia|systems|sistemas|eletronic|eletro|televisores|set-?top|conversor|\bAmericanet\b|\bsindisat\b'
R[6]=r'\bEBC\b|empresa brasil de comunicacao|correios|telebras|conselho fiscal|radiobras|\bRNCP\b|tv brasil|agencia brasil|\bSERPRO\b|\bDATAPREV\b|\bEPL\b|\bBNDES\b|\bFUNTTEL\b|\bFNDCT\b'
R[7]=r'deputad|senador|senado|camara dos deputados|camara municipal|prefeit|governador|assembleia|vereador|secretari[ao] (de estado|municipal|estadual)|governo do estado|governo estadual|municipio|\bFPA\b|bancada|frente parlamentar|comissao de comunicacao|comissao de ciencia|\bCCTCI\b|congresso nacional|parlamentar|\bsecretaria de estado|estado d[eo] '
R[8]=r'\bTCU\b|tribunal de contas|\bCGU\b|controladoria|casa civil|\bAECI\b|ministerio publico|\bMPF\b|\bAGU\b|advocacia-geral|auditoria|\bCADE\b|controle interno|presidencia da republica|\bSRI\b|planalto'
T={}
T[0]=r'tv ?3\.?0|dtv\+|\bsbtvd\b|gt tv 3|comite executivo'
T[1]=r'brasil digital|digitaliza|tv digital|seja digital|digitalizacao'
T[2]=r'gaispi|\bgired\b|3[,\.]5 ?ghz|interferencia'
T[3]=r'5g[ -]?broadcast|5g broadcasting|5g mbs|lte broadcast'
T[4]=r'outorga|licitac|edital|concessao|envelopes|habilitacao|selecao de municipios|chamamento|ted/unb|precificacao'
T[5]=r'radio[s]? comunitaria|radcom|educativ|radiodifusao comunitaria|comunitari|radio[s]? publica'
T[6]=r'conselho fiscal|\bROCF\b|correios'
T[7]=r'siga antenado|brasil antenado|antenado'
T[8]=r'600 ?mhz|faixa de 600|faixa de frequencia para radiodifusao'
def classificar(reg):
    """reg: dict com assunto, local, objetivos, info, participantes. Retorna (interlocutores, temas) como listas de índices."""
    subj=N(reg.get('assunto')); loc=N(reg.get('local')); obj=N(reg.get('objetivos')); info=N(reg.get('info'))
    pub,priv=split_part(reg.get('participantes'))
    cats=set()
    f=lambda k,txt,fl=re.I: re.search(R[k],txt,fl) is not None
    if re.search(R['min'],' | '.join([subj,loc,pub,info]),re.I): cats.add(0)
    if f(1,' | '.join([subj,pub,priv,info,obj])): cats.add(1)
    if f(2,' | '.join([subj,priv,info,obj,loc])): cats.add(2)
    if f(3,' | '.join([subj,priv,info,obj])): cats.add(3)
    sp=' | '.join([subj,priv])
    if re.search(R[4],sp,re.I) or re.search(r'\b(TV|Radio|Rede) (?!3|RO\b|Digital|Satelite|Aberta|Paga|Brasil|Por|a cabo|Comunitaria|Universitaria|Legislativa|Anatel|Nacional|Mundial)[A-Z]',sp): cats.add(4)
    if re.search(R[5],sp,re.I) or re.search(r'\b(Oi|Astral)\b',sp): cats.add(5)
    if f(6,' | '.join([subj,pub,priv,info])): cats.add(6)
    if f(7,' | '.join([subj,pub,priv,info])): cats.add(7)
    if f(8,' | '.join([subj,pub,priv,info,obj])): cats.add(8)
    tem=set(k for k,r in T.items() if re.search(r,' | '.join([subj,info,obj]),re.I))
    return sorted(cats),sorted(tem)
