import os, io, re, traceback, tempfile
from itertools import cycle
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

SLIDE_W, SLIDE_H = 12192000, 6858000
VINHO = RGBColor(0x70, 0x00, 0x1C)
PRETO = RGBColor(0x00, 0x00, 0x00)
LOGO_MASTER_X,LOGO_MASTER_Y,LOGO_MASTER_CX,LOGO_MASTER_CY = 9282544,385975,2507673,776504
LOGO_CAPA_X,LOGO_CAPA_Y,LOGO_CAPA_CX,LOGO_CAPA_CY = 4429838,558156,3251122,764208
LOGO_ENC_X,LOGO_ENC_Y,LOGO_ENC_CX,LOGO_ENC_CY = 2756357,2683952,6679286,1490095
TEXTO_X,TEXTO_Y,TEXTO_CX,TEXTO_H = 347272,1074509,11497455,5200000
RODAPE_X,RODAPE_Y,RODAPE_CX,RODAPE_CY = 3382781,6298579,8461946,400110
CAPA_X,CAPA_Y,CAPA_CX,CAPA_CY = 2548009,2013228,7465441,2985433
AREA_UTIL = (RODAPE_Y - TEXTO_Y) - int(40 * 12700)
LARGURA = 11497455
FATOR = 0.42
CITACOES = [
    '"A forca esta em se erguer com mais poder, como a aguia."',
    '"A aguia foca no futuro, nao no que esta atras."',
    '"Para voar alto, deixe as correntes para tras, como a aguia."',
    '"Nas tempestades, como a aguia, encontre o voo mais alto."',
]
ASSETS = os.path.join(os.path.dirname(__file__), "assets")
WS_NS = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def _img(k):
    p = {"capa_background": "capa_background.png",
         "logo_carranza": "logo_carranza.png",
         "encerramento": "encerramento.png"}[k]
    with open(os.path.join(ASSETS, p), "rb") as f:
        return io.BytesIO(f.read())

def _blank(prs):
    for lay in prs.slide_layouts:
        if lay.name.lower() in ("em branco","blank",""): return lay
    return prs.slide_layouts[6]

def _pic(slide, k, l, t, w, h):
    b = _img(k); b.seek(0)
    return slide.shapes.add_picture(b, Emu(l), Emu(t), Emu(w), Emu(h))

def _run(para, txt, bold=False, sz=635000, color=None):
    r = para.add_run(); r.text = txt
    r.font.bold = bold; r.font.size = Emu(sz); r.font.name = "Calibri"
    if color: r.font.color.rgb = color
    return r

def _slide_capa(prs, dados):
    s = prs.slides.add_slide(_blank(prs))
    _pic(s,"capa_background",0,0,SLIDE_W,SLIDE_H)
    _pic(s,"logo_carranza",LOGO_CAPA_X,LOGO_CAPA_Y,LOGO_CAPA_CX,LOGO_CAPA_CY)
    tb = s.shapes.add_textbox(Emu(CAPA_X),Emu(CAPA_Y),Emu(CAPA_CX),Emu(CAPA_CY))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for txt, bold, sz in [
        (dados.get("disciplina","").upper(), True, 635000),
        (dados.get("assunto","").upper(), True, 635000),
        (dados.get("tipo","QUESTOES").upper(), True, 508000),
        (dados.get("professor",""), False, 355600),
    ]:
        if not txt: continue
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False; p.alignment = PP_ALIGN.CENTER
        _run(p, txt, bold=bold, sz=sz, color=VINHO)

def _slide_conteudo(prs, paragrafos, citacao=None):
    s = prs.slides.add_slide(_blank(prs))
    _pic(s,"logo_carranza",LOGO_MASTER_X,LOGO_MASTER_Y,LOGO_MASTER_CX,LOGO_MASTER_CY)
    tb = s.shapes.add_textbox(Emu(TEXTO_X),Emu(TEXTO_Y),Emu(TEXTO_CX),Emu(TEXTO_H))
    tf = tb.text_frame; tf.word_wrap = True
    first = True
    for p in paragrafos:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False; para.alignment = PP_ALIGN.JUSTIFY
        _run(para, p.get("text",""), bold=p.get("bold",False),
             sz=p.get("sz",571500), color=p.get("color",PRETO))
    if citacao:
        tb2 = s.shapes.add_textbox(Emu(RODAPE_X),Emu(RODAPE_Y),Emu(RODAPE_CX),Emu(RODAPE_CY))
        p2 = tb2.text_frame.paragraphs[0]; p2.alignment = PP_ALIGN.RIGHT
        _run(p2, citacao, bold=True, sz=254000, color=VINHO)

def _slide_gabarito(prs, qs, rs):
    s = prs.slides.add_slide(_blank(prs))
    _pic(s,"logo_carranza",LOGO_MASTER_X,LOGO_MASTER_Y,LOGO_MASTER_CX,LOGO_MASTER_CY)
    tb = s.shapes.add_textbox(Emu(4000000),Emu(2400000),Emu(4200000),Emu(600000))
    p = tb.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _run(p, "GABARITO", bold=True, sz=304800, color=VINHO)
    n = len(qs)
    if not n: return
    tbl = s.shapes.add_table(2, n, Emu(2031997), Emu(3058160), Emu(8128005), Emu(741680)).table
    for j, (q, r) in enumerate(zip(qs, rs)):
        c0 = tbl.cell(0,j); c0.text = str(q).zfill(2)
        p0 = c0.text_frame.paragraphs[0]; p0.alignment = PP_ALIGN.CENTER
        for ru in p0.runs:
            ru.font.bold = True; ru.font.size = Pt(18); ru.font.color.rgb = VINHO
        c1 = tbl.cell(1,j); c1.text = str(r).upper()
        p1 = c1.text_frame.paragraphs[0]; p1.alignment = PP_ALIGN.CENTER
        for ru in p1.runs:
            ru.font.bold = True; ru.font.size = Pt(18); ru.font.color.rgb = PRETO

def _slide_enc(prs):
    s = prs.slides.add_slide(_blank(prs))
    _pic(s,"capa_background",0,0,SLIDE_W,SLIDE_H)
    _pic(s,"logo_carranza",LOGO_ENC_X,LOGO_ENC_Y,LOGO_ENC_CX,LOGO_ENC_CY)

def _h(txt, sz):
    if not txt.strip(): return int(sz*0.4)
    cpp = LARGURA / (sz * FATOR)
    return int(sz * 1.15 * max(1.0, len(txt)/cpp))

def _sz(te, alts):
    total = len(te) + sum(len(a) for a in alts)
    if total > 1200: return 406400
    if total > 800:  return 444500
    if total > 500:  return 482600
    return 533400

def _distribuir(te, alts, sz):
    grupos, grupo, altura = [], [], _h(te, sz)
    grupo.append({"text": te, "bold": True, "sz": sz})
    for alt in alts:
        ha = _h(alt, sz)
        if altura + ha > AREA_UTIL and grupo:
            grupos.append(grupo); grupo = []; altura = 0
        grupo.append({"text": alt, "bold": False, "sz": sz})
        altura += ha
    if grupo: grupos.append(grupo)
    return grupos

def _build_pptx(payload):
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W); prs.slide_height = Emu(SLIDE_H)
    dados = {k: payload.get(k,"") for k in ("disciplina","assunto","tipo","professor")}
    if not dados["tipo"]: dados["tipo"] = "QUESTOES"
    _slide_capa(prs, dados)
    cit = cycle(CITACOES)
    for s in payload.get("slides", []):
        if s.get("tipo") == "contexto":
            _slide_conteudo(prs, [{"text": s.get("texto",""), "bold": True, "sz": 508000}])
            continue
        num = s.get("numero",""); enc = s.get("enunciado","")
        alts = s.get("alternativas",[]); ce = s.get("certo_errado", False)
        snum = str(num).zfill(2) if num else ""
        te = enc if (snum and (enc.startswith(snum+".") or enc.startswith(str(num)+"."))) else (snum+". "+enc if snum else enc)
        if ce:
            sz = _sz(te, ["Certo (  )","Errado (  )"])
            _slide_conteudo(prs, [{"text":te,"bold":True,"sz":sz},{"text":"Certo (  )","bold":False,"sz":sz},{"text":"Errado (  )","bold":False,"sz":sz}], citacao=next(cit))
        elif alts:
            sz = _sz(te, alts)
            grupos = _distribuir(te, alts, sz)
            for gi, grupo in enumerate(grupos):
                _slide_conteudo(prs, grupo, citacao=next(cit) if gi==len(grupos)-1 else None)
        else:
            sz = _sz(te, [])
            _slide_conteudo(prs, [{"text":te,"bold":True,"sz":sz}], citacao=next(cit))
    gab = payload.get("gabarito")
    if gab and gab.get("questoes"):
        _slide_gabarito(prs, gab["questoes"], gab["respostas"])
    _slide_enc(prs)
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf

def _parse_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    slides, gqs, grs = [], [], []
    RE_NUM  = re.compile(r'^(\d{1,2})[.)]\s+(.+)', re.DOTALL)
    RE_LETRA = re.compile(r'^[A-Ea-e]$')

def _parse_texto(texto):
    linhas = [l.rstrip() for l in texto.splitlines()]
    slides, gqs, grs = [], [], []
    RE_Q = re.compile(r'^(\d{1,2})[.\-]\s+(.+)')
    RE_A = re.compile(r'^[A-Ea-e][).]')
    RE_G = re.compile(r'^GABARITO', re.IGNORECASE)
    RE_GI = re.compile(r'(\d{1,2})\s*[-]\s*([A-Ea-eCcEe])\b')
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        if RE_G.match(l):
            bloco = l; i += 1
            while i < len(linhas): bloco += ' ' + linhas[i].strip(); i += 1
            for m in RE_GI.finditer(bloco): gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
            continue
        m = RE_Q.match(l)
        if m:
            num = int(m.group(1)); ep = [m.group(2).strip()]; i += 1; alts = []; ce = False
            while i < len(linhas):
                ll = linhas[i].strip()
                if not ll:
                    i += 1
                    if i < len(linhas) and (RE_Q.match(linhas[i].strip()) or RE_G.match(linhas[i].strip())): break
                    continue
                if RE_A.match(ll): alts.append(ll); i += 1
                elif ll.lower().startswith('certo') or ll.lower().startswith('errado'): ce = True; i += 1
                elif RE_Q.match(ll) or RE_G.match(ll): break
                else: ep.append(ll); i += 1
            slides.append({"tipo":"questao","numero":num,"enunciado":" ".join(ep),"certo_errado":ce,"alternativas":alts})
            continue
        if l and not RE_G.match(l):
            cp = [l]; i += 1
            while i < len(linhas):
                ll = linhas[i].strip()
                if RE_Q.match(ll) or RE_G.match(ll): break
                if not ll:
                    i += 1
                    if i < len(linhas) and not linhas[i].strip(): break
                    continue
                cp.append(ll); i += 1
            slides.append({"tipo":"contexto","texto":" ".join(cp)})
            continue
        i += 1
    gab = {"questoes": gqs, "respostas": grs} if gqs else None
    return slides, gab

app = Flask(__name__)
CORS(app, origins="*")

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        if request.files or request.form:
            arq = request.files.get("arquivo")
            disc = request.form.get("disciplina","DISCIPLINA")
            ass  = request.form.get("assunto","ASSUNTO")
            prof = request.form.get("professor","")
            tipo = request.form.get("tipo","QUESTOES")
            if not arq: return jsonify({"erro":"Arquivo nao enviado"}), 400
            if arq.filename.lower().endswith(".docx"):
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                    arq.save(tmp.name); path = tmp.name
                try: sl, gab = _parse_docx(path)
                finally: os.unlink(path)
            else:
                texto = arq.read().decode("utf-8", errors="ignore")
                sl, gab = _parse_texto(texto)
            payload = {"disciplina":disc,"assunto":ass,"tipo":tipo,"professor":prof,"slides":sl,"gabarito":gab}
        else:
            payload = request.get_json(force=True)
            if not payload: return jsonify({"erro":"Payload vazio"}), 400
        buf = _build_pptx(payload)
        d = payload.get("disciplina","apresentacao").replace(" ","_")
        a = payload.get("assunto","").replace(" ","_")
        fn = "Carranza_" + d + "_" + a + ".pptx" if a else "Carranza_" + d + ".pptx"
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                         as_attachment=True, download_name=fn)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
)
    RE_ALT  = re.compile(r'^\([A-Ea-e]\)\s+.+|^[A-Ea-e][).]\s+.+')
    RE_CE   = re.compile(r'^(certo|errado)[\s.(]', re.IGNORECASE)
    RE_GAB_CELL = re.compile(r'^(\d{1,2})\.\s*([A-Ea-eCcEe])')
    RE_GAB_TX   = re.compile(r'(\d{1,2})\s*([A-Ea-e])\b')
    def ib(p): return any(r.bold for r in p.runs if r.text.strip())
    def hn(p): return p._element.find('.//' + WS_NS + 'numPr') is not None
    # Gabarito em tabela
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                m = RE_GAB_CELL.match(cell.text.strip())
                if m: gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
    # Agrupar parágrafos em blocos separados por linha vazia
    blocos = []
    bloco = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if not txt:
            if bloco: blocos.append(bloco); bloco = []
        else:
            bloco.append(para)
    if bloco: blocos.append(bloco)
    # Detectar padrão: se todos blocos têm alternativas claras → parsear por bloco
    tem_alts = sum(1 for b in blocos if any(RE_ALT.match(p.text.strip()) for p in b))
    usa_blocos = tem_alts >= 2
    if usa_blocos:
        qnum = 0
        for bloco in blocos:
            txts = [p.text.strip() for p in bloco]
            bolds = [ib(p) for p in bloco]
            # GABARITO
            if any(t.upper() == 'GABARITO' for t in txts):
                for t in txts:
                    for m in RE_GAB_TX.finditer(t):
                        gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
                continue
            alts_idx = [i for i, t in enumerate(txts) if RE_ALT.match(t)]
            if not alts_idx:
                slides.append({'tipo':'contexto','texto':' '.join(txts)})
                continue
            first_alt = alts_idx[0]
            enc_parts = txts[:first_alt]
            alts = [txts[j] for j in alts_idx]
            enc_txt = ' '.join(enc_parts) if enc_parts else txts[0] if txts else ''
            m_num = RE_NUM.match(enc_txt) if enc_txt else None
            if m_num: qnum = int(m_num.group(1))
            else: qnum += 1
            ce = len(alts) > 0 and all(RE_CE.match(a) for a in alts)
            slides.append({'tipo':'questao','numero':qnum,'enunciado':enc_txt,'certo_errado':ce,'alternativas':alts if not ce else []})
        gab = {"questoes": gqs, "respostas": grs} if gqs else None
        return slides, gab
    # Parsear parágrafo a parágrafo (padrões 1 e 2)
    paras = doc.paragraphs
    i, qnum = 0, 0
    while i < len(paras):
        para = paras[i]; txt = para.text.strip()
        if not txt: i += 1; continue
        bold = ib(para); m = RE_NUM.match(txt)
        if m and bold:
            qnum = int(m.group(1)); enc = txt; i += 1; alts = []; extras = []
            while i < len(paras):
                p2 = paras[i]; t2 = p2.text.strip()
                if not t2: i += 1; continue
                b2 = ib(p2); m2 = RE_NUM.match(t2)
                if m2 and b2: break
                if p2.style.name == 'List Paragraph' and hn(p2): break
                if RE_CE.match(t2): alts.append(t2); i += 1
                elif RE_ALT.match(t2): alts.append(t2); i += 1
                elif b2 and RE_LETRA.match(t2):
                    letra = t2; i += 1
                    if i < len(paras) and paras[i].text.strip():
                        alts.append(letra + ') ' + paras[i].text.strip()); i += 1
                else: extras.append(t2); i += 1
            enc2 = enc + ('\n' + '\n'.join(extras) if extras else '')
            ce = len(alts) > 0 and all(RE_CE.match(a) for a in alts)
            slides.append({'tipo':'questao','numero':qnum,'enunciado':enc2,'certo_errado':ce,'alternativas':alts if not ce else []})
            continue
        if para.style.name == 'List Paragraph' and hn(para) and txt:
            qnum += 1; enc = txt; i += 1; alts = []; extras = []
            while i < len(paras):
                p2 = paras[i]; t2 = p2.text.strip()
                if not t2: i += 1; continue
                b2 = ib(p2)
                if p2.style.name == 'List Paragraph' and hn(p2): break
                if RE_NUM.match(t2) and b2: break
                if b2 and RE_LETRA.match(t2):
                    letra = t2; i += 1
                    if i < len(paras):
                        t3 = paras[i].text.strip(); b3 = ib(paras[i])
                        if t3 and not (b3 and RE_LETRA.match(t3)):
                            alts.append(letra + ') ' + t3); i += 1
                    continue
                if RE_CE.match(t2): alts.append(t2); i += 1; continue
                if RE_ALT.match(t2): alts.append(t2); i += 1; continue
                if not alts: extras.append(t2)
                i += 1
            enc2 = enc + ('\n' + '\n'.join(extras) if extras else '')
            ce = len(alts) > 0 and all(RE_CE.match(a) for a in alts)
            slides.append({'tipo':'questao','numero':qnum,'enunciado':enc2,'certo_errado':ce,'alternativas':alts if not ce else []})
            continue
        i += 1
    gab = {"questoes": gqs, "respostas": grs} if gqs else None
    return slides, gab

def _parse_texto(texto):
    linhas = [l.rstrip() for l in texto.splitlines()]
    slides, gqs, grs = [], [], []
    RE_Q = re.compile(r'^(\d{1,2})[.\-]\s+(.+)')
    RE_A = re.compile(r'^[A-Ea-e][).]')
    RE_G = re.compile(r'^GABARITO', re.IGNORECASE)
    RE_GI = re.compile(r'(\d{1,2})\s*[-]\s*([A-Ea-eCcEe])\b')
    i = 0
    while i < len(linhas):
        l = linhas[i].strip()
        if RE_G.match(l):
            bloco = l; i += 1
            while i < len(linhas): bloco += ' ' + linhas[i].strip(); i += 1
            for m in RE_GI.finditer(bloco): gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
            continue
        m = RE_Q.match(l)
        if m:
            num = int(m.group(1)); ep = [m.group(2).strip()]; i += 1; alts = []; ce = False
            while i < len(linhas):
                ll = linhas[i].strip()
                if not ll:
                    i += 1
                    if i < len(linhas) and (RE_Q.match(linhas[i].strip()) or RE_G.match(linhas[i].strip())): break
                    continue
                if RE_A.match(ll): alts.append(ll); i += 1
                elif ll.lower().startswith('certo') or ll.lower().startswith('errado'): ce = True; i += 1
                elif RE_Q.match(ll) or RE_G.match(ll): break
                else: ep.append(ll); i += 1
            slides.append({"tipo":"questao","numero":num,"enunciado":" ".join(ep),"certo_errado":ce,"alternativas":alts})
            continue
        if l and not RE_G.match(l):
            cp = [l]; i += 1
            while i < len(linhas):
                ll = linhas[i].strip()
                if RE_Q.match(ll) or RE_G.match(ll): break
                if not ll:
                    i += 1
                    if i < len(linhas) and not linhas[i].strip(): break
                    continue
                cp.append(ll); i += 1
            slides.append({"tipo":"contexto","texto":" ".join(cp)})
            continue
        i += 1
    gab = {"questoes": gqs, "respostas": grs} if gqs else None
    return slides, gab

app = Flask(__name__)
CORS(app, origins="*")

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok"})

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        if request.files or request.form:
            arq = request.files.get("arquivo")
            disc = request.form.get("disciplina","DISCIPLINA")
            ass  = request.form.get("assunto","ASSUNTO")
            prof = request.form.get("professor","")
            tipo = request.form.get("tipo","QUESTOES")
            if not arq: return jsonify({"erro":"Arquivo nao enviado"}), 400
            if arq.filename.lower().endswith(".docx"):
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                    arq.save(tmp.name); path = tmp.name
                try: sl, gab = _parse_docx(path)
                finally: os.unlink(path)
            else:
                texto = arq.read().decode("utf-8", errors="ignore")
                sl, gab = _parse_texto(texto)
            payload = {"disciplina":disc,"assunto":ass,"tipo":tipo,"professor":prof,"slides":sl,"gabarito":gab}
        else:
            payload = request.get_json(force=True)
            if not payload: return jsonify({"erro":"Payload vazio"}), 400
        buf = _build_pptx(payload)
        d = payload.get("disciplina","apresentacao").replace(" ","_")
        a = payload.get("assunto","").replace(" ","_")
        fn = "Carranza_" + d + "_" + a + ".pptx" if a else "Carranza_" + d + ".pptx"
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                         as_attachment=True, download_name=fn)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT",5000)), debug=False)
