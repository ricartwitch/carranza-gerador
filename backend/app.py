import os, io, re, traceback, tempfile, base64
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
FAIXA_H = 190000   # faixa vinho no topo (~15pt)
FAIXA_Y = 0        # colada no topo, sem espacamento
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
NS_R  = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_A  = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def _img_asset(k):
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
    b = _img_asset(k); b.seek(0)
    return slide.shapes.add_picture(b, Emu(l), Emu(t), Emu(w), Emu(h))

def _run(para, txt, bold=False, sz=635000, color=None):
    r = para.add_run(); r.text = txt
    r.font.bold = bold; r.font.size = Emu(sz); r.font.name = "Calibri"
    if color: r.font.color.rgb = color
    return r

def _faixa_rodape(slide):
    """Adiciona faixa vinho no rodapé de slides de conteúdo."""
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Emu(0), Emu(FAIXA_Y), Emu(SLIDE_W), Emu(FAIXA_H)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0x70, 0x00, 0x1C)
    shape.line.fill.background()  # sem borda

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
    _faixa_rodape(s)

def _slide_imagem(prs, img_bytes_b64, img_ext="png"):
    s = prs.slides.add_slide(_blank(prs))
    _pic(s,"logo_carranza",LOGO_MASTER_X,LOGO_MASTER_Y,LOGO_MASTER_CX,LOGO_MASTER_CY)
    margin_x = TEXTO_X
    margin_y = TEXTO_Y
    max_w = SLIDE_W - 2 * margin_x
    max_h = RODAPE_Y - margin_y - int(20 * 12700)
    img_bytes = base64.b64decode(img_bytes_b64)
    try:
        from PIL import Image as PILImage
        pil = PILImage.open(io.BytesIO(img_bytes))
        orig_w, orig_h = pil.size
        scale = min(max_w / (orig_w * 9144), max_h / (orig_h * 9144), 1.0)
        w_emu = int(orig_w * 9144 * scale)
        h_emu = int(orig_h * 9144 * scale)
    except Exception:
        w_emu = int(max_w * 0.8); h_emu = int(max_h * 0.8)
    left = (SLIDE_W - w_emu) // 2
    top = margin_y + (max_h - h_emu) // 2
    buf = io.BytesIO(img_bytes); buf.seek(0)
    s.shapes.add_picture(buf, Emu(left), Emu(top), Emu(w_emu), Emu(h_emu))
    _faixa_rodape(s)

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
    _faixa_rodape(s)

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
        tipo = s.get("tipo")
        if tipo == "contexto":
            _slide_conteudo(prs, [{"text": s.get("texto",""), "bold": True, "sz": 508000}])
            continue
        if tipo == "imagem":
            img_b64 = s.get("img_b64",""); img_ext = s.get("img_ext","png")
            if img_b64: _slide_imagem(prs, img_b64, img_ext)
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

def _get_img_from_para(para, doc_part):
    if '<w:drawing' not in para._element.xml: return None
    blip = para._element.find('.//' + NS_A + 'blip')
    if blip is None: return None
    rId = blip.get(NS_R + 'embed')
    if rId is None: return None
    try:
        rel = doc_part.rels.get(rId)
        if rel is None: return None
        ext = rel.target_ref.split('.')[-1].lower()
        if ext not in ('png','jpg','jpeg','gif','bmp','webp'): ext = 'png'
        return base64.b64encode(rel.target_part.blob).decode(), ext
    except Exception: return None

def _extrair_texto_docx(filepath):
    """Extrai texto bruto do docx para enviar ao Claude."""
    from docx import Document
    doc = Document(filepath)
    linhas = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            linhas.append(txt)
        else:
            linhas.append("")
    return "\n".join(linhas)

def _parse_via_claude(texto_bruto, disciplina="", assunto=""):
    """Usa Claude API para extrair questões de qualquer formato de texto."""
    import urllib.request, json as jsonlib
    prompt = f"""Você é um extrator de questões de concurso. Analise o texto abaixo e extraia TODAS as questões.

Retorne APENAS um JSON válido, sem texto antes ou depois, sem marcações markdown, com esta estrutura exata:
{{
  "questoes": [
    {{
      "numero": 1,
      "enunciado": "texto completo do enunciado",
      "alternativas": ["A) texto", "B) texto", "C) texto", "D) texto", "E) texto"],
      "certo_errado": false
    }}
  ],
  "gabarito": {{
    "questoes": [1, 2, 3],
    "respostas": ["A", "B", "C"]
  }}
}}

Regras:
- Capture o enunciado COMPLETO incluindo textos de apoio/excertos antes das alternativas
- Alternativas sempre no formato "A) texto", "B) texto" etc
- Se não houver gabarito no texto, retorne gabarito com listas vazias
- Se for questão Certo/Errado, coloque certo_errado=true e alternativas=[]
- Numere sequencialmente se não houver numeração explícita

TEXTO:
{texto_bruto[:8000]}
"""
    payload = jsonlib.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 4000,
        "messages": [{"role": "user", "content": prompt}]
    }).encode()
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "Content-Type": "application/json",
            "anthropic-version": "2023-06-01",
            "x-api-key": os.environ.get("ANTHROPIC_API_KEY", "")
        },
        method="POST"
    )
    with urllib.request.urlopen(req, timeout=60) as resp:
        data = jsonlib.loads(resp.read())
    txt = data["content"][0]["text"].strip()
    # Limpar markdown se houver
    if txt.startswith("```"): txt = "\n".join(txt.split("\n")[1:])
    if txt.endswith("```"): txt = "\n".join(txt.split("\n")[:-1])
    return jsonlib.loads(txt.strip())

def _parse_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    doc_part = doc.part
    slides, gqs, grs = [], [], []
    RE_NUM      = re.compile(r'^(\d{1,2})\.\s+(.+)', re.DOTALL)
    RE_NUM_ONLY = re.compile(r'^(\d{1,2})\.\s*')
    RE_LETRA    = re.compile(r'^[A-Ea-e]' + '$')
    RE_LETRA_SOLTA = re.compile(r'^([A-Ea-e])\s+(.+)', re.DOTALL)
    RE_ALT      = re.compile(r'^\([A-Ea-e]\)\s+.+|^[A-Ea-e][).]\s+.+')
    RE_CE       = re.compile(r'^(certo|errado)[\s.(]', re.IGNORECASE)
    RE_GAB_CELL = re.compile(r'^(\d{1,2})\.\s*([A-Ea-eCcEe])')
    RE_GAB_TX   = re.compile(r'(\d{1,2})\s*([A-Ea-e])\b')
    RE_ALT_SEP  = re.compile(r'^Alternativas$', re.IGNORECASE)
    def ib(p): return any(r.bold for r in p.runs if r.text.strip())
    def hn(p): return p._element.find('.//' + WS_NS + 'numPr') is not None
    def has_img(p): return '<w:drawing' in p._element.xml
    # Gabarito em tabela
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                m = RE_GAB_CELL.match(txt)
                if m: gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
                else:
                    for mm in RE_GAB_TX.finditer(txt):
                        gqs.append(int(mm.group(1))); grs.append(mm.group(2).upper())
    # Agrupar em blocos por linha vazia
    blocos = []
    bloco = []
    for para in doc.paragraphs:
        txt = para.text.strip(); is_img = has_img(para)
        if not txt and not is_img:
            if bloco: blocos.append(bloco); bloco = []
        else: bloco.append(para)
    if bloco: blocos.append(bloco)

    # --- DETECÇÃO DE PADRÃO ---

    # Padrão 4: "Alternativas" como separador
    tem_alt_sep = sum(1 for p in doc.paragraphs if RE_ALT_SEP.match(p.text.strip()))
    if tem_alt_sep >= 2:
        qnum = 0
        state_num = None; state_enc = []; state_in_alts = False; state_alts = []
        def flush(slides, state_num, state_enc, state_alts, qnum):
            enc = ' '.join(e for e in state_enc if e)
            if state_num: qnum = state_num
            else: qnum += 1
            alts = [a['letra'] + ') ' + (a['texto'] or '') for a in state_alts if a.get('texto')]
            slides.append({'tipo':'questao','numero':qnum,'enunciado':enc,'certo_errado':False,'alternativas':alts})
            return qnum
        for para in doc.paragraphs:
            txt = para.text.strip()
            if has_img(para):
                if state_in_alts or state_enc:
                    qnum = flush(slides, state_num, state_enc, state_alts, qnum)
                    state_num = None; state_enc = []; state_in_alts = False; state_alts = []
                ir = _get_img_from_para(para, doc_part)
                if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
                continue
            if not txt: continue
            m = RE_NUM_ONLY.match(txt)
            if m:
                if state_in_alts or state_enc:
                    qnum = flush(slides, state_num, state_enc, state_alts, qnum)
                state_num = int(m.group(1))
                state_enc = [txt]; state_in_alts = False; state_alts = []
                continue
            if RE_ALT_SEP.match(txt):
                state_in_alts = True; continue
            if state_in_alts:
                if RE_LETRA.match(txt): state_alts.append({'letra': txt, 'texto': None})
                elif state_alts and state_alts[-1]['texto'] is None: state_alts[-1]['texto'] = txt
                continue
            if state_num is not None: state_enc.append(txt)
        if state_in_alts or state_enc:
            flush(slides, state_num, state_enc, state_alts, qnum)
        gab = {"questoes": gqs, "respostas": grs} if gqs else None
        return slides, gab

    # Padrão 5: alternativas "A texto", "B texto" em blocos separados por linha vazia
    def is_alt_block(bloco):
        txts = [p.text.strip() for p in bloco]
        if len(txts) < 4: return False
        letras = []
        for t in txts:
            m = RE_LETRA_SOLTA.match(t)
            if m: letras.append(m.group(1).upper())
            else: return False
        return letras == list('ABCDE')[:len(letras)]

    tem_alt_blocos = sum(1 for b in blocos if is_alt_block(b))
    if tem_alt_blocos >= 2:
        qnum = 0; i = 0
        while i < len(blocos):
            bloco = blocos[i]
            txts = [p.text.strip() for p in bloco]
            if any('gabarito' in t.lower() for t in txts):
                for t in txts:
                    for m in RE_GAB_TX.finditer(t):
                        gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
                i += 1; continue
            if is_alt_block(bloco): i += 1; continue
            enc = ' '.join(txts)
            if i + 1 < len(blocos) and is_alt_block(blocos[i+1]):
                qnum += 1
                alts = []
                for t in [p.text.strip() for p in blocos[i+1]]:
                    m = RE_LETRA_SOLTA.match(t)
                    if m: alts.append(m.group(1).upper() + ') ' + m.group(2).strip())
                slides.append({'tipo':'questao','numero':qnum,'enunciado':enc,'certo_errado':False,'alternativas':alts})
                i += 2
            else:
                slides.append({'tipo':'contexto','texto':enc}); i += 1
        gab = {"questoes": gqs, "respostas": grs} if gqs else None
        return slides, gab

    # Padrão 3: blocos por linha vazia com alternativas (A), (B)...
    tem_alts_par = sum(1 for b in blocos
                       if any(p.text.strip().startswith('(') and RE_ALT.match(p.text.strip()) for p in b))
    if tem_alts_par >= 2:
        qnum = 0
        for bloco in blocos:
            txts = [p.text.strip() for p in bloco]
            imgs_bloco = [(_get_img_from_para(p, doc_part) if has_img(p) else None) for p in bloco]
            if any(t.upper() == 'GABARITO' for t in txts):
                for t in txts:
                    for m in RE_GAB_TX.finditer(t):
                        gqs.append(int(m.group(1))); grs.append(m.group(2).upper())
                continue
            alts_idx = [i for i, t in enumerate(txts) if RE_ALT.match(t)]
            if not alts_idx:
                txt_c = ' '.join(t for t in txts if t)
                if txt_c: slides.append({'tipo':'contexto','texto':txt_c})
                for ir in imgs_bloco:
                    if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
                continue
            first_alt = alts_idx[0]
            enc_txt = ' '.join(txts[:first_alt]) if txts[:first_alt] else (txts[0] if txts else '')
            alts = [txts[j] for j in alts_idx]
            m_num = RE_NUM.match(enc_txt) if enc_txt else None
            if m_num: qnum = int(m_num.group(1))
            else: qnum += 1
            ce = len(alts) > 0 and all(RE_CE.match(a) for a in alts)
            slides.append({'tipo':'questao','numero':qnum,'enunciado':enc_txt,'certo_errado':ce,'alternativas':alts if not ce else []})
            for ir in imgs_bloco:
                if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
        gab = {"questoes": gqs, "respostas": grs} if gqs else None
        return slides, gab

    # Padrão 6 (CEBRASPE): questões numeradas "01. texto" sem alternativas, certo/errado
    # Gabarito na tabela como "01.V" ou "01.F"
    RE_GAB_CE = re.compile(r'^(\d{1,2})\.\s*([VFvf])\b')
    # Coletar gabarito V/F da tabela
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for mm in RE_GAB_CE.finditer(cell.text.strip()):
                    q_n = int(mm.group(1))
                    letra = mm.group(2).upper()
                    if q_n not in gqs:  # evitar duplicata com RE_GAB_CELL
                        gqs.append(q_n)
                        grs.append('C' if letra == 'V' else 'E')
    tem_ce_tabela = any(RE_GAB_CE.search(c.text) for tbl in doc.tables for row in tbl.rows for c in row.cells)
    tem_num_seq = sum(1 for p in doc.paragraphs if RE_NUM_ONLY.match(p.text.strip()))
    nao_tem_alts = not (tem_alt_sep >= 2 or tem_alts_par >= 2 or tem_alt_blocos >= 2)
    usa_cebraspe = tem_ce_tabela and tem_num_seq >= 2 and nao_tem_alts
    if usa_cebraspe:
        state_enc = []; state_qnum = None
        def flush_ce(state_qnum, state_enc, slides):
            if state_qnum and state_enc:
                enc = ' '.join(e for e in state_enc if e)
                slides.append({'tipo':'questao','numero':state_qnum,'enunciado':enc,'certo_errado':True,'alternativas':[]})
        for para in doc.paragraphs:
            txt = para.text.strip()
            if has_img(para):
                flush_ce(state_qnum, state_enc, slides)
                state_qnum = None; state_enc = []
                ir = _get_img_from_para(para, doc_part)
                if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
                continue
            if not txt or txt.upper() == 'GABARITO': continue
            m = RE_NUM_ONLY.match(txt)
            if m:
                flush_ce(state_qnum, state_enc, slides)
                state_qnum = int(m.group(1)); state_enc = [txt]
            elif state_qnum is not None:
                state_enc.append(txt)
        flush_ce(state_qnum, state_enc, slides)
        gab = {"questoes": gqs, "respostas": grs} if gqs else None
        return slides, gab

    # Padrões 1 e 2: bold / List Paragraph
    paras = doc.paragraphs; i, qnum = 0, 0
    while i < len(paras):
        para = paras[i]; txt = para.text.strip()
        if has_img(para):
            ir = _get_img_from_para(para, doc_part)
            if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
            i += 1; continue
        if not txt: i += 1; continue
        bold = ib(para); m = RE_NUM.match(txt)
        if m and bold:
            qnum = int(m.group(1)); enc = txt; i += 1; alts = []; extras = []
            while i < len(paras):
                p2 = paras[i]; t2 = p2.text.strip()
                if has_img(p2):
                    ir = _get_img_from_para(p2, doc_part)
                    if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
                    i += 1; continue
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
                if has_img(p2):
                    ir = _get_img_from_para(p2, doc_part)
                    if ir: slides.append({'tipo':'imagem','img_b64':ir[0],'img_ext':ir[1]})
                    i += 1; continue
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

def _parse_pptx(filepath):
    """Extrai texto e imagens de um PPTX não formatado para gerar slides formatados."""
    prs_in = Presentation(filepath)
    # Coleta todo o texto bruto, slide a slide
    all_text = []
    img_slides = []  # (slide_index, img_bytes_b64, img_ext)

    for si, slide in enumerate(prs_in.slides):
        slide_texts = []
        for shape in slide.shapes:
            # Extrair imagens
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                try:
                    img_blob = shape.image.blob
                    content_type = shape.image.content_type or "image/png"
                    ext = content_type.split("/")[-1].replace("jpeg", "jpg")
                    if ext not in ("png", "jpg", "gif", "bmp", "webp"):
                        ext = "png"
                    img_b64 = base64.b64encode(img_blob).decode()
                    img_slides.append((si, img_b64, ext))
                except Exception:
                    pass
            # Extrair texto de textboxes e tabelas
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if txt:
                        slide_texts.append(txt)
            if shape.has_table:
                for row in shape.table.rows:
                    row_texts = []
                    for cell in row.cells:
                        ct = cell.text.strip()
                        if ct:
                            row_texts.append(ct)
                    if row_texts:
                        slide_texts.append(" | ".join(row_texts))
        if slide_texts:
            all_text.append((si, slide_texts))

    # Juntar todo o texto extraído em um bloco único
    linhas = []
    for si, texts in all_text:
        for t in texts:
            linhas.append(t)
        linhas.append("")  # separador entre slides

    texto_bruto = "\n".join(linhas)

    # Usar o mesmo parser de texto para extrair questões
    slides_parsed, gab = _parse_texto(texto_bruto)

    # Se o parser de texto não encontrou questões suficientes, inserir imagens como slides
    # Mapear imagens: inserir cada imagem como slide de imagem
    img_queue = list(img_slides)

    # Se não conseguiu parsear nenhuma questão via texto, tenta cada slide como contexto/imagem
    if not slides_parsed:
        for si, texts in all_text:
            txt_combined = " ".join(texts)
            if txt_combined.strip():
                slides_parsed.append({"tipo": "contexto", "texto": txt_combined})
        for si, img_b64, img_ext in img_queue:
            slides_parsed.append({"tipo": "imagem", "img_b64": img_b64, "img_ext": img_ext})
    else:
        # Adicionar imagens que não foram capturadas pelo parser de texto
        for si, img_b64, img_ext in img_queue:
            slides_parsed.append({"tipo": "imagem", "img_b64": img_b64, "img_ext": img_ext})

    return slides_parsed, gab

def _parse_pptx_via_claude(filepath, disciplina="", assunto=""):
    """Extrai texto de PPTX e usa Claude para interpretar questões."""
    prs_in = Presentation(filepath)
    linhas = []
    img_slides = []

    for si, slide in enumerate(prs_in.slides):
        slide_texts = []
        for shape in slide.shapes:
            if shape.shape_type == 13:
                try:
                    img_blob = shape.image.blob
                    content_type = shape.image.content_type or "image/png"
                    ext = content_type.split("/")[-1].replace("jpeg", "jpg")
                    if ext not in ("png", "jpg", "gif", "bmp", "webp"):
                        ext = "png"
                    img_b64 = base64.b64encode(img_blob).decode()
                    img_slides.append((si, img_b64, ext))
                except Exception:
                    pass
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if txt:
                        slide_texts.append(txt)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        ct = cell.text.strip()
                        if ct:
                            slide_texts.append(ct)
        for t in slide_texts:
            linhas.append(t)
        linhas.append("")

    texto_bruto = "\n".join(linhas)
    resultado = _parse_via_claude(texto_bruto, disciplina, assunto)

    sl = []
    for q in resultado.get("questoes", []):
        sl.append({
            "tipo": "questao",
            "numero": q.get("numero", 0),
            "enunciado": q.get("enunciado", ""),
            "certo_errado": q.get("certo_errado", False),
            "alternativas": q.get("alternativas", [])
        })

    # Adicionar imagens extraídas
    for si, img_b64, img_ext in img_slides:
        sl.append({"tipo": "imagem", "img_b64": img_b64, "img_ext": img_ext})

    gab_raw = resultado.get("gabarito", {})
    gab = gab_raw if gab_raw and gab_raw.get("questoes") else None
    return sl, gab


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
            formato = request.form.get("formato","word_to_slides")
            usar_ia = request.form.get("usar_ia","0") == "1"
            if not arq: return jsonify({"erro":"Arquivo nao enviado"}), 400
            fname = arq.filename.lower()

            if formato == "slides_to_slides":
                # --- SLIDES → SLIDES FORMATADOS ---
                if not fname.endswith(".pptx"):
                    return jsonify({"erro":"Para o formato Slides→Slides, envie um arquivo .pptx"}), 400
                with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
                    arq.save(tmp.name); path = tmp.name
                try:
                    if usar_ia:
                        sl, gab = _parse_pptx_via_claude(path, disc, ass)
                    else:
                        sl, gab = _parse_pptx(path)
                finally: os.unlink(path)

            elif fname.endswith(".docx"):
                # --- WORD → SLIDES (fluxo original) ---
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                    arq.save(tmp.name); path = tmp.name
                try:
                    if usar_ia:
                        texto_bruto = _extrair_texto_docx(path)
                        resultado = _parse_via_claude(texto_bruto, disc, ass)
                        sl = []
                        for q in resultado.get("questoes", []):
                            sl.append({
                                "tipo": "questao",
                                "numero": q.get("numero", 0),
                                "enunciado": q.get("enunciado", ""),
                                "certo_errado": q.get("certo_errado", False),
                                "alternativas": q.get("alternativas", [])
                            })
                        gab_raw = resultado.get("gabarito", {})
                        gab = gab_raw if gab_raw and gab_raw.get("questoes") else None
                    else:
                        sl, gab = _parse_docx(path)
                finally: os.unlink(path)
            else:
                # --- TXT → SLIDES (fluxo original) ---
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
