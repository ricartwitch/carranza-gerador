"""
Carranza Cursos — Gerador de PPTX
Backend Flask para o Render
"""

import os
import io
import re
import traceback
import tempfile
from itertools import cycle

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ─────────────────────────────────────────────
# Constantes visuais (medidas exatas do template)
# ─────────────────────────────────────────────

SLIDE_W = 12192000
SLIDE_H = 6858000

VINHO = RGBColor(0x70, 0x00, 0x1C)
PRETO = RGBColor(0x00, 0x00, 0x00)

LOGO_MASTER_X  = 9282544;  LOGO_MASTER_Y  = 385975
LOGO_MASTER_CX = 2507673;  LOGO_MASTER_CY = 776504

LOGO_CAPA_X  = 4429838;  LOGO_CAPA_Y  = 558156
LOGO_CAPA_CX = 3251122;  LOGO_CAPA_CY = 764208

LOGO_ENC_X  = 2756357;  LOGO_ENC_Y  = 2683952
LOGO_ENC_CX = 6679286;  LOGO_ENC_CY = 1490095

TEXTO_X  = 347272;  TEXTO_Y  = 1074509
TEXTO_CX = 11497455;  TEXTO_H = 5200000

RODAPE_X  = 3382781;  RODAPE_Y  = 6298579
RODAPE_CX = 8461946;  RODAPE_CY = 400110

CAPA_TEXTO_X  = 2548009;  CAPA_TEXTO_Y  = 2013228
CAPA_TEXTO_CX = 7465441;  CAPA_TEXTO_CY = 2985433

CITACOES = [
    '"A força está em se erguer com mais poder, como a águia." ​',
    '"A águia foca no futuro, não no que está atrás." ​',
    '"Para voar alto, deixe as correntes para trás, como a águia." ​',
    '"Nas tempestades, como a águia, encontre o voo mais alto." ​',
]

# ─────────────────────────────────────────────
# Assets
# ─────────────────────────────────────────────

ASSETS_DIR = os.path.join(os.path.dirname(__file__), "assets")

def _asset(filename):
    return os.path.join(ASSETS_DIR, filename)

def _img(key):
    paths = {
        "capa_background": _asset("capa_background.png"),
        "logo_carranza":   _asset("logo_carranza.png"),
        "encerramento":    _asset("encerramento.png"),
    }
    with open(paths[key], "rb") as f:
        return io.BytesIO(f.read())

# ─────────────────────────────────────────────
# Helpers pptx
# ─────────────────────────────────────────────

def _blank_layout(prs):
    for layout in prs.slide_layouts:
        if layout.name.lower() in ("em branco", "blank", ""):
            return layout
    return prs.slide_layouts[6]

def _add_picture(slide, img_key, left, top, width, height):
    buf = _img(img_key)
    buf.seek(0)
    return slide.shapes.add_picture(buf, Emu(left), Emu(top), Emu(width), Emu(height))

def _add_run(para, text, bold=False, sz_emu=635000, color=None, font_name="Calibri"):
    run = para.add_run()
    run.text = text
    run.font.bold = bold
    run.font.size = Emu(sz_emu)
    run.font.name = font_name
    if color:
        run.font.color.rgb = color
    return run

# ─────────────────────────────────────────────
# Slides
# ─────────────────────────────────────────────

def _slide_capa(prs, dados):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_picture(slide, "capa_background", 0, 0, SLIDE_W, SLIDE_H)
    _add_picture(slide, "logo_carranza", LOGO_CAPA_X, LOGO_CAPA_Y, LOGO_CAPA_CX, LOGO_CAPA_CY)
    tb = slide.shapes.add_textbox(Emu(CAPA_TEXTO_X), Emu(CAPA_TEXTO_Y), Emu(CAPA_TEXTO_CX), Emu(CAPA_TEXTO_CY))
    tf = tb.text_frame
    tf.word_wrap = True
    campos = [
        (dados.get("disciplina", "").upper(), True,  635000),
        (dados.get("assunto",    "").upper(), True,  635000),
        (dados.get("tipo", "QUESTÕES").upper(), True, 508000),
        (dados.get("professor",  ""),          False, 355600),
    ]
    first = True
    for texto, bold, sz in campos:
        if not texto:
            continue
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = PP_ALIGN.CENTER
        _add_run(para, texto, bold=bold, sz_emu=sz, color=VINHO)
    return slide

def _slide_conteudo(prs, paragrafos, citacao=None):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_picture(slide, "logo_carranza", LOGO_MASTER_X, LOGO_MASTER_Y, LOGO_MASTER_CX, LOGO_MASTER_CY)
    tb = slide.shapes.add_textbox(Emu(TEXTO_X), Emu(TEXTO_Y), Emu(TEXTO_CX), Emu(TEXTO_H))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for p in paragrafos:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        para.alignment = PP_ALIGN.JUSTIFY
        _add_run(para, p.get("text", ""), bold=p.get("bold", False),
                 sz_emu=p.get("sz_emu", 571500), color=p.get("color", PRETO))
    if citacao:
        tb2 = slide.shapes.add_textbox(Emu(RODAPE_X), Emu(RODAPE_Y), Emu(RODAPE_CX), Emu(RODAPE_CY))
        tf2 = tb2.text_frame
        para2 = tf2.paragraphs[0]
        para2.alignment = PP_ALIGN.RIGHT
        _add_run(para2, citacao, bold=True, sz_emu=254000, color=VINHO)
    return slide

def _slide_gabarito(prs, questoes, respostas):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_picture(slide, "logo_carranza", LOGO_MASTER_X, LOGO_MASTER_Y, LOGO_MASTER_CX, LOGO_MASTER_CY)
    tb = slide.shapes.add_textbox(Emu(4000000), Emu(2400000), Emu(4200000), Emu(600000))
    tf = tb.text_frame
    para = tf.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    _add_run(para, "GABARITO", bold=True, sz_emu=304800, color=VINHO)
    n = len(questoes)
    if n == 0:
        return slide
    table = slide.shapes.add_table(2, n, Emu(2031997), Emu(3058160), Emu(8128005), Emu(741680)).table
    for j, (q, r) in enumerate(zip(questoes, respostas)):
        c0 = table.cell(0, j)
        c0.text = str(q).zfill(2)
        p0 = c0.text_frame.paragraphs[0]
        p0.alignment = PP_ALIGN.CENTER
        for run in p0.runs:
            run.font.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = VINHO
        c1 = table.cell(1, j)
        c1.text = str(r).upper()
        p1 = c1.text_frame.paragraphs[0]
        p1.alignment = PP_ALIGN.CENTER
        for run in p1.runs:
            run.font.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = PRETO
    return slide

def _slide_encerramento(prs):
    slide = prs.slides.add_slide(_blank_layout(prs))
    _add_picture(slide, "capa_background", 0, 0, SLIDE_W, SLIDE_H)
    _add_picture(slide, "logo_carranza", LOGO_ENC_X, LOGO_ENC_Y, LOGO_ENC_CX, LOGO_ENC_CY)
    return slide


# ─────────────────────────────────────────────
# Parser DOCX (python-docx, usa bold para identificar enunciados)
# ─────────────────────────────────────────────

def _parse_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    slides = []
    gabarito_qs = []
    gabarito_rs = []
    RE_QUESTAO = re.compile(r'^(\d{1,2})\.\s+(.+)', re.DOTALL)
    RE_ALT = re.compile(r'^[A-Ea-e]\)\s+')
    RE_CERTO_ERRADO = re.compile(r'^(certo|errado)[\s\.\(]', re.IGNORECASE)
    RE_GAB_CELL = re.compile(r'^(\d{1,2})\.\s*([A-Ea-eCcEe])')
    # Gabarito em tabela
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                txt = cell.text.strip()
                m = RE_GAB_CELL.match(txt)
                if m:
                    gabarito_qs.append(int(m.group(1)))
                    gabarito_rs.append(m.group(2).upper())
    # Parsear parágrafos
    i = 0
    paras = doc.paragraphs
    while i < len(paras):
        para = paras[i]
        txt = para.text.strip()
        if not txt:
            i += 1
            continue
        is_bold = any(run.bold for run in para.runs if run.text.strip())
        m_q = RE_QUESTAO.match(txt)
        if m_q and is_bold:
            numero = int(m_q.group(1))
            enunciado = txt
            i += 1
            alternativas = []
            extras = []
            while i < len(paras):
                p2 = paras[i]
                t2 = p2.text.strip()
                if not t2:
                    i += 1
                    continue
                is_bold2 = any(run.bold for run in p2.runs if run.text.strip())
                m_q2 = RE_QUESTAO.match(t2)
                if m_q2 and is_bold2:
                    break
                elif RE_ALT.match(t2):
                    alternativas.append(t2)
                    i += 1
                elif RE_CERTO_ERRADO.match(t2):
                    alternativas.append(t2)
                    i += 1
                else:
                    extras.append(t2)
                    i += 1
            enunciado_completo = enunciado + ("\n" + "\n".join(extras) if extras else "")
            certo_errado = len(alternativas) > 0 and all(RE_CERTO_ERRADO.match(a) for a in alternativas)
            slides.append({
                "tipo": "questao",
                "numero": numero,
                "enunciado": enunciado_completo,
                "certo_errado": certo_errado,
                "alternativas": alternativas if not certo_errado else [],
            })
        else:
            i += 1
    gabarito = {"questoes": gabarito_qs, "respostas": gabarito_rs} if gabarito_qs else None
    return slides, gabarito


# ─────────────────────────────────────────────
# Parser TXT
# ─────────────────────────────────────────────

def _parse_texto(texto):
    linhas = [l.rstrip() for l in texto.splitlines()]
    slides = []
    gabarito_qs = []
    gabarito_rs = []
    i = 0
    RE_QUESTAO  = re.compile(r'^(\d{1,2})[.\-]\s+(.+)')
    RE_ALT      = re.compile(r'^[A-Ea-e][).]')
    RE_GABARITO = re.compile(r'^GABARITO', re.IGNORECASE)
    RE_GAB_ITEM = re.compile(r'(\d{1,2})\s*[-–]\s*([A-Ea-eCcEe])\b')
    while i < len(linhas):
        linha = linhas[i].strip()
        if RE_GABARITO.match(linha):
            bloco = linha
            i += 1
            while i < len(linhas):
                bloco += " " + linhas[i].strip()
                i += 1
            for m in RE_GAB_ITEM.finditer(bloco):
                gabarito_qs.append(int(m.group(1)))
                gabarito_rs.append(m.group(2).upper())
            continue
        m_q = RE_QUESTAO.match(linha)
        if m_q:
            numero = int(m_q.group(1))
            enunciado_parts = [m_q.group(2).strip()]
            i += 1
            alternativas = []
            certo_errado = False
            while i < len(linhas):
                l = linhas[i].strip()
                if not l:
                    i += 1
                    if i < len(linhas):
                        prox = linhas[i].strip()
                        if RE_QUESTAO.match(prox) or RE_GABARITO.match(prox):
                            break
                    continue
                if RE_ALT.match(l):
                    alternativas.append(l)
                    i += 1
                elif l.lower().startswith("certo") or l.lower().startswith("errado"):
                    certo_errado = True
                    i += 1
                elif RE_QUESTAO.match(l) or RE_GABARITO.match(l):
                    break
                else:
                    enunciado_parts.append(l)
                    i += 1
            slides.append({"tipo": "questao", "numero": numero, "enunciado": " ".join(enunciado_parts), "certo_errado": certo_errado, "alternativas": alternativas})
            continue
        if linha and not RE_GABARITO.match(linha):
            ctx_parts = [linha]
            i += 1
            while i < len(linhas):
                l = linhas[i].strip()
                if RE_QUESTAO.match(l) or RE_GABARITO.match(l):
                    break
                if not l:
                    i += 1
                    if i < len(linhas) and not linhas[i].strip():
                        break
                    continue
                ctx_parts.append(l)
                i += 1
            slides.append({"tipo": "contexto", "texto": " ".join(ctx_parts)})
            continue
        i += 1
    gabarito = {"questoes": gabarito_qs, "respostas": gabarito_rs} if gabarito_qs else None
    return slides, gabarito


# ─────────────────────────────────────────────
# Builder principal
# ─────────────────────────────────────────────

def _build_pptx(payload):
    prs = Presentation()
    prs.slide_width  = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    dados = {
        "disciplina": payload.get("disciplina", "DISCIPLINA"),
        "assunto":    payload.get("assunto",    "ASSUNTO"),
        "tipo":       payload.get("tipo",       "QUESTÕES"),
        "professor":  payload.get("professor",  ""),
    }
    _slide_capa(prs, dados)
    citacoes_iter = cycle(CITACOES)
    for s in payload.get("slides", []):
        tipo = s.get("tipo", "questao")
        if tipo == "contexto":
            paras = [{"text": s.get("texto", ""), "bold": True, "sz_emu": 508000}]
            _slide_conteudo(prs, paras)
        elif tipo == "questao":
            numero       = s.get("numero", "")
            enunciado    = s.get("enunciado", "")
            certo_errado = s.get("certo_errado", False)
            alternativas = s.get("alternativas", [])
            prefixo = f"{str(numero).zfill(2)}. " if numero else ""
            # Remover prefixo duplicado se o enunciado já começa com o número
            if enunciado.startswith(f"{str(numero).zfill(2)}.") or enunciado.startswith(f"{numero}."):
                texto_enunciado = enunciado
            else:
                texto_enunciado = prefixo + enunciado
            # Fonte adaptativa
            total_chars = len(texto_enunciado) + sum(len(a) for a in alternativas)
            if total_chars > 900:
                sz = 444500
            elif total_chars > 600:
                sz = 482600
            else:
                sz = 533400
            paras = [{"text": texto_enunciado, "bold": True, "sz_emu": sz}]
            if certo_errado:
                paras.append({"text": "Certo (  )",  "bold": False, "sz_emu": sz})
                paras.append({"text": "Errado (  )", "bold": False, "sz_emu": sz})
                _slide_conteudo(prs, paras, citacao=next(citacoes_iter))
            elif alternativas:
                chars_enunciado = len(texto_enunciado)
                if chars_enunciado > 400:
                    max_slide1 = 3
                elif chars_enunciado > 250:
                    max_slide1 = 4
                else:
                    max_slide1 = 5
                alts_1 = alternativas[:max_slide1]
                alts_2 = alternativas[max_slide1:]
                paras_1 = paras + [{"text": a, "bold": False, "sz_emu": sz} for a in alts_1]
                _slide_conteudo(prs, paras_1, citacao=next(citacoes_iter) if not alts_2 else None)
                if alts_2:
                    paras_2 = [{"text": a, "bold": False, "sz_emu": sz} for a in alts_2]
                    _slide_conteudo(prs, paras_2, citacao=next(citacoes_iter))
            else:
                _slide_conteudo(prs, paras, citacao=next(citacoes_iter))
    gab = payload.get("gabarito")
    if gab and gab.get("questoes"):
        _slide_gabarito(prs, gab["questoes"], gab["respostas"])
    _slide_encerramento(prs)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────────────────────────
# Flask App
# ─────────────────────────────────────────────

app = Flask(__name__)
CORS(app, origins="*")

@app.route("/", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "Carranza PPTX Generator"})

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        if request.files or request.form:
            arquivo    = request.files.get("arquivo")
            disciplina = request.form.get("disciplina", "DISCIPLINA")
            assunto    = request.form.get("assunto",    "ASSUNTO")
            professor  = request.form.get("professor",  "")
            tipo       = request.form.get("tipo",       "QUESTÕES")
            if not arquivo:
                return jsonify({"erro": "Arquivo não enviado"}), 400
            filename = arquivo.filename.lower()
            if filename.endswith(".docx"):
                with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                    arquivo.save(tmp.name)
                    tmp_path = tmp.name
                try:
                    slides_data, gabarito = _parse_docx(tmp_path)
                finally:
                    os.unlink(tmp_path)
            else:
                texto = arquivo.read().decode("utf-8", errors="ignore")
                slides_data, gabarito = _parse_texto(texto)
            payload = {"disciplina": disciplina, "assunto": assunto, "tipo": tipo, "professor": professor, "slides": slides_data, "gabarito": gabarito}
        else:
            payload = request.get_json(force=True)
            if not payload:
                return jsonify({"erro": "Payload vazio"}), 400
        buf = _build_pptx(payload)
        disc  = payload.get("disciplina", "apresentacao").replace(" ", "_")
        ass   = payload.get("assunto", "").replace(" ", "_")
        fname = f"Carranza_{disc}_{ass}.pptx" if ass else f"Carranza_{disc}.pptx"
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name=fname)
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500

@app.route("/gerar-texto", methods=["POST"])
def gerar_texto():
    try:
        payload = request.get_json(force=True)
        texto = payload.get("texto", "")
        if not texto:
            return jsonify({"erro": "Campo 'texto' é obrigatório"}), 400
        slides_data, gabarito = _parse_texto(texto)
        estrutura = {"disciplina": payload.get("disciplina", "DISCIPLINA"), "assunto": payload.get("assunto", "ASSUNTO"), "tipo": payload.get("tipo", "QUESTÕES"), "professor": payload.get("professor", ""), "slides": slides_data, "gabarito": gabarito}
        buf = _build_pptx(estrutura)
        disc = estrutura["disciplina"].replace(" ", "_")
        ass  = estrutura["assunto"].replace(" ", "_")
        return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation", as_attachment=True, download_name=f"Carranza_{disc}_{ass}.pptx")
    except Exception as e:
        traceback.print_exc()
        return jsonify({"erro": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
