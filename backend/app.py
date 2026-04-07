import os, re, tempfile, traceback
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

app = Flask(__name__)
CORS(app)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
UPLOAD_FOLDER = tempfile.mkdtemp()
ASSETS = os.path.join(os.path.dirname(os.path.dirname(__file__)))
ALLOWED_EXTENSIONS = {'docx', 'txt'}

VINHO = RGBColor(0x70, 0x00, 0x1C)
PRETO = RGBColor(0x00, 0x00, 0x00)
SLIDE_W, SLIDE_H = 12192000, 6858000
LOGO_X, LOGO_Y = 9282544, 385975
LOGO_CX, LOGO_CY = 2507673, 776504
TEXTO_X, TEXTO_Y = 347272, 1074509
TEXTO_CX = 11497455
RODAPE_X, RODAPE_Y = 3382781, 6298579
RODAPE_CX, RODAPE_CY = 8461946, 400110
CITACOES = [
    '"A forca esta em se erguer com mais poder, como a aguia."',
    '"A aguia foca no futuro, nao no que esta atras."',
    '"Para voar alto, deixe as correntes para tras, como a aguia."',
    '"Nas tempestades, como a aguia, encontre o voo mais alto."',
]

def nova_prs():
    prs = Presentation()
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)
    return prs

def layout_branco(prs):
    for layout in prs.slide_layouts:
        if layout.name.lower() in ('blank', 'em branco', ''):
            return layout
    return prs.slide_layouts[6]

def add_logo(slide):
    logo = os.path.join(ASSETS, 'logo_carranza.png')
    if os.path.exists(logo):
        slide.shapes.add_picture(logo, Emu(LOGO_X), Emu(LOGO_Y), Emu(LOGO_CX), Emu(LOGO_CY))

def add_citacao(slide, citacao):
    txBox = slide.shapes.add_textbox(Emu(RODAPE_X), Emu(RODAPE_Y), Emu(RODAPE_CX), Emu(RODAPE_CY))
    tf = txBox.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = citacao
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.color.rgb = VINHO
    r.font.name = 'Calibri'

def add_textbox_conteudo(slide, blocos, sz_pt=40):
    txBox = slide.shapes.add_textbox(Emu(TEXTO_X), Emu(TEXTO_Y), Emu(TEXTO_CX), Emu(100000))
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, bloco in enumerate(blocos):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = PP_ALIGN.JUSTIFY
        r = para.add_run()
        r.text = bloco['text']
        r.font.bold = bloco.get('bold', False)
        r.font.size = Pt(bloco.get('sz_pt', sz_pt))
        r.font.color.rgb = PRETO
        r.font.name = 'Calibri'

def slide_capa(prs, disciplina, assunto, tipo, professor):
    slide = prs.slides.add_slide(layout_branco(prs))
    bg = os.path.join(ASSETS, 'capa_background.png')
    if os.path.exists(bg):
        slide.shapes.add_picture(bg, Emu(0), Emu(0), Emu(SLIDE_W), Emu(SLIDE_H))
    logo = os.path.join(ASSETS, 'logo_carranza.png')
    if os.path.exists(logo):
        slide.shapes.add_picture(logo, Emu(4429838), Emu(558156), Emu(3251122), Emu(764208))
    linhas = [disciplina.upper(), assunto.upper(), tipo.upper(), '', professor]
    tx = slide.shapes.add_textbox(Emu(347272), Emu(2000000), Emu(11497455), Emu(2500000))
    tf = tx.text_frame
    tf.word_wrap = True
    for i, linha in enumerate(linhas):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = linha
        r.font.bold = (linha != '' and linha != professor)
        r.font.size = Pt(32) if linha != professor else Pt(22)
        r.font.color.rgb = PRETO
        r.font.name = 'Calibri'
    return slide

def slide_conteudo(prs, blocos, citacao=None, sz=40):
    slide = prs.slides.add_slide(layout_branco(prs))
    add_logo(slide)
    add_textbox_conteudo(slide, blocos, sz_pt=sz)
    if citacao:
        add_citacao(slide, citacao)
    return slide

def slide_gabarito(prs, respostas):
    slide = prs.slides.add_slide(layout_branco(prs))
    add_logo(slide)
    tx = slide.shapes.add_textbox(Emu(TEXTO_X), Emu(TEXTO_Y), Emu(TEXTO_CX), Emu(600000))
    tf = tx.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = 'GABARITO'
    r.font.bold = True
    r.font.size = Pt(32)
    r.font.color.rgb = VINHO
    r.font.name = 'Calibri'
    nums = sorted(respostas.keys())
    cols = len(nums)
    if cols > 0:
        table = slide.shapes.add_table(2, cols, Emu(TEXTO_X), Emu(TEXTO_Y + 700000), Emu(TEXTO_CX), Emu(900000)).table
        for i, num in enumerate(nums):
            cell = table.cell(0, i)
            cell.text = str(num).zfill(2)
            cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
            cell2 = table.cell(1, i)
            cell2.text = str(respostas[num])
            cell2.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
    return slide

def slide_encerramento(prs):
    slide = prs.slides.add_slide(layout_branco(prs))
    enc = os.path.join(ASSETS, 'encerramento.png')
    if os.path.exists(enc):
        slide.shapes.add_picture(enc, Emu(0), Emu(0), Emu(SLIDE_W), Emu(SLIDE_H))
    return slide

def extrair_paragrafos_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    return [p.text for p in doc.paragraphs]

def extrair_paragrafos_txt(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read().split('\n')

def parse_documento(paragrafos):
    """
    Parser robusto para documentos de questoes.
    Identifica: titulo, questoes com alternativas, gabarito.
    """
    blocos = []
    gabarito = {}
    i = 0
    n = len(paragrafos)

    while i < n:
        linha = paragrafos[i].strip()

        if not linha:
            i += 1
            continue

        # Detecta gabarito
        if re.match(r'^gabarito', linha, re.IGNORECASE):
            i += 1
            # Tenta pegar respostas nas linhas seguintes
            while i < n:
                l = paragrafos[i].strip()
                matches = re.findall(r'(\d{1,3})\s*[-.:)]\s*([A-Ea-e])', l)
                for num_str, resp in matches:
                    gabarito[int(num_str)] = resp.upper()
                i += 1
            break

        # Detecta questao numerada: "01.", "1.", "01)", "Questao 01"
        match_q = re.match(r'^(\d{1,3})\s*[.)\-]\s*(.+)', linha)
        if not match_q:
            match_q = re.match(r'^[Qq]uest[aãoõ]+\s*(\d{1,3})\s*[.)\-:]\s*(.*)', linha)

        if match_q:
            numero = int(match_q.group(1))
            enunciado = match_q.group(2).strip()
            i += 1

            # Continua lendo o enunciado ate achar alternativa
            alternativas = []
            while i < n:
                l = paragrafos[i].strip()
                if not l:
                    i += 1
                    continue

                # Detecta alternativa: "A)", "A.", "(A)", "a)"
                match_alt = re.match(r'^[(\s]*([A-Ea-e])\s*[.)]\s*(.+)', l)
                if match_alt:
                    letra = match_alt.group(1).upper()
                    texto_alt = match_alt.group(2).strip()
                    alternativas.append(letra + ') ' + texto_alt)
                    i += 1
                    continue

                # Se comeca com numero, e proxima questao
                if re.match(r'^\d{1,3}\s*[.)]\s*', l):
                    break
                if re.match(r'^[Qq]uest[aãoõ]+\s*\d', l):
                    break

                # Detecta gabarito
                if re.match(r'^gabarito', l, re.IGNORECASE):
                    break

                # Se e texto sem ser alternativa, faz parte do enunciado
                if not match_alt:
                    enunciado += ' ' + l
                i += 1

            blocos.append({
                'tipo': 'multipla_escolha' if alternativas else 'certo_errado',
                'numero': numero,
                'enunciado': enunciado,
                'alternativas': alternativas,
            })
        else:
            # Texto avulso (titulo, contexto, etc)
            if len(linha) > 10:
                blocos.append({
                    'tipo': 'contexto',
                    'texto': linha,
                })
            i += 1

    return blocos, gabarito

def gerar_pptx(blocos, gabarito, disciplina, assunto, professor):
    prs = nova_prs()

    # 1. Capa
    slide_capa(prs, disciplina, assunto, 'Questoes', professor)

    # 2. Slides de questoes
    citacao_idx = 0
    for bloco in blocos:
        if bloco['tipo'] == 'contexto':
            slide_conteudo(prs, [{'text': bloco['texto'], 'bold': False, 'sz_pt': 32}])
        elif bloco['tipo'] == 'certo_errado':
            items = [
                {'text': str(bloco['numero']).zfill(2) + '. ' + bloco['enunciado'], 'bold': True, 'sz_pt': 36},
                {'text': '', 'bold': False, 'sz_pt': 20},
                {'text': 'Certo (  )', 'bold': False, 'sz_pt': 36},
                {'text': 'Errado (  )', 'bold': False, 'sz_pt': 36},
            ]
            slide_conteudo(prs, items, citacao=CITACOES[citacao_idx % len(CITACOES)])
            citacao_idx += 1
        elif bloco['tipo'] == 'multipla_escolha':
            alts = bloco.get('alternativas', [])
            # Calcular tamanho da fonte baseado no conteudo
            total_chars = len(bloco['enunciado']) + sum(len(a) for a in alts)
            if total_chars > 800:
                sz_enum = 28
                sz_alt = 28
            elif total_chars > 500:
                sz_enum = 32
                sz_alt = 32
            else:
                sz_enum = 36
                sz_alt = 36

            # Se o conteudo e muito grande, dividir em 2 slides
            if total_chars > 900:
                # Slide 1: enunciado
                items1 = [
                    {'text': str(bloco['numero']).zfill(2) + '. ' + bloco['enunciado'], 'bold': True, 'sz_pt': sz_enum},
                ]
                slide_conteudo(prs, items1)

                # Slide 2: alternativas
                items2 = []
                for alt in alts:
                    items2.append({'text': alt, 'bold': False, 'sz_pt': sz_alt})
                slide_conteudo(prs, items2, citacao=CITACOES[citacao_idx % len(CITACOES)])
            else:
                # Tudo em 1 slide
                items = [
                    {'text': str(bloco['numero']).zfill(2) + '. ' + bloco['enunciado'], 'bold': True, 'sz_pt': sz_enum},
                    {'text': '', 'bold': False, 'sz_pt': 14},
                ]
                for alt in alts:
                    items.append({'text': alt, 'bold': False, 'sz_pt': sz_alt})
                slide_conteudo(prs, items, citacao=CITACOES[citacao_idx % len(CITACOES)])

            citacao_idx += 1

    # 3. Gabarito
    if gabarito:
        slide_gabarito(prs, gabarito)

    # 4. Encerramento
    slide_encerramento(prs)

    return prs

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/api/gerar', methods=['POST'])
def gerar_material():
    try:
        if 'arquivo' not in request.files:
            return jsonify({'error': 'Nenhum arquivo enviado'}), 400
        arquivo = request.files['arquivo']
        if not arquivo.filename:
            return jsonify({'error': 'Arquivo vazio'}), 400
        ext = arquivo.filename.rsplit('.', 1)[-1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return jsonify({'error': 'Formato nao suportado. Use .docx ou .txt'}), 400

        disciplina = request.form.get('disciplina', 'DISCIPLINA')
        assunto = request.form.get('assunto', 'ASSUNTO')
        professor = request.form.get('professor', 'Professor')

        filename = secure_filename(arquivo.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        arquivo.save(filepath)

        if ext == 'docx':
            paragrafos = extrair_paragrafos_docx(filepath)
        else:
            paragrafos = extrair_paragrafos_txt(filepath)

        blocos, gabarito = parse_documento(paragrafos)

        if not blocos:
            return jsonify({'error': 'Nao foi possivel identificar questoes.'}), 400

        prs = gerar_pptx(blocos, gabarito, disciplina, assunto, professor)

        nome_saida = 'Carranza_' + assunto.replace(' ', '_') + '.pptx'
        output_path = os.path.join(UPLOAD_FOLDER, nome_saida)
        prs.save(output_path)

        try:
            os.remove(filepath)
        except:
            pass

        return send_file(
            output_path,
            as_attachment=True,
            download_name=nome_saida,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': 'Erro: ' + str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
