
def slide_contexto(prs, texto):
    slide = prs.slides.add_slide(layout_branco(prs))
    add_logo(slide)
    add_textbox_conteudo(slide, [{'text': texto, 'bold': False, 'sz_pt': 32}], sz_pt=32)
    return slide

def slide_certo_errado(prs, numero, enunciado, citacao_idx=0):
    slide = prs.slides.add_slide(layout_branco(prs))
    add_logo(slide)
    blocos = [
        {'text': f'{numero:02d}. {enunciado}', 'bold': True, 'sz_pt': 40},
        {'text': '', 'bold': False, 'sz_pt': 40},
        {'text': 'Certo (  )', 'bold': False, 'sz_pt': 40},
        {'text': 'Errado (  )', 'bold': False, 'sz_pt': 40},
    ]
    add_textbox_conteudo(slide, blocos)
    add_citacao(slide, CITACOES[citacao_idx % len(CITACOES)])
    return slide

def slide_multipla_escolha(prs, numero, enunciado, alternativas, citacao_idx=0):
    slide = prs.slides.add_slide(layout_branco(prs))
    add_logo(slide)
    blocos = [
        {'text': f'{numero:02d}. {enunciado}', 'bold': True, 'sz_pt': 40},
        {'text': '', 'bold': False, 'sz_pt': 28},
    ]
    for alt in alternativas:
        blocos.append({'text': alt, 'bold': False, 'sz_pt': 40})
    sz = 40 if len(alternativas) <= 3 else 34
    add_textbox_conteudo(slide, blocos, sz_pt=sz)
    add_citacao(slide, CITACOES[citacao_idx % len(CITACOES)])
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
    table = slide.shapes.add_table(2, cols, Emu(TEXTO_X), Emu(TEXTO_Y + 700000), Emu(TEXTO_CX), Emu(900000)).table
    for i, num in enumerate(nums):
        cell = table.cell(0, i)
        cell.text = f'{num:02d}'
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.runs[0].font.bold = True
        p.runs[0].font.size = Pt(20)
        p.runs[0].font.name = 'Calibri'
        cell2 = table.cell(1, i)
        cell2.text = str(respostas[num])
        p2 = cell2.text_frame.paragraphs[0]
        p2.alignment = PP_ALIGN.CENTER
        p2.runs[0].font.bold = True
        p2.runs[0].font.size = Pt(20)
        p2.runs[0].font.name = 'Calibri'
    return slide

def slide_encerramento(prs):
    slide = prs.slides.add_slide(layout_branco(prs))
    enc = os.path.join(ASSETS, 'encerramento.png')
    if os.path.exists(enc):
        slide.shapes.add_picture(enc, Emu(0), Emu(0), Emu(SLIDE_W), Emu(SLIDE_H))
    return slide

def extrair_texto_docx(filepath):
    from docx import Document
    doc = Document(filepath)
    return '\n'.join([p.text for p in doc.paragraphs])

def extrair_texto_txt(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        return f.read()

def parse_questoes(texto):
    blocos = []
    linhas = texto.split('\n')
    i = 0
    while i < len(linhas):
        linha = linhas[i].strip()
        match_q = re.match(r'^(?:Quest[aA]o\s*)?(\d{1,3})[.)\s]\s*(.+)', linha, re.IGNORECASE)
        if match_q:
            numero = int(match_q.group(1))
            enunciado = match_q.group(2).strip()
            i += 1
            alternativas = []
            while i < len(linhas):
                l = linhas[i].strip()
                match_alt = re.match(r'^[(\s]*([A-Ea-e])[).\s]+(.+)', l)
                if match_alt:
                    letra = match_alt.group(1).upper()
                    texto_alt = match_alt.group(2).strip()
                    alternativas.append(f'{letra}) {texto_alt}')
                    i += 1
                    continue
                if re.match(r'^(Certo|Errado)', l, re.IGNORECASE):
                    break
                if re.match(r'^(?:Quest[aA]o\s*)?\d{1,3}[.)\s]', l, re.IGNORECASE):
                    break
                if not l:
                    i += 1
                    continue
                enunciado += ' ' + l
                i += 1
            if alternativas:
                blocos.append({'tipo': 'multipla_escolha', 'numero': numero, 'enunciado': enunciado, 'alternativas': alternativas})
            else:
                blocos.append({'tipo': 'certo_errado', 'numero': numero, 'enunciado': enunciado})
        else:
            if linha and not re.match(r'^(Gabarito|GABARITO)', linha, re.IGNORECASE):
                contexto = linha
                i += 1
                while i < len(linhas):
                    l = linhas[i].strip()
                    if not l:
                        i += 1
                        break
                    if re.match(r'^(?:Quest[aA]o\s*)?\d{1,3}[.)\s]', l, re.IGNORECASE):
                        break
                    contexto += ' ' + l
                    i += 1
                if len(contexto) > 20:
                    blocos.append({'tipo': 'contexto', 'texto': contexto})
            else:
                i += 1
    return blocos

def parse_gabarito(texto):
    respostas = {}
    matches = re.findall(r'(\d{1,3})\s*[-.)]\s*([A-Ea-eCE]|Certo|Errado)', texto, re.IGNORECASE)
    for num, resp in matches:
        n = int(num)
        r = resp.strip().upper()
        if r == 'CERTO': r = 'C'
        elif r == 'ERRADO': r = 'E'
        respostas[n] = r
    return respostas

@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'message': 'Carranza Gerador API rodando!'})

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
            return jsonify({'error': f'Formato .{ext} nao suportado'}), 400
        disciplina = request.form.get('disciplina', 'DISCIPLINA')
        assunto = request.form.get('assunto', 'ASSUNTO')
        professor = request.form.get('professor', 'Professor')
        filename = secure_filename(arquivo.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        arquivo.save(filepath)
        if ext == 'docx':
            texto = extrair_texto_docx(filepath)
        elif ext == 'txt':
            texto = extrair_texto_txt(filepath)
        else:
            return jsonify({'error': 'PDF requer OCR - use .docx ou .txt'}), 400
        blocos = parse_questoes(texto)
        gabarito = parse_gabarito(texto)
        if not blocos:
            return jsonify({'error': 'Nao foi possivel identificar questoes no documento.'}), 400
        prs = nova_prs()
        slide_capa(prs, disciplina, assunto, 'Questoes', professor)
        citacao_idx = 0
        for bloco in blocos:
            if bloco['tipo'] == 'contexto':
                slide_contexto(prs, bloco['texto'])
            elif bloco['tipo'] == 'certo_errado':
                slide_certo_errado(prs, bloco['numero'], bloco['enunciado'], citacao_idx)
                citacao_idx += 1
            elif bloco['tipo'] == 'multipla_escolha':
                slide_multipla_escolha(prs, bloco['numero'], bloco['enunciado'], bloco['alternativas'], citacao_idx)
                citacao_idx += 1
        if gabarito:
            slide_gabarito(prs, gabarito)
        slide_encerramento(prs)
        nome_saida = f"Carranza_{assunto.replace(' ', '_')}.pptx"
        output_path = os.path.join(UPLOAD_FOLDER, nome_saida)
        prs.save(output_path)
        try:
            os.remove(filepath)
        except:
            pass
        return send_file(output_path, as_attachment=True, download_name=nome_saida, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Erro ao gerar material: {str(e)}'}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
