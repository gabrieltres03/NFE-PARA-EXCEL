import os
import uuid
from pathlib import Path
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from importar_nfe import extrair_texto_pdf, parsear_cabecalho, parsear_itens, gerar_excel

app = Flask(__name__)

# Libera chamadas vindas do GitHub Pages
CORS(app)

UPLOAD_FOLDER = Path('/tmp/uploads')
OUTPUT_FOLDER = Path('/tmp/outputs')
UPLOAD_FOLDER.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)

app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024


@app.route('/processar', methods=['POST'])
def processar():
    if 'pdf' not in request.files:
        return jsonify({'sucesso': False, 'erro': 'Nenhum arquivo enviado.'})

    arquivo = request.files['pdf']
    if not arquivo.filename.lower().endswith('.pdf'):
        return jsonify({'sucesso': False, 'erro': 'Apenas PDFs são aceitos.'})

    nome_base = str(uuid.uuid4())
    pdf_path  = UPLOAD_FOLDER / f'{nome_base}.pdf'
    xlsx_path = OUTPUT_FOLDER / f'{nome_base}.xlsx'

    arquivo.save(pdf_path)

    try:
        texto     = extrair_texto_pdf(str(pdf_path))
        cabecalho = parsear_cabecalho(texto)
        itens     = parsear_itens(texto)

        if not itens:
            return jsonify({'sucesso': False, 'erro': 'Nenhum item encontrado. Verifique se é um DANFE válido.'})

        total_qtde, total_vliq = gerar_excel(cabecalho, itens, str(xlsx_path))

        return jsonify({
            'sucesso':  True,
            'arquivo':  nome_base + '.xlsx',
            'mensagem': f'{len(itens)} itens extraídos com sucesso!',
            'nfe':      cabecalho.get('numero_nfe', ''),
            'emissao':  cabecalho.get('data_emissao', ''),
            'valor':    f"R$ {total_vliq:,.2f}".replace(',','X').replace('.', ',').replace('X', '.'),
        })

    except Exception as e:
        return jsonify({'sucesso': False, 'erro': str(e)})

    finally:
        pdf_path.unlink(missing_ok=True)


@app.route('/download/<nome_arquivo>')
def download(nome_arquivo):
    if not nome_arquivo.endswith('.xlsx') or '/' in nome_arquivo or '\\' in nome_arquivo:
        return 'Arquivo inválido.', 400

    caminho = OUTPUT_FOLDER / nome_arquivo
    if not caminho.exists():
        return 'Arquivo não encontrado ou expirado.', 404

    return send_file(
        caminho,
        as_attachment=True,
        download_name=f'NFe_{nome_arquivo}',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == '__main__':

    app.run(debug=False)

