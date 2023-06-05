import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = os.path.join(app.root_path, 'static', 'barcode')
app.config['DATA_FOLDER'] = os.path.join(app.root_path, 'static')

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])


@app.route("/")
def homepage():
    return render_template("homepage.html")


@app.route('/formulario', methods=['POST'])
def process_form():
    nome = request.form['nome']
    cpf = request.form['cpf']
    email = request.form['email']

    file_path = os.path.join(app.config['DATA_FOLDER'], 'dados.xlsx')

    # Verificar se a planilha existe
    if not os.path.isfile(file_path):
        # Criar uma nova planilha
        df = pd.DataFrame(columns=['Nome', 'CPF', 'E-mail', 'Código de Barras'])
    else:
        # Carregar a planilha existente
        df = pd.read_excel(file_path)

    # Criar um DataFrame com os dados do formulário atual
    data = {'Nome': [nome], 'CPF': [cpf], 'E-mail': [email]}
    df_new = pd.DataFrame(data)

    # Gerar o código de barras para o formulário atual
    code = Code128(cpf, writer=ImageWriter())
    barcode_image_filename = f'barcode_{cpf}'
    barcode_image_path = os.path.join(app.config['UPLOAD_FOLDER'], barcode_image_filename)
    code.save(barcode_image_path)

    # Concatenar os dados existentes e os novos dados do formulário
    df_new['Código de Barras'] = barcode_image_filename
    df_combined = pd.concat([df, df_new], ignore_index=True)

    try:
        # Salvar o DataFrame na planilha
        df_combined.to_excel(file_path, index=False)

        return redirect(url_for('download_page', cpf=cpf))
    except Exception as e:
        return f'Erro: {str(e)}'


@app.route('/download_page/<cpf>')
def download_page(cpf):
    return render_template('download.html', cpf=cpf, barcode_image_filename=f'barcode_{cpf}.png')


@app.route('/download_barcode/<filename>')
def download_barcode(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
