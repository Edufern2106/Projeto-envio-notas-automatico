import os
import tempfile
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, jsonify
import win32com.client
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx'}
ALLOWED_ATTACHMENTS = {'pdf', 'xml'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Estado simples em memória (para demo)
excel_data = []
enviados = set()


def allowed_file(filename, allowed):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed

@app.route('/', methods=['GET', 'POST'])
def index():
    global excel_data, enviados
    if request.method == 'POST':
        file = request.files.get('file')
        if file and allowed_file(file.filename, ALLOWED_EXTENSIONS):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            df = pd.read_excel(filepath)
            print('Colunas lidas do Excel:', list(df.columns))
            # Verificar duplicidade de notas fiscais
            duplicadas = df['Nº Nota'].duplicated(keep=False)
            df['Duplicada'] = duplicadas
            alerta_duplicidade = None
            if duplicadas.any():
                alerta_duplicidade = 'Atenção: Existem notas fiscais com numeração duplicada! Confira antes de enviar.'
            # Adicionar comentário nas duplicadas
            df['Comentario'] = df.apply(lambda row: 'Nota fiscal em duplicidade' if row['Duplicada'] else '', axis=1)
            # Ordenar por cliente
            df.sort_values(by='Cliente', inplace=True)
            excel_data = df.to_dict(orient='records')
            # Agrupar por cliente
            grouped = {}
            for idx, row in enumerate(excel_data):
                cliente = row['Cliente']
                if cliente not in grouped:
                    grouped[cliente] = []
                grouped[cliente].append((idx, row))
            enviados = set()
            return render_template('index.html', rows=excel_data, enviados=enviados, grouped=grouped, excel_columns=list(df.columns), alerta_duplicidade=alerta_duplicidade)
        else:
            return render_template('index.html', rows=[], enviados=enviados, error='Arquivo inválido!')
    # Agrupar por cliente para GET também
    grouped = {}
    for idx, row in enumerate(excel_data):
        cliente = row.get('Cliente', 'Desconhecido')
        if cliente not in grouped:
            grouped[cliente] = []
        grouped[cliente].append((idx, row))
    return render_template('index.html', rows=excel_data, enviados=enviados, grouped=grouped)

@app.route('/send_email', methods=['POST'])
def send_email():
    global excel_data, enviados
    idx = int(request.form['idx'])
    folder = request.form.get('folder', '')
    subject = request.form.get('subject', '')
    attachment = request.files.get('attachment')
    # Não exige mais o anexo obrigatório; permite envio apenas com anexos da pasta
    row = excel_data[idx]
    try:
        import pythoncom
        pythoncom.CoInitialize()
        att_path = None
        if 'attachment' in request.files and request.files['attachment'].filename:
            attachment = request.files['attachment']
            att_filename = secure_filename(attachment.filename)
            att_path = os.path.join(app.config['UPLOAD_FOLDER'], att_filename)
            attachment.save(att_path)
        # Buscar arquivos na pasta pelo número da nota
        nota_num = str(row['Nº Nota'])
        arquivos_encontrados = []
        if folder and os.path.isdir(folder):
            for fname in os.listdir(folder):
                if nota_num in fname and (fname.lower().endswith('.pdf') or fname.lower().endswith('.xml')):
                    arquivos_encontrados.append(os.path.join(folder, fname))
        # Envia e-mail via Outlook
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        # Usar o e-mail fornecido no preview, se disponível
        destinatario = request.form.get('email', row.get('E-mail', ''))
        mail.To = destinatario
        # Definir CC
        cc_padrao = f"faturamento.adb@autodefesabrasil; {destinatario}; contasareceber@autodefesabrasil.com.br; medicao@autodefesabrasil.com.br"
        cc_preview = request.form.get('cc', cc_padrao)
        mail.CC = cc_preview
        # Definir assunto: FAT - NOME DO CLIENTE - MÊS/ANO
        nome_cliente = row.get('Cliente', 'CLIENTE')
        mes_ano = 'MÊS/ANO'
        # Tenta obter mês/ano do subject enviado pelo front-end
        if subject:
            import re
            match = re.search(r'- ([^\-]+)$', subject)
            if match:
                mes_ano = match.group(1).strip().upper()
        mail.Subject = f"FAT - {nome_cliente} - {mes_ano}"
        # Corpo padrão solicitado
        mail.Body = f"Prezado(a) {nome_cliente}!\n       \nSegue em anexo as Notas Fiscais referentes a {mes_ano}.\nCaso identifique qualquer divergência ou problema, solicitamos que nos informe dentro dos seguintes prazos:\n- Notas Fiscais de Venda: Em até 24 horas.\n- Notas Fiscais de Serviço: Até o dia 28 do mês corrente.\n\nAgradecemos desde já pela atenção e solicitamos a gentileza de confirmar o recebimento deste e-mail e dos documentos anexos.\n\nEstamos à disposição para qualquer esclarecimento ou dúvida adicional.\n\nAtenciosamente,\nEquipe Faturamento"
        if request.form.get('separado') == '1':
            enviados_count = 0
            # Enviar anexo manual (upload) separadamente, se houver
            if att_path:
                mail_sep = outlook.CreateItem(0)
                mail_sep.To = destinatario
                mail_sep.CC = cc_preview
                mail_sep.Subject = mail.Subject
                mail_sep.Body = mail.Body
                namespace = outlook.GetNamespace("MAPI")
                sent_folder = namespace.GetDefaultFolder(5)
                mail_sep.SaveSentMessageFolder = sent_folder
                mail_sep.Attachments.Add(att_path)
                mail_sep.Send()
                enviados_count += 1
            # Enviar cada arquivo encontrado separadamente
            for arq in arquivos_encontrados:
                mail_sep = outlook.CreateItem(0)
                mail_sep.To = destinatario
                mail_sep.CC = cc_preview
                mail_sep.Subject = mail.Subject
                mail_sep.Body = mail.Body
                namespace = outlook.GetNamespace("MAPI")
                sent_folder = namespace.GetDefaultFolder(5)
                mail_sep.SaveSentMessageFolder = sent_folder
                mail_sep.Attachments.Add(arq)
                mail_sep.Send()
                enviados_count += 1
            enviados.add(idx)
            return jsonify({'success': True, 'msg': f'{enviados_count} e-mail(s) enviados separadamente com 1 anexo cada.'})
        else:
            # NOVO: Anexar todos os arquivos enviados no campo anexos[]
            anexos_form = request.form.getlist('anexos[]')
            if att_path:
                mail.Attachments.Add(att_path)
            # Se vier anexos do form, anexa todos eles (com caminho completo)
            if anexos_form and folder:
                for nome_arq in anexos_form:
                    caminho = os.path.join(folder, nome_arq)
                    if os.path.isfile(caminho):
                        mail.Attachments.Add(caminho)
            else:
                for arq in arquivos_encontrados:
                    mail.Attachments.Add(arq)
            # Garantir que o e-mail seja salvo nos Itens Enviados
            namespace = outlook.GetNamespace("MAPI")
            sent_folder = namespace.GetDefaultFolder(5)  # 5 = olFolderSentMail
            mail.SaveSentMessageFolder = sent_folder
            mail.Send()
            enviados.add(idx)
            return jsonify({'success': True, 'msg': f'E-mail enviado com sucesso! {len(anexos_form)} arquivo(s) anexados.'})
    except Exception as e:
        return jsonify({'success': False, 'msg': f'Erro ao enviar: {e}'})

@app.route('/buscar_anexos', methods=['POST'])
def buscar_anexos():
    global excel_data
    data = request.get_json()
    pasta = data.get('pasta', '')
    resultados = []
    print('Caminho recebido para busca:', pasta)
    if not pasta or not os.path.isdir(pasta):
        msg = f"Pasta inválida! Recebido: {pasta} | Existe? {os.path.exists(pasta)} | isdir? {os.path.isdir(pasta)}"
        print(msg)
        return jsonify({'success': False, 'msg': msg}), 400
    import re
    for row in excel_data:
        nota_num = str(row.get('Nº Nota', '')).strip()
        encontrados = []
        if nota_num:
            # Regex: número isolado (não dentro de outros números), mas pode estar junto de letras, símbolos ou espaços
            padrao = re.compile(r'(?<!\d)' + re.escape(nota_num) + r'(?!\d)', re.IGNORECASE)
            for fname in os.listdir(pasta):
                if (fname.lower().endswith('.pdf') or fname.lower().endswith('.xml')) and padrao.search(fname):
                    encontrados.append(fname)
        resultados.append(encontrados)
    return jsonify({'success': True, 'resultados': resultados})

if __name__ == '__main__':
    app.run(debug=True)
