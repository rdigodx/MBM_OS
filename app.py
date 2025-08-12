import os
import webbrowser
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_from_directory
import mysql.connector
import uuid
import datetime
import bcrypt
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

mysql.connector.locales.LANG = 'en_US'

app = Flask(__name__)
app.secret_key = 'chave_secreta_qualquer'

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # Limite 16MB

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def get_connection():
    return mysql.connector.connect(
        host='localhost',
        user='root',
        password='Mbm@20251234',
        database='sistema_os'
    )

def enviar_email(destinatarios, assunto, mensagem, caminhos_anexos=None):
    remetente = 'ti2@supricopygyn.com.br'
    senha = 'Mbm@2025'

    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = assunto
    msg.attach(MIMEText(mensagem, 'html'))

    # Se houver anexos, adiciona
    if caminhos_anexos:
        for caminho in caminhos_anexos:
            nome_arquivo = os.path.basename(caminho)
            with open(caminho, 'rb') as file:
                parte = MIMEBase('application', 'octet-stream')
                parte.set_payload(file.read())
                encoders.encode_base64(parte)
                parte.add_header('Content-Disposition', f'attachment; filename={nome_arquivo}')
                msg.attach(parte)

    try:
        servidor = smtplib.SMTP('smtp.supricopygyn.com.br',587)
        servidor.ehlo()
        servidor.starttls()
        servidor.ehlo()
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatarios, msg.as_string())
        servidor.quit()
        print("Email enviado com sucesso!")
    except Exception as e:
        print("Erro ao enviar email:", e)


@app.route('/')
def index():
    token = session.pop('token_os', None)  
    ordem = None

    if token:
        conn = get_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute("""
            SELECT os.id, u.nome AS solicitante, os.setor, os.tipo_servico,
                   os.descricao, os.status, os.data_criacao, os.token
            FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.token = %s
        """, (token,))
        ordem = cursor.fetchone()
        cursor.close()
        conn.close()

    return render_template('index.html', ordem=ordem)

@app.route('/nova-os', methods=['GET', 'POST'])
def nova_os():
    if request.method == 'POST':
        nome = request.form['nome']
        setor = request.form['setor']
        tipo = request.form['tipo_servico']
        descricao = request.form['descricao']
        token = str(uuid.uuid4())[:8].upper()
        data_criacao = datetime.datetime.now()

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)

        # Verifica usuário
        cursor.execute("SELECT id FROM usuarios WHERE nome = %s", (nome,))
        user = cursor.fetchone()
        if not user:
            cursor.execute("INSERT INTO usuarios (nome) VALUES (%s)", (nome,))
            conn.commit()
            solicitante_id = cursor.lastrowid
        else:
            solicitante_id = user['id']

        # Cria OS
        cursor.execute("""
    INSERT INTO ordens_servico (solicitante_id, setor, tipo_servico, descricao, status, token, data_criacao)
    VALUES (%s, %s, %s, %s, 'Nova', %s, %s)
""", (solicitante_id, setor, tipo, descricao, token, data_criacao))

        conn.commit()

        os_id = cursor.lastrowid

        caminhos_anexos = []

        # ✅ Upload de arquivo (INDENTADO CORRETO)
        if 'arquivo' in request.files:
            arquivo = request.files['arquivo']
            if arquivo and arquivo.filename != '':
                nome_arquivo = f"{token}_{arquivo.filename}"
                caminho_completo = os.path.join(app.config['UPLOAD_FOLDER'], nome_arquivo)
                arquivo.save(caminho_completo)

                # Salva anexo
                cursor.execute(
                    "INSERT INTO anexos_os (os_id, nome_arquivo, caminho_arquivo) VALUES (%s, %s, %s)",
                    (os_id, nome_arquivo, caminho_completo)
                )
                conn.commit()
                caminhos_anexos.append(caminho_completo)

        cursor.close()
        conn.close()

        mensagem = f"""
            <p><strong>Nova OS Criada</strong></p>
            <p><strong>Token:</strong> {token}</p>
            <p><strong>Solicitante:</strong> {nome}</p>
            <p><strong>Setor:</strong> {setor}</p>
            <p><strong>Tipo:</strong> {tipo}</p>
            <p><strong>Descrição:</strong> {descricao}</p>
            <p><strong>Data:</strong> {data_criacao.strftime('%d/%m/%Y %H:%M')}</p>
        """

        enviar_email(
            ["depto.ti1@mbmcopy.com.br", "depto.ti2@mbmcopy.com.br", "paulo.faraone@mbmcopy.com.br"],
            "Nova OS Criada - Sistema MBM",
            mensagem,
            caminhos_anexos
        )

        return redirect(url_for('sucesso', token=token))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT nome FROM usuarios ORDER BY nome ASC")
    usuarios = cursor.fetchall()
    cursor.close()
    conn.close()
    return render_template('nova_os.html', usuarios=usuarios)

@app.route('/sucesso')
def sucesso():
    token = request.args.get('token', 'N/A')
    return render_template('sucesso.html', token=token)

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        usuario = request.form['usuario']
        senha = request.form['senha']

        conn = get_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM tecnicos WHERE usuario = %s", (usuario,))
        tecnico = cursor.fetchone()

        if tecnico:
            senha_bd = tecnico['senha']

            if senha_bd.startswith('$2b$'):
                # Senha já está criptografada
                if bcrypt.checkpw(senha.encode('utf-8'), senha_bd.encode('utf-8')):
                    session['tecnico'] = tecnico['usuario']
                    cursor.close()
                    conn.close()
                    return redirect(url_for('painel'))
            else:
                # Senha antiga em texto puro: faz login e converte para bcrypt
                if senha == senha_bd:
                    nova_hash = bcrypt.hashpw(senha.encode('utf-8'), bcrypt.gensalt()).decode()
                    cursor.execute("UPDATE tecnicos SET senha = %s WHERE id = %s", (nova_hash, tecnico['id']))
                    conn.commit()
                    session['tecnico'] = tecnico['usuario']
                    cursor.close()
                    conn.close()
                    return redirect(url_for('painel'))

        cursor.close()
        conn.close()
        flash('Usuário ou senha inválidos.')
        return redirect(url_for('login'))

    return render_template('login.html')

@app.route('/painel')
def painel():
    if 'tecnico' not in session:
        return redirect(url_for('login'))

    conn = get_connection()
    cursor = conn.cursor(dictionary=True)

    # ✅ Atualiza status para 'Pendente' após 24h
    limite_data_pendente = datetime.datetime.now() - datetime.timedelta(days=1)
    cursor.execute("""
        UPDATE ordens_servico
        SET status = 'Pendente'
        WHERE status = 'Nova' AND data_criacao < %s
    """, (limite_data_pendente,))
    conn.commit()

    filtro = request.args.get('filtro')

    # Filtros
    if filtro == 'pendentes':
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.status = 'Pendente'
        """)
    elif filtro == 'concluidas':
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.status = 'Concluída'
        """)
    elif filtro == 'novas':
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.status = 'Nova'
        """)
    elif filtro == 'fora_prazo':
        limite_data = datetime.datetime.now() - datetime.timedelta(days=3)
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.status != 'Concluída' AND os.data_criacao < %s
        """, (limite_data,))
    elif filtro:
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
            WHERE os.token = %s
        """, (filtro,))
    else:
        cursor.execute("""
            SELECT os.*, u.nome AS solicitante_nome FROM ordens_servico os
            JOIN usuarios u ON os.solicitante_id = u.id
        """)

    ordens = cursor.fetchall()

    # Anexos
    cursor_anexo = conn.cursor(dictionary=True)
    for os_item in ordens:
        cursor_anexo.execute(
            "SELECT id, nome_arquivo FROM anexos_os WHERE os_id = %s",
            (os_item['id'],)
        )
        anexos = cursor_anexo.fetchall()
        os_item['anexos'] = anexos
    cursor_anexo.close()

    # Estatísticas
    cursor.execute("SELECT COUNT(*) AS qtd FROM ordens_servico WHERE status = 'Nova'")
    qtd_nova = cursor.fetchone()['qtd']

    cursor.execute("SELECT COUNT(*) AS qtd FROM ordens_servico WHERE status = 'Pendente'")
    qtd_pendente = cursor.fetchone()['qtd']

    cursor.execute("SELECT COUNT(*) AS qtd FROM ordens_servico WHERE status = 'Concluída'")
    qtd_concluida = cursor.fetchone()['qtd']

    limite_data = datetime.datetime.now() - datetime.timedelta(days=3)
    cursor.execute("""
        SELECT COUNT(*) AS qtd FROM ordens_servico
        WHERE status != 'Concluída' AND data_criacao < %s
    """, (limite_data,))
    qtd_fora_prazo = cursor.fetchone()['qtd']

    cursor.execute("""
        SELECT u.nome, COUNT(*) AS total
        FROM ordens_servico os
        JOIN usuarios u ON os.solicitante_id = u.id
        GROUP BY u.nome
        ORDER BY total DESC
        LIMIT 1
    """)
    usuario_top = cursor.fetchone()

    cursor.execute("""
        SELECT u.nome, COUNT(*) AS total
        FROM ordens_servico os
        JOIN usuarios u ON os.solicitante_id = u.id
        GROUP BY u.nome
        ORDER BY total DESC
        LIMIT 5
    """)
    usuarios_top5 = cursor.fetchall()

    cursor.close()
    conn.close()

    qtd_total = len(ordens)


    return render_template(
    'painel.html',
    ordens=ordens,
    qtd_nova=qtd_nova,
    qtd_pendente=qtd_pendente,
    qtd_concluida=qtd_concluida,
    qtd_fora_prazo=qtd_fora_prazo,
    qtd_total=qtd_total,  
    usuario_top=usuario_top,
    usuarios_top5=usuarios_top5,
    tecnico=session['tecnico']
)


@app.route('/fechar-os/<int:id>', methods=['POST'])
def fechar_os(id):
    if 'tecnico' not in session:
        return redirect(url_for('login'))

    resolucao = request.form['resolucao']
    tecnico = session['tecnico']
    data_fechamento = datetime.datetime.now()

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE ordens_servico
        SET status = 'Concluída', resolucao = %s, tecnico = %s, data_fechamento = %s
        WHERE id = %s
    """, (resolucao, tecnico, data_fechamento, id))
    conn.commit()
    cursor.close()
    conn.close()

    return redirect(url_for('painel'))

@app.route('/anexo/<int:arquivo_id>')
def baixar_anexo(arquivo_id):
    conn = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT nome_arquivo, caminho_arquivo FROM anexos_os WHERE id = %s", (arquivo_id,))
    anexo = cursor.fetchone()
    cursor.close()
    conn.close()

    if anexo:
        return send_from_directory(
            directory=app.config['UPLOAD_FOLDER'],
            path=os.path.basename(anexo['caminho_arquivo']),
            as_attachment=True,
            download_name=anexo['nome_arquivo']
        )
    else:
        return "Arquivo não encontrado", 404


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    webbrowser.open("http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
