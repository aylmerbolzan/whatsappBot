import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
import sqlite3
import pywhatkit as kit
import time
import pyautogui  # Importação adicionada

# Função para criar o banco de dados
def criar_banco_de_dados():
    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    # Tabela de destinatários
    c.execute('''CREATE TABLE IF NOT EXISTS destinatarios
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  nome TEXT NOT NULL,
                  sobrenome TEXT NOT NULL,
                  telefone TEXT NOT NULL,
                  observacao TEXT)''')
    # Tabela de mensagens
    c.execute('''CREATE TABLE IF NOT EXISTS mensagens
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  titulo TEXT NOT NULL,
                  corpo TEXT NOT NULL,
                  anexo TEXT)''')  # Novo campo para o caminho do anexo
    conn.commit()
    conn.close()

criar_banco_de_dados()

# Função para abrir a tela de cadastro/edição de destinatários
def abrir_tela_cadastro_destinatario(editar=False, destinatario_id=None):
    tela_cadastro = ctk.CTkToplevel(app)
    tela_cadastro.title("Cadastrar Destinatário" if not editar else "Editar Destinatário")
    tela_cadastro.geometry("400x450")

    # Campos de entrada
    label_nome = ctk.CTkLabel(tela_cadastro, text="Nome:")
    label_nome.pack(pady=5)
    entry_nome = ctk.CTkEntry(tela_cadastro, width=300)
    entry_nome.pack(pady=5)

    label_sobrenome = ctk.CTkLabel(tela_cadastro, text="Sobrenome:")
    label_sobrenome.pack(pady=5)
    entry_sobrenome = ctk.CTkEntry(tela_cadastro, width=300)
    entry_sobrenome.pack(pady=5)

    label_telefone = ctk.CTkLabel(tela_cadastro, text="Telefone:")
    label_telefone.pack(pady=5)
    entry_telefone = ctk.CTkEntry(tela_cadastro, width=300)
    entry_telefone.pack(pady=5)

    label_observacao = ctk.CTkLabel(tela_cadastro, text="Observação:")
    label_observacao.pack(pady=5)
    text_observacao = ctk.CTkTextbox(tela_cadastro, width=300, height=100)
    text_observacao.pack(pady=5)

    # Preencher campos se estiver editando
    if editar and destinatario_id:
        conn = sqlite3.connect('destinatarios.db')
        c = conn.cursor()
        c.execute("SELECT nome, sobrenome, telefone, observacao FROM destinatarios WHERE id = ?", (destinatario_id,))
        destinatario = c.fetchone()
        conn.close()
        if destinatario:
            entry_nome.insert(0, destinatario[0])
            entry_sobrenome.insert(0, destinatario[1])
            entry_telefone.insert(0, destinatario[2])
            text_observacao.insert("1.0", destinatario[3])

    # Função para salvar ou editar destinatário
    def salvar_destinatario():
        nome = entry_nome.get()
        sobrenome = entry_sobrenome.get()
        telefone = entry_telefone.get()
        observacao = text_observacao.get("1.0", "end-1c")

        if nome and sobrenome and telefone:
            conn = sqlite3.connect('destinatarios.db')
            c = conn.cursor()
            if editar and destinatario_id:
                c.execute("UPDATE destinatarios SET nome = ?, sobrenome = ?, telefone = ?, observacao = ? WHERE id = ?",
                          (nome, sobrenome, telefone, observacao, destinatario_id))
            else:
                c.execute("INSERT INTO destinatarios (nome, sobrenome, telefone, observacao) VALUES (?, ?, ?, ?)",
                          (nome, sobrenome, telefone, observacao))
            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "Destinatário salvo com sucesso!")
            tela_cadastro.destroy()
            carregar_destinatarios()
        else:
            messagebox.showwarning("Erro", "Nome, Sobrenome e Telefone são obrigatórios!")

    # Botão de salvar
    button_salvar = ctk.CTkButton(tela_cadastro, text="Salvar", command=salvar_destinatario)
    button_salvar.pack(pady=10)

# Função para abrir a tela de cadastro/edição de mensagens
def abrir_tela_cadastro_mensagem(editar=False, mensagem_id=None):
    tela_cadastro = ctk.CTkToplevel(app)
    tela_cadastro.title("Cadastrar Mensagem" if not editar else "Editar Mensagem")
    tela_cadastro.geometry("400x400")

    # Campos de entrada
    label_titulo = ctk.CTkLabel(tela_cadastro, text="Título da Mensagem:")
    label_titulo.pack(pady=5)
    entry_titulo = ctk.CTkEntry(tela_cadastro, width=300)
    entry_titulo.pack(pady=5)

    label_corpo = ctk.CTkLabel(tela_cadastro, text="Corpo da Mensagem:")
    label_corpo.pack(pady=5)
    text_corpo = ctk.CTkTextbox(tela_cadastro, width=300, height=100)
    text_corpo.pack(pady=5)

    # Campo para anexo
    label_anexo = ctk.CTkLabel(tela_cadastro, text="Anexo (apenas imagens):")
    label_anexo.pack(pady=5)
    entry_anexo = ctk.CTkEntry(tela_cadastro, width=300)
    entry_anexo.pack(pady=5)
    button_procurar_anexo = ctk.CTkButton(tela_cadastro, text="Procurar Anexo", command=lambda: selecionar_anexo(entry_anexo))
    button_procurar_anexo.pack(pady=5)

    # Preencher campos se estiver editando
    if editar and mensagem_id:
        conn = sqlite3.connect('destinatarios.db')
        c = conn.cursor()
        c.execute("SELECT titulo, corpo, anexo FROM mensagens WHERE id = ?", (mensagem_id,))
        mensagem = c.fetchone()
        conn.close()
        if mensagem:
            entry_titulo.insert(0, mensagem[0])
            text_corpo.insert("1.0", mensagem[1])
            entry_anexo.insert(0, mensagem[2])

    # Função para salvar ou editar mensagem
    def salvar_mensagem():
        titulo = entry_titulo.get()
        corpo = text_corpo.get("1.0", "end-1c")
        anexo = entry_anexo.get()

        if titulo and corpo:
            conn = sqlite3.connect('destinatarios.db')
            c = conn.cursor()
            if editar and mensagem_id:
                c.execute("UPDATE mensagens SET titulo = ?, corpo = ?, anexo = ? WHERE id = ?",
                          (titulo, corpo, anexo, mensagem_id))
            else:
                c.execute("INSERT INTO mensagens (titulo, corpo, anexo) VALUES (?, ?, ?)",
                          (titulo, corpo, anexo))
            conn.commit()
            conn.close()
            messagebox.showinfo("Sucesso", "Mensagem salva com sucesso!")
            tela_cadastro.destroy()
            carregar_mensagens()
        else:
            messagebox.showwarning("Erro", "Título e Corpo da Mensagem são obrigatórios!")

    # Botão de salvar
    button_salvar = ctk.CTkButton(tela_cadastro, text="Salvar", command=salvar_mensagem)
    button_salvar.pack(pady=10)

# Função para selecionar anexo (apenas imagens)
def selecionar_anexo(entry_anexo):
    arquivo = filedialog.askopenfilename(
        title="Selecione uma imagem",
        filetypes=[("Imagens", "*.jpg *.jpeg *.png")]
    )
    if arquivo:
        entry_anexo.delete(0, "end")
        entry_anexo.insert(0, arquivo)

# Função para carregar destinatários no grid
def carregar_destinatarios():
    for row in tree_destinatarios.get_children():
        tree_destinatarios.delete(row)

    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    c.execute("SELECT id, nome, sobrenome, telefone, observacao FROM destinatarios")
    destinatarios = c.fetchall()
    conn.close()

    for destinatario in destinatarios:
        tree_destinatarios.insert("", "end", values=destinatario)

# Função para carregar mensagens no grid
def carregar_mensagens():
    for row in tree_mensagens.get_children():
        tree_mensagens.delete(row)

    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    c.execute("SELECT id, titulo, corpo, anexo FROM mensagens")
    mensagens = c.fetchall()
    conn.close()

    for mensagem in mensagens:
        tree_mensagens.insert("", "end", values=mensagem)

# Função para editar destinatário
def editar_destinatario():
    selecionado = tree_destinatarios.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhum destinatário selecionado!")
        return
    destinatario_id = tree_destinatarios.item(selecionado[0], "values")[0]
    abrir_tela_cadastro_destinatario(editar=True, destinatario_id=destinatario_id)

# Função para excluir destinatário
def excluir_destinatario():
    selecionado = tree_destinatarios.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhum destinatário selecionado!")
        return
    destinatario_id = tree_destinatarios.item(selecionado[0], "values")[0]

    confirmacao = messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir este destinatário?")
    if confirmacao:
        conn = sqlite3.connect('destinatarios.db')
        c = conn.cursor()
        c.execute("DELETE FROM destinatarios WHERE id = ?", (destinatario_id,))
        conn.commit()
        conn.close()
        carregar_destinatarios()

# Função para editar mensagem
def editar_mensagem():
    selecionado = tree_mensagens.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhuma mensagem selecionada!")
        return
    mensagem_id = tree_mensagens.item(selecionado[0], "values")[0]
    abrir_tela_cadastro_mensagem(editar=True, mensagem_id=mensagem_id)

# Função para excluir mensagem
def excluir_mensagem():
    selecionado = tree_mensagens.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhuma mensagem selecionada!")
        return
    mensagem_id = tree_mensagens.item(selecionado[0], "values")[0]

    confirmacao = messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir esta mensagem?")
    if confirmacao:
        conn = sqlite3.connect('destinatarios.db')
        c = conn.cursor()
        c.execute("DELETE FROM mensagens WHERE id = ?", (mensagem_id,))
        conn.commit()
        conn.close()
        carregar_mensagens()

# Função para enviar mensagens usando pywhatkit (automático)
def enviar_mensagens():
    selecionado = tree_mensagens.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhuma mensagem selecionada!")
        return
    mensagem_id = tree_mensagens.item(selecionado[0], "values")[0]

    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    c.execute("SELECT corpo, anexo FROM mensagens WHERE id = ?", (mensagem_id,))
    mensagem = c.fetchone()
    conn.close()

    if not mensagem:
        messagebox.showwarning("Erro", "Mensagem não encontrada!")
        return

    mensagem_principal = mensagem[0]
    anexo = mensagem[1]

    if not mensagem_principal:
        messagebox.showwarning("Erro", "A mensagem principal é obrigatória!")
        return

    # Obter destinatários selecionados
    selecionados = []
    for item in tree_destinatarios.selection():
        destinatario = tree_destinatarios.item(item, "values")
        selecionados.append(destinatario)

    if not selecionados:
        messagebox.showwarning("Erro", "Nenhum destinatário selecionado!")
        return

    for destinatario in selecionados:
        id_destinatario, nome, sobrenome, telefone, observacao = destinatario
        mensagem_final = mensagem_principal.replace("{nome}", nome)  # Usa apenas o nome
        if observacao:
            mensagem_final += f"\n\n{observacao}"  # Adiciona espaço entre a mensagem principal e a observação

        # Envia a mensagem usando pywhatkit
        try:
            if anexo and anexo.lower().endswith(('.jpg', '.jpeg', '.png')):  # Verifica se o anexo é uma imagem
                kit.sendwhats_image(
                    receiver=telefone,
                    img_path=anexo,
                    caption=mensagem_final,
                    wait_time=25
                )
            else:  # Se não houver anexo ou não for uma imagem
                kit.sendwhatmsg_instantly(telefone, mensagem_final, wait_time=25)
            print(f"Mensagem enviada para {nome} ({telefone})")  # Log no console
            time.sleep(3)  # Aguarda 3 segundos antes de enviar a próxima mensagem
            pyautogui.hotkey('ctrl', 'w')  # Fecha a aba do navegador
        except Exception as e:
            print(f"Erro ao enviar mensagem para {nome}: {e}")  # Log de erro no console

# Configuração da interface gráfica
app = ctk.CTk()
app.title("WhatsApp Sender")

# Frame para listagem de destinatários
frame_destinatarios = ctk.CTkFrame(app)
frame_destinatarios.pack(pady=10, padx=10, fill="both", expand=True)

label_destinatarios = ctk.CTkLabel(frame_destinatarios, text="Destinatários:")
label_destinatarios.pack(pady=5)

tree_destinatarios = ttk.Treeview(frame_destinatarios, columns=("ID", "Nome", "Sobrenome", "Telefone", "Observação"), show="headings")
tree_destinatarios.heading("Nome", text="Nome")
tree_destinatarios.heading("Sobrenome", text="Sobrenome")
tree_destinatarios.heading("Telefone", text="Telefone")
tree_destinatarios.heading("Observação", text="Observação")
tree_destinatarios.column("ID", width=0, stretch=False)  # Oculta a coluna de ID
tree_destinatarios.column("Nome", width=150)
tree_destinatarios.column("Sobrenome", width=150)
tree_destinatarios.column("Telefone", width=100)
tree_destinatarios.column("Observação", width=200)
tree_destinatarios.pack(fill="both", expand=True)

# Frame para botões de destinatários (centralizado)
frame_botoes_destinatarios = ctk.CTkFrame(frame_destinatarios)
frame_botoes_destinatarios.pack(pady=10, fill="x")

# Usar um frame interno para centralizar os botões
frame_botoes_destinatarios_interno = ctk.CTkFrame(frame_botoes_destinatarios)
frame_botoes_destinatarios_interno.pack(expand=True)

button_cadastrar_destinatario = ctk.CTkButton(frame_botoes_destinatarios_interno, text="Cadastrar Destinatário", command=abrir_tela_cadastro_destinatario)
button_cadastrar_destinatario.pack(side="left", padx=5)

button_editar_destinatario = ctk.CTkButton(frame_botoes_destinatarios_interno, text="Editar Destinatário", command=editar_destinatario)
button_editar_destinatario.pack(side="left", padx=5)

button_excluir_destinatario = ctk.CTkButton(frame_botoes_destinatarios_interno, text="Excluir Destinatário", command=excluir_destinatario)
button_excluir_destinatario.pack(side="left", padx=5)

# Frame para listagem de mensagens
frame_mensagens = ctk.CTkFrame(app)
frame_mensagens.pack(pady=10, padx=10, fill="both", expand=True)

label_mensagens = ctk.CTkLabel(frame_mensagens, text="Mensagens:")
label_mensagens.pack(pady=5)

tree_mensagens = ttk.Treeview(frame_mensagens, columns=("ID", "Título", "Corpo", "Anexo"), show="headings")
tree_mensagens.heading("Título", text="Título")
tree_mensagens.heading("Corpo", text="Corpo")
tree_mensagens.heading("Anexo", text="Anexo")
tree_mensagens.column("ID", width=0, stretch=False)  # Oculta a coluna de ID
tree_mensagens.column("Título", width=150)
tree_mensagens.column("Corpo", width=300)
tree_mensagens.column("Anexo", width=200)
tree_mensagens.pack(fill="both", expand=True)

# Frame para botões de mensagens (centralizado)
frame_botoes_mensagens = ctk.CTkFrame(frame_mensagens)
frame_botoes_mensagens.pack(pady=10, fill="x")

# Usar um frame interno para centralizar os botões
frame_botoes_mensagens_interno = ctk.CTkFrame(frame_botoes_mensagens)
frame_botoes_mensagens_interno.pack(expand=True)

button_cadastrar_mensagem = ctk.CTkButton(frame_botoes_mensagens_interno, text="Cadastrar Mensagem", command=abrir_tela_cadastro_mensagem)
button_cadastrar_mensagem.pack(side="left", padx=5)

button_editar_mensagem = ctk.CTkButton(frame_botoes_mensagens_interno, text="Editar Mensagem", command=editar_mensagem)
button_editar_mensagem.pack(side="left", padx=5)

button_excluir_mensagem = ctk.CTkButton(frame_botoes_mensagens_interno, text="Excluir Mensagem", command=excluir_mensagem)
button_excluir_mensagem.pack(side="left", padx=5)

# Frame para enviar mensagens
frame_enviar = ctk.CTkFrame(app)
frame_enviar.pack(pady=10, padx=10, fill="x")

button_enviar = ctk.CTkButton(frame_enviar, text="Enviar Mensagens", command=enviar_mensagens)
button_enviar.pack(pady=10)

# Carregar destinatários e mensagens ao iniciar
carregar_destinatarios()
carregar_mensagens()

app.mainloop()