import customtkinter as ctk
from tkinter import messagebox, ttk
import sqlite3
import pyautogui
import pywhatkit as kit
import time

# Função para criar o banco de dados
def criar_banco_de_dados():
    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS destinatarios
                 (id INTEGER PRIMARY KEY AUTOINCREMENT,
                  nome TEXT NOT NULL,
                  sobrenome TEXT NOT NULL,
                  telefone TEXT NOT NULL,
                  observacao TEXT)''')
    conn.commit()
    conn.close()

criar_banco_de_dados()

# Função para abrir a tela de cadastro/edição
def abrir_tela_cadastro(editar=False, destinatario_id=None):
    tela_cadastro = ctk.CTkToplevel(app)
    tela_cadastro.title("Cadastrar Destinatário" if not editar else "Editar Destinatário")
    tela_cadastro.geometry("400x450")  # Aumentei a altura para 450

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

# Função para carregar destinatários no grid
def carregar_destinatarios():
    for row in tree.get_children():
        tree.delete(row)

    conn = sqlite3.connect('destinatarios.db')
    c = conn.cursor()
    c.execute("SELECT id, nome, sobrenome, telefone, observacao FROM destinatarios")
    destinatarios = c.fetchall()
    conn.close()

    for destinatario in destinatarios:
        tree.insert("", "end", values=destinatario)

# Função para editar destinatário
def editar_destinatario():
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhum destinatário selecionado!")
        return
    destinatario_id = tree.item(selecionado[0], "values")[0]
    abrir_tela_cadastro(editar=True, destinatario_id=destinatario_id)

# Função para excluir destinatário
def excluir_destinatario():
    selecionado = tree.selection()
    if not selecionado:
        messagebox.showwarning("Erro", "Nenhum destinatário selecionado!")
        return
    destinatario_id = tree.item(selecionado[0], "values")[0]

    confirmacao = messagebox.askyesno("Confirmar", "Tem certeza que deseja excluir este destinatário?")
    if confirmacao:
        conn = sqlite3.connect('destinatarios.db')
        c = conn.cursor()
        c.execute("DELETE FROM destinatarios WHERE id = ?", (destinatario_id,))
        conn.commit()
        conn.close()
        carregar_destinatarios()

# Função para enviar mensagens usando pywhatkit (automático)
def enviar_mensagens():
    mensagem_principal = text_mensagem.get("1.0", "end-1c")
    if not mensagem_principal:
        messagebox.showwarning("Erro", "A mensagem principal é obrigatória!")
        return

    # Obter destinatários selecionados
    selecionados = []
    for item in tree.selection():
        destinatario = tree.item(item, "values")
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
            kit.sendwhatmsg_instantly(telefone, mensagem_final, wait_time=25)
            print(f"Mensagem enviada para {nome} ({telefone})")  # Log no console
            time.sleep(3)  # Aguarda 3 segundos antes de enviar a próxima mensagem
            pyautogui.hotkey('ctrl', 'w')
        except Exception as e:
            print(f"Erro ao enviar mensagem para {nome}: {e}")  # Log de erro no console

# Configuração da interface gráfica
app = ctk.CTk()
app.title("WhatsApp Sender")

# Frame para listagem de destinatários
frame_listagem = ctk.CTkFrame(app)
frame_listagem.pack(pady=10, padx=10, fill="both", expand=True)

tree = ttk.Treeview(frame_listagem, columns=("ID", "Nome", "Sobrenome", "Telefone", "Observação"), show="headings")
tree.heading("Nome", text="Nome")
tree.heading("Sobrenome", text="Sobrenome")
tree.heading("Telefone", text="Telefone")
tree.heading("Observação", text="Observação")
tree.column("ID", width=0, stretch=False)  # Oculta a coluna de ID
tree.column("Nome", width=150)
tree.column("Sobrenome", width=150)
tree.column("Telefone", width=100)
tree.column("Observação", width=200)
tree.pack(fill="both", expand=True)

# Frame para botões (centralizado)
frame_botoes = ctk.CTkFrame(frame_listagem)
frame_botoes.pack(pady=10, fill="x")

# Usar um frame interno para centralizar os botões
frame_botoes_interno = ctk.CTkFrame(frame_botoes)
frame_botoes_interno.pack(expand=True)

button_cadastrar = ctk.CTkButton(frame_botoes_interno, text="Cadastrar", command=abrir_tela_cadastro)
button_cadastrar.pack(side="left", padx=5)

button_editar = ctk.CTkButton(frame_botoes_interno, text="Editar", command=editar_destinatario)
button_editar.pack(side="left", padx=5)

button_excluir = ctk.CTkButton(frame_botoes_interno, text="Excluir", command=excluir_destinatario)
button_excluir.pack(side="left", padx=5)

# Frame para mensagem principal
frame_mensagem = ctk.CTkFrame(app)
frame_mensagem.pack(pady=10, padx=10, fill="x")

label_mensagem = ctk.CTkLabel(frame_mensagem, text="Mensagem Principal:")
label_mensagem.pack(pady=5)
text_mensagem = ctk.CTkTextbox(frame_mensagem, height=100)
text_mensagem.pack(pady=5, padx=5, fill="x")

# Botão de Enviar Mensagens
button_enviar = ctk.CTkButton(app, text="Enviar Mensagens", command=enviar_mensagens)
button_enviar.pack(pady=10)

# Carregar destinatários ao iniciar
carregar_destinatarios()

app.mainloop()