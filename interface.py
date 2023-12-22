import os
import time
import win32com.client as win32
import datetime
import shutil
from tkinter import Tk, Label, Entry, Button, filedialog, messagebox


def listar_nomes_de_arquivos(caminho_da_pasta):
    try:
        nomes_de_arquivos = os.listdir(caminho_da_pasta)
        return nomes_de_arquivos
    except FileNotFoundError:
        return []


def criar_ou_verificar_pasta(pasta):
    if not os.path.exists(pasta):
        os.makedirs(pasta)


def enviar_email_com_anexo(caminho_arquivo, destinatario, remetente):
    try:
        outlook = win32.Dispatch("outlook.application")
        email = outlook.CreateItem(0)
        email.To = destinatario
        hora_atual = datetime.datetime.now().hour
        saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
        nome_arquivo = os.path.basename(caminho_arquivo)
        email.Subject = nome_arquivo
        email.HTMLBody = f"""  
         <p>Olá, {saudacao}.</p> 
         <p>Envio em anexo FICHA DE OPME.</p>

         <p style="margin-bottom: 80px;">&nbsp;</p>
         
         <p>Atenciosamente,</p> 
         <p>{remetente}</p> 
         <p>Solution Consultoria e Soluções em Medicina</p> 
         <p>http://www.solutionltda.com</p> 
         <p>Tel.: (21) 2233-7031 / 3083-7606 </p>  
        """
        email.Attachments.Add(caminho_arquivo)
        email.Send()
        print(f"Arquivo enviado: {nome_arquivo}")
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")


def selecionar_pasta():
    pasta = filedialog.askdirectory()
    entry_folder_path.delete(0, "end")
    entry_folder_path.insert(0, pasta)


def enviar_emails_e_mover_arquivos():
    pasta = entry_folder_path.get()
    destinatario = entry_email_address.get()
    remetente = entry_remetente.get()

    if not pasta or not destinatario or not remetente:
        messagebox.showerror("Erro", "Preencha todos os campos!")
        return

    nomes = listar_nomes_de_arquivos(pasta)

    if not nomes:
        messagebox.showinfo("Aviso", "Nenhum arquivo encontrado na pasta.")
        return

    for nome_a_copiar in nomes:
        caminho_arquivo_a_copiar = os.path.join(pasta, nome_a_copiar)

        if os.path.exists(caminho_arquivo_a_copiar):
            enviar_email_com_anexo(caminho_arquivo_a_copiar, destinatario, remetente)
            time.sleep(3)
        else:
            print(f"Arquivo não encontrado: {nome_a_copiar}")

    pasta_origem = pasta
    pasta_destino = os.path.join(pasta, "ARQUIVOS ENVIADOS")
    criar_ou_verificar_pasta(pasta_destino)

    for nome_arquivo in nomes:
        caminho_arquivo_origem = os.path.join(pasta_origem, nome_arquivo)
        caminho_arquivo_destino = os.path.join(pasta_destino, nome_arquivo)

        if os.path.exists(caminho_arquivo_origem):
            shutil.move(caminho_arquivo_origem, caminho_arquivo_destino)

    messagebox.showinfo(
        "Concluído",
        "Todos os emails foram enviados e os arquivos foram movidos com sucesso.",
    )


# Criar a janela principal
root = Tk()
root.title("Envio de Emails")

# Cores mais atraentes
cor_principal = "#3498db"  # Azul
cor_secundaria = "#2ecc71"  # Verde

# Estilo para botões com bordas arredondadas
estilo_botoes = {
    "borderwidth": 5,
    "highlightthickness": 0,
    "relief": "flat",
    "bg": cor_secundaria,
    "fg": "white",
}

# Componentes da interface
label_folder_path = Label(root, text="Caminho da pasta dos arquivos:", pady=5)
entry_folder_path = Entry(root, width=40)
button_browse = Button(root, text="Procurar", command=selecionar_pasta, **estilo_botoes)

label_email_address = Label(root, text="Email do destinatário:", pady=5)
entry_email_address = Entry(root, width=40)

label_remetente = Label(root, text="Nome e Sobrenome do remetente:", pady=5)
entry_remetente = Entry(root, width=40)

button_send = Button(
    root,
    text="Enviar Emails e Mover Arquivos",
    command=enviar_emails_e_mover_arquivos,
    **estilo_botoes,
)

# Posicionamento dos componentes
label_folder_path.grid(row=0, column=0, padx=10, pady=5, sticky="w")
entry_folder_path.grid(row=0, column=1, padx=5, pady=5)
button_browse.grid(row=0, column=2, pady=5)

label_email_address.grid(row=1, column=0, padx=10, pady=5, sticky="w")
entry_email_address.grid(row=1, column=1, padx=5, pady=5)

label_remetente.grid(row=2, column=0, padx=10, pady=5, sticky="w")
entry_remetente.grid(row=2, column=1, padx=5, pady=5)

button_send.grid(row=3, column=0, columnspan=3, pady=10)

# Iniciar a interface gráfica
root.mainloop()
