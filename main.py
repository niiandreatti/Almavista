import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os

cor_fundo = "#f3d8cb"  # Cor de fundo para todos os elementos
cor_texto = "#000000"  # Cor do texto, escolha uma que contraste bem com o fundo
cor = '#3a0913'
branco = '#ffffff'

def abrir_novo_cadastro():
    def salvar_dados():
        nome = entry_nome.get().strip()
        cpf = entry_cpf.get().strip().replace(".", "").replace("-", "")
        email = entry_email.get().strip()
        telefone = entry_telefone.get().strip()

        if nome == "" or cpf == "" or email == "" or telefone == "":
            messagebox.showwarning("Campos vazios", "Preencha todos os campos antes de salvar.")
        else:
            if os.path.exists('clientes.xlsx'):
                df = pd.read_excel('clientes.xlsx')
            else:
                df = pd.DataFrame(columns=['Nome', 'CPF', 'Email', 'Telefone'])

            # Verifica se o CPF já está cadastrado
            if cpf in df['CPF'].values:
                messagebox.showwarning("Duplicado", "CPF já cadastrado.")
            else:
                novo_cliente = pd.DataFrame({'Nome': [nome], 'CPF': [cpf], 'Email': [email], 'Telefone': [telefone]})
                df = pd.concat([df, novo_cliente], ignore_index=True)
                df.to_excel('clientes.xlsx', index=False)

                entry_nome.delete(0, tk.END)
                entry_cpf.delete(0, tk.END)
                entry_email.delete(0, tk.END)
                entry_telefone.delete(0, tk.END)

                messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
                janela_novo_cadastro.destroy()

    def voltar():
        janela_novo_cadastro.destroy()

    janela_novo_cadastro = tk.Toplevel(root)
    janela_novo_cadastro.title("Novo Cadastro")
    centralizar_janela(janela_novo_cadastro)

    janela_novo_cadastro.configure(bg=cor_fundo)

    tk.Label(janela_novo_cadastro, text="Nome", bg=cor_fundo, fg=cor_texto).grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_nome = tk.Entry(janela_novo_cadastro)
    entry_nome.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

    tk.Label(janela_novo_cadastro, text="CPF", bg=cor_fundo, fg=cor_texto).grid(row=1, column=0, padx=10, pady=5, sticky="e")
    entry_cpf = tk.Entry(janela_novo_cadastro)
    entry_cpf.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

    tk.Label(janela_novo_cadastro, text="Email", bg=cor_fundo, fg=cor_texto).grid(row=2, column=0, padx=10, pady=5, sticky="e")
    entry_email = tk.Entry(janela_novo_cadastro)
    entry_email.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

    tk.Label(janela_novo_cadastro, text="Telefone", bg=cor_fundo, fg=cor_texto).grid(row=3, column=0, padx=10, pady=5, sticky="e")
    entry_telefone = tk.Entry(janela_novo_cadastro)
    entry_telefone.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

    tk.Button(janela_novo_cadastro, text="Salvar", command=salvar_dados, bg=cor_fundo, fg=cor_texto).grid(row=4, column=0, columnspan=2, pady=10, sticky="")
    tk.Button(janela_novo_cadastro, text="Voltar", command=voltar, bg=cor_fundo, fg=cor_texto).grid(row=5, column=0, columnspan=2, pady=10, sticky="")

    janela_novo_cadastro.grid_columnconfigure(1, weight=1)


def abrir_consulta_cadastro():
    def voltar():
        janela_consulta_cadastro.destroy()

    def buscar_cadastro():
        cpf_busca = entry_cpf_busca.get().strip().replace(".", "").replace("-", "")
        if os.path.exists('clientes.xlsx'):
            df = pd.read_excel('clientes.xlsx')
            resultado = df[df['CPF'].astype(str) == cpf_busca]

            if not resultado.empty:
                dados_cliente = resultado.iloc[0]

                janela_detalhes = tk.Toplevel(janela_consulta_cadastro)
                janela_detalhes.title("Detalhes do Cadastro")
                centralizar_janela(janela_detalhes)

                janela_detalhes.configure(bg=cor_fundo)

                tk.Label(janela_detalhes, text="Nome:", bg=cor_fundo, fg=cor_texto).grid(row=0, column=0, padx=10, pady=5, sticky="w")
                tk.Label(janela_detalhes, text=dados_cliente['Nome'], bg=cor_fundo, fg=cor_texto).grid(row=0, column=1, padx=10, pady=5, sticky="w")

                tk.Label(janela_detalhes, text="CPF:", bg=cor_fundo, fg=cor_texto).grid(row=1, column=0, padx=10, pady=5, sticky="w")
                tk.Label(janela_detalhes, text=dados_cliente['CPF'], bg=cor_fundo, fg=cor_texto).grid(row=1, column=1, padx=10, pady=5, sticky="w")

                tk.Label(janela_detalhes, text="Email:", bg=cor_fundo, fg=cor_texto).grid(row=2, column=0, padx=10, pady=5, sticky="w")
                tk.Label(janela_detalhes, text=dados_cliente['Email'], bg=cor_fundo, fg=cor_texto).grid(row=2, column=1, padx=10, pady=5, sticky="w")

                tk.Label(janela_detalhes, text="Telefone:", bg=cor_fundo, fg=cor_texto).grid(row=3, column=0, padx=10, pady=5, sticky="w")
                tk.Label(janela_detalhes, text=dados_cliente['Telefone'], bg=cor_fundo, fg=cor_texto).grid(row=3, column=1, padx=10, pady=5, sticky="w")

                def editar_cadastro():
                    janela_detalhes.destroy()
                    editar_cadastro_janela(dados_cliente, atualizar_lista_pacientes)

                tk.Button(janela_detalhes, text="Editar", command=editar_cadastro, bg=cor_fundo, fg=cor_texto).grid(row=4, column=0, columnspan=2, pady=10)
                tk.Button(janela_detalhes, text="Voltar", command=janela_detalhes.destroy, bg=cor_fundo, fg=cor_texto).grid(row=5, column=0, columnspan=2, pady=10)

            else:
                messagebox.showinfo("Resultado da Busca", "Nenhum cadastro encontrado para o CPF informado.")
        else:
            messagebox.showwarning("Arquivo não encontrado", "Nenhum cadastro foi realizado ainda.")

    def atualizar_lista_pacientes():
        for item in tree.get_children():
            tree.delete(item)
            
        if os.path.exists('clientes.xlsx'):
            df = pd.read_excel('clientes.xlsx')
            for i, row in df.iterrows():
                tree.insert("", "end", values=(row['Nome'], row['CPF'], row['Email'], row['Telefone']))

    janela_consulta_cadastro = tk.Toplevel(root)
    janela_consulta_cadastro.title("Consulta de Cadastro")
    centralizar_janela(janela_consulta_cadastro)

    janela_consulta_cadastro.configure(bg=cor_fundo)

    tk.Label(janela_consulta_cadastro, text="CPF", bg=cor_fundo, fg=cor_texto).grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_cpf_busca = tk.Entry(janela_consulta_cadastro)
    entry_cpf_busca.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

    tk.Button(janela_consulta_cadastro, text="Buscar", command=buscar_cadastro, bg=cor_fundo, fg=cor_texto).grid(row=1, column=0, columnspan=2, pady=10, sticky="")

    cols = ('Nome', 'CPF', 'Email', 'Telefone')
    tree = ttk.Treeview(janela_consulta_cadastro, columns=cols, show='headings')
    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    # Customizando o estilo do Treeview
    style = ttk.Style()
    style.configure("Treeview",
                    background=cor,
                    foreground=branco,
                    fieldbackground=cor)
    style.configure("Treeview.Heading", background=branco, foreground=cor_fundo)
    style.map("Treeview",
              background=[('selected', cor)],
              foreground=[('selected', branco)])

    tree.grid(row=2, column=0, columnspan=2, pady=10, sticky='nsew')

    tk.Button(janela_consulta_cadastro, text="Voltar", command=voltar, bg=cor_fundo, fg=cor_texto).grid(row=1, column=0, columnspan=2, pady=10, padx=(170,10))

    atualizar_lista_pacientes()

    janela_consulta_cadastro.grid_columnconfigure(1, weight=1)
    janela_consulta_cadastro.grid_rowconfigure(2, weight=1)


def editar_cadastro_janela(dados_cliente, callback=None):
    janela_editar = tk.Toplevel(root)
    janela_editar.title("Editar Cadastro")
    centralizar_janela(janela_editar)

    janela_editar.configure(bg=cor_fundo)

    tk.Label(janela_editar, text="Nome", bg=cor_fundo, fg=cor_texto).grid(row=0, column=0)
    entry_nome_editar = tk.Entry(janela_editar)
    entry_nome_editar.grid(row=0, column=1)
    entry_nome_editar.insert(0, dados_cliente['Nome'])

    tk.Label(janela_editar, text="CPF", bg=cor_fundo, fg=cor_texto).grid(row=1, column=0)
    entry_cpf_editar = tk.Entry(janela_editar)
    entry_cpf_editar.grid(row=1, column=1)
    entry_cpf_editar.insert(0, dados_cliente['CPF'])
    entry_cpf_editar.config(state=tk.DISABLED)

    tk.Label(janela_editar, text="Email", bg=cor_fundo, fg=cor_texto).grid(row=2, column=0)
    entry_email_editar = tk.Entry(janela_editar)
    entry_email_editar.grid(row=2, column=1)
    entry_email_editar.insert(0, dados_cliente['Email'])

    tk.Label(janela_editar, text="Telefone", bg=cor_fundo, fg=cor_texto).grid(row=3, column=0)
    entry_telefone_editar = tk.Entry(janela_editar)
    entry_telefone_editar.grid(row=3, column=1)
    entry_telefone_editar.insert(0, dados_cliente['Telefone'])

    def salvar_edicao():
        nome_editado = entry_nome_editar.get().strip()
        email_editado = entry_email_editar.get().strip()
        telefone_editado = entry_telefone_editar.get().strip()

        if os.path.exists('clientes.xlsx'):
            df = pd.read_excel('clientes.xlsx')

            df.loc[df['CPF'] == dados_cliente['CPF'], 'Nome'] = nome_editado
            df.loc[df['CPF'] == dados_cliente['CPF'], 'Email'] = email_editado
            df.loc[df['CPF'] == dados_cliente['CPF'], 'Telefone'] = telefone_editado

            df.to_excel('clientes.xlsx', index=False)

            messagebox.showinfo("Sucesso", "Cadastro atualizado com sucesso!")
            janela_editar.destroy()

            if callback:
                callback()

        else:
            messagebox.showwarning("Erro", "Arquivo de dados não encontrado!")
    def remover_cliente():
        if messagebox.askyesno("Confirmar Exclusão", "Tem certeza de que deseja remover este cadastro?"):
            if os.path.exists('clientes.xlsx'):
                df = pd.read_excel('clientes.xlsx')
                df = df[df['CPF'] != dados_cliente['CPF']]
                df.to_excel('clientes.xlsx', index=False)

                messagebox.showinfo("Sucesso", "Cadastro removido com sucesso!")
                janela_editar.destroy()

                if callback:
                    callback()
            else:
                messagebox.showwarning("Erro", "Arquivo de dados não encontrado!")

    tk.Button(janela_editar, text="Salvar", command=salvar_edicao, bg=cor_fundo, fg=cor_texto).grid(row=4, column=0, columnspan=2)
    tk.Button(janela_editar, text="Cancelar", command=janela_editar.destroy, bg=cor_fundo, fg=cor_texto).grid(row=5, column=0, columnspan=2)
    tk.Button(janela_editar, text="Remover", command=remover_cliente, bg="#ff4d4d", fg=cor_texto).grid(row=5, column=0, columnspan=2, pady=10)

    janela_editar.grid_columnconfigure(1, weight=1)


def centralizar_janela(janela, centralizar_vertical=True):
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    largura_janela = 600
    altura_janela = 500
    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2 if centralizar_vertical else 50
    janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")


root = tk.Tk()
root.title("Sistema de Cadastro")
centralizar_janela(root, centralizar_vertical=True)

root.configure(bg=cor_fundo)

tk.Button(root, text="Novo Cadastro", command=abrir_novo_cadastro, bg=cor_fundo, fg=cor_texto).pack(pady=10)
tk.Button(root, text="Consulta de Cadastro", command=abrir_consulta_cadastro, bg=cor_fundo, fg=cor_texto).pack(pady=10)

root.mainloop()
