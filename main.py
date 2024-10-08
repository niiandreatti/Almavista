import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
import os
from PIL import ImageTk, Image 
import sys
from tkinter import filedialog
from datetime import datetime 

cor_fundo = "#ECBEAF"  
cor_texto = "#000000"  
cor = '#590016'
branco = '#ffffff'
arquivo_excel = 'clientes.xlsx'
entradas_ultimo_od = []
entradas_ultimo_oe = []
entradas_rx_od = []
entradas_rx_oe = []

def abrir_novo_cadastro():
    def formatar_data(event):
        data = entry_data.get().replace("/", "") 
        
        if len(data) > 2 and len(data) <= 4:
            data = data[:2] + '/' + data[2:]
        elif len(data) > 4:
            data = data[:2] + '/' + data[2:4] + '/' + data[4:]

        entry_data.delete(0, tk.END)
        entry_data.insert(0, data)    

        entry_data.icursor(tk.END)

    def criar_tabela_ultimo():
        global entradas_ultimo_od, entradas_ultimo_oe
        
        entradas_ultimo_od.clear()
        entradas_ultimo_oe.clear()

        frame_tabela = tk.Frame(scrollable_frame, bg=cor_fundo)
        frame_tabela.grid(row=16, column=1, columnspan=4, padx=10, pady=5, sticky="w")

        tk.Label(frame_tabela, text="PRESCRIÇÃO DO ÚLTIMO ÓCULOS", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=7, padx=10, pady=5)

        labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
        for col, label in enumerate(labels, start=1):
            tk.Label(frame_tabela, text=label, bg=cor_fundo, fg=cor_texto).grid(row=1, column=col, padx=2)

        tk.Label(frame_tabela, text="O.D.", bg=cor_fundo, fg=cor_texto).grid(row=2, column=0, padx=0)
        tk.Label(frame_tabela, text="O.E.", bg=cor_fundo, fg=cor_texto).grid(row=3, column=0, padx=0)

        for row, prefix in zip([2, 3], ['od', 'oe']):
            for col in range(1, 7):
                entry = tk.Entry(frame_tabela, width=4, bg=branco, fg=cor_texto,justify="center")
                entry.grid(row=row, column=col, padx=0, pady=2)
                if prefix == 'od':
                    entradas_ultimo_od.append(entry)
                else:
                    entradas_ultimo_oe.append(entry)

        tk.Label(scrollable_frame, text="TIPO DE LENTE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=20, column=0, padx=2)
        
        global entry_tipo_lente_ultimo
        entry_tipo_lente_ultimo = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo)
        entry_tipo_lente_ultimo.grid(row=20, column=1, padx=10, pady=2, sticky="w")

    def criar_tabela_RX():
        global entradas_rx_od, entradas_rx_oe
        
        entradas_rx_od.clear()
        entradas_rx_oe.clear()

        frame_tabela = tk.Frame(scrollable_frame, bg=cor_fundo)
        frame_tabela.grid(row=16, column=4, columnspan=4, padx=10, pady=5, sticky="w")

        tk.Label(frame_tabela, text="RX FINAL", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=7, padx=10, pady=5)

        labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
        for col, label in enumerate(labels, start=1):
            tk.Label(frame_tabela, text=label, bg=cor_fundo, fg=cor_texto).grid(row=1, column=col, padx=2)

        tk.Label(frame_tabela, text="O.D.", bg=cor_fundo, fg=cor_texto).grid(row=2, column=0, padx=0)
        tk.Label(frame_tabela, text="O.E.", bg=cor_fundo, fg=cor_texto).grid(row=3, column=0, padx=0)

        for row, prefix in zip([2, 3], ['od', 'oe']):
            for col in range(1, 7):
                entry = tk.Entry(frame_tabela, width=4, bg=branco, fg=cor_texto, justify="center")
                entry.grid(row=row, column=col, padx=0, pady=2)
                if prefix == 'od':
                    entradas_rx_od.append(entry)
                else:
                    entradas_rx_oe.append(entry)

        tk.Label(scrollable_frame, text="TIPO DE LENTE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=20, column=4, padx=2)
        
        global entry_tipo_lente_rx
        entry_tipo_lente_rx = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo)
        entry_tipo_lente_rx.grid(row=20, column=5, padx=10, pady=2, sticky="w")

    def obter_valor(entry, valor_padrao="0.0"):
        valor = entry.get().strip()
        return valor if valor else valor_padrao
    
    def obter_campo(entry):
        valor = entry.get().strip()
        return valor if valor else "Nenhum"

    def salvar_dados():
        try:
            nome = entry_nome.get().strip()
            data = entry_data.get().strip()
            idade = entry_idade.get().strip()
            telefone = entry_telefone.get().strip()

            if not all([nome, data, idade, telefone]):
                messagebox.showwarning("Campos vazios", "Preencha todos os campos antes de salvar.")
                return
            
        
            responsavel = obter_campo(entry_responsavel)
            profissao = entry_profissao.get().strip()
            motivo = entry_motivo.get().strip()
            avaliacao = obter_campo(entry_avaliacao)
            opcao_var = obter_campo(entry_opcao_var)
            obs = obter_campo(entry_obs)
            problema_sistemico = obter_campo(entry_problema_sistemico)
            uso_medicamento = obter_campo(entry_uso_medicamento)
            procedimento_olho = obter_campo(entry_procedimento_olho)
            hist_familia = obter_campo(entry_hist_familia)
            lente_contato = obter_campo(entry_lente_contato)
            dificuldade = obter_campo(entry_dificuldade)
            sintomas = obter_campo(entry_sintomas)
            outros = entry_outros.get("1.0", "end-1c") 
            esf_od_ultimo = obter_valor(entradas_ultimo_od[0])
            cil_od_ultimo = obter_valor(entradas_ultimo_od[1])
            eixo_od_ultimo = obter_valor(entradas_ultimo_od[2])
            adicao_od_ultimo = obter_valor(entradas_ultimo_od[3])
            dnp_od_ultimo = obter_valor(entradas_ultimo_od[4])
            alt_od_ultimo = obter_valor(entradas_ultimo_od[5])

            esf_oe_ultimo = obter_valor(entradas_ultimo_oe[0])
            cil_oe_ultimo = obter_valor(entradas_ultimo_oe[1])
            eixo_oe_ultimo = obter_valor(entradas_ultimo_oe[2])
            adicao_oe_ultimo = obter_valor(entradas_ultimo_oe[3])
            dnp_oe_ultimo = obter_valor(entradas_ultimo_oe[4])
            alt_oe_ultimo = obter_valor(entradas_ultimo_oe[5])

            tipo_lente_ultimo = obter_campo(entry_tipo_lente_ultimo)

            esf_od_rx = obter_valor(entradas_rx_od[0])
            cil_od_rx = obter_valor(entradas_rx_od[1])
            eixo_od_rx = obter_valor(entradas_rx_od[2])
            adicao_od_rx = obter_valor(entradas_rx_od[3])
            dnp_od_rx = obter_valor(entradas_rx_od[4])
            alt_od_rx = obter_valor(entradas_rx_od[5])

            esf_oe_rx = obter_valor(entradas_rx_oe[0])
            cil_oe_rx = obter_valor(entradas_rx_oe[1])
            eixo_oe_rx = obter_valor(entradas_rx_oe[2])
            adicao_oe_rx = obter_valor(entradas_rx_oe[3])
            dnp_oe_rx = obter_valor(entradas_rx_oe[4])
            alt_oe_rx = obter_valor(entradas_rx_oe[5])

            tipo_lente_rx = obter_campo(entry_tipo_lente_rx)


        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar os dados: {e}")


        novo_cliente = pd.DataFrame({
            'Nome': [nome],
            'Data': [data],
            'Idade': [idade],
            'Telefone': [telefone],
            'Responsável': [responsavel],
            'Profissão': [profissao],
            'Motivo': [motivo],
            'Avaliação': [avaliacao],
            'Opção': [opcao_var],
            'Obs': [obs],
            'Problema Sistêmico': [problema_sistemico],
            'Uso Medicamento': [uso_medicamento],
            'Procedimento Olho': [procedimento_olho],
            'Hist. Familiar': [hist_familia],
            'Lente de Contato': [lente_contato],
            'Dificuldade': [dificuldade],
            'Sintomas': [sintomas],
            'Outros': [outros],
            'ESF. O.D. Último': [esf_od_ultimo],
            'CIL. O.D. Último': [cil_od_ultimo],
            'EIXO O.D. Último': [eixo_od_ultimo],
            'ADIÇÃO O.D. Último': [adicao_od_ultimo],
            'DNP O.D. Último': [dnp_od_ultimo],
            'ALT O.D. Último': [alt_od_ultimo],
            'ESF. O.E. Último': [esf_oe_ultimo],
            'CIL. O.E. Último': [cil_oe_ultimo],
            'EIXO O.E. Último': [eixo_oe_ultimo],
            'ADIÇÃO O.E. Último': [adicao_oe_ultimo],
            'DNP O.E. Último': [dnp_oe_ultimo],
            'ALT O.E. Último': [alt_oe_ultimo],
            'Tipo Lente Último': [tipo_lente_ultimo],
            'ESF. O.D. RX': [esf_od_rx],
            'CIL. O.D. RX': [cil_od_rx],
            'EIXO O.D. RX': [eixo_od_rx],
            'ADIÇÃO O.D. RX': [adicao_od_rx],
            'DNP O.D. RX': [dnp_od_rx],
            'ALT O.D. RX': [alt_od_rx],
            'ESF. O.E. RX': [esf_oe_rx],
            'CIL. O.E. RX': [cil_oe_rx],
            'EIXO O.E. RX': [eixo_oe_rx],
            'ADIÇÃO O.E. RX': [adicao_oe_rx],
            'DNP O.E. RX': [dnp_oe_rx],
            'ALT O.E. RX': [alt_oe_rx],
            'Tipo Lente RX': [tipo_lente_rx]
        })

        try: 
            df_existente = pd.read_excel(arquivo_excel)
            df_atualizado = pd.concat([df_existente, novo_cliente], ignore_index=True)
        except FileNotFoundError:
            df_atualizado = novo_cliente

        df_atualizado.to_excel(arquivo_excel, index=False)

        entry_nome.delete(0, tk.END)
        entry_data.delete(0, tk.END)
        entry_idade.delete(0, tk.END)
        entry_telefone.delete(0, tk.END)
        entry_responsavel.delete(0, tk.END)
        entry_profissao.delete(0, tk.END)
        entry_motivo.delete(0, tk.END)
        entry_avaliacao.delete(0, tk.END)
        entry_opcao_var.delete(0, tk.END)
        entry_obs.delete(0, tk.END)
        entry_problema_sistemico.delete(0, tk.END)
        entry_uso_medicamento.delete(0, tk.END)
        entry_procedimento_olho.delete(0, tk.END)
        entry_hist_familia.delete(0, tk.END)
        entry_lente_contato.delete(0, tk.END)
        entry_dificuldade.delete(0, tk.END)
        entry_sintomas.delete(0, tk.END)
        entry_outros.delete("1.0", "end")

        

        for entry in entradas_ultimo_od + entradas_ultimo_oe + entradas_rx_od + entradas_rx_oe:
            entry.delete(0, tk.END)
        
        entry_tipo_lente_ultimo.delete(0, tk.END)
        entry_tipo_lente_rx.delete(0, tk.END)

        messagebox.showinfo("Sucesso", "Dados salvos com sucesso!")
        janela_novo_cadastro.destroy()

    def voltar():
        janela_novo_cadastro.destroy()

    janela_novo_cadastro = tk.Toplevel(root)
    janela_novo_cadastro.title("Novo Cadastro")
    centralizar_janela(janela_novo_cadastro)

    janela_novo_cadastro.configure(bg=cor_fundo)

    canvas = tk.Canvas(janela_novo_cadastro, bg=cor_fundo)
    scrollbar = tk.Scrollbar(janela_novo_cadastro, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=cor_fundo)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0,column=1, sticky="ns")

    janela_novo_cadastro.grid_rowconfigure(0, weight=1)
    janela_novo_cadastro.grid_columnconfigure(0, weight=1)

    logo_path1 = obter_caminho_imagem("logo_horizontal_branco.png")  
    logo_image1 = Image.open(logo_path1)
    logo_image1 = logo_image1.resize((500, 100), Image.Resampling.LANCZOS)  
    logo_photo1 = ImageTk.PhotoImage(logo_image1)

    logo_label_novo_cadastro = tk.Label(scrollable_frame, image=logo_photo1, bg=cor_fundo)
    logo_label_novo_cadastro.image = logo_photo1 
    logo_label_novo_cadastro.grid(row=0, column=0, columnspan=8, padx=5, pady=10, sticky="nsew")

    tk.Label(scrollable_frame, text="NOME:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
    entry_nome = tk.Entry(scrollable_frame, bg=branco, fg = cor_texto,highlightbackground=cor_fundo)
    entry_nome.grid(row=1, column=1, columnspan=5, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="DATA:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    entry_data = tk.Entry(scrollable_frame,bg=branco, fg = cor_texto, highlightbackground=cor_fundo, width= 15, justify='center')
    entry_data.grid(row=2, column=1, padx=10, pady=5, sticky="w", columnspan=1)

    entry_data.bind("<KeyRelease>", formatar_data)

    tk.Label(scrollable_frame, text="IDADE:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=1, padx=10, pady=5, sticky="e")
    entry_idade = tk.Entry(scrollable_frame, width=10,bg=branco, fg = cor_texto,highlightbackground=cor_fundo)  
    entry_idade.grid(row=2, column=2, padx=0, pady=5, sticky="w", columnspan=1)

    tk.Label(scrollable_frame, text="RESPONSÁVEL:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=3, column=2, padx= 10, pady=5, sticky="w")
    entry_responsavel = tk.Entry(scrollable_frame, bg=branco, fg = cor_texto,highlightbackground=cor_fundo)
    entry_responsavel.grid(row=3, column=3, padx=10, pady=5, sticky="we", columnspan=4)  

    tk.Label(scrollable_frame, text="TELEFONE:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="w")
    entry_telefone = tk.Entry(scrollable_frame, width=25,bg=branco, fg = cor_texto,highlightbackground=cor_fundo, justify='center')
    entry_telefone.grid(row=3,columnspan = 3,column=1, padx=10, pady=5, sticky="w")

    tk.Label(scrollable_frame, text="ÚLTIMA AVALIAÇÃO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=4, padx=10, pady=5, sticky="w")
    entry_avaliacao = tk.Entry(scrollable_frame, bg=branco, fg = cor_texto,highlightbackground=cor_fundo, justify='center')
    entry_avaliacao.grid(row=2, column=5, padx=10, pady=5, sticky="w")

    tk.Label(scrollable_frame, text="PROFISSÃO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=5, column=0, padx=10, pady=5, sticky="w")
    entry_profissao = tk.Entry(scrollable_frame,bg=branco, fg = cor_texto,highlightbackground=cor_fundo,justify='center')
    entry_profissao.grid(row=5, column=1, columnspan=5, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="MOTIVO DA CONSULTA:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=6, column=0, padx=10, pady=5, sticky="w")
    entry_motivo = tk.Entry(scrollable_frame,bg=branco, fg = cor_texto,highlightbackground=cor_fundo,justify='center')
    entry_motivo.grid(row=6, columnspan = 5,column=1, padx=10, pady=5, sticky="we")


    opcao_var = tk.StringVar()
    tk.Label(scrollable_frame, text="USUÁRIO DE RX:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=7, column=0, padx=10, pady=5, sticky="w")
    entry_opcao_var = ttk.Combobox(scrollable_frame, textvariable=opcao_var, width=20)  
    entry_opcao_var['values'] = ('LC', 'Óculos','Ambos', 'Nenhum')
    entry_opcao_var.grid(row=7, column=1, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="Opções de Assinalar:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=8, column=0, padx=10, pady=5, sticky="w")

    opcoes = ['BEBE', 'FUMA', 'GESTANTE', 'LACTANTE', 'DORMIU BEM', 'DORMIU MAL' ,'NENHUM']
    selected_values = []  

    def toggle_selection(event):
        value = entry_obs.get()
        if value and value in opcoes:
            if value in selected_values:
                selected_values.remove(value) 
                print(f"Removido: {value}")  
            else:
                selected_values.append(value)
                print(f"Adicionado: {value}")  
            atualiza_opcoes()

    def atualiza_opcoes():
        opcoes_disponiveis = [opcao for opcao in opcoes if opcao not in selected_values]
        entry_obs['values'] = opcoes_disponiveis
        obs.set(', '.join(selected_values))

    def limpar_selecao():
        removed_value = selected_values.pop() 
        print(f"Removido último: {removed_value}")  
        atualiza_opcoes()

    obs = tk.StringVar()
    entry_obs = ttk.Combobox(scrollable_frame, textvariable=obs, values=opcoes, width=15)
    entry_obs.grid(row=8, column=1, columnspan=4, padx=10, pady=5, sticky="we")

    botao_limpar = tk.Button(scrollable_frame, text="Removido último", command=limpar_selecao, highlightbackground=cor_fundo)
    botao_limpar.grid(row=8, column=5, padx=10, pady=5, sticky="w")

    entry_obs.bind("<<ComboboxSelected>>", toggle_selection)
    
    tk.Label(scrollable_frame, text="PROBLEMA SISTÊMICO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=9, column=0, padx=10, pady=5, sticky="w")
    entry_problema_sistemico = tk.Entry(scrollable_frame, bg=branco, fg = cor_texto,highlightbackground=cor_fundo)
    entry_problema_sistemico.grid(row=9, column=1, columnspan=5, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="USO DE MEDICAMENTO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=10, column=0, padx=10, pady=5, sticky="w")
    entry_uso_medicamento = tk.Entry(scrollable_frame, bg=branco, fg = cor_texto,highlightbackground=cor_fundo)
    entry_uso_medicamento.grid(row=10, column=1, columnspan=5, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="PROCEDIMENTO NO OLHO?", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=11, column=0, padx=10, pady=5, sticky="w")
    entry_procedimento_olho = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=15)
    entry_procedimento_olho.grid(row=11, column=1, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="GLAUCOMA NA FAMILIA?", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=12, column=0, padx=10, pady=5, sticky="w")
    entry_hist_familia = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=15)
    entry_hist_familia.grid(row=12, column=1, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="LENTES DE CONTATO APÓS 24 HORAS DE 'NÃO' USO", 
    bg=cor_fundo, fg=cor_texto, font=("Arial", 11, "bold")).grid(row=13, column=0, columnspan=2, padx=10, pady=10, sticky="w")
    entry_lente_contato = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=10)
    entry_lente_contato.grid(row=13, column=1, padx=100, pady=5, sticky="w")

    tk.Label(scrollable_frame, text="DIFICULDADE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=14, column=0, padx=10, pady=5, sticky="w")
    entry_dificuldade = ttk.Combobox(scrollable_frame, values=('PERTO', 'LONGE', 'OS DOIS', 'CLARIDADE', 'NENHUM'), width=15)
    entry_dificuldade.grid(row=14, column=1, padx=10, pady=5, sticky="we")

    tk.Label(scrollable_frame, text="SINTOMAS:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=15, column=0, padx=10, pady=5, sticky="w")

    sintomas_opcoes = ('ARDE', 'COÇA', 'PONTADA', 'DOR DE CABEÇA', 'DÓI NO SOL', 'LAGRIMEJA MUITO','VERMELHIDÃO', 'DOR NOS OLHOS' ,'RESSECADO', 'HIPEREMIA', 'NENHUM')
    selected_sintomas = []  

    
    def toggle_sintomas_selection(event):
        value = entry_sintomas.get()
        if value and value in sintomas_opcoes:
            if value in selected_sintomas:
                selected_sintomas.remove(value)  
                print(f"Sintoma Removido: {value}")  
            else:
                selected_sintomas.append(value) 
                print(f"Sintoma Adicionado: {value}")  
            atualiza_sintomas_opcoes()

    def atualiza_sintomas_opcoes():
        opcoes_disponiveis = [opcao for opcao in sintomas_opcoes if opcao not in selected_sintomas]
        entry_sintomas['values'] = opcoes_disponiveis
        sintomas.set(', '.join(selected_sintomas))

    def limpar_sintomas_selecao():
            removed_value = selected_sintomas.pop()  
            print(f"Sintoma Removido: {removed_value}") 
            atualiza_sintomas_opcoes()

    
    sintomas = tk.StringVar()
    entry_sintomas = ttk.Combobox(scrollable_frame, textvariable=sintomas, values=sintomas_opcoes, width=15)
    entry_sintomas.grid(row=15, column=1, columnspan=4,padx=10, pady=5, sticky="we")

    botao_limpar_sintomas = tk.Button(scrollable_frame, text="Remover último", command=limpar_sintomas_selecao, highlightbackground=cor_fundo)
    botao_limpar_sintomas.grid(row=15, column=5, padx=10, pady=5, sticky="w")

    entry_sintomas.bind("<<ComboboxSelected>>", toggle_sintomas_selection)

    tk.Label(scrollable_frame, text="OUTROS:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=23, column=0, padx=10, pady=5, sticky="w")
    entry_outros = tk.Text(scrollable_frame, bg=branco, fg=cor_texto, height=4,wrap="word", highlightbackground=cor_fundo)
    entry_outros.grid(row=23, column=1, columnspan=5, padx=10, pady=5, sticky="we")


    frame_botoes = tk.Frame(scrollable_frame, bg=cor_fundo)
    frame_botoes.grid(row=24, column=2, columnspan=2, pady=10)

    tk.Button(frame_botoes, text="Salvar", command=salvar_dados, bg=cor_fundo, fg=cor_texto, highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Voltar", command=voltar, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)


    criar_tabela_ultimo()
    criar_tabela_RX()

def abrir_consulta_cadastro():
    def voltar():
        janela_consulta_cadastro.destroy()
        
    def buscar_cadastro(event=None):
        busca = entry_busca.get().strip().lower()

        if os.path.exists('clientes.xlsx'):
            df = pd.read_excel('clientes.xlsx')
            
            df['Nome'] = df['Nome'].astype(str)
            df['Telefone'] = df['Telefone'].astype(str)
            df['Data'] = df['Data'].astype(str)

            if busca == "":
                resultado = df
            else:
                resultado = df[(df['Nome'].str.lower().str.contains(busca)) | 
                            (df['Telefone'].str.contains(busca)) | 
                            (df['Data'].str.contains(busca))]
            
            if resultado.empty:
                messagebox.showinfo("Resultado da Busca", "Nenhum cadastro encontrado para o nome ou telefone informado.")
                entry_busca.delete(0, tk.END)  
                atualizar_lista_tree(df)  
            else:
                atualizar_lista_tree(resultado)  
        else:
            messagebox.showwarning("Arquivo não encontrado", "Nenhum cadastro foi realizado ainda.")

    def abrir_janela_detalhes(dados_cliente):
        janela_detalhes = tk.Toplevel(root)
        janela_detalhes.title("Detalhes do Cadastro")
        centralizar_janela(janela_detalhes)

        
        janela_detalhes.configure(bg=cor_fundo)
        logo_label_novo_cadastro = tk.Label(janela_detalhes, image=logo_photo, bg=cor_fundo)
        logo_label_novo_cadastro.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")

        canvas = tk.Canvas(janela_detalhes, bg=cor_fundo)
        scrollbar = tk.Scrollbar(janela_detalhes, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=cor_fundo)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0,column=1, sticky="ns")

        janela_detalhes.grid_rowconfigure(0, weight=1)
        janela_detalhes.grid_columnconfigure(0, weight=1)

        logo_path2 = obter_caminho_imagem("logo_horizontal_branco.png")  
        logo_image2 = Image.open(logo_path2)
        logo_image2 = logo_image2.resize((500, 100), Image.Resampling.LANCZOS)  
        logo_photo2 = ImageTk.PhotoImage(logo_image2)

        logo_label_novo_cadastro = tk.Label(scrollable_frame, image=logo_photo2, bg=cor_fundo)
        logo_label_novo_cadastro.image = logo_photo2
        logo_label_novo_cadastro.grid(row=0, column=0, columnspan=6, padx=5, pady=10, sticky="nsew")

        tk.Label(scrollable_frame, text="NOME:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Nome'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=1, column=1, padx=10,columnspan=5, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="DATA:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Data'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=2, column=1, columnspan=1,padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="IDADE:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=2, padx=10, pady=5, sticky="e")
        tk.Label(scrollable_frame, text=dados_cliente['Idade'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=2, column=3, columnspan=1, padx=10, pady=5, sticky="w")

        tk.Label(scrollable_frame, text="RESPONSÁVEL:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=3, column=2, padx= 10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Responsável'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=3, column=3,columnspan=4, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="TELEFONE:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Telefone'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=3, column=1, columnspan=1, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="ÚLTIMA AVALIAÇÃO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=2, column=4, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Avaliação'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=2, column=5,columnspan=1, padx=10, pady=5, sticky="we")
    
        tk.Label(scrollable_frame, text="PROFISSÃO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=5, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Profissão'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=5, column=1, columnspan=5, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="MOTIVO DA CONSULTA:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=6, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Motivo'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=6, column=1, columnspan=5, padx=10, pady=5, sticky="we")


        tk.Label(scrollable_frame, text="USUÁRIO DE RX:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=7, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Opção'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=7, column=1, columnspan=5,padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="OPÇÃO DE ASSINALAR:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=8, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Obs'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=8, column=1, columnspan=5, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="PROBLEMA SISTÊMICO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=9, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Problema Sistêmico'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=9, column=1,columnspan=5, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="USO DE MEDICAMENTO:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=10, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Uso Medicamento'],bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=10, column=1,columnspan=5, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="PROCEDIMENTO NO OLHO?", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=11, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Procedimento Olho'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=11, columnspan=5,column=1, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="GLAUCOMA NA FAMILIA?", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=12, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Hist. Familiar'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=12, column=1, columnspan=5,padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="LENTES DE CONTATO APÓS 24 HORAS DE 'NÃO' USO", 
        bg=cor_fundo, fg=cor_texto, font=("Arial", 10, "bold")).grid(row=13, column=0, columnspan=2, padx=10, pady=10, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Lente de Contato'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=13, column=1, padx=100, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="DIFICULDADE:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=14, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Dificuldade'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=14, column=1,columnspan=5, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="SINTOMAS:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold")).grid(row=15, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Sintomas'], bg=branco, fg = cor_texto,highlightbackground=cor_fundo).grid(row=15,columnspan=5, column=1, padx=10, pady=5, sticky="we")

        tk.Label(scrollable_frame, text="OUTROS:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=20, column=0, padx=10, pady=5, sticky="w")
        tk.Label(scrollable_frame, text=dados_cliente['Outros'], bg=branco, fg=cor_texto, highlightbackground=cor_fundo,  height=4,wraplength=700).grid(row=20, column=1, columnspan=5, padx=10, pady=5, sticky="we")

        def editar_cadastro():
            janela_detalhes.destroy()
            editar_cadastro_janela(dados_cliente, atualizar_lista_tree)

        frame_tabela = tk.Frame(scrollable_frame, bg=cor_fundo)
        frame_tabela.grid(row=17, column=0, columnspan=3, padx=10, pady=5, sticky="e")

        titulo = tk.Label(frame_tabela, text="PRESCRIÇÃO DO ÚLTIMO ÓCULOS", font=("Arial", 16, "bold"), bg=cor_fundo, fg=cor_texto)
        titulo.grid(row=0, column=0, columnspan=7)

        labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
        for col, label in enumerate(labels, start=1):
            tk.Label(frame_tabela, text=label, bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=1, column=col, padx=10, pady=5)

        tk.Label(frame_tabela, text="O.D.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=2, column=0, padx=10, pady=5)
        tk.Label(frame_tabela, text="O.E.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=3, column=0, padx=10, pady=5)

        campos_od_oe = [
            ("Esfericidade OD", 'ESF. O.D. Último'),  
            ("Cilindro OD", 'CIL. O.D. Último'),
            ("Eixo OD", 'EIXO O.D. Último'),
            ("Adição OD", 'ADIÇÃO O.D. Último'),
            ("DNP OD", 'DNP O.D. Último'),
            ("ALT OD", 'ALT O.D. Último'),
            ("Esfericidade OE", 'ESF. O.E. Último'),
            ("Cilindro OE", 'CIL. O.E. Último'),
            ("Eixo OE", 'EIXO O.E. Último'),
            ("Adição OE", 'ADIÇÃO O.E. Último'),
            ("DNP OE", 'DNP O.E. Último'),
            ("ALT OE", 'ALT O.E. Último'),
            ("Tipo Lente", 'Tipo Lente Último')
        ]
        
        
            
        for col, (label_text, key) in enumerate(campos_od_oe[:6], start=1):
            label = tk.Label(frame_tabela, text=dados_cliente[key], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid") 
            label.grid(row=2, column=col, padx=3, pady=5, sticky="nsew")  
            frame_tabela.grid_columnconfigure(col, weight=1)  

        for col, (label_text, key) in enumerate(campos_od_oe[6:12], start=1):
            label = tk.Label(frame_tabela, text=dados_cliente[key], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid")  
            label.grid(row=3, column=col, padx=3, pady=5, sticky="nsew")  
            frame_tabela.grid_columnconfigure(col, weight=1)  

        label_tipo_lente = tk.Label(frame_tabela, text="TIPO LENTE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold"))
        label_tipo_lente.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        label_tipo_lente_valor = tk.Label(frame_tabela, text=dados_cliente['Tipo Lente Último'], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid")  # Adiciona borda
        label_tipo_lente_valor.grid(row=4, column=1, columnspan=6, padx=5, pady=5, sticky="we")

        frame_tabela1 = tk.Frame(scrollable_frame, bg=cor_fundo)
        frame_tabela1.grid(row=17, column=3, columnspan=3, padx=10, pady=5, sticky="w")

        titulo = tk.Label(frame_tabela1, text="RX FINAL", font=("Arial", 16, "bold"), bg=cor_fundo, fg=cor_texto)
        titulo.grid(row=0, column=0, columnspan=7)

        labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
        for col, label in enumerate(labels, start=1):
            tk.Label(frame_tabela1, text=label, bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=1, column=col, padx=10, pady=5)

        tk.Label(frame_tabela1, text="O.D.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=2, column=0, padx=10, pady=5)
        tk.Label(frame_tabela1, text="O.E.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=3, column=0, padx=10, pady=5)

        campos_rx = [
            ("Esfericidade OD RX", 'ESF. O.D. RX'),
            ("Cilindro OD RX", 'CIL. O.D. RX'),
            ("Eixo OD RX", 'EIXO O.D. RX'),
            ("Adição OD RX", 'ADIÇÃO O.D. RX'),
            ("DNP OD RX", 'DNP O.D. RX'),
            ("Altura OD RX", 'ALT O.D. RX'),
            ("Esfericidade OE RX", 'ESF. O.E. RX'),
            ("Cilindro OE RX", 'CIL. O.E. RX'),
            ("Eixo OE RX", 'EIXO O.E. RX'),
            ("Adição OE RX", 'ADIÇÃO O.E. RX'),
            ("DNP OE RX", 'DNP O.E. RX'),
            ("Altura OE RX", 'ALT O.E. RX'),
            ("Tipo Lente RX", 'Tipo Lente RX')
        ]

        for col, (label_text, key) in enumerate(campos_rx[:6], start=1):
            label = tk.Label(frame_tabela1, text=dados_cliente[key], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid")  # Adiciona borda
            label.grid(row=2, column=col, padx=3, pady=5, sticky="nsew") 
              

        for col, (label_text, key) in enumerate(campos_rx[6:12], start=1):
            label = tk.Label(frame_tabela1, text=dados_cliente[key], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid")  # Adiciona borda
            label.grid(row=3, column=col, padx=3, pady=5, sticky="nsew")
           

        label_tipo_lente = tk.Label(frame_tabela1, text="TIPO LENTE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold"))
        label_tipo_lente.grid(row=4, column=0, padx=5, pady=5, sticky="w")

        label_tipo_lente_valor = tk.Label(frame_tabela1, text=dados_cliente['Tipo Lente RX'], bg=branco, fg=cor_texto, justify="center", borderwidth=1, relief="solid")  # Adiciona borda
        label_tipo_lente_valor.grid(row=4, column=1, columnspan=6, padx=5, pady=5, sticky="we")

        frame_botoes = tk.Frame(scrollable_frame, bg=cor_fundo)
        frame_botoes.grid(row=21, column=2, columnspan=2, pady=10)

        tk.Button(frame_botoes, text="Editar", command=editar_cadastro, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_botoes, text="Voltar", command=janela_detalhes.destroy, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)

    def selecionar_paciente(event):
        item_selecionado = tree.selection()
        if item_selecionado:
            valores = tree.item(item_selecionado, 'values')
            df = pd.read_excel('clientes.xlsx')
            resultado = df[df['Nome'] == valores[0]]
            if not resultado.empty:
                dados_cliente = resultado.iloc[0]
                abrir_janela_detalhes(dados_cliente)

    def atualizar_lista_tree(df):
        df = df.sort_values(by='Nome')

        for item in tree.get_children():
            tree.delete(item)

        for i, row in df.iterrows():
            tree.insert("", "end", values=(row['Nome'], row['Data'], row['Telefone']))

    janela_consulta_cadastro = tk.Toplevel(root)
    janela_consulta_cadastro.title("Consulta de Cadastro")
    centralizar_janela(janela_consulta_cadastro)

    janela_consulta_cadastro.configure(bg=cor_fundo)

    frame_busca = tk.Frame(janela_consulta_cadastro, bg=cor_fundo)
    frame_busca.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    tk.Label(frame_busca, text="Nome, Telefone ou Data:", bg=cor_fundo, fg=cor_texto,font=("Arial", 14, "bold") ).grid(row=0, column=0, padx=10, pady=5, sticky="w")
    entry_busca = tk.Entry(frame_busca, width=70,bg=branco, fg = cor_texto,highlightbackground=cor_fundo)
    entry_busca.grid(row=0, column=1, padx=10, pady=5, sticky="we")
    entry_busca.bind('<KeyRelease>', buscar_cadastro) 

    tk.Button(frame_busca, text="Buscar", command=buscar_cadastro, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).grid(row=0, column=2, padx=10, pady=5)
    tk.Button(frame_busca, text="Voltar", command=janela_consulta_cadastro.destroy, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).grid(row=0, column=3, padx=10, pady=5)

    frame_tabela = tk.Frame(janela_consulta_cadastro, bg=cor_fundo)
    frame_tabela.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")

    cols = ('Nome', 'Data de avaliação', 'Telefone')
    tree = ttk.Treeview(frame_tabela, columns=cols, show='headings')

    for col in cols:
        tree.heading(col, text=col)
        tree.column(col, anchor="w", width=330)  

    style = ttk.Style()
    style.configure("Treeview", background=cor, foreground=branco, fieldbackground=cor)
    style.configure("Treeview.Heading", background=branco, foreground=cor_fundo)
    style.map("Treeview", background=[('selected', cor)], foreground=[('selected', branco)])
    
    tree.grid(row=0, column=0, sticky="nsew")
    tree.bind("<<TreeviewSelect>>", selecionar_paciente)

    scrollbar = ttk.Scrollbar(frame_tabela, orient="vertical", command=tree.yview)
    tree.configure(yscroll=scrollbar.set, height=40)
    scrollbar.grid(row=0, column=1, sticky='ns')

    frame_tabela.grid_columnconfigure(0, weight=1)
    frame_tabela.grid_rowconfigure(0, weight=1)
    
    if os.path.exists('clientes.xlsx'):
        df = pd.read_excel('clientes.xlsx')
        atualizar_lista_tree(df)

    janela_consulta_cadastro.grid_columnconfigure(0, weight=1)
    janela_consulta_cadastro.grid_rowconfigure(1, weight=1)

    janela_consulta_cadastro.update_idletasks()
    largura = janela_consulta_cadastro.winfo_width() - 20  
    altura = janela_consulta_cadastro.winfo_height() - 20  
    janela_consulta_cadastro.geometry(f"{largura}x{altura}")


def editar_cadastro_janela(dados_cliente, callback=None):

    janela_editar = tk.Toplevel(root)
    janela_editar.title("Editar Cadastro")
    centralizar_janela(janela_editar)

    janela_editar.configure(bg=cor_fundo)
    canvas = tk.Canvas(janela_editar, bg=cor_fundo)
    scrollbar = tk.Scrollbar(janela_editar, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg=cor_fundo)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar.grid(row=0,column=1, sticky="ns")

    janela_editar.grid_rowconfigure(0, weight=1)
    janela_editar.grid_columnconfigure(0, weight=1)
    scrollable_frame.grid_columnconfigure(1, weight=1)

    logo_path3 = obter_caminho_imagem("logo_horizontal_branco.png")  
    logo_image3 = Image.open(logo_path3)
    logo_image3 = logo_image3.resize((500, 100), Image.Resampling.LANCZOS)  
    logo_photo3 = ImageTk.PhotoImage(logo_image3)
    
    logo_label_novo_cadastro = tk.Label(scrollable_frame, image=logo_photo3, bg=cor_fundo)
    logo_label_novo_cadastro.image = logo_photo3    

    logo_label_novo_cadastro.grid(row=0, column=0, columnspan=6, padx=5, pady=10, sticky="nsew")

    tk.Label(scrollable_frame, text="NOME:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=1, column=0, padx=10, pady=5, sticky="w")
    entry_nome_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_nome_editar.grid(row=1, column=1, columnspan=5, padx=10, pady=5, sticky="we")
    entry_nome_editar.insert(0, dados_cliente['Nome'])

    tk.Label(scrollable_frame, text="DATA:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=2, column=0, padx=10, pady=5, sticky="w")
    entry_data_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo, justify='center')
    entry_data_editar.grid(row=2, column=1, padx=10, pady=5, sticky="we")
    entry_data_editar.insert(0, dados_cliente['Data'])

    def formatar_data(event):
        data = entry_data_editar.get().replace("/", "")  
        
        if len(data) > 2 and len(data) <= 4:
            data = data[:2] + '/' + data[2:]
        elif len(data) > 4:
            data = data[:2] + '/' + data[2:4] + '/' + data[4:]

        entry_data_editar.delete(0, tk.END)
        entry_data_editar.insert(0, data)

    entry_data_editar.bind("<KeyRelease>", formatar_data)

    tk.Label(scrollable_frame, text="IDADE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=2, column=2, padx=10, pady=5, sticky="e")
    entry_idade_editar = tk.Entry(scrollable_frame, width=5, bg=branco, fg=cor_texto, highlightbackground=cor_fundo, justify='center')
    entry_idade_editar.grid(row=2, column=3, padx=10, pady=5, sticky="w")
    entry_idade_editar.insert(0, dados_cliente['Idade'])

    tk.Label(scrollable_frame, text="RESPONSÁVEL:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=3, column=2, padx=10, pady=5, sticky="e")
    entry_responsavel_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_responsavel_editar.grid(row=3, column=3, padx=10, columnspan=4,pady=5, sticky="we")
    entry_responsavel_editar.insert(0, dados_cliente['Responsável'])

    tk.Label(scrollable_frame, text="TELEFONE:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=3, column=0, padx=10, pady=5, sticky="w")
    entry_telefone_editar = tk.Entry(scrollable_frame, width=22, bg=branco, fg=cor_texto, highlightbackground=cor_fundo, justify='center')
    entry_telefone_editar.grid(row=3, columnspan=3, column=1, padx=10, pady=5, sticky="w")
    entry_telefone_editar.insert(0, dados_cliente['Telefone']) 
    
    tk.Label(scrollable_frame, text="ÚLTIMA AVALIACÃO:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=2, column=4, padx=10, pady=5, sticky="w")
    entry_avaliacao_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_avaliacao_editar.grid(row=2, column=5, padx=10, pady=5, sticky="w")
    entry_avaliacao_editar.insert(0, dados_cliente['Avaliação'])

    tk.Label(scrollable_frame, text="PROFISSÃO:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=5, column=0, padx=10, pady=5, sticky="w")
    entry_profissao_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_profissao_editar.grid(row=5, column=1, columnspan=5, padx=10, pady=5, sticky="we")
    entry_profissao_editar.insert(0, dados_cliente['Profissão'])

    tk.Label(scrollable_frame, text="MOTIVO DA CONSULTA:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=6, column=0, padx=10, pady=5, sticky="w")
    entry_motivo_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_motivo_editar.grid(row=6, columnspan=5, column=1, padx=10, pady=5, sticky="we")
    entry_motivo_editar.insert(0, dados_cliente['Motivo'])

    tk.Label(scrollable_frame, text="USUÁRIO DE RX:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=7, column=0, padx=10, pady=5, sticky="w")
    entry_opcao_var_editar = ttk.Combobox(scrollable_frame, width=20, values=('LC', 'Óculos','Ambos', 'Nenhum'))
    entry_opcao_var_editar.grid(row=7, column=1, padx=10, pady=5, sticky="we")
    entry_opcao_var_editar.insert(0, dados_cliente['Opção'])

    tk.Label(scrollable_frame, text="OPÇÕES DE ASSINALAR:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=8, column=0, padx=10, pady=5, sticky="w")

    opcoes = ['BEBE', 'FUMA', 'GESTANTE', 'LACTANTE', 'DORMIU BEM', 'DORMIU MAL', 'NENHUM']
    selected_values = [] 

    entry_obs_editar = ttk.Combobox(scrollable_frame, values=opcoes, width=15)
    entry_obs_editar.grid(row=8, columnspan=4, column=1, padx=10, pady=5, sticky="we")
    entry_obs_editar.insert(0, dados_cliente['Obs'])

    if 'Obs' in dados_cliente:
        observacoes_cliente = dados_cliente['Obs'].split(', ')  
        selected_values.extend(observacoes_cliente) 
        entry_obs_editar.set(', '.join(selected_values))  


    def toggle_editar_selection(event):
        value = entry_obs_editar.get()
        if value and value in opcoes:
            if value in selected_values:
                selected_values.remove(value)  
                print(f"Removido: {value}")  
            else:
                selected_values.append(value)  
                print(f"Adicionado: {value}")  
            atualiza_opcoes_editar()

    def atualiza_opcoes_editar():
        opcoes_disponiveis = [opcao for opcao in opcoes if opcao not in selected_values]
        entry_obs_editar['values'] = opcoes_disponiveis
        entry_obs_editar.set(', '.join(selected_values))

    def limpar_editar_selecao():
        removed_value = selected_values.pop()  
        print(f"Removido último: {removed_value}")  
        atualiza_opcoes_editar()

    botao_limpar_editar = tk.Button(scrollable_frame, text="Remover último", command=limpar_editar_selecao, highlightbackground=cor_fundo)
    botao_limpar_editar.grid(row=8, column=5, padx=10, pady=5, sticky="w")

    entry_obs_editar.bind("<<ComboboxSelected>>", toggle_editar_selection)

    tk.Label(scrollable_frame, text="PROBLEMA SISTÊMICO:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=9, column=0, padx=10, pady=5, sticky="w")
    entry_problema_sistemico_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_problema_sistemico_editar.grid(row=9, column=1, columnspan=5, padx=10, pady=5, sticky="we")
    entry_problema_sistemico_editar.insert(0, dados_cliente['Problema Sistêmico'])
 
    tk.Label(scrollable_frame, text="USO DE MEDICAMENTOS:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=10, column=0, padx=10, pady=5, sticky="w")
    entry_uso_medicamento_editar = tk.Entry(scrollable_frame, bg=branco, fg=cor_texto, highlightbackground=cor_fundo,justify='center')
    entry_uso_medicamento_editar.grid(row=10, column=1, columnspan=5, padx=10, pady=5, sticky="we")
    entry_uso_medicamento_editar.insert(0, dados_cliente['Uso Medicamento'])

    tk.Label(scrollable_frame, text="PROCEDIMENTOS NO OLHO:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=11, column=0, padx=10, pady=5, sticky="w")
    entry_procedimento_olho_editar = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=15)
    entry_procedimento_olho_editar.grid(row=11, column=1, columnspan=1, padx=10, pady=5, sticky="we")
    entry_procedimento_olho_editar.insert(0, dados_cliente['Procedimento Olho'])

    tk.Label(scrollable_frame, text="GLAUCOMA NA FAMILIA?", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=12, column=0, padx=10, pady=5, sticky="w")
    entry_hist_familia_editar = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=15)
    entry_hist_familia_editar.grid(row=12, column=1, columnspan=1, padx=10, pady=5, sticky="we")
    entry_hist_familia_editar.insert(0, dados_cliente['Hist. Familiar'])

    tk.Label(scrollable_frame, text="LENTES DE CONTATO APÓS 24 HORAS DE 'NÃO' USO", 
    bg=cor_fundo, fg=cor_texto, font=("Arial", 11, "bold")).grid(row=13, column=0, columnspan=2, padx=10, pady=10, sticky="w")
    entry_lente_contato_editar = ttk.Combobox(scrollable_frame, values=('Sim', 'Não'), width=10)
    entry_lente_contato_editar.grid(row=13, column=1, columnspan=5, padx=55, pady=5, sticky="w")
    entry_lente_contato_editar.insert(0, dados_cliente['Lente de Contato'])

    tk.Label(scrollable_frame, text="Dificuldade:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=14, column=0, padx=10, pady=5, sticky="w")
    entry_dificuldade_editar = ttk.Combobox(scrollable_frame, values=('PERTO', 'LONGE', 'OS DOIS', 'CLARIDADE', 'NENHUM'), width=15)
    entry_dificuldade_editar.grid(row=14, column=1, columnspan=1, padx=10, pady=5, sticky="we")
    entry_dificuldade_editar.insert(0, dados_cliente['Dificuldade'])

    tk.Label(scrollable_frame, text="Sintomas:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=15, column=0, padx=10, pady=5, sticky="w")
    sintomas_opcoes = ['ARDE', 'COÇA', 'PONTADA', 'DOR DE CABEÇA', 'DÓI NO SOL', 'LAGRIMEJA MUITO','VERMELHIDÃO', 'DOR NOS OLHOS' ,'RESSECADO', 'HIPEREMIA', 'NENHUM']
    selected_sintomas = []  
    entry_sintomas_editar = ttk.Combobox(scrollable_frame, values=sintomas_opcoes, width=15)
    entry_sintomas_editar.grid(row=15, column=1, columnspan=4, padx=10, pady=5, sticky="we")
    entry_sintomas_editar.insert(0, dados_cliente['Sintomas'])

    if 'Sintomas' in dados_cliente:
        sintomas_cliente = dados_cliente['Sintomas'].split(', ') 
        selected_sintomas.extend(sintomas_cliente)  
        entry_sintomas_editar.set(', '.join(selected_sintomas))  

    def toggle_sintomas_selection(event):
        value = entry_sintomas_editar.get()
        if value and value in sintomas_opcoes:
            if value in selected_sintomas:
                selected_sintomas.remove(value)  
                print(f"Sintoma Removido: {value}") 
            else:
                selected_sintomas.append(value) 
                print(f"Sintoma Adicionado: {value}")  
            atualiza_sintomas_opcoes()

    def atualiza_sintomas_opcoes():
        opcoes_disponiveis = [opcao for opcao in sintomas_opcoes if opcao not in selected_sintomas]
        entry_sintomas_editar['values'] = opcoes_disponiveis
        entry_sintomas_editar.set(', '.join(selected_sintomas))

    def limpar_sintomas_selecao():
            removed_value = selected_sintomas.pop() 
            print(f"Sintoma Removido Último: {removed_value}")  
            atualiza_sintomas_opcoes()


    botao_limpar_sintomas = tk.Button(scrollable_frame, text="Remover último", command=limpar_sintomas_selecao, highlightbackground=cor_fundo)
    botao_limpar_sintomas.grid(row=15, column=5, padx=10, pady=5, sticky="w")

    entry_sintomas_editar.bind("<<ComboboxSelected>>", toggle_sintomas_selection)

    tk.Label(scrollable_frame, text="OUTROS:", bg=cor_fundo, fg=cor_texto, font=("Arial", 14, "bold")).grid(row=18, column=0, padx=10, pady=5, sticky="w")
    entry_outros_editar = tk.Text(scrollable_frame, bg=branco, fg=cor_texto, height=4,wrap="word", highlightbackground=cor_fundo)
    entry_outros_editar.grid(row=18, column=1, columnspan=5, padx=10, pady=5, sticky="we")
    entry_outros_editar.insert("1.0", dados_cliente['Outros'])


    frame_tabela = tk.Frame(scrollable_frame, bg=cor_fundo)
    frame_tabela.grid(row=17, column=0, columnspan=3, padx=10, pady=5, sticky="e")

    titulo = tk.Label(frame_tabela, text="PRESCRIÇÃO DO ÚLTIMO ÓCULOS", font=("Arial", 16, "bold"), bg=cor_fundo, fg=cor_texto)
    titulo.grid(row=0, column=0, columnspan=7)

    labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
    for col, label in enumerate(labels, start=1):
        tk.Label(frame_tabela, text=label, bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=1, column=col, padx=10, pady=5)

    tk.Label(frame_tabela, text="O.D.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=2, column=0, padx=10, pady=5)
    tk.Label(frame_tabela, text="O.E.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=3, column=0, padx=10, pady=5)

    campos_od_oe = [
            ("Esfericidade OD", 'ESF. O.D. Último'),  
            ("Cilindro OD", 'CIL. O.D. Último'),
            ("Eixo OD", 'EIXO O.D. Último'),
            ("Adição OD", 'ADIÇÃO O.D. Último'),
            ("DNP OD", 'DNP O.D. Último'),
            ("ALT OD", 'ALT O.D. Último'),
            ("Esfericidade OE", 'ESF. O.E. Último'),
            ("Cilindro OE", 'CIL. O.E. Último'),
            ("Eixo OE", 'EIXO O.E. Último'),
            ("Adição OE", 'ADIÇÃO O.E. Último'),
            ("DNP OE", 'DNP O.E. Último'),
            ("ALT OE", 'ALT O.E. Último'),
            ("Tipo Lente", 'Tipo Lente Último')
        ]
    
    entries_od_ulitmo = {}

    for col, (label_text, key) in enumerate(campos_od_oe[:6], start=1):
        entry = tk.Entry(frame_tabela, bg=branco, fg=cor_texto, justify="center", width=4)
        entry.insert(0, dados_cliente[key])
        entry.grid(row=2, column=col,padx=3, pady=5, sticky="nsew")
        entries_od_ulitmo[key] = entry
       
    entries_oe_ulitmo = {}

    for col, (label_text, key) in enumerate(campos_od_oe[6:12], start=1):
        entry = tk.Entry(frame_tabela, bg=branco, fg=cor_texto, justify="center", width=4)
        entry.insert(0, dados_cliente[key])
        entry.grid(row=3, column=col,padx=3, pady=5, sticky="nsew")
        entries_oe_ulitmo[key] = entry

    label_tipo_lente = tk.Label(frame_tabela, text="Tipo Lente:", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold"))
    label_tipo_lente.grid(row=4, column=0, padx=5, pady=5, sticky="w")

    entry_tipo_lente = tk.Entry(frame_tabela, bg=branco, fg=cor_texto, justify="center")
    entry_tipo_lente.insert(0, dados_cliente['Tipo Lente Último'])
    entry_tipo_lente.grid(row=4, column=1, columnspan=6, padx=3, pady=5, sticky="we")


    frame_tabela1 = tk.Frame(scrollable_frame, bg=cor_fundo)
    frame_tabela1.grid(row=17, column=3, columnspan=3, padx=10, pady=5, sticky="w")

    titulo = tk.Label(frame_tabela1, text="RX FINAL", font=("Arial", 16, "bold"), bg=cor_fundo, fg=cor_texto)
    titulo.grid(row=0, column=0, columnspan=7)

    labels = ["ESF.", "CIL.", "EIXO", "ADIÇÃO", "DNP", "ALT"]
    for col, label in enumerate(labels, start=1):
        tk.Label(frame_tabela1, text=label, bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=1, column=col, padx=10, pady=5)

    tk.Label(frame_tabela1, text="O.D.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=2, column=0, padx=10, pady=5)
    tk.Label(frame_tabela1, text="O.E.", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold")).grid(row=3, column=0, padx=10, pady=5)

    # Campos para O.D. e O.E.
    campos_rx = [
        ("Esfericidade OD RX", 'ESF. O.D. RX'),
        ("Cilindro OD RX", 'CIL. O.D. RX'),
        ("Eixo OD RX", 'EIXO O.D. RX'),
        ("Adição OD RX", 'ADIÇÃO O.D. RX'),
        ("DNP OD RX", 'DNP O.D. RX'),
        ("Altura OD RX", 'ALT O.D. RX'),
        ("Esfericidade OE RX", 'ESF. O.E. RX'),
        ("Cilindro OE RX", 'CIL. O.E. RX'),
        ("Eixo OE RX", 'EIXO O.E. RX'),
        ("Adição OE RX", 'ADIÇÃO O.E. RX'),
        ("DNP OE RX", 'DNP O.E. RX'),
        ("Altura OE RX", 'ALT O.E. RX'),
        ("Tipo Lente RX", 'Tipo Lente RX')
    ]

    entries_od = {}
    for col, (label_text, key) in enumerate(campos_rx[:6], start=1):
        entry = tk.Entry(frame_tabela1, bg=branco, fg=cor_texto, justify="center", width=4)
        entry.grid(row=2, column=col, padx=3, pady=5, sticky="nsew")
        entry.insert(0, dados_cliente[key]) 
        entries_od[key] = entry 
        

    entries_oe = {}
    for col, (label_text, key) in enumerate(campos_rx[6:12], start=1):
        entry = tk.Entry(frame_tabela1, bg=branco, fg=cor_texto, justify="center",  width=4)
        entry.grid(row=3, column=col, padx=3, pady=5, sticky="nsew")
        entry.insert(0, dados_cliente[key]) 
        entries_oe[key] = entry  
    
    label_tipo_lente = tk.Label(frame_tabela1, text="Tipo Lente:", bg=cor_fundo, fg=cor_texto, font=("Arial", 12, "bold"))
    label_tipo_lente.grid(row=4, column=0, padx=5, pady=5, sticky="w")

    entry_tipo_lente_rx_editado = tk.Entry(frame_tabela1, bg=branco, fg=cor_texto, justify="center")
    entry_tipo_lente_rx_editado.grid(row=4, column=1, columnspan=6, padx=3, pady=5, sticky="we")
    entry_tipo_lente_rx_editado.insert(0, dados_cliente['Tipo Lente RX'])


    def salvar_edicao():
        def valor_ou_padrao(entry, padrao="0.0"):
            return entry.get().strip() if entry.get().strip() else padrao

        def valor_ou_default(entry, default="Nenhum"):
            return entry.get().strip() if entry.get().strip() else default

        nome_editado = valor_ou_default(entry_nome_editar)
        data_editado = valor_ou_default(entry_data_editar)
        idade_editado = valor_ou_default(entry_idade_editar)
        responsavel_editado = valor_ou_default(entry_responsavel_editar)
        telefone_editado = valor_ou_default(entry_telefone_editar)
        profissao_editado = valor_ou_default(entry_profissao_editar)
        motivo_editado = valor_ou_default(entry_motivo_editar)
        avaliacao_editado = valor_ou_default(entry_avaliacao_editar)
        opcao_var_editado = valor_ou_default(entry_opcao_var_editar)
        obs_editado = valor_ou_default(entry_obs_editar)
        problema_sistemico_editado = valor_ou_default(entry_problema_sistemico_editar)
        uso_medicamento_editado = valor_ou_default(entry_uso_medicamento_editar)
        procedimento_olho_editado = valor_ou_default(entry_procedimento_olho_editar)
        hist_familia_editado = valor_ou_default(entry_hist_familia_editar)
        lente_contato_editado = valor_ou_default(entry_lente_contato_editar)
        dificuldade_editado = valor_ou_default(entry_dificuldade_editar)
        sintomas_editado = valor_ou_default(entry_sintomas_editar)
        outros_editado = entry_outros_editar.get("1.0", "end-1c").strip() if entry_outros_editar.get("1.0", "end-1c").strip() else "Nenhum"

        esf_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['ESF. O.D. Último'])
        cil_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['CIL. O.D. Último'])
        eixo_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['EIXO O.D. Último'])
        adicao_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['ADIÇÃO O.D. Último'])
        dnp_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['DNP O.D. Último'])
        alt_od_ultimo_editado = valor_ou_padrao(entries_od_ulitmo['ALT O.D. Último'])

        esf_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['ESF. O.E. Último'])
        cil_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['CIL. O.E. Último'])
        eixo_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['EIXO O.E. Último'])
        adicao_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['ADIÇÃO O.E. Último'])
        dnp_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['DNP O.E. Último'])
        alt_oe_ultimo_editado = valor_ou_padrao(entries_oe_ulitmo['ALT O.E. Último'])

        tipo_lente_ultimo_editado = valor_ou_padrao(entry_tipo_lente, "Nenhum")

        esf_od_rx_editado = valor_ou_padrao(entries_od['ESF. O.D. RX'])
        cil_od_rx_editado = valor_ou_padrao(entries_od['CIL. O.D. RX'])
        eixo_od_rx_editado = valor_ou_padrao(entries_od['EIXO O.D. RX'])
        adicao_od_rx_editado = valor_ou_padrao(entries_od['ADIÇÃO O.D. RX'])
        dnp_od_rx_editado = valor_ou_padrao(entries_od['DNP O.D. RX'])
        alt_od_rx_editado = valor_ou_padrao(entries_od['ALT O.D. RX'])

        esf_oe_rx_editado = valor_ou_padrao(entries_oe['ESF. O.E. RX'])
        cil_oe_rx_editado = valor_ou_padrao(entries_oe['CIL. O.E. RX'])
        eixo_oe_rx_editado = valor_ou_padrao(entries_oe['EIXO O.E. RX'])
        adicao_oe_rx_editado = valor_ou_padrao(entries_oe['ADIÇÃO O.E. RX'])
        dnp_oe_rx_editado = valor_ou_padrao(entries_oe['DNP O.E. RX'])
        alt_oe_rx_editado = valor_ou_padrao(entries_oe['ALT O.E. RX'])

        tipo_lente_rx_editado = valor_ou_padrao(entry_tipo_lente_rx_editado, "Nenhum")

        if os.path.exists('clientes.xlsx'):
            df = pd.read_excel('clientes.xlsx')

            df.loc[df['Nome'] == dados_cliente['Nome'], [
                'Nome',
                'Data',
                'Idade',
                'Telefone',
                'Responsável',
                'Profissão',
                'Motivo',
                'Avaliação',
                'Opção',
                'Obs',
                'Problema Sistêmico',
                'Uso Medicamento',
                'Procedimento Olho',
                'Hist. Familiar',
                'Lente de Contato',
                'Dificuldade',
                'Sintomas',
                'Outros',
                'ESF. O.D. Último',
                'CIL. O.D. Último',
                'EIXO O.D. Último',
                'ADIÇÃO O.D. Último',
                'DNP O.D. Último',
                'ALT O.D. Último',
                'ESF. O.E. Último',
                'CIL. O.E. Último',
                'EIXO O.E. Último',
                'ADIÇÃO O.E. Último',
                'DNP O.E. Último',
                'ALT O.E. Último',
                'Tipo Lente Último',
                'ESF. O.D. RX',
                'CIL. O.D. RX',
                'EIXO O.D. RX',
                'ADIÇÃO O.D. RX',
                'DNP O.D. RX',
                'ALT O.D. RX',
                'ESF. O.E. RX',
                'CIL. O.E. RX',
                'EIXO O.E. RX',
                'ADIÇÃO O.E. RX',
                'DNP O.E. RX',
                'ALT O.E. RX',
                'Tipo Lente RX'
            ]] = [
                nome_editado, data_editado, idade_editado, telefone_editado, responsavel_editado, 
                profissao_editado, motivo_editado, avaliacao_editado, 
                opcao_var_editado, obs_editado, problema_sistemico_editado, uso_medicamento_editado, 
                procedimento_olho_editado, hist_familia_editado, lente_contato_editado, 
                dificuldade_editado, sintomas_editado, outros_editado, 
                esf_od_ultimo_editado, cil_od_ultimo_editado, eixo_od_ultimo_editado, 
                adicao_od_ultimo_editado, dnp_od_ultimo_editado, alt_od_ultimo_editado, 
                esf_oe_ultimo_editado, cil_oe_ultimo_editado, eixo_oe_ultimo_editado, 
                adicao_oe_ultimo_editado, dnp_oe_ultimo_editado, alt_oe_ultimo_editado, 
                tipo_lente_ultimo_editado, 
                esf_od_rx_editado, cil_od_rx_editado, eixo_od_rx_editado, 
                adicao_od_rx_editado, dnp_od_rx_editado, alt_od_rx_editado, 
                esf_oe_rx_editado, cil_oe_rx_editado, eixo_oe_rx_editado, 
                adicao_oe_rx_editado, dnp_oe_rx_editado, alt_oe_rx_editado, 
                tipo_lente_rx_editado
            ]

        
        df.to_excel('clientes.xlsx', index=False)

        messagebox.showinfo("Sucesso", "Cadastro atualizado com sucesso!")
        janela_editar.destroy()

        if callback:
            callback(df)

        else:
            messagebox.showwarning("Erro", "Arquivo de dados não encontrado!")

    def remover_cliente():
        if messagebox.askyesno("Confirmar Exclusão", "Tem certeza de que deseja remover este cadastro?"):
            if os.path.exists('clientes.xlsx'):
                df = pd.read_excel('clientes.xlsx')
                df = df[(df['Nome'] != dados_cliente['Nome']) | (df['Telefone'] != dados_cliente['Telefone'])]

                df.to_excel('clientes.xlsx', index=False)

                messagebox.showinfo("Sucesso", "Cadastro removido com sucesso!")
                janela_editar.destroy()

                if callback:
                    callback(df)
            else:
                messagebox.showwarning("Erro", "Arquivo de dados não encontrado!")

    frame_botoes = tk.Frame(scrollable_frame, bg=cor_fundo)
    frame_botoes.grid(row=21, column=2, columnspan=2, pady=10)

    tk.Button(frame_botoes, text="Salvar", command=salvar_edicao, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Remover", command=remover_cliente, bg="#ff4d4d", fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)
    tk.Button(frame_botoes, text="Voltar", command=janela_editar.destroy, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo).pack(side=tk.LEFT, padx=5)

def centralizar_janela(janela, centralizar_vertical=True):
    janela.update_idletasks()
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()
    largura_janela = 1200
    altura_janela = 900
    pos_x = (largura_tela - largura_janela) // 2
    pos_y = (altura_tela - altura_janela) // 2 if centralizar_vertical else 50
    janela.geometry(f"{largura_janela}x{altura_janela}+{pos_x}+{pos_y}")

def abrir_arquivo_excel():
    caminho_excel = "clientes.xlsx"  
    os.system(f'open "{caminho_excel}"')

def obter_caminho_imagem(nome_imagem):
    if getattr(sys, 'frozen', False):
        caminho = os.path.join(sys._MEIPASS, nome_imagem)
    else:
        caminho = os.path.join(os.path.dirname(__file__), nome_imagem)
    return caminho

root = tk.Tk()
root.title("Sistema de Cadastro")
centralizar_janela(root, centralizar_vertical=True)

root.configure(bg=cor_fundo)

logo_path = obter_caminho_imagem("logo_horizontal_branco.png")  
logo_image = Image.open(logo_path)
logo_image = logo_image.resize((1000, 300), Image.Resampling.LANCZOS)  
logo_photo = ImageTk.PhotoImage(logo_image)

logo_label = tk.Label(root, image=logo_photo, bg=cor_fundo)
logo_label.pack(pady=10)


tk.Button(root, text="Novo Cadastro", command=abrir_novo_cadastro, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo, font=("Arial", 18, "bold"),width=60).pack(pady=10)
tk.Button(root, text="Consulta de Cadastro", command=abrir_consulta_cadastro, bg=cor_fundo, fg=cor_texto,highlightbackground=cor_fundo, font=("Arial", 18, "bold"),width=60).pack(pady=10)
tk.Button(root, text="Abrir Excel", command=abrir_arquivo_excel, bg=cor_fundo, fg=cor_texto, highlightbackground=cor_fundo, font=("Arial", 18, "bold"),width=60).pack(pady=10)

root.mainloop()

