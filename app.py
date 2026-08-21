# -*- coding: utf-8 -*-
"""
app.py
-------------------------------------------------
Interface gráfica de usuário (GUI) para a aplicação Gestor do Bolsão.
Utiliza tkinter e a biblioteca ttkbootstrap para um visual moderno.
Este arquivo lida com a apresentação e a interação com o usuário,
enquanto o backend.py cuida de toda a lógica de negócio.
"""
# --- Importações de Módulos ---
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as bs
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import ScrolledFrame
import pandas as pd
from pathlib import Path
import json # Módulo para lidar com o arquivo de salvamento offline (JSON)
import os 
import ctypes
import requests
from packaging.version import parse as parse_version
import subprocess
import sys
import re

# Importa todas as funções de lógica do nosso outro arquivo
import backend as be

class App(bs.Window):
    def __init__(self, title, size):
        super().__init__(themename="minty")
        
        # --- DEFINIÇÃO CENTRAL DA VERSÃO ---
        # Alterar aqui atualiza para todo o sistema (Update, Título, ID)
        self.APP_VERSION = "3.0.6"
        
        myappid = f'MatrizEducacao.GestorBolsao.Desktop.{self.APP_VERSION}' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        try:
            self.icon_path = be.resource_path(os.path.join("images", "matriz.ico"))
            self.iconbitmap(self.icon_path)
        except tk.TclError:
            print("Aviso: Ícone 'images/matriz.ico' não encontrado ou inválido.")

        # Verifica atualização usando a versão definida acima
        if self.check_for_updates():
            return 
            
        self.title(f"Gestor do Bolsão {self.APP_VERSION}")
        
        # Estado dos Filtros
        self.geometry(f'{size[0]}x{size[1]}')
        self.minsize(size[0], size[1])
        
        self.snapshot_data = None
        self.hubspot_df = None

        self.setup_main_ui()

        self.loading_frame = ttk.Frame(self)
        self.loading_frame.place(relx=0.5, rely=0.5, anchor='center')
        self.loading_label_var = tk.StringVar(value="Conectando e carregando dados...")
        self.loading_label_var_style = ttk.Label(self.loading_frame, textvariable=self.loading_label_var, font=("-size 12"))
        self.loading_label_var_style.pack(pady=10)
        self.progress_bar = ttk.Progressbar(self.loading_frame, mode='indeterminate')
        self.progress_bar.pack(pady=10, fill='x', padx=20)
        self.progress_bar.start()
        
        self.after(200, self.load_initial_data)

    def check_for_updates(self):
        """Verifica, extrai o updater para um local seguro, e inicia a atualização."""
        # CORREÇÃO: Usa a versão definida no __init__ em vez de valor fixo
        VERSION_URL = "https://raw.githubusercontent.com/Inteligencia-Matriz/BolsaoDesktop/main/version.json"

        try:
            response = requests.get(VERSION_URL, timeout=5)
            response.raise_for_status()
            data = response.json()
            server_version_str = data["version"]
            
            # Compara a versão do servidor com a versão atual da classe (self.APP_VERSION)
            if parse_version(server_version_str) > parse_version(self.APP_VERSION):
                if messagebox.askyesno("Atualização Disponível", 
                                       f"Uma nova versão ({server_version_str}) está disponível.\nDeseja atualizar agora?",
                                       parent=self):
                    
                    embedded_updater_path = be.resource_path("updater.exe")
                    temp_dir = os.getenv('TEMP')
                    stable_updater_path = os.path.join(temp_dir, "updater.exe")

                    with open(embedded_updater_path, 'rb') as f_in:
                        with open(stable_updater_path, 'wb') as f_out:
                            f_out.write(f_in.read())
                    
                    zip_url = data["url"]
                    current_exe_path = sys.executable

                    subprocess.Popen([stable_updater_path, zip_url, current_exe_path])
                    self.destroy()
                    return True
        
        except requests.RequestException:
            print("Não foi possível verificar por atualizações (sem conexão ou timeout).")
        except Exception as e:
            messagebox.showerror("Erro na Verificação", f"Ocorreu um erro ao verificar por atualizações:\n{e}", parent=self)
        
        return False

    def load_initial_data(self):
        """Carrega os dados em segundo plano. Se falhar, exibe um erro claro na UI."""
        try:
            self.hubspot_df = be.get_hubspot_data_for_activation()
            self.snapshot_data = be.load_resultados_snapshot()
            
            self.progress_bar.stop()
            self.loading_frame.destroy()
            self.enable_ui_components(True)
            self.populate_form_filters_initial()
            self.sync_offline_data(silent=True)
            self.update_status_bar()

        except Exception as e:
            self.progress_bar.stop()
            self.loading_label_var.set("Falha na conexão com a planilha.")
            self.progress_bar.pack_forget()

            retry_button = ttk.Button(self.loading_frame, text="Tentar Novamente", command=self.retry_load, style="success.TButton")
            retry_button.pack(pady=10)

    def retry_load(self):
        """Função para o botão 'Tentar Novamente'."""
        for widget in self.loading_frame.winfo_children():
            if isinstance(widget, ttk.Button):
                widget.destroy()
        
        self.progress_bar.pack(pady=10, fill='x', padx=20)
        self.loading_label_var.set("Conectando e carregando dados...")
        self.progress_bar.start()
        self.after(100, self.load_initial_data)

    def setup_main_ui(self):
        """Constrói a UI principal, mas começa com os widgets desabilitados."""
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)
        
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self, textvariable=self.status_var, anchor='w', padding=(5, 2))
        self.status_bar.pack(side='bottom', fill='x')

        self.create_carta_tab()
        self.create_negociacao_tab()
        self.create_formulario_tab()
        self.create_valores_tab()
        
        self.enable_ui_components(False)

    def enable_ui_components(self, enabled: bool):
        """Habilita ou desabilita todos os widgets interativos da aplicação."""
        state = 'normal' if enabled else 'disabled'
        for tab in self.notebook.winfo_children():
            for widget in tab.winfo_children():
                self.set_widget_state(widget, state)

    def set_widget_state(self, parent_widget, state):
        """Função recursiva para alterar o estado de um widget e de todos os seus filhos."""
        try:
            if not isinstance(parent_widget, ttk.Label):
                parent_widget.config(state=state)
        except tk.TclError:
            pass
        
        for child in parent_widget.winfo_children():
            self.set_widget_state(child, state)

    def _configure_combobox_click(self, combobox_widget):
        """Configura um Combobox para abrir a lista ao ser clicado em qualquer lugar."""
        def open_dropdown(event):
            combobox_widget.focus_set()
            combobox_widget.event_generate('<Down>')
        combobox_widget.bind("<Button-1>", open_dropdown)

    # --- ABA 1: GERAR CARTA ---
    def create_carta_tab(self):
        carta_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(carta_frame, text='Gerar Carta')

        self.load_frame = ttk.LabelFrame(carta_frame, text="Filtrar Candidato", padding=15)
        self.load_frame.pack(fill='x', padx=10, pady=10)
        self.load_frame.grid_columnconfigure(1, weight=1)
        self.c_load_unidade_var = tk.StringVar(value=be.UNIDADES_LIMPAS[0])
        self.c_load_candidato_var = tk.StringVar()
        
        ttk.Label(self.load_frame, text="Filtrar por Unidade:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.c_unidade_filter_combo = ttk.Combobox(self.load_frame, textvariable=self.c_load_unidade_var, values=be.UNIDADES_LIMPAS, state="readonly")
        self.c_unidade_filter_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self.c_unidade_filter_combo.bind("<<ComboboxSelected>>", self.filter_hubspot_candidates_by_unit)
        self._configure_combobox_click(self.c_unidade_filter_combo)

        ttk.Label(self.load_frame, text="Selecione o Candidato:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.c_candidato_combo = ttk.Combobox(self.load_frame, textvariable=self.c_load_candidato_var, state="readonly", height=10)
        self.c_candidato_combo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self.c_candidato_combo.bind("<<ComboboxSelected>>", self.populate_from_hubspot)
        self._configure_combobox_click(self.c_candidato_combo)
        
        self.filter_hubspot_candidates_by_unit()

        self.form_frame = ttk.LabelFrame(carta_frame, text="Dados do Candidato", padding=15)
        self.form_frame.pack(fill='x', padx=10, pady=10)
        self.form_frame.grid_columnconfigure(1, weight=1)
        self.form_frame.grid_columnconfigure(3, weight=1)

        self.c_nome_var = tk.StringVar()
        self.c_unidade_var = tk.StringVar(value=be.UNIDADES_LIMPAS[0])
        self.c_turma_var = tk.StringVar(value=list(be.TURMA_DE_INTERESSE_MAP.keys())[0])
        self.c_ac_mat_var = tk.IntVar(value=0)
        self.c_ac_port_var = tk.IntVar(value=0)
        self.c_serie_var = tk.StringVar()
        self.c_bolsa_resultado_var = tk.StringVar(value="➔ Bolsa obtida: -")

        def update_serie_and_limits(*args):
            turma = self.c_turma_var.get()
            serie = be.TURMA_DE_INTERESSE_MAP.get(turma, "")
            self.c_serie_var.set(serie)
            max_acertos = 10 if serie == "1º ao 5º Ano" else 20
            self.c_ac_mat_spinbox.config(to=max_acertos)
            self.c_ac_port_spinbox.config(to=max_acertos)
            if self.c_ac_mat_var.get() > max_acertos: self.c_ac_mat_var.set(max_acertos)
            if self.c_ac_port_var.get() > max_acertos: self.c_ac_port_var.set(max_acertos)
        
        self.c_turma_var.trace_add("write", update_serie_and_limits)
        self.c_ac_mat_var.trace_add("write", self.calcular_bolsa_display)
        self.c_ac_port_var.trace_add("write", self.calcular_bolsa_display)
        self.c_unidade_var.trace_add("write", self.calcular_bolsa_display)

        ttk.Label(self.form_frame, text="Nome do Candidato:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(self.form_frame, textvariable=self.c_nome_var, width=50).grid(row=0, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        ttk.Label(self.form_frame, text="Unidade:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        unidade_combo = ttk.Combobox(self.form_frame, textvariable=self.c_unidade_var, values=be.UNIDADES_LIMPAS, state="readonly")
        unidade_combo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(unidade_combo)
        ttk.Label(self.form_frame, text="Turma de Interesse:").grid(row=1, column=2, padx=5, pady=5, sticky='w')
        turma_combo = ttk.Combobox(self.form_frame, textvariable=self.c_turma_var, values=list(be.TURMA_DE_INTERESSE_MAP.keys()), state="readonly")
        turma_combo.grid(row=1, column=3, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(turma_combo)
        ttk.Label(self.form_frame, text="Acertos - Matemática:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.c_ac_mat_spinbox = ttk.Spinbox(self.form_frame, from_=0, to=12, textvariable=self.c_ac_mat_var, width=8)
        self.c_ac_mat_spinbox.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(self.form_frame, text="Acertos - Português:").grid(row=2, column=2, padx=5, pady=5, sticky='w')
        self.c_ac_port_spinbox = ttk.Spinbox(self.form_frame, from_=0, to=12, textvariable=self.c_ac_port_var, width=8)
        self.c_ac_port_spinbox.grid(row=2, column=3, padx=5, pady=5, sticky='w')
        ttk.Label(self.form_frame, text="Série/Modalidade:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(self.form_frame, textvariable=self.c_serie_var, state="readonly").grid(row=3, column=1, columnspan=3, padx=5, pady=5, sticky='ew')
        
        form_actions_frame = ttk.Frame(self.form_frame)
        form_actions_frame.grid(row=4, column=0, columnspan=4, pady=10)
        ttk.Button(form_actions_frame, text="Limpar Campos", command=self.clear_carta_form).pack(side='left', padx=5)
        ttk.Label(form_actions_frame, textvariable=self.c_bolsa_resultado_var, font=("Helvetica", 12, "bold")).pack(side='left', padx=20)

        action_frame = ttk.Frame(carta_frame)
        action_frame.pack(fill='x', padx=10, pady=10)
        ttk.Button(action_frame, text="Gerar e Salvar Carta PDF", command=self.gerar_carta, style='success.TButton').pack(side='left', padx=10, expand=True)
        ttk.Button(action_frame, text="Sincronizar Dados Offline", command=self.sync_offline_data, style='info.TButton').pack(side='left', padx=10, expand=True)
        
        update_serie_and_limits()

    def filter_hubspot_candidates_by_unit(self, event=None):
        """Filtra os candidatos já em memória com base na unidade selecionada."""
        if self.hubspot_df is None:
            return

        unidade_sel = self.c_load_unidade_var.get()
        unidade_completa = be.UNIDADES_MAP.get(unidade_sel)
        df_filtrado = self.hubspot_df[self.hubspot_df['Unidade'] == unidade_completa]
        
        if df_filtrado.empty:
            self.c_candidato_combo['values'] = []
        else:
            nomes_candidatos = sorted(df_filtrado['Nome do Candidato'].tolist())
            self.c_candidato_combo['values'] = nomes_candidatos
        
        self.c_load_candidato_var.set("")

    def populate_from_hubspot(self, event=None):
        """Preenche os campos do formulário com os dados do candidato selecionado."""
        nome_selecionado = self.c_load_candidato_var.get()
        if not nome_selecionado or self.hubspot_df is None:
            return

        candidato_data = self.hubspot_df[self.hubspot_df['Nome do Candidato'] == nome_selecionado].iloc[0]
        
        self.c_nome_var.set(candidato_data.get('Nome do Candidato', ''))
        unidade_completa = candidato_data.get('Unidade', '')
        unidade_limpa = next((key for key, value in be.UNIDADES_MAP.items() if value == unidade_completa), be.UNIDADES_LIMPAS[0])
        self.c_unidade_var.set(unidade_limpa)

        serie_modalidade = candidato_data.get('Turma de Interesse - Geral', '')
        turma_interesse = next(
            (key for key, value in be.TURMA_DE_INTERESSE_MAP.items() if value == serie_modalidade), 
            list(be.TURMA_DE_INTERESSE_MAP.keys())[0]
        )
        self.c_turma_var.set(turma_interesse)

    def clear_carta_form(self):
        """Limpa todos os campos do formulário da aba Gerar Carta."""
        self.c_nome_var.set("")
        self.c_unidade_var.set(be.UNIDADES_LIMPAS[0])
        self.c_turma_var.set(list(be.TURMA_DE_INTERESSE_MAP.keys())[0])
        self.c_ac_mat_var.set(0)
        self.c_ac_port_var.set(0)
        self.c_load_candidato_var.set("")
        self.calcular_bolsa_display()

    def calcular_bolsa_display(self, *args):
        """Calcula e exibe o percentual da bolsa na tela."""
        try:
            total = self.c_ac_mat_var.get() + self.c_ac_port_var.get()
            serie = self.c_serie_var.get()
            unidade = self.c_unidade_var.get()
            pct = be.calcula_bolsa(total, serie, unidade)
            self.c_bolsa_resultado_var.set(f"➔ Bolsa obtida: {pct*100:.0f}% ({total} acertos)")
        except Exception as e:
            messagebox.showerror("Erro de Cálculo", str(e))

    def gerar_carta(self):
        """Coleta dados, prepara o nome do arquivo, gera e salva o PDF."""
        aluno = self.c_nome_var.get()
        if not aluno:
            messagebox.showerror("Erro de Validação", "O nome do candidato é obrigatório.")
            return
        try:
            unidade_limpa = self.c_unidade_var.get()
            turma = self.c_turma_var.get()
            ac_mat = self.c_ac_mat_var.get()
            ac_port = self.c_ac_port_var.get()
            serie_modalidade = self.c_serie_var.get()
            total_acertos = ac_mat + ac_port

            brasilia_datetime = be.get_current_brasilia_datetime()
            hoje = brasilia_datetime.date()
            nome_bolsao = be.get_bolsao_name_for_date(hoje)

            pct_bolsa = be.calcula_bolsa(total_acertos, serie_modalidade, unidade_limpa)
            precos = be.precos_2027(serie_modalidade)
            
            # --- CÁLCULOS PÁGINA 1: ANUIDADE E PARCELAMENTO NORMAL ---
            # Anuidade Total com o desconto da Bolsa (Normal)
            anuidade_com_bolsa = precos["anuidade"] * (1 - pct_bolsa)
            
            # Parcelamento padrão (Matrícula R$ 300, Saldo em 12x)
            entrada_normal = 300.00
            saldo_restante_normal = anuidade_com_bolsa - entrada_normal
            if saldo_restante_normal < 0: saldo_restante_normal = 0
            val_12x_normal = saldo_restante_normal / 12

            # --- CÁLCULOS PÁGINA 1: CONDIÇÃO ESPECIAL (HOJE) ---
            # Regra: Seguir o percentual obtido sem +5%
            
            pct_especial_hoje = pct_bolsa
            if pct_especial_hoje > 1.0: 
                pct_especial_hoje = 1.0
            
            # Valor total da anuidade com o desconto especial
            anuidade_especial_total = precos["anuidade"] * (1 - pct_especial_hoje)
            
            # Valor fixo da 1ª parcela
            entrada_especial = 300.00
            
            # Abate a entrada para achar o saldo
            saldo_restante_especial = anuidade_especial_total - entrada_especial
            if saldo_restante_especial < 0: saldo_restante_especial = 0
            
            # O saldo é dividido em 12 parcelas
            val_12x_especial = saldo_restante_especial / 12

            # --- CRIAÇÃO DO TEXTO DINÂMICO PARA O CABEÇALHO ---
            base_int = int(round(pct_bolsa * 100))
            total_int = int(round(pct_especial_hoje * 100))
            texto_condicao = f"Condições de hoje {total_int}%"

            # --- CÁLCULOS PÁGINA 3: PROPOSTA ESPECIAL (GENÉRICA) ---
            # Seguir a risca a bolsa, sem +5%
            proposta_pct = pct_bolsa
            if proposta_pct > 1.0: proposta_pct = 1.0
            anuidade_proposta = precos["anuidade"] * (1 - proposta_pct)
            
            entrada_proposta = 300.00
            saldo_proposta = anuidade_proposta - entrada_proposta
            if saldo_proposta < 0: saldo_proposta = 0
            prop12_val = saldo_proposta / 12
            
            aluno_safe = re.sub(r'[\\/*?:"<>|]', "", aluno.strip())
            bolsao_safe = re.sub(r'[\\/*?:"<>|]', "", nome_bolsao.strip()).replace(" ", "_")
            initial_filename = f"Carta_{aluno_safe}_{bolsao_safe}.pdf"
            initial_dir = str(Path.home() / "Downloads")
            
            html_tabelas_material = be.gerar_html_material_didatico(unidade_limpa)
            
            # Contexto enviado para o HTML (carta.html)
            ctx = {
                "ano": "2027", 
                "unidade": f"Colégio Matriz – {unidade_limpa}",
                "aluno": aluno.strip().title(), 
                "bolsa_pct": f"{pct_bolsa * 100:.0f}",
                "acertos_mat": ac_mat, 
                "acertos_port": ac_port, 
                "turma": turma,
                "data_limite": (hoje + be.timedelta(days=7)).strftime("%d/%m/%Y"),
                
                # Valores Página 1 - Tabela Superior (Normal)
                "anuidade_total_bolsa": be.format_currency(anuidade_com_bolsa),
                "entrada_normal": be.format_currency(entrada_normal),
                "val_12x_normal": be.format_currency(val_12x_normal),
                
                # Valores Página 1 - Tabela Inferior (Condição de Hoje)
                # Passamos o texto dinâmico gerado aqui
                "texto_condicao_hoje": texto_condicao,
                "entrada_especial": be.format_currency(entrada_especial),
                "val_12x_especial": be.format_currency(val_12x_especial),
                
                # Valores Página 3 (Proposta Especial +5% genérica)
                "proposta_pct": f"{proposta_pct * 100:.0f}",
                "entrada_proposta": be.format_currency(entrada_proposta),
                "prop12_val": be.format_currency(prop12_val),
                
                "unidades_html": "".join(f"<span class='unidade-item'>{u}</span>" for u in be.UNIDADES_LIMPAS),
                "tabelas_material_didatico": html_tabelas_material,
            }
            
            pdf_bytes = be.gera_pdf_html(ctx)

            file_path = filedialog.asksaveasfilename(
                initialdir=initial_dir,
                initialfile=initial_filename,
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Salvar Carta PDF"
            )
            
            if file_path:
                with open(file_path, "wb") as f: f.write(pdf_bytes)
                messagebox.showinfo("Sucesso", f"Carta PDF salva com sucesso em:\n{file_path}")
                if messagebox.askyesno("Registrar na Planilha?", "Deseja registrar este resultado na planilha online?"):
                    self.registrar_na_planilha(aluno, unidade_limpa, turma, ac_mat, ac_port, total_acertos, pct_bolsa, serie_modalidade, ctx, brasilia_datetime, nome_bolsao)
        except Exception as e:
            messagebox.showerror("Erro ao Gerar Carta", str(e))

    def registrar_na_planilha(self, aluno, unidade_limpa, turma, ac_mat, ac_port, total, pct, serie, ctx, brasilia_dt, nome_bolsao):
        """Envia os dados gerados para a planilha Resultados_Bolsao."""
        row_data_map = {
            "Data/Hora": brasilia_dt.strftime("%d/%m/%Y %H:%M:%S"),
            "Nome do Aluno": aluno.strip().title(),
            "Unidade": be.UNIDADES_MAP[unidade_limpa],
            "Turma de Interesse": turma,
            "Acertos Matemática": ac_mat,
            "Acertos Português": ac_port,
            "Total de Acertos": total,
            "% Bolsa": f"{pct*100:.0f}%",
            "Série / Modalidade": serie,
            "Valor Anuidade à Vista": ctx["anuidade_total_bolsa"], 
            "Valor da 1ª Cota": ctx["val_12x_normal"], # Registrando o valor da parcela normal como referência
            "Valor da Mensalidade com Bolsa": ctx["val_12x_normal"],
            "REGISTRO_ID": be.new_uuid(),
            "Bolsão": nome_bolsao
        }
        
        try:
            ws_res = be.get_ws("Resultados_Bolsao")
            hmap_res = be.header_map("Resultados_Bolsao")
            header_list = sorted(hmap_res, key=hmap_res.get)
            nova_linha = [row_data_map.get(col_name, "") for col_name in header_list]
            ws_res.append_row(nova_linha, value_input_option="USER_ENTERED")
            messagebox.showinfo("Sucesso", "Dados registrados na planilha online!")

            try:
                self.status_var.set("Sincronizando dados atualizados...")
                self.update() 
                self.snapshot_data = be.load_resultados_snapshot()
                self.status_var.set("Dados sincronizados.")
            except Exception as sync_error:
                messagebox.showwarning("Aviso de Sincronização", f"O registro foi salvo, mas a sincronização automática falhou. Pode ser necessário reiniciar para editar.\nErro: {sync_error}")

        except Exception as e:
            messagebox.showwarning(
                "Falha na Conexão",
                f"Não foi possível registrar na planilha online.\nErro: {e}\n\nOs dados serão salvos localmente e enviados mais tarde."
            )
            self.save_to_offline_queue(row_data_map)
            if self.snapshot_data:
                self.snapshot_data['rows'].append(row_data_map)
        
        if self.snapshot_data:
            self.populate_form_filters_initial()

    # --- ABA 2: NEGOCIAÇÃO ---
    def create_negociacao_tab(self):
        neg_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(neg_frame, text='Negociação')
        self.n_unidade_var = tk.StringVar(value=be.UNIDADES_LIMPAS[0])
        self.n_serie_var = tk.StringVar(value=list(be.TUITION.keys())[0])
        self.n_valor_minimo_var = tk.StringVar(value="Valor Mínimo: R$ 0,00")
        self.n_modo_sim_var = tk.StringVar(value="Bolsa (%)")
        self.n_bolsa_sim_var = tk.IntVar(value=30)
        self.n_valor_neg_var = tk.DoubleVar(value=1500.0)
        self.n_resultado_var = tk.StringVar()
        self.n_bolsa_percent_var = tk.StringVar(value="30%")

        def update_negociacao_calculo(*args):
            try:
                self.n_bolsa_percent_var.set(f"{self.n_bolsa_sim_var.get()}%")
                unidade = self.n_unidade_var.get()
                serie = self.n_serie_var.get()
                valor_minimo = be.calcula_valor_minimo(unidade, serie)
                self.n_valor_minimo_var.set(f"Valor Mínimo Negociável: {be.format_currency(valor_minimo)}")
                precos = be.precos_2027(serie)
                valor_integral = precos["parcela_mensal"]
                resultado_str = ""
                if self.n_modo_sim_var.get() == "Bolsa (%)":
                    bolsa_pct = self.n_bolsa_sim_var.get() / 100
                    valor_final = valor_integral * (1 - bolsa_pct)
                    resultado_str = f"Valor da Parcela: {be.format_currency(valor_final)}"
                    if valor_final < valor_minimo: resultado_str += " (Abaixo do mínimo!)"
                else:
                    valor_desejado = self.n_valor_neg_var.get()
                    bolsa_necessaria = (1 - (valor_desejado / valor_integral)) * 100 if valor_integral > 0 else 0
                    resultado_str = f"Bolsa Necessária: {bolsa_necessaria:.2f}%"
                    if valor_desejado < valor_minimo: resultado_str += " (Abaixo do mínimo!)"
                self.n_resultado_var.set(resultado_str)
            except Exception as e:
                self.n_resultado_var.set("Erro no cálculo.")
                messagebox.showerror("Erro de Cálculo", str(e))

        top_frame = ttk.Frame(neg_frame)
        top_frame.pack(fill='x', padx=10, pady=5)
        top_frame.grid_columnconfigure(1, weight=1)
        ttk.Label(top_frame, text="Unidade:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        unidade_neg_combo = ttk.Combobox(top_frame, textvariable=self.n_unidade_var, values=be.UNIDADES_LIMPAS, state='readonly')
        unidade_neg_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(unidade_neg_combo)
        ttk.Label(top_frame, text="Série/Modalidade:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        serie_neg_combo = ttk.Combobox(top_frame, textvariable=self.n_serie_var, values=list(be.TUITION.keys()), state='readonly')
        serie_neg_combo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(serie_neg_combo)
        ttk.Separator(neg_frame).pack(fill='x', padx=10, pady=10)
        ttk.Label(neg_frame, textvariable=self.n_valor_minimo_var, font=("-size 12 -weight bold")).pack(pady=5)
        ttk.Separator(neg_frame).pack(fill='x', padx=10, pady=10)
        sim_frame = ttk.LabelFrame(neg_frame, text="Simulador", padding=15)
        sim_frame.pack(fill='x', padx=10, pady=10)
        
        slider_frame = ttk.Frame(sim_frame)
        slider_frame.pack(fill='x', padx=20, pady=5)
        ttk.Radiobutton(sim_frame, text="Calcular por Bolsa (%)", variable=self.n_modo_sim_var, value="Bolsa (%)").pack(anchor='w')
        ttk.Scale(slider_frame, from_=0, to=100, variable=self.n_bolsa_sim_var, orient='horizontal', length=350, command=update_negociacao_calculo).pack(side='left', fill='x', expand=True)
        ttk.Label(slider_frame, textvariable=self.n_bolsa_percent_var, font=("-size 10 -weight bold")).pack(side='left', padx=10)

        ttk.Radiobutton(sim_frame, text="Calcular por Valor da Parcela (R$)", variable=self.n_modo_sim_var, value="Valor da Parcela (R$)").pack(anchor='w', pady=(10,0))
        valor_input_frame = ttk.Frame(sim_frame)
        valor_input_frame.pack(fill='x', padx=20, pady=5)
        ttk.Label(valor_input_frame, text="R$").pack(side='left')
        ttk.Entry(valor_input_frame, textvariable=self.n_valor_neg_var).pack(side='left', fill='x', expand=True)
        
        ttk.Label(neg_frame, textvariable=self.n_resultado_var, font=("-size 14 -weight bold"), style='success.TLabel').pack(pady=20)
        
        self.n_unidade_var.trace_add('write', update_negociacao_calculo)
        self.n_serie_var.trace_add('write', update_negociacao_calculo)
        self.n_modo_sim_var.trace_add('write', update_negociacao_calculo)
        ttk.Button(sim_frame, text="Calcular", command=update_negociacao_calculo).pack(pady=10)
        update_negociacao_calculo()

    # --- ABA 3: FORMULÁRIO BÁSICO ---
    def create_formulario_tab(self):
        self.f_unidade_var = tk.StringVar()
        self.f_bolsao_var = tk.StringVar()
        self.f_candidato_var = tk.StringVar()
        self.f_info_var = tk.StringVar()
        self.f_escola_var = tk.StringVar()
        self.f_resp_fin_var = tk.StringVar()
        self.f_tel_var = tk.StringVar()
        
        self.f_valor_neg_var = tk.StringVar()
        self.f_expectativa_var = tk.StringVar()

        self.f_matriculou_var = tk.StringVar()
        self.f_obs_var = None
        
        tab_container = ttk.Frame(self.notebook)
        self.notebook.add(tab_container, text='Formulário Básico')
        f_scrolled_frame = ScrolledFrame(tab_container, autohide=True)
        f_scrolled_frame.pack(fill="both", expand=True)
        
        filter_frame = ttk.LabelFrame(f_scrolled_frame, text="Filtros", padding=10)
        filter_frame.pack(fill='x', padx=10, pady=5)
        filter_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(filter_frame, text="Unidade:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.f_unidade_combo = ttk.Combobox(filter_frame, textvariable=self.f_unidade_var, values=["Carregue os dados..."], state='readonly')
        self.f_unidade_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(self.f_unidade_combo)
        ttk.Label(filter_frame, text="Bolsão:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.f_bolsao_combo = ttk.Combobox(filter_frame, textvariable=self.f_bolsao_var, values=["Filtre por unidade..."], state='readonly')
        self.f_bolsao_combo.grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(self.f_bolsao_combo)
        ttk.Label(filter_frame, text="Candidato:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        self.f_candidato_combo = ttk.Combobox(filter_frame, textvariable=self.f_candidato_var, values=["Filtre por bolsão..."], state='readonly', height=15)
        self.f_candidato_combo.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(self.f_candidato_combo)
        
        edit_frame = ttk.LabelFrame(f_scrolled_frame, text="Editar Registro", padding=10)
        edit_frame.pack(fill='both', expand=True, padx=10, pady=10)
        edit_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(edit_frame, textvariable=self.f_info_var, font=("-size 10 -weight bold")).grid(row=0, column=0, columnspan=2, pady=5, sticky='w')
        ttk.Label(edit_frame, text="Escola de Origem:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(edit_frame, textvariable=self.f_escola_var).grid(row=1, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(edit_frame, text="Responsável Financeiro:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(edit_frame, textvariable=self.f_resp_fin_var).grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(edit_frame, text="Telefone:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(edit_frame, textvariable=self.f_tel_var).grid(row=3, column=1, padx=5, pady=5, sticky='ew')
        ttk.Label(edit_frame, text="Valor Negociado (R$):").grid(row=4, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(edit_frame, textvariable=self.f_valor_neg_var).grid(row=4, column=1, padx=5, pady=5, sticky='ew')
        
        ttk.Label(edit_frame, text="Expectativa de mensalidade (R$):").grid(row=5, column=0, padx=5, pady=5, sticky='w')
        ttk.Entry(edit_frame, textvariable=self.f_expectativa_var).grid(row=5, column=1, padx=5, pady=5, sticky='ew')

        ttk.Label(edit_frame, text="Aluno Matriculou?").grid(row=6, column=0, padx=5, pady=5, sticky='w')
        matriculou_combo = ttk.Combobox(edit_frame, textvariable=self.f_matriculou_var, values=["", "Sim", "Não"], state='readonly')
        matriculou_combo.grid(row=6, column=1, padx=5, pady=5, sticky='ew')
        self._configure_combobox_click(matriculou_combo)
        ttk.Label(edit_frame, text="Observações:").grid(row=7, column=0, padx=5, pady=5, sticky='nw')
        self.f_obs_var = tk.Text(edit_frame, height=4, wrap='word')
        self.f_obs_var.grid(row=7, column=1, padx=5, pady=5, sticky='ew')
        ttk.Button(edit_frame, text="Salvar Formulário", command=self.save_form_data, style='success.TButton').grid(row=8, column=0, columnspan=2, pady=15)
        
        self.f_unidade_combo.bind("<<ComboboxSelected>>", self.update_form_filters)
        self.f_bolsao_combo.bind("<<ComboboxSelected>>", self.update_form_filters)
        self.f_candidato_combo.bind("<<ComboboxSelected>>", self.populate_form_fields)
        
        self.populate_form_filters_initial()

        self.f_valor_neg_var.trace_add("write", lambda *args, var=self.f_valor_neg_var: self._validate_and_format_currency(var, *args))
        self.f_expectativa_var.trace_add("write", lambda *args, var=self.f_expectativa_var: self._validate_and_format_currency(var, *args))

    def populate_form_filters_initial(self):
        """Popula os filtros do formulário com os dados já carregados."""
        if self.snapshot_data:
            unidades_completas = sorted({r.get("Unidade") for r in self.snapshot_data['rows'] if r.get("Unidade")})
            unidades_limpas = sorted([u.replace("COLEGIO E CURSO MATRIZ EDUCACAO", "").replace("COLEGIO E CURSO MATRIZ EDUCAÇÃO", "").strip() for u in unidades_completas])
            self.f_unidade_combo['values'] = ["Todas"] + unidades_limpas
            self.f_unidade_var.set("Todas")
            self.update_form_filters()
    
    def update_form_filters(self, event=None):
        """Filtra os dados com base nas seleções de unidade e bolsão."""
        if not self.snapshot_data: return
        unidade_sel = self.f_unidade_var.get()
        bolsao_sel = self.f_bolsao_var.get()
        rows_unit = []
        if unidade_sel == "Todas": rows_unit = self.snapshot_data['rows']
        else:
            unidade_completa = be.UNIDADES_MAP.get(unidade_sel)
            rows_unit = [r for r in self.snapshot_data['rows'] if r.get("Unidade") == unidade_completa]
        bolsoes = sorted({r.get("Bolsão") for r in rows_unit if r.get("Bolsão")})
        self.f_bolsao_combo['values'] = ["Todos"] + bolsoes
        self.filtered_rows = []
        if bolsao_sel == "Todos" or not bolsao_sel: self.filtered_rows = rows_unit
        else: self.filtered_rows = [r for r in rows_unit if r.get("Bolsão") == bolsao_sel]
        candidatos = [f"{r.get('Nome do Aluno')} ({r.get('REGISTRO_ID')})" for r in self.filtered_rows]
        self.f_candidato_combo['values'] = sorted(candidatos)
        self.f_candidato_var.set("")
        self.clear_form_fields()

    def populate_form_fields(self, event=None):
        """Preenche o formulário com os dados do candidato selecionado."""
        selecao = self.f_candidato_var.get()
        if not selecao:
            self.clear_form_fields()
            return
        reg_id = selecao.split('(')[-1][:-1]
        self.selected_reg_id = reg_id
        row = next((r for r in self.filtered_rows if str(r.get("REGISTRO_ID")) == reg_id), None)
        if not row:
            self.clear_form_fields()
            return

        col_expectativa = "Expectativa de mensalidade"
        col_expectativa_fallback = "Valor Limite (PIA)"
        expectativa_val = row.get(col_expectativa, row.get(col_expectativa_fallback, 0.0))

        self.f_info_var.set(f"Aluno: {row.get('Nome do Aluno')} | Bolsa: {row.get('% Bolsa')} | Parcela: {row.get('Valor da Mensalidade com Bolsa')}")
        self.f_escola_var.set(row.get("Escola de Origem", ""))
        self.f_resp_fin_var.set(row.get("Responsável Financeiro", ""))
        self.f_tel_var.set(be.format_phone_mask(row.get("Telefone", "")))
        
        valor_neg_float = be.parse_brl_to_float(row.get("Valor Negociado", 0.0))
        expectativa_float = be.parse_brl_to_float(expectativa_val)
        self.f_valor_neg_var.set(be.format_currency(valor_neg_float))
        self.f_expectativa_var.set(be.format_currency(expectativa_float))
        
        self.f_matriculou_var.set(row.get("Aluno Matriculou?", ""))
        self.f_obs_var.delete('1.0', END)
        self.f_obs_var.insert('1.0', row.get("Observações (Form)", ""))

    def clear_form_fields(self):
        """Limpa todos os campos do formulário de edição."""
        self.selected_reg_id = None
        self.f_info_var.set("")
        self.f_escola_var.set("")
        self.f_resp_fin_var.set("")
        self.f_tel_var.set("")
        
        self.f_valor_neg_var.set("")
        self.f_expectativa_var.set("")

        self.f_matriculou_var.set("")
        if self.f_obs_var: self.f_obs_var.delete('1.0', END)

    def save_form_data(self):
        """Salva os dados do formulário na planilha."""
        if not self.selected_reg_id:
            messagebox.showwarning("Aviso", "Nenhum candidato selecionado para salvar.")
            return

        escola_origem = self.f_escola_var.get().strip()
        if not escola_origem:
            messagebox.showerror("Campo Obrigatório", "O campo 'Escola de Origem' não pode estar vazio.")
            return

        try:
            rownum = self.snapshot_data['id_to_rownum'].get(str(self.selected_reg_id))
            if not rownum:
                messagebox.showerror("Erro", "Não foi possível encontrar o número da linha para este registro. Sincronize novamente.")
                return
            ws_res = be.get_ws("Resultados_Bolsao")
            hmap = be.header_map("Resultados_Bolsao")
            
            valor_neg_float = be.parse_brl_to_float(self.f_valor_neg_var.get())
            expectativa_float = be.parse_brl_to_float(self.f_expectativa_var.get())

            updates_dict = {
                "Escola de Origem": self.f_escola_var.get(),
                "Responsável Financeiro": self.f_resp_fin_var.get(),
                "Telefone": self.f_tel_var.get(),
                "Valor Negociado": be.format_currency(valor_neg_float),
                "Aluno Matriculou?": self.f_matriculou_var.get(),
                "Observações (Form)": self.f_obs_var.get('1.0', 'end-1c'),
            }
            col_expectativa = "Expectativa de mensalidade"
            col_expectativa_fallback = "Valor Limite (PIA)"
            if col_expectativa in hmap:
                updates_dict[col_expectativa] = be.format_currency(expectativa_float)
            elif col_expectativa_fallback in hmap:
                updates_dict[col_expectativa_fallback] = be.format_currency(expectativa_float)

            updates_to_batch = []
            for col_name, value in updates_dict.items():
                col_idx = hmap.get(col_name)
                if col_idx:
                    a1_notation = be.gspread.utils.rowcol_to_a1(rownum, col_idx)
                    updates_to_batch.append({"range": a1_notation, "values": [[value]]})
            if updates_to_batch:
                be.batch_update_cells(ws_res, updates_to_batch)
                messagebox.showinfo("Sucesso", "Dados do formulário salvos com sucesso na planilha!")
                self.snapshot_data = be.load_resultados_snapshot()
                self.update_form_filters()
            else:
                messagebox.showinfo("Informação", "Nenhuma alteração para salvar.")
        except Exception as e:
            messagebox.showerror("Erro ao Salvar", str(e))

    # --- ABA 4: VALORES ---
    def create_valores_tab(self):
        val_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(val_frame, text='Valores 2027')
        ttk.Label(val_frame, text="Tabela de Valores", font=("-size 14 -weight bold")).pack(pady=10)
        cols = ["Curso", "Série", "Primeira Cota", "12 parcelas de"]
        tree = ttk.Treeview(val_frame, columns=cols, show='headings', style='info.Treeview')
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, anchor=CENTER, width=150)
        linhas = [
            ("EFI",  "1º Ano", 2251.24, 2251.24), ("EFI",  "2º Ano", 2251.24, 2251.24),
            ("EFI",  "3º Ano", 2251.24, 2251.24), ("EFI",  "4º Ano", 2251.24, 2251.24),
            ("EFI",  "5º Ano", 2251.24, 2251.24), ("EFII", "6º Ano", 2648.21, 2648.21),
            ("EFII", "7º Ano", 2648.21, 2648.21), ("EFII", "8º Ano", 2648.21, 2648.21),
            ("EFII", "9º Ano - Militar",     2884.02, 2884.02),
            ("EFII", "9º Ano - Vestibular", 2884.02, 2884.02),
            ("EM",   "1ª Série - Militar",     3097.20, 3097.20),
            ("EM",   "1ª Série - Vestibular", 3097.20, 3097.20),
            ("EM",   "2ª Série - Militar",     3097.20, 3097.20),
            ("EM",   "2ª Série - Vestibular", 3097.20, 3097.20),
            ("EM",   "3ª série - Medicina",   3109.20, 3109.20),
            ("EM",   "3ª Série - Militar",     3109.20, 3109.20),
            ("EM",   "3ª Série - Vestibular", 3109.20, 3109.20),
            ("PM",   "AFA/EN/EFOMM", 1250.20, 1250.20), ("PM",   "CN/EPCAr", 748.58,  748.58),
            ("PM",   "ESA", 603.48,  603.48), ("PM",   "EsPCEx", 1250.20, 1250.20),
            ("PM",   "IME/ITA", 1250.20, 1250.20), ("PV",   "Medicina", 1250.20, 1250.20),
            ("PV",   "Pré-Vestibular", 1250.20, 1250.20),
        ]
        for linha in linhas:
            formatted_linha = (linha[0], linha[1], be.format_currency(linha[2]), be.format_currency(linha[3]))
            tree.insert("", END, values=formatted_linha)
        tree.pack(expand=True, fill='both', padx=10, pady=10)
    
    def _validate_and_format_currency(self, var: tk.StringVar, *args):
        # Flag para evitar recursão infinita
        if hasattr(self, '_formatting_in_progress') and self._formatting_in_progress:
            return
        
        self._formatting_in_progress = True
        
        try:
            current_value = var.get()
            digits = "".join(filter(str.isdigit, str(current_value)))
            
            if not digits:
                var.set("")
            else:
                float_value = float(digits) / 100
                formatted_value = be.format_currency(float_value)
                var.set(formatted_value)
        except (ValueError, tk.TclError):
            pass
        finally:
            self._formatting_in_progress = False

    # --- FUNÇÕES PARA A FILA OFFLINE ---
    def load_offline_queue(self):
        """Carrega a fila de registros offline do arquivo JSON."""
        try:
            with open("offline_queue.json", "r") as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return []

    def save_to_offline_queue(self, data):
        """Salva um novo registro na fila offline."""
        queue = self.load_offline_queue()
        queue.append(data)
        with open("offline_queue.json", "w") as f:
            json.dump(queue, f, indent=4)
        self.update_status_bar()

    def update_status_bar(self):
        """Atualiza o texto da barra de status com o número de itens na fila."""
        queue = self.load_offline_queue()
        count = len(queue)
        if count > 0:
            self.status_var.set(f"{count} registro(s) na fila para sincronizar.")
        else:
            self.status_var.set("Todos os dados estão sincronizados.")

    def sync_offline_data(self, silent=False):
        """Tenta enviar todos os registros da fila offline para a planilha online."""
        queue = self.load_offline_queue()
        if not queue:
            if not silent:
                messagebox.showinfo("Sincronização", "Não há dados offline para sincronizar.")
            return

        if not silent:
            if not messagebox.askyesno("Sincronização", f"Deseja enviar {len(queue)} registro(s) pendente(s) agora?"):
                return

        try:
            ws_res = be.get_ws("Resultados_Bolsao")
            hmap_res = be.header_map("Resultados_Bolsao")
            header_list = sorted(hmap_res, key=hmap_res.get)
            
            linhas_para_enviar = []
            for record in queue:
                nova_linha = [record.get(col_name, "") for col_name in header_list]
                linhas_para_enviar.append(nova_linha)

            if linhas_para_enviar:
                ws_res.append_rows(linhas_para_enviar, value_input_option="USER_ENTERED")
                with open("offline_queue.json", "w") as f:
                    json.dump([], f)
                
                if not silent:
                    messagebox.showinfo("Sincronização Concluída", f"{len(linhas_para_enviar)} registro(s) enviados com sucesso.")
            
            self.update_status_bar()

        except Exception as e:
            if not silent:
                messagebox.showerror("Erro de Sincronização", f"Não foi possível conectar à planilha. Tente novamente mais tarde.\nErro: {e}")

if __name__ == '__main__':
    app = App(title="Gestor do Bolsão", size=(800, 650))
    app.mainloop()