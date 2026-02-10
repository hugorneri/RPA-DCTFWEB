"""
Interface grafica para a automacao DCTF usando CustomTkinter.
Visual moderno e limpo, mantendo o fluxo funcional existente.
"""
import logging
import queue
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import ttk
import webbrowser

import customtkinter as ctk
import pandas as pd

from src.automacao import configurar_driver, transmissao
from src.config import Config, get_config, save_config


COLORS = {
    "bg": "#111827",
    "card": "#1f2937",
    "card_alt": "#243244",
    "text": "#e5e7eb",
    "text_dim": "#9ca3af",
    "accent": "#3b82f6",
    "accent_hover": "#2563eb",
    "success": "#10b981",
    "warning": "#f59e0b",
    "danger": "#ef4444",
    "border": "#334155",
    "log_bg": "#0b1220",
    "log_fg": "#7dd3fc",
}


class TextHandler(logging.Handler):
    """Handler de logging que envia logs para uma fila."""

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(self.format(record))


class AutomacaoDCTFApp:
    """Aplicacao principal da interface grafica."""

    def __init__(self):
        self.config = get_config()

        # Janela principal
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.root = ctk.CTk()
        self.root.title("Automacao DCTF")
        self.root.geometry("1200x840")
        self.root.minsize(1000, 760)

        # Estados de execucao
        self.running = False
        self.should_stop = False
        self.waiting_login = False
        self.driver = None
        self.worker_thread = None
        self.log_queue = queue.Queue()

        # Dados da planilha
        self.df = None
        self.cnpjs = []
        self.codigos = []
        self.planilha_carregada = False

        # Variaveis de UI
        self.planilha_path_var = tk.StringVar(value=str(self.config.planilha))
        self.resumo_var = tk.StringVar(value="Nenhuma planilha carregada")
        self.status_var = tk.StringVar(value="Aguardando...")
        self.progress_pct_var = tk.IntVar(value=0)

        self.field_vars = {
            "data_inicial": tk.StringVar(),
            "data_final": tk.StringVar(),
            "competencia": tk.StringVar(),
            "timeout": tk.StringVar(),
            "tentativas_cnpj": tk.StringVar(),
            "tentativas_gerais": tk.StringVar(),
        }

        self.setup_logging()
        self.create_widgets()
        self.load_config_to_fields()
        self.check_log_queue()

    # =========================================================================
    # WIDGETS
    # =========================================================================
    def create_widgets(self):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        self.main = ctk.CTkScrollableFrame(self.root, fg_color=COLORS["bg"], corner_radius=0)
        self.main.grid(row=0, column=0, sticky="nsew")
        self.main.grid_columnconfigure(0, weight=2)
        self.main.grid_columnconfigure(1, weight=1)

        self._build_header()
        self._build_planilha_card()
        self._build_config_card()
        self._build_control_card()
        self._build_log_card()
        self._build_footer()

    def _build_header(self):
        header = ctk.CTkFrame(self.main, fg_color="transparent")
        header.grid(row=0, column=0, columnspan=2, sticky="ew", padx=18, pady=(12, 8))
        header.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            header,
            text="AUTOMACAO DCTF",
            font=ctk.CTkFont(family="Segoe UI", size=30, weight="bold"),
            text_color=COLORS["text"],
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkLabel(
            header,
            text="Sistema de Download Automatico DCTFWeb",
            font=ctk.CTkFont(family="Segoe UI", size=14),
            text_color=COLORS["text_dim"],
        ).grid(row=1, column=0, sticky="w", pady=(2, 4))

        separator = ctk.CTkFrame(self.main, fg_color=COLORS["accent"], height=2)
        separator.grid(row=1, column=0, columnspan=2, sticky="ew", padx=18, pady=(0, 10))

    def _card(self, parent, title):
        outer = ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=14, border_width=1, border_color=COLORS["border"])
        outer.grid_columnconfigure(0, weight=1)

        title_bar = ctk.CTkFrame(outer, fg_color=COLORS["card_alt"], corner_radius=10, height=42)
        title_bar.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 0))
        ctk.CTkLabel(
            title_bar,
            text=title,
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=COLORS["text"],
        ).pack(anchor="w", padx=12, pady=10)

        content = ctk.CTkFrame(outer, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=12, pady=12)
        content.grid_columnconfigure(0, weight=1)
        return outer, content

    def _build_planilha_card(self):
        card_outer, card = self._card(self.main, "PLANILHA DE DADOS")
        card_outer.grid(row=2, column=0, sticky="nsew", padx=(18, 9), pady=8)
        card.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(card, text="Arquivo:", text_color=COLORS["text"]).grid(row=0, column=0, sticky="w", padx=(2, 8), pady=(0, 8))

        self.path_entry = ctk.CTkEntry(
            card,
            textvariable=self.planilha_path_var,
            state="disabled",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            fg_color="#0f172a",
            border_color=COLORS["border"],
            text_color=COLORS["text_dim"],
        )
        self.path_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=0, column=2, padx=(8, 0), pady=(0, 8))
        ctk.CTkButton(actions, text="Selecionar", width=100, command=self.select_planilha).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(actions, text="Carregar Dados", width=120, command=self.load_planilha, fg_color=COLORS["success"], hover_color="#059669").pack(side=tk.LEFT, padx=4)

        ctk.CTkLabel(
            card,
            textvariable=self.resumo_var,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            text_color=COLORS["warning"],
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(2, 10))

        table_wrap = ctk.CTkFrame(card, fg_color="#0b1220", corner_radius=10)
        table_wrap.grid(row=2, column=0, columnspan=3, sticky="nsew")
        table_wrap.grid_columnconfigure(0, weight=1)
        table_wrap.grid_rowconfigure(0, weight=1)
        card.grid_rowconfigure(2, weight=1)

        style = ttk.Style(self.root)
        style.theme_use("default")
        style.configure(
            "Dark.Treeview",
            background="#0b1220",
            foreground=COLORS["text"],
            fieldbackground="#0b1220",
            rowheight=28,
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10),
        )
        style.configure(
            "Dark.Treeview.Heading",
            background=COLORS["accent"],
            foreground="#ffffff",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
        )
        style.map("Dark.Treeview", background=[("selected", "#1d4ed8")], foreground=[("selected", "#ffffff")])
        style.map("Dark.Treeview.Heading", background=[("active", COLORS["accent_hover"])])

        y_scroll = ttk.Scrollbar(table_wrap, orient="vertical")
        y_scroll.grid(row=0, column=1, sticky="ns")

        self.data_table = ttk.Treeview(
            table_wrap,
            columns=("cod", "cnpj", "nome", "status"),
            show="headings",
            style="Dark.Treeview",
            yscrollcommand=y_scroll.set,
            height=8,
        )
        y_scroll.configure(command=self.data_table.yview)
        self.data_table.grid(row=0, column=0, sticky="nsew")

        self.data_table.heading("cod", text="Codigo")
        self.data_table.heading("cnpj", text="CNPJ")
        self.data_table.heading("nome", text="Nome / Razao Social")
        self.data_table.heading("status", text="Status")
        self.data_table.column("cod", width=80, minwidth=60, anchor=tk.CENTER)
        self.data_table.column("cnpj", width=180, minwidth=130, anchor=tk.W)
        self.data_table.column("nome", width=370, minwidth=180, anchor=tk.W)
        self.data_table.column("status", width=230, minwidth=140, anchor=tk.W)

    def _build_config_card(self):
        card_outer, card = self._card(self.main, "CONFIGURACOES")
        card_outer.grid(row=3, column=0, sticky="ew", padx=(18, 9), pady=8)
        card.grid_columnconfigure((0, 2, 4), weight=1)

        fields = [
            ("Data Inicial", "data_inicial", "DDMMAAAA", 0, 0),
            ("Data Final", "data_final", "DDMMAAAA", 0, 2),
            ("Competencia", "competencia", "MM AAAA", 1, 0),
            ("Timeout (seg)", "timeout", "", 1, 2),
            ("Tentativas/CNPJ", "tentativas_cnpj", "", 2, 0),
            ("Tentativas Gerais", "tentativas_gerais", "", 2, 2),
        ]

        for label, key, hint, row, col in fields:
            ctk.CTkLabel(card, text=label, text_color=COLORS["text"]).grid(row=row * 2, column=col, sticky="w", padx=6, pady=(4, 0))
            entry = ctk.CTkEntry(
                card,
                textvariable=self.field_vars[key],
                fg_color="#0f172a",
                border_color=COLORS["border"],
                text_color=COLORS["text"],
            )
            entry.grid(row=row * 2 + 1, column=col, sticky="ew", padx=6, pady=(0, 6))
            if hint:
                ctk.CTkLabel(card, text=hint, text_color=COLORS["text_dim"], font=ctk.CTkFont(size=11)).grid(
                    row=row * 2 + 1, column=col + 1, sticky="w", padx=(0, 6), pady=(0, 6)
                )

        ctk.CTkButton(
            card,
            text="Salvar Configuracoes",
            command=self.save_config,
            width=190,
            fg_color="#334155",
            hover_color="#475569",
        ).grid(row=7, column=2, columnspan=2, sticky="e", padx=6, pady=(10, 2))

    def _build_control_card(self):
        card_outer, card = self._card(self.main, "CONTROLE DA AUTOMACAO")
        card_outer.grid(row=2, column=1, sticky="new", padx=(9, 18), pady=8)
        card.grid_columnconfigure(0, weight=1)

        actions = ctk.CTkFrame(card, fg_color="transparent")
        actions.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        actions.grid_columnconfigure((0, 1, 2), weight=1)

        self.start_btn = ctk.CTkButton(
            actions,
            text="Iniciar Automacao",
            command=self.start_automation,
            fg_color=COLORS["success"],
            hover_color="#059669",
            height=40,
        )
        self.start_btn.grid(row=0, column=0, padx=4, sticky="ew")

        self.stop_btn = ctk.CTkButton(
            actions,
            text="Parar",
            command=self.stop_automation,
            fg_color=COLORS["danger"],
            hover_color="#dc2626",
            state="disabled",
            height=40,
        )
        self.stop_btn.grid(row=0, column=1, padx=4, sticky="ew")

        self.login_btn = ctk.CTkButton(
            actions,
            text="Confirmar Login",
            command=self.confirm_login,
            fg_color=COLORS["warning"],
            hover_color="#d97706",
            text_color="#111827",
            state="disabled",
            height=40,
        )
        self.login_btn.grid(row=0, column=2, padx=4, sticky="ew")

        status_row = ctk.CTkFrame(card, fg_color="transparent")
        status_row.grid(row=1, column=0, sticky="ew", pady=(2, 8))
        status_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(status_row, text="Status:", text_color=COLORS["text_dim"]).grid(row=0, column=0, sticky="w", padx=(2, 8))
        ctk.CTkLabel(
            status_row,
            textvariable=self.status_var,
            text_color=COLORS["success"],
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=1, sticky="w")

        self.progress_bar = ctk.CTkProgressBar(card, progress_color=COLORS["success"], height=16)
        self.progress_bar.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(card, text="0/0 | 0%", text_color=COLORS["text"])
        self.progress_label.grid(row=3, column=0, sticky="e")

    def _build_log_card(self):
        card_outer, card = self._card(self.main, "LOG DE EXECUCAO")
        card_outer.grid(row=3, column=1, sticky="nsew", padx=(9, 18), pady=8)
        self.main.grid_rowconfigure(3, weight=1)
        card_outer.grid_rowconfigure(1, weight=1)
        card.grid_rowconfigure(0, weight=1)
        card.grid_columnconfigure(0, weight=1)

        self.log_text = ctk.CTkTextbox(
            card,
            height=300,
            corner_radius=8,
            fg_color=COLORS["log_bg"],
            text_color=COLORS["log_fg"],
            font=ctk.CTkFont(family="Consolas", size=12),
            border_color=COLORS["border"],
            border_width=1,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        log_actions = ctk.CTkFrame(card, fg_color="transparent")
        log_actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ctk.CTkButton(
            log_actions,
            text="Limpar Log",
            width=110,
            command=self.clear_log,
            fg_color="#334155",
            hover_color="#475569",
        ).pack(side=tk.RIGHT)

    def _build_footer(self):
        footer = ctk.CTkFrame(self.main, fg_color="transparent")
        footer.grid(row=4, column=0, columnspan=2, sticky="ew", padx=18, pady=(2, 16))
        footer.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            footer,
            text="v1.0.0 - Automacao DCTF",
            text_color=COLORS["text_dim"],
            font=ctk.CTkFont(size=11),
        ).grid(row=0, column=0, sticky="w")

        ctk.CTkButton(
            footer,
            text="Manual de Instrucoes",
            command=self.open_manual,
            width=170,
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
        ).grid(row=0, column=1, sticky="e")

    # =========================================================================
    # LOGGING
    # =========================================================================
    def setup_logging(self):
        logging.basicConfig(
            filename=str(self.config.log_file),
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )
        root_logger = logging.getLogger()
        has_handler = any(isinstance(handler, TextHandler) for handler in root_logger.handlers)
        if not has_handler:
            handler = TextHandler(self.log_queue)
            handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
            root_logger.addHandler(handler)

    # =========================================================================
    # CARREGAR/SALVAR CONFIG
    # =========================================================================
    def load_config_to_fields(self):
        self.field_vars["data_inicial"].set(self.config.data_inicial)
        self.field_vars["data_final"].set(self.config.data_final)
        self.field_vars["competencia"].set(self.config.competencia)
        self.field_vars["timeout"].set(str(self.config.timeout_elemento))
        self.field_vars["tentativas_cnpj"].set(str(self.config.tentativas_por_cnpj))
        self.field_vars["tentativas_gerais"].set(str(self.config.tentativas_gerais))
        self.planilha_path_var.set(str(self.config.planilha))

    def get_config_from_fields(self) -> Config:
        return Config(
            data_inicial=self.field_vars["data_inicial"].get().strip(),
            data_final=self.field_vars["data_final"].get().strip(),
            competencia=self.field_vars["competencia"].get().strip(),
            timeout_elemento=int(self.field_vars["timeout"].get()),
            tentativas_por_cnpj=int(self.field_vars["tentativas_cnpj"].get()),
            tentativas_gerais=int(self.field_vars["tentativas_gerais"].get()),
            planilha_path=self.planilha_path_var.get().strip(),
        )

    def save_config(self):
        try:
            config = self.get_config_from_fields()
            save_config(config)
            self.config = config
            self.log_message("Configuracoes salvas com sucesso!")
            messagebox.showinfo("Sucesso", "Configuracoes salvas com sucesso!")
        except ValueError as e:
            messagebox.showerror("Erro", f"Valores invalidos: {e}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {e}")

    # =========================================================================
    # PLANILHA
    # =========================================================================
    def select_planilha(self):
        filepath = filedialog.askopenfilename(
            title="Selecionar Planilha",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls"), ("Todos", "*.*")],
            initialdir=str(self.config.pasta_base),
        )
        if filepath:
            self.planilha_path_var.set(filepath)
            self.log_message(f"Planilha selecionada: {filepath}")
            self.df = None
            self.cnpjs = []
            self.codigos = []
            self.planilha_carregada = False
            self.resumo_var.set("Clique em 'Carregar Dados' para visualizar")
            for item in self.data_table.get_children():
                self.data_table.delete(item)

    def load_planilha(self):
        planilha_path = self.planilha_path_var.get()
        if not planilha_path:
            messagebox.showwarning("Aviso", "Nenhuma planilha selecionada!")
            return
        if not Path(planilha_path).exists():
            messagebox.showerror("Erro", f"Arquivo nao encontrado: {planilha_path}")
            return

        try:
            self.log_message(f"Carregando planilha: {planilha_path}")
            df = pd.read_excel(planilha_path)

            required = ["CNPJ", "COD"]
            missing = [c for c in required if c not in df.columns]
            if missing:
                raise ValueError(f"Colunas obrigatorias nao encontradas: {', '.join(missing)}")

            df["CNPJ"] = df["CNPJ"].astype(str).str.strip()
            df["COD"] = df["COD"].astype(str).str.strip()
            if "STATUS" not in df.columns:
                df["STATUS"] = ""
            else:
                df["STATUS"] = df["STATUS"].fillna("")

            self.df = df
            self.cnpjs = df["CNPJ"].tolist()
            self.codigos = df["COD"].tolist()
            self.planilha_carregada = True

            self._populate_table()

            total = len(self.cnpjs)
            baixados = len(df[df["STATUS"].str.contains("Guia baixada", na=False)])
            pendentes = total - baixados
            self.log_message(f"Planilha carregada: {total} CNPJs ({pendentes} pendentes)")
            messagebox.showinfo("Sucesso", f"Planilha carregada!\n\nTotal: {total}\nPendentes: {pendentes}")

        except Exception as e:
            self.log_message(f"Erro ao carregar planilha: {e}")
            messagebox.showerror("Erro", f"Erro ao carregar planilha:\n{e}")
            self.planilha_carregada = False

    def _get_nome_col(self):
        if self.df is None:
            return None
        for col in ["NOME", "RAZAO", "RAZAO_SOCIAL", "RAZAO SOCIAL", "EMPRESA"]:
            if col in self.df.columns:
                return col
        return None

    def _populate_table(self):
        for item in self.data_table.get_children():
            self.data_table.delete(item)
        if self.df is None:
            return

        nome_col = self._get_nome_col()
        for _, row in self.df.iterrows():
            cod = str(row["COD"])
            cnpj = str(row["CNPJ"])
            nome = str(row[nome_col]) if nome_col else ""
            status = str(row["STATUS"]) if row["STATUS"] else "Pendente"
            self.data_table.insert("", "end", values=(cod, cnpj, nome, status))

        total = len(self.cnpjs)
        baixados = len(self.df[self.df["STATUS"].str.contains("Guia baixada", na=False)])
        pendentes = total - baixados
        self.resumo_var.set(f"Total: {total}  |  Baixados: {baixados}  |  Pendentes: {pendentes}")

    def refresh_table(self):
        self._populate_table()

    # =========================================================================
    # VALIDACAO
    # =========================================================================
    def validate_config(self) -> bool:
        try:
            cfg = self.get_config_from_fields()
            if len(cfg.data_inicial) != 8 or not cfg.data_inicial.isdigit():
                raise ValueError("Data inicial deve ter 8 digitos (DDMMAAAA)")
            if len(cfg.data_final) != 8 or not cfg.data_final.isdigit():
                raise ValueError("Data final deve ter 8 digitos (DDMMAAAA)")
            if not cfg.competencia:
                raise ValueError("Competencia nao pode estar vazia")
            if not self.planilha_carregada:
                raise ValueError("Planilha nao foi carregada! Clique em 'Carregar Dados' primeiro.")
            if not self.cnpjs:
                raise ValueError("Nenhum CNPJ encontrado na planilha!")
            return True
        except ValueError as e:
            messagebox.showerror("Erro de Validacao", str(e))
            return False

    # =========================================================================
    # LOG
    # =========================================================================
    def log_message(self, message: str):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n")
        self.log_text.see("end")

    def clear_log(self):
        self.log_text.delete("1.0", "end")

    def open_manual(self):
        manual_path = self.config.pasta_base / "manual.html"
        if not manual_path.exists():
            messagebox.showerror("Erro", f"Manual nao encontrado:\n{manual_path}")
            return
        try:
            webbrowser.open(f"file://{manual_path.absolute()}")
            self.log_message("Manual aberto no navegador")
        except Exception as e:
            messagebox.showerror("Erro", f"Nao foi possivel abrir o manual:\n{e}")

    def check_log_queue(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
                self.log_text.insert("end", msg + "\n")
                self.log_text.see("end")
            except queue.Empty:
                break
        self.root.after(120, self.check_log_queue)

    # =========================================================================
    # PROGRESSO
    # =========================================================================
    def update_progress(self, message: str, current: int, total: int):
        if total > 0:
            pct = int((current / total) * 100)
            self.progress_pct_var.set(pct)
            self.progress_bar.set(pct / 100)
            self.progress_label.configure(text=f"{current}/{total} | {pct}%")
        self.status_var.set(message)

    # =========================================================================
    # AUTOMACAO
    # =========================================================================
    def start_automation(self):
        if self.running:
            messagebox.showwarning("Aviso", "A automacao ja esta em execucao!")
            return
        if not self.validate_config():
            return

        self.config = self.get_config_from_fields()
        self.running = True
        self.should_stop = False
        self.start_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self.login_btn.configure(state="disabled")
        self.progress_pct_var.set(0)
        self.progress_bar.set(0)
        self.progress_label.configure(text="0/0 | 0%")

        self.status_var.set("Iniciando...")
        self.log_message("Iniciando automacao DCTF...")

        self.worker_thread = threading.Thread(target=self.run_automation, daemon=True)
        self.worker_thread.start()

    def run_automation(self):
        import time

        try:
            config = self.config
            pasta = config.pasta_download
            if not pasta.exists():
                pasta.mkdir(parents=True)

            self.root.after(0, lambda: self.status_var.set("Configurando navegador..."))
            self.log_message("Configurando driver do Chrome...")
            self.driver = configurar_driver(pasta)

            self.root.after(0, lambda: self.status_var.set("Aguardando login manual..."))
            self.root.after(0, lambda: self.login_btn.configure(state="normal"))
            self.log_message("Navegador aberto. Faca o login e clique em 'Confirmar Login'.")

            self.waiting_login = True
            while self.waiting_login and not self.should_stop:
                time.sleep(0.5)

            if self.should_stop:
                raise Exception("Automacao interrompida pelo usuario")

            self.root.after(0, lambda: self.login_btn.configure(state="disabled"))
            self.root.after(0, lambda: self.status_var.set("Processando..."))

            cnpjs = self.cnpjs
            codigos = self.codigos
            df = self.df
            planilha_path = self.planilha_path_var.get()
            total = len(cnpjs)
            self.log_message(f"Iniciando processamento de {total} CNPJs")

            def progress_callback(msg, current, total_count):
                self.root.after(0, lambda: self.update_progress(msg, current, total_count))
                self.root.after(0, self.refresh_table)

            transmissao(
                cnpjs=cnpjs,
                codigos=codigos,
                df=df,
                driver=self.driver,
                competencia=config.competencia,
                pasta_competencia=pasta,
                data_inicial=config.data_inicial,
                data_final=config.data_final,
                timeout_elemento=config.timeout_elemento,
                tentativas_por_cnpj=config.tentativas_por_cnpj,
                callback=progress_callback,
                should_stop=lambda: self.should_stop,
                planilha_path=planilha_path,
            )

            df.to_excel(planilha_path, index=False)
            self.root.after(0, self.refresh_table)
            self.root.after(0, lambda: self.status_var.set("Concluido!"))
            self.log_message("Automacao concluida com sucesso!")
            self.root.after(0, lambda: messagebox.showinfo("Sucesso", "Automacao concluida!"))

        except Exception as e:
            msg = str(e)
            self.log_message(f"Erro: {msg}")
            self.root.after(0, lambda: self.status_var.set(f"Erro: {msg[:60]}..."))
            self.root.after(0, lambda: messagebox.showerror("Erro", f"Ocorreu um erro:\n\n{msg}"))

        finally:
            if self.driver:
                try:
                    self.driver.quit()
                except Exception:
                    pass
                self.driver = None
            self.running = False
            self.root.after(0, lambda: self.start_btn.configure(state="normal"))
            self.root.after(0, lambda: self.stop_btn.configure(state="disabled"))
            self.root.after(0, lambda: self.login_btn.configure(state="disabled"))

    def confirm_login(self):
        self.waiting_login = False
        self.log_message("Login confirmado. Iniciando processamento...")

    def stop_automation(self):
        if not self.running:
            return
        if messagebox.askyesno("Confirmar", "Deseja realmente parar a automacao?"):
            self.should_stop = True
            self.waiting_login = False
            self.status_var.set("Parando...")
            self.log_message("Solicitacao de parada enviada...")

    def on_closing(self):
        if self.running:
            if not messagebox.askyesno("Confirmar", "A automacao esta em execucao. Deseja sair?"):
                return
            self.should_stop = True
            self.waiting_login = False
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.root.destroy()

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()


def run_gui():
    app = AutomacaoDCTFApp()
    app.run()
