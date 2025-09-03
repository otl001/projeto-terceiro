import customtkinter as ctk
from tkinter import messagebox, ttk
import sqlite3
from datetime import datetime, date
import csv
import os
import configparser
import tkinter as tk
from tkinter import ttk
from datetime import datetime, date, timedelta

def corrigir_status_existentes():
    # print("DEBUG: Corrigindo status existentes no banco...")
    db = DatabaseManager(DB_NAME)
    
    # Atualizar status antigos para os novos
    db.execute("UPDATE movimentacoes SET status = 'Em terceiro' WHERE status = 'em uso'")
    db.execute("UPDATE movimentacoes SET status = 'Atrasado' WHERE status = 'atrasado'")
    db.execute("UPDATE movimentacoes SET status = 'Devolvido' WHERE status = 'devolvido'")
    
    # print("DEBUG: Status corrigidos")


def apply_ttk_styles():
    style = ttk.Style()
    style.theme_use("default")

    # Detecta o tema atual do customtkinter
    if ctk.get_appearance_mode() == "Dark":
        bg_color = "#2B2B2B"
        fg_color = "#FFFFFF"
        alt_row = "#333333"
        select_bg = "#444444"
    else:
        bg_color = "#FFFFFF"
        fg_color = "#000000"
        alt_row = "#F0F0F0"
        select_bg = "#D9D9D9"

    # Treeview geral
    style.configure(
        "Treeview",
        background=bg_color,
        foreground=fg_color,
        fieldbackground=bg_color,
        borderwidth=0,
        relief="flat"
    )

    style.map("Treeview", background=[("selected", select_bg)])

    # Cabeçalho do Treeview
    style.configure(
        "Treeview.Heading",
        background=bg_color,
        foreground=fg_color,
        borderwidth=0
    )

    # Atualiza cores alternadas nas linhas
    for tree in style.master.winfo_children():
        if isinstance(tree, ttk.Treeview):
            for item in tree.get_children():
                tree.tag_configure("oddrow", background=alt_row)
                tree.tag_configure("evenrow", background=bg_color)


DB_NAME = "controle_terceiros.db"
EXPORT_DIR = "exports"
THEME_FILE= "theme.ini"


def verificar_movimentacoes_detalhadas():
    print("\n=== DETALHES DAS MOVIMENTAÇÕES ===")
    try:
        db = DatabaseManager(DB_NAME)
        movimentacoes = db.fetch_all("SELECT * FROM movimentacoes")
        for m in movimentacoes:
            print(f"Mov ID: {m['id_mov']}, Equip: {m['id_equipamento']}, Local: {m['id_local']}")
            print(f"  Data envio: {m['data_envio']}, Data prevista: {m['data_prevista_retorno']}")
            print(f"  Data retorno: {m['data_retorno']}, Status: '{m['status']}'")
            print(f"  Observações: '{m['observacoes']}'")
            print("---")
    except Exception as e:
        print(f"Erro ao ver movimentações detalhadas: {e}")

# -----------------------------
# Camada de Abstração do Banco de Dados
# -----------------------------
class DatabaseManager:
    def __init__(self, db_name):
        self.db_name = db_name
        self.conn = None

    def __enter__(self):
        try:
            self.conn = sqlite3.connect(self.db_name)
            self.conn.row_factory = sqlite3.Row
            # print(f"DEBUG: Conectado ao banco {self.db_name}")
            return self.conn
        except Exception as e:
            # print(f"DEBUG: Erro ao conectar com banco: {e}")
            raise

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.conn:
            self.conn.close()
            # print("DEBUG: Conexão com banco fechada")

    def execute(self, query, params=()):
        with self as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            # print(f"DEBUG: Query executada: {query} com params: {params}")
            return cursor

    def fetch_all(self, query, params=()):
        with self as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            results = cursor.fetchall()
            # print(f"DEBUG: Fetch all: {query} retornou {len(results)} resultados")
            return results

    def fetch_one(self, query, params=()):
        with self as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            result = cursor.fetchone()
            # print(f"DEBUG: Fetch one: {query} retornou {result}")
            return result

# def verificar_dados():
#     print("\n=== VERIFICAÇÃO DE DADOS ===")
#     try:
#         db = DatabaseManager(DB_NAME)
        
#         # Verificar equipamentos
#         try:
#             equipamentos = db.fetch_all("SELECT * FROM equipamentos")
#             print(f"Equipamentos: {len(equipamentos)}")
#             for e in equipamentos:
#                 print(f"  - {e['tipo']} {e['modelo']} (ID: {e['id_equipamento']})")
#         except Exception as e:
#             print(f"Erro ao ver equipamentos: {e}")
        
#         # Verificar locais
#         try:
#             locais = db.fetch_all("SELECT * FROM locais")
#             print(f"Locais: {len(locais)}")
#             for l in locais:
#                 print(f"  - {l['nome_local']} (ID: {l['id_local']})")
#         except Exception as e:
#             print(f"Erro ao ver locais: {e}")
        
#         # Verificar movimentações
#         try:
#             movimentacoes = db.fetch_all("SELECT * FROM movimentacoes")
#             print(f"Movimentações: {len(movimentacoes)}")
#             for m in movimentacoes:
#                 print(f"  - Mov ID: {m['id_mov']}, Equip: {m['id_equipamento']}, Local: {m['id_local']}")
#         except Exception as e:
#             print(f"Erro ao ver movimentações: {e}")
            
#     except Exception as e:
#         print(f"Erro geral ao verificar dados: {e}")

# -----------------------------
# Banco de Dados
# -----------------------------

def init_db():
    # print("DEBUG: Inicializando banco de dados...")
    os.makedirs(EXPORT_DIR, exist_ok=True)
    
    #Primeiro verificar se o banco existe e quais tabelas tem
    try:
        db = DatabaseManager(DB_NAME)
        tables = db.fetch_all("SELECT name FROM sqlite_master WHERE type='table'")
        table_names = [t['name'] for t in tables]
        print(f"DEBUG: Tabelas existentes: {table_names}")
    except Exception as e:
        print(f"DEBUG: Erro ao verificar tabelas: {e}")
        table_names = []

    #Criar tabelas se não existirem
    db = DatabaseManager(DB_NAME)
    
    # Verificar e criar tabela locais
    if 'locais' not in table_names:
        print("DEBUG: Criando tabela locais")
        db.execute(
            """
            CREATE TABLE locais (
                id_local INTEGER PRIMARY KEY,
                nome_local TEXT NOT NULL,
                endereco TEXT,
                contato TEXT
            )
            """
        )
    else:
        print("DEBUG: Tabela 'locais' já existe")
    
    # Verificar e criar tabela equipamentos
    if 'equipamentos' not in table_names:
        print("DEBUG: Criando tabela equipamentos")
        db.execute(
            """
            CREATE TABLE equipamentos (
                id_equipamento INTEGER PRIMARY KEY,
                tipo TEXT NOT NULL,
                modelo TEXT,
                numero_serie TEXT UNIQUE,
                numero_os TEXT,
                condicao TEXT
            )
            """
        )
    else:
        print("DEBUG: Tabela 'equipamentos' já existe")
    
    # Verificar e criar tabela movimentacoes
    if 'movimentacoes' not in table_names:
        print("DEBUG: Criando tabela movimentacoes")
        db.execute(
            """
            CREATE TABLE movimentacoes (
                id_mov INTEGER PRIMARY KEY,
                id_equipamento INTEGER NOT NULL,
                id_local INTEGER NOT NULL,
                data_envio TEXT NOT NULL,
                data_prevista_retorno TEXT,
                data_retorno TEXT,
                status TEXT NOT NULL,
                observacoes TEXT,
                FOREIGN KEY(id_equipamento) REFERENCES equipamentos(id_equipamento),
                FOREIGN KEY(id_local) REFERENCES locais(id_local)
            )
            """
        )
    else:
        print("DEBUG: Tabela 'movimentacoes' já existe")
    
    # Criar índices se não existirem
    try:
        db.execute("CREATE INDEX IF NOT EXISTS idx_mov_equip ON movimentacoes(id_equipamento)")
        db.execute("CREATE INDEX IF NOT EXISTS idx_mov_status ON movimentacoes(status)")
        print("DEBUG: Índices verificados/criados")
    except Exception as e:
        print(f"DEBUG: Erro ao criar índices: {e}")
    
    print("DEBUG: Banco de dados inicializado com sucesso")


# -----------------------------
# Utilitários
# -----------------------------

def today_str():
    return date.today().strftime("%Y-%m-%d")


def parse_date(s: str):
    if not s:
        return None
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except ValueError:
        return None

def load_theme_pref():
    """Lê o último tema salvo ('dark' ou 'light'). Padrão: 'dark'."""
    cfg = configparser.ConfigParser()
    if os.path.exists(THEME_FILE):
        cfg.read(THEME_FILE, encoding='utf-8')
        return cfg.get("ui", "appearance", fallback="dark").lower()
    return "dark"

def save_theme_pref(mode: str):
    """Salva o tema atual ('dark' ou 'light')."""
    cfg = configparser.ConfigParser()
    cfg["ui"] = {"appearance": mode.lower()}
    with open(THEME_FILE, "w", encoding='utf-8') as f:
        cfg.write(f)

# -----------------------------
# Camada de Dados
# -----------------------------

def add_local(nome, endereco, contato):
    if not nome.strip():
        raise ValueError("Nome do local é obrigatório.")
    if not contato.strip():
        raise ValueError("Contato do local é obrigatório.")
    db = DatabaseManager(DB_NAME)
    db.execute(
        "INSERT INTO locais (nome_local, endereco, contato) VALUES (?, ?, ?)",
        (nome.strip(), endereco.strip(), contato.strip()),
    )


def listar_locais():
    db = DatabaseManager(DB_NAME)
    return db.fetch_all("SELECT id_local, nome_local, endereco, contato FROM locais ORDER BY nome_local")


def add_equipamento(tipo, modelo, numero_serie, numero_os, condicao):
    if not tipo.strip():
        raise ValueError("Tipo é obrigatório.")
    numero_serie = numero_serie.strip() if numero_serie else None
    numero_os = numero_os.strip() if numero_os else None

    db = DatabaseManager(DB_NAME)
    # Validação de unicidade para numero_serie
    if numero_serie:
        existing_equip = db.fetch_one("SELECT id_equipamento FROM equipamentos WHERE numero_serie = ?", (numero_serie,))
        if existing_equip:
            raise ValueError(f"Número de série '{numero_serie}' já existe.")

    db.execute(
        "INSERT INTO equipamentos (tipo, modelo, numero_serie, numero_os, condicao) VALUES (?, ?, ?, ?, ?)",
        (tipo.strip(), (modelo or "").strip(), numero_serie, numero_os, (condicao or "").strip()),
    )


def listar_equipamentos():
    db = DatabaseManager(DB_NAME)
    return db.fetch_all(
        "SELECT id_equipamento, tipo, modelo, numero_serie, numero_os, condicao FROM equipamentos ORDER BY tipo, modelo"
    )


def listar_equipamentos_em_movimento():
    """Retorna equipamentos que possuem movimentação aberta (data_retorno IS NULL)."""
    db = DatabaseManager(DB_NAME)
    return db.fetch_all(
        """
        SELECT DISTINCT e.id_equipamento, e.tipo, e.modelo, e.numero_serie, e.numero_os, e.condicao
        FROM equipamentos e
        JOIN movimentacoes m ON m.id_equipamento = e.id_equipamento
        WHERE m.data_retorno IS NULL
        ORDER BY e.tipo, e.modelo
        """
    )


def equipamentos_disponiveis(apenas_nao_movimentados: bool = False):
    """
    Retorna equipamentos disponíveis.
    - Se apenas_nao_movimentados == True: retorna apenas equipamentos que NUNCA tiveram movimentação.
    - Se False (padrão): retorna equipamentos que não estão em movimentação aberta.
    """
    db = DatabaseManager(DB_NAME)
    if apenas_nao_movimentados:
        return db.fetch_all(
            """
            SELECT e.id_equipamento, e.tipo, e.modelo, e.numero_serie, e.numero_os
            FROM equipamentos e
            WHERE e.id_equipamento NOT IN (
                SELECT m.id_equipamento FROM movimentacoes m
            )
            ORDER BY e.tipo, e.modelo
            """
        )
    else:
        return db.fetch_all(
            """
            SELECT e.id_equipamento, e.tipo, e.modelo, e.numero_serie, e.numero_os
            FROM equipamentos e
            WHERE e.id_equipamento NOT IN (
                SELECT m.id_equipamento FROM movimentacoes m 
                WHERE m.status IN ('Em terceiro','Atrasado') AND m.data_retorno IS NULL
            )
            ORDER BY e.tipo, e.modelo
            """
        )


def registrar_envio(id_equipamento, id_local, data_envio, data_prevista, observacoes):
    if not id_equipamento or not id_local:
        raise ValueError("Selecione equipamento e local.")
    d_envio = parse_date(data_envio)
    if not d_envio:
        raise ValueError("Data de envio inválida. Use AAAA-MM-DD.")
    d_prev = parse_date(data_prevista)
    status = "Em terceiro"  # ← CORRIGIDO
    if d_prev and date.today() > d_prev:
        status = "Atrasado"  # ← CORRIGIDO
    db = DatabaseManager(DB_NAME)
    db.execute(
        """
        INSERT INTO movimentacoes
        (id_equipamento, id_local, data_envio, data_prevista_retorno, data_retorno, status, observacoes)
        VALUES (?, ?, ?, ?, NULL, ?, ?)
        """,
        (id_equipamento, id_local, data_envio, data_prevista, status, (observacoes or "").strip()),
    )


def atualizar_status_atrasados():
    db = DatabaseManager(DB_NAME)
    hoje = today_str()
    db.execute(
        """
        UPDATE movimentacoes
        SET status = 'Atrasado'
        WHERE status = 'Em terceiro'
          AND data_retorno IS NULL
          AND data_prevista_retorno IS NOT NULL
          AND date(data_prevista_retorno) < date(?)
        """,
        (hoje,),
    )


def listar_mov_abertas():
    atualizar_status_atrasados()
    db = DatabaseManager(DB_NAME)
    return db.fetch_all(
        """
        SELECT m.id_mov,
               e.id_equipamento,
               e.tipo,
               e.modelo,
               e.numero_serie,
               e.numero_os,
               l.nome_local,
               m.data_envio,
               m.data_prevista_retorno,
               m.status,
               m.observacoes
        FROM movimentacoes m
        JOIN equipamentos e ON e.id_equipamento = m.id_equipamento
        JOIN locais l ON l.id_local = m.id_local
        WHERE m.data_retorno IS NULL AND m.status IN ('Em terceiro', 'Atrasado')
        ORDER BY m.status DESC, m.data_envio DESC
        """
    )


def registrar_retorno(id_mov, data_retorno=None):
    data_retorno = data_retorno or today_str()
    if not parse_date(data_retorno):
        raise ValueError("Data de retorno inválida. Use AAAA-MM-DD.")
    db = DatabaseManager(DB_NAME)
    db.execute(
        "UPDATE movimentacoes SET data_retorno = ?, status = 'Retornado' WHERE id_mov = ?",
        (data_retorno, id_mov),
    )

def marcar_como_devolvido(id_mov):
    """Marca uma movimentação como devolvida (processo finalizado)"""
    db = DatabaseManager(DB_NAME)
    db.execute(
        "UPDATE movimentacoes SET status = 'Devolvido' WHERE id_mov = ?",
        (id_mov,),
    )


def listar_retornados():
    """Retorna equipamentos com status Retornado"""
    db = DatabaseManager(DB_NAME)
    return db.fetch_all(
        """
        SELECT m.id_mov,
               e.tipo,
               e.modelo,
               e.numero_serie,
               e.numero_os,
               l.nome_local,
               m.data_envio,
               m.data_retorno,
               m.status,
               m.observacoes
        FROM movimentacoes m
        JOIN equipamentos e ON e.id_equipamento = m.id_equipamento
        JOIN locais l ON l.id_local = m.id_local
        WHERE m.status = 'Retornado'
        ORDER BY m.data_retorno DESC
        """
    )

def filtro_relatorio(id_local=None, tipo=None, status=None):
    atualizar_status_atrasados()
    db = DatabaseManager(DB_NAME)
    query = [
        "SELECT m.id_mov, e.tipo, e.modelo, e.numero_serie, e.numero_os, l.nome_local, m.data_envio, m.data_prevista_retorno, m.data_retorno, m.status, m.observacoes",
        "FROM movimentacoes m",
        "JOIN equipamentos e ON e.id_equipamento = m.id_equipamento",
        "JOIN locais l ON l.id_local = m.id_local",
        "WHERE 1=1",
    ]
    params = []
    if id_local:
        query.append("AND l.id_local = ?")
        params.append(id_local)
    if tipo and tipo.strip():
        query.append("AND e.tipo = ?")
        params.append(tipo.strip())
    if status and status.strip():
        query.append("AND m.status = ?")
        params.append(status.strip())
    query.append("ORDER BY m.data_envio DESC")
    return db.fetch_all("\n".join(query), params)


def exportar_csv(rows, filepath):
    header = [
        "ID Mov",
        "Tipo",
        "Modelo",
        "Número de Série",
        "Número OS",
        "Local",
        "Data Envio",
        "Prev. Retorno",
        "Data Retorno",
        "Status",
        "Observações",
    ]
    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        for r in rows:
            writer.writerow(r)


# -----------------------------
# Funções para exclusão
# -----------------------------

def has_movimentacoes_equip(id_equip):
    db = DatabaseManager(DB_NAME)
    cnt = db.fetch_one("SELECT COUNT(1) FROM movimentacoes WHERE id_equipamento = ?", (id_equip,))
    return cnt[0] if cnt else 0


def delete_equipamento_db(id_equip, delete_movs=False):
    db = DatabaseManager(DB_NAME)
    if delete_movs:
        db.execute("DELETE FROM movimentacoes WHERE id_equipamento = ?", (id_equip,))
    else:
        # se existir movimentação, previne a exclusão
        if db.fetch_one("SELECT 1 FROM movimentacoes WHERE id_equipamento = ? LIMIT 1", (id_equip,)):
            raise ValueError("Existem movimentações associadas a este equipamento. Use a opção para excluir também as movimentações.")
    db.execute("DELETE FROM equipamentos WHERE id_equipamento = ?", (id_equip,))


def has_movimentacoes_local(id_local):
    db = DatabaseManager(DB_NAME)
    cnt = db.fetch_one("SELECT COUNT(1) FROM movimentacoes WHERE id_local = ?", (id_local,))
    return cnt[0] if cnt else 0

def delete_local_db(id_local):
    db = DatabaseManager(DB_NAME)
    if db.fetch_one("SELECT 1 FROM movimentacoes WHERE id_local = ? LIMIT 1", (id_local,)):
        raise ValueError("Existem movimentações associadas a este local. Não é possível excluí-lo diretamente.")
    db.execute("DELETE FROM locais WHERE id_local = ?", (id_local,))


# -----------------------------
# Interface (melhorada - dark total)
# -----------------------------

TIPOS_PADRAO = ["Rádio", "Repetidora", "Duplexador", "Wattímetro", "Fonte", "Antena", "Outro"]
CONDICOES = ["Funcionando", "Manutenção", "Danificado", "Desconhecida"]
STATUS_OPCOES = ["Em terceiro", "Atrasado", "Devolvido", "Retornado"]  # ← CORRIGIDO


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Controle de Equipamentos em Terceiros")
        self.geometry("1250x780")
        ctk.set_default_color_theme("dark-blue")
        ctk.set_appearance_mode(load_theme_pref())

        # Top bar
        top = ctk.CTkFrame(self, corner_radius=0, fg_color="#071028")
        top.pack(fill="x")
        self.logo = self._create_ctk_label(top, text="📡 TERCEIROS", font=("Segoe UI", 20, "bold"))
        self.logo.pack(side="left", padx=(16, 8), pady=12)
        self.search_var = ctk.StringVar()
        self.search = self._create_ctk_entry(top, placeholder_text="Pesquisar por modelo / SN / O.S...", width=420, textvariable=self.search_var)
        self.search.pack(side="left", padx=8)
        self._create_ctk_button(top, text="Pesquisar", command=self.on_search).pack(side="left", padx=8)
        self._create_ctk_button(top, text="Alternar Tema", command=self.toggle_theme, fg_color="#2b2f38", hover_color="#3a3f47").pack(side="right", padx=12)

        # ===== Main content (stats + tabs) =====
        main = ctk.CTkFrame(self, fg_color="#0b1220")
        main.pack(expand=True, fill="both", padx=12, pady=12)

        # ---- Stats strip ----
        stats = ctk.CTkFrame(main, fg_color="#0b1220")
        stats.pack(fill="x", pady=(4, 8))

        self.lbl_total_equip   = ctk.CTkLabel(stats, text="Equipamentos: 0", text_color="#e6eef8")
        self.lbl_disponiveis   = ctk.CTkLabel(stats, text="Disponíveis: 0",  text_color="#a7f3d0")
        self.lbl_em_uso        = ctk.CTkLabel(stats, text="Em terceiro: 0",  text_color="#93c5fd")
        self.lbl_atrasados     = ctk.CTkLabel(stats, text="Atrasados: 0",    text_color="#fecaca")

        self.lbl_total_equip.pack(side="left", padx=(4,12))
        self.lbl_disponiveis.pack(side="left", padx=12)
        self.lbl_em_uso.pack(side="left", padx=12)
        self.lbl_atrasados.pack(side="left", padx=12)

        # ---- Tabview principal ----
        # CORREÇÃO: Adicionar a nova aba "Retornados"
        self.tabs = ctk.CTkTabview(main, segmented_button_selected_color="#1f6feb")
        for name in ["Equipamentos", "Locais", "Enviar/Movimentar", "Devolução", "Retornados", "Relatórios"]:
            self.tabs.add(name)
        self.tabs.pack(expand=True, fill="both")

        # Constrói cada aba
        self.build_tab_equip()
        self.build_tab_locais()
        self.build_tab_envio()
        self.build_tab_devolucao()
        self.build_tab_retornados() 
        self.build_tab_relatorios()

        # Estiliza e carrega dados
        self.apply_ttk_styles()
        self.refresh_all()


    # -----------------------------
    # Métodos auxiliares
    # -----------------------------
    def _create_ctk_label(self, parent, text, text_color="#e6eef8", **kwargs):
        return ctk.CTkLabel(parent, text=text, text_color=text_color, **kwargs)

    def _create_ctk_entry(self, parent, textvariable=None, placeholder_text=None, **kwargs):
        return ctk.CTkEntry(parent, textvariable=textvariable, placeholder_text=placeholder_text,
                            fg_color="#0b1220", text_color="#e6eef8", **kwargs)

    def _create_ctk_button(self, parent, text, command, fg_color="#1f6feb", hover_color="#1554b6", **kwargs):
        return ctk.CTkButton(parent, text=text, command=command, fg_color=fg_color, hover_color=hover_color, **kwargs)

    def _all_treeviews(self):
        """Retorna todas as Treeviews existentes (se já criadas)."""
        names = ("tree_equip", "tree_locais", "tree_mov", "tree_devol", "tree_rel")
        return [getattr(self, n) for n in names if hasattr(self, n)]

    def apply_ttk_styles(self):
        """Aplica/atualiza estilos do ttk Treeview de acordo com o tema atual do customtkinter."""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        dark = ctk.get_appearance_mode().lower() == "dark"
        if dark:
            bg = "#0b1220"; fg = "#e6eef8"; head_bg = "#091025"
            select_bg = "#2541b2"; select_fg = "#ffffff"; odd = "#0f1724"; even = "#071025"
        else:
            bg = "#ffffff"; fg = "#111827"; head_bg = "#e5e7eb"
            select_bg = "#93c5fd"; select_fg = "#111827"; odd = "#f9fafb"; even = "#ffffff"

        style.configure("Treeview", background=bg, foreground=fg,
                        fieldbackground=bg, rowheight=26, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", background=head_bg, foreground=fg,
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview", background=[("selected", select_bg)], foreground=[("selected", select_fg)])

        for tv in self._all_treeviews():
            try:
                tv.tag_configure("oddrow", background=odd)
                tv.tag_configure("evenrow", background=even)
                for item in tv.get_children():
                    tags = tv.item(item).get("tags", ())
                    if tags:
                        tv.item(item, tags=tags)
            except Exception:
                pass

    def toggle_theme(self):
        if ctk.get_appearance_mode() == "Dark":
            ctk.set_appearance_mode("Light")
        else:
            ctk.set_appearance_mode("Dark")
        self.apply_ttk_styles()


    def on_search(self):
        q = self.search_var.get().strip()
        if not q:
            self.refresh_equip_list()
            return

        db = DatabaseManager(DB_NAME)
        q_lower = f"%{q.lower()}%"
        rows = db.fetch_all(
            """
            SELECT id_equipamento, tipo, modelo, numero_serie, numero_os, condicao
            FROM equipamentos
            WHERE LOWER(tipo) LIKE ? OR LOWER(modelo) LIKE ? OR LOWER(numero_serie) LIKE ? OR LOWER(numero_os) LIKE ?
            ORDER BY tipo, modelo
            """,
            (q_lower, q_lower, q_lower, q_lower),
        )

        for iid in self.tree_equip.get_children():
            self.tree_equip.delete(iid)

        for idx, e in enumerate(rows):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.tree_equip.insert("", "end", iid=str(e["id_equipamento"]),
                                   values=(e["tipo"], e["modelo"] or "-", e["numero_serie"] or "-",
                                    e["numero_os"] or "-", (e["condicao"] if "condicao" in e.keys() else "-") or "-"),
                                   tags=(tag,))

    # -----------------------------
    # ✅ Também estavam fora e movi para dentro
    # -----------------------------
    def refresh_stats(self):
        try:
            # Total de equipamentos
            all_eq = listar_equipamentos()
            total = len(all_eq)
            
            # Equipamentos disponíveis
            disp = len(equipamentos_disponiveis())
            
            # Movimentações abertas
            abertas = listar_mov_abertas()
            
            # Contar por status
            em_terceiro = sum(1 for m in abertas if m["status"] == 'Em terceiro')
            atrasados = sum(1 for m in abertas if m["status"] == 'Atrasado')
            
            # Atualizar os labels
            self.lbl_total_equip.configure(text=f"Equipamentos: {total}")
            self.lbl_disponiveis.configure(text=f"Disponíveis: {disp}")
            self.lbl_em_uso.configure(text=f"Em terceiro: {em_terceiro}")
            self.lbl_atrasados.configure(text=f"Atrasados: {atrasados}")
        
        except Exception as e:
            print(f"Erro ao atualizar stats: {e}")

    def refresh_all(self):
        self.refresh_stats()
        self.refresh_equip_list()
        self.refresh_locais_list()
        self.refresh_envio_combos()
        self.refresh_devolucao()
        self.refresh_retornados()  # ← Nova linha
        self.refresh_relatorios_combos()


    # -----------------------------
    # Tab Equipamentos
    # -----------------------------
    def build_tab_equip(self):
        tab = self.tabs.tab("Equipamentos")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        form = ctk.CTkFrame(frame, fg_color="#071025")
        form.pack(fill="x", padx=6, pady=6)

        self.tipo_var = ctk.StringVar(value=TIPOS_PADRAO[0])
        self.modelo_var = ctk.StringVar()
        self.serial_var = ctk.StringVar()
        self.os_var = ctk.StringVar()
        self.cond_var = ctk.StringVar(value=CONDICOES[0])

        # por padrão, vai mostrar os equipamentos que estão em movimento e os nunca movimentados;  se marcado coloca os retornados0
        self.show_all_var = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(
            form,
            text = "Retornados",
            variable= self.show_all_var,
            command = self.refresh_equip_list
        ).grid (row=0, column=0, sticky="w", padx=6, pady=6)


        self._create_ctk_label(form, text="Tipo:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ctk.CTkOptionMenu(form, values=TIPOS_PADRAO, variable=self.tipo_var, width=180).grid(row=1, column=1, sticky="w", padx=6, pady=6)
        self._create_ctk_label(form, text="Modelo:").grid(row=1, column=2, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.modelo_var).grid(row=1, column=3, sticky="we", padx=6, pady=6)

        self._create_ctk_label(form, text="Serial:").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.serial_var).grid(row=2, column=1, sticky="we", padx=6, pady=6)
        self._create_ctk_label(form, text="O.S:").grid(row=2, column=2, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.os_var).grid(row=2, column=3, sticky="we", padx=6, pady=6)

        self._create_ctk_label(form, text="Condição:").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        ctk.CTkOptionMenu(form, values=CONDICOES, variable=self.cond_var, width=180).grid(row=3, column=1, sticky="w", padx=6, pady=6)

        # Add and Delete buttons
        btn_frame = ctk.CTkFrame(form, fg_color="#071025")
        btn_frame.grid(row=3, column=3, sticky="e", padx=6, pady=6)
        self._create_ctk_button(btn_frame, text="Adicionar", command=self.on_add_equip, width=110).pack(side="left", padx=(0,6))
        self._create_ctk_button(btn_frame, text="Excluir", command=self.on_delete_equip, width=110, fg_color="#b43b3b", hover_color="#912b2b").pack(side="left")

        form.grid_columnconfigure(3, weight=1)

        # list
        self.tree_equip = ttk.Treeview(frame, columns=("Tipo", "Modelo", "SN", "OS", "Cond"), show="headings", height=14)
        for h in ("Tipo", "Modelo", "SN", "OS", "Cond"):
            self.tree_equip.heading(h, text=h)
            self.tree_equip.column(h, width=160, anchor="center")
        self.tree_equip.pack(fill="both", expand=True, padx=6, pady=(8,6))

        # style alternating rows
        self.tree_equip.tag_configure("oddrow", background="#0f1724")
        self.tree_equip.tag_configure("evenrow", background="#071025")

    def refresh_equip_list(self):
        if not hasattr(self, "tree_equip") or self.tree_equip is None:
            return

        if self.show_all_var.get():
            # Mostrar todos (inclui retornados)
            rows = listar_equipamentos()
        else:
            # Mostrar apenas: em movimento + nunca movimentados
            em_mov = {e["id_equipamento"] for e in listar_equipamentos_em_movimento()}
            todos = {e["id_equipamento"]: e for e in listar_equipamentos()}
            disponiveis = {e["id_equipamento"]: e for e in equipamentos_disponiveis(apenas_nao_movimentados=True)}

            # Combina em movimento + nunca movimentados
            rows = []
            for id_eq in em_mov.union(disponiveis.keys()):
                if id_eq in todos:
                    rows.append(todos[id_eq])
                else:
                    rows.append(disponiveis[id_eq])

            # Ordena por tipo/modelo
            rows.sort(key=lambda x: (x["tipo"], x["modelo"] or ""))
        
        # Limpa e repopula a tree
        for i in self.tree_equip.get_children():
            self.tree_equip.delete(i)

        for idx, e in enumerate(rows):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.tree_equip.insert(
                "", "end", iid=str(e["id_equipamento"]),
                values=(
                    e["tipo"] or "-",
                    e["modelo"] or "-",
                    e["numero_serie"] or "-",
                    e["numero_os"] or "-",  # ← CORREÇÃO: Adicionar este campo
                    (e["condicao"] if "condicao" in e.keys() else "-") or "-"
                ),
                tags=(tag,)
            )
        
    def on_add_equip(self):
        try:
            add_equipamento(self.tipo_var.get(), self.modelo_var.get(), self.serial_var.get(), self.os_var.get(), self.cond_var.get())
            messagebox.showinfo("Sucesso", "Equipamento adicionado!")
            self.modelo_var.set("")
            self.serial_var.set("")
            self.os_var.set("")
            self.refresh_all()
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    def on_delete_equip(self):
        sel = self.tree_equip.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione um equipamento para excluir.")
            return
        try:
            id_equip = int(sel[0])
        except Exception:
            messagebox.showerror("Erro", "ID do equipamento inválido.")
            return

        cnt = has_movimentacoes_equip(id_equip)
        if cnt > 0:
            # existe movimentações: perguntar se quer excluir também as movimentações
            confirmado = messagebox.askyesno(
                "Confirmar exclusão",
                f"O equipamento selecionado possui {cnt} movimentação(ões).\n"
                "Deseja excluir o equipamento **e** todas as movimentações associadas? (A ação é irreversível)"
            )
            if not confirmado:
                return
            try:
                delete_equipamento_db(id_equip, delete_movs=True)
                messagebox.showinfo("Sucesso", "Equipamento e movimentações excluídos.")
                self.refresh_all()
            except Exception as ex:
                messagebox.showerror("Erro", str(ex))
        else:
            # sem movimentações: exclusão simples
            confirmado = messagebox.askyesno("Confirmar exclusão", "Deseja excluir este equipamento? Esta ação é irreversível.")
            if not confirmado:
                return
            try:
                delete_equipamento_db(id_equip, delete_movs=False)
                messagebox.showinfo("Sucesso", "Equipamento excluído.")
                self.refresh_all()
            except Exception as ex:
                messagebox.showerror("Erro", str(ex))

    # -----------------------------
    # Tab Locais
    # -----------------------------
    def build_tab_locais(self):
        tab = self.tabs.tab("Locais")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        form = ctk.CTkFrame(frame, fg_color="#071025")
        form.pack(fill="x")

        self.local_nome = ctk.StringVar()
        self.local_end = ctk.StringVar()
        self.local_contato = ctk.StringVar()

        self._create_ctk_label(form, text="Nome:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.local_nome).grid(row=0, column=1, sticky="we", padx=6, pady=6)
        self._create_ctk_label(form, text="Endereço:").grid(row=0, column=2, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.local_end).grid(row=0, column=3, sticky="we", padx=6, pady=6)
        self._create_ctk_label(form, text="Contato:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.local_contato).grid(row=1, column=1, sticky="we", padx=6, pady=6)
        
        btn_frame = ctk.CTkFrame(form, fg_color="#071025")
        btn_frame.grid(row=1, column=3, sticky="e", padx=6, pady=6)
        self._create_ctk_button(btn_frame, text="Adicionar Local", command=self.on_add_local).pack(side="left", padx=(0,6))
        self._create_ctk_button(btn_frame, text="Excluir Local", command=self.on_delete_local, fg_color="#b43b3b", hover_color="#912b2b").pack(side="left")

        form.grid_columnconfigure(1, weight=1)
        form.grid_columnconfigure(3, weight=1)

        self.tree_locais = ttk.Treeview(frame, columns=("Nome", "Endereço", "Contato"), show="headings", height=12)
        for h in ("Nome", "Endereço", "Contato"):
            self.tree_locais.heading(h, text=h)
            self.tree_locais.column(h, width=260, anchor="w")
        self.tree_locais.pack(fill="both", expand=True, pady=(8,6))

        self.tree_locais.tag_configure("oddrow", background="#0f1724")
        self.tree_locais.tag_configure("evenrow", background="#071025")

    def refresh_locais_list(self):
        for i in self.tree_locais.get_children():
            self.tree_locais.delete(i)
        for idx, l in enumerate(listar_locais()):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            self.tree_locais.insert("", "end", iid=str(l["id_local"]), values=(l["nome_local"], l["endereco"] or "-", l["contato"] or "-"), tags=(tag,))

    def on_add_local(self):
        try:
            add_local(self.local_nome.get(), self.local_end.get(), self.local_contato.get())
            messagebox.showinfo("Sucesso", "Local adicionado!")
            self.local_nome.set("")
            self.local_end.set("")
            self.local_contato.set("")
            self.refresh_all()
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    def on_delete_local(self):
        sel = self.tree_locais.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione um local para excluir.")
            return
        try:
            id_local = int(sel[0])
        except Exception:
            messagebox.showerror("Erro", "ID do local inválido.")
            return

        cnt = has_movimentacoes_local(id_local)
        if cnt > 0:
            messagebox.showwarning("Atenção", f"Existem {cnt} movimentação(ões) associadas a este local. Não é possível excluí-lo diretamente.")
            return
        
        confirmado = messagebox.askyesno("Confirmar exclusão", "Deseja excluir este local? Esta ação é irreversível.")
        if not confirmado:
            return
        try:
            delete_local_db(id_local)
            messagebox.showinfo("Sucesso", "Local excluído!")
            self.refresh_all()
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    # -----------------------------
    # Tab Enviar/Movimentar
    # -----------------------------
    def build_tab_envio(self):
        tab = self.tabs.tab("Enviar/Movimentar")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        form = ctk.CTkFrame(frame, fg_color="#071025")
        form.pack(fill="x")

        self.envio_equip_var = ctk.StringVar()
        self.envio_local_var = ctk.StringVar()
        self.envio_data_envio_var = ctk.StringVar(value=today_str())
        self.envio_data_prev_var = ctk.StringVar()
        self.envio_obs_var = ctk.StringVar()
        
        # Adicionar trace para preenchimento automático
        self.envio_data_envio_var.trace_add("write", self.auto_calcular_previsao)

        self._create_ctk_label(form, text="Equipamento:").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.envio_equip_combo = ctk.CTkOptionMenu(form, values=[], variable=self.envio_equip_var, width=300)
        self.envio_equip_combo.grid(row=0, column=1, sticky="we", padx=6, pady=6)

        self._create_ctk_label(form, text="Local:").grid(row=1, column=0, sticky="w", padx=6, pady=6)
        self.envio_local_combo = ctk.CTkOptionMenu(form, values=[], variable=self.envio_local_var, width=300)
        self.envio_local_combo.grid(row=1, column=1, sticky="we", padx=6, pady=6)

        self._create_ctk_label(form, text="Data Envio (AAAA-MM-DD):").grid(row=2, column=0, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.envio_data_envio_var).grid(row=2, column=1, sticky="we", padx=6, pady=6)

        # Nova linha para data prevista com botão de previsão automática
        self._create_ctk_label(form, text="Data Prev. Retorno (AAAA-MM-DD):").grid(row=3, column=0, sticky="w", padx=6, pady=6)
        
        # Frame para agrupar o campo de entrada e o botão
        prev_frame = ctk.CTkFrame(form, fg_color="#071025")
        prev_frame.grid(row=3, column=1, sticky="we", padx=6, pady=6)
        
        self._create_ctk_entry(prev_frame, textvariable=self.envio_data_prev_var).pack(side="left", fill="x", expand=True)
        self._create_ctk_button(prev_frame, text="+10 dias", command=self.calcular_previsao_10_dias, 
                            width=80, fg_color="#2D4F8B", hover_color="#1E3A5F").pack(side="right", padx=(6, 0))

        self._create_ctk_label(form, text="Observações:").grid(row=4, column=0, sticky="w", padx=6, pady=6)
        self._create_ctk_entry(form, textvariable=self.envio_obs_var).grid(row=4, column=1, sticky="we", padx=6, pady=6)

        self._create_ctk_button(form, text="Registrar Envio", command=self.on_registrar_envio).grid(row=5, column=1, sticky="e", padx=6, pady=6)
        form.grid_columnconfigure(1, weight=1)

        # CORREÇÃO: Mudar "Obs" para "Observações" nas colunas
        self.tree_mov = ttk.Treeview(frame, columns=("Equipamento", "Local", "Envio", "Prev. Retorno", "Status", "Observações"), show="headings", height=10)
        
        # Definir os cabeçalhos com textos mais claros
        headings = {
            "Equipamento": "Equipamento",
            "Local": "Local", 
            "Envio": "Data Envio",
            "Prev. Retorno": "Prev. Retorno",
            "Status": "Status",
            "Observações": "Observações"
        }
        
        for h in headings:
            self.tree_mov.heading(h, text=headings[h])
            self.tree_mov.column(h, width=160, anchor="center")
        
        self.tree_mov.pack(fill="both", expand=True, pady=(8,6))

        self.tree_mov.tag_configure("oddrow", background="#0f1724")
        self.tree_mov.tag_configure("evenrow", background="#071025")
        
        # Inicializar os comboboxes
        self.refresh_envio_combos()
    
    def calcular_previsao_10_dias(self):
        """Calcula a data de previsão de retorno como data atual + 10 dias"""
        try:
            # Pega a data de envio
            data_envio_str = self.envio_data_envio_var.get().strip()
            
            # Se não houver data de envio, usa a data atual
            if not data_envio_str:
                data_envio = date.today()
                self.envio_data_envio_var.set(data_envio.strftime("%Y-%m-%d"))
            else:
                # Tenta parsear a data
                data_envio = datetime.strptime(data_envio_str, "%Y-%m-%d").date()
            
            # Calcula data de previsão (envio + 10 dias)
            data_prevista = data_envio + timedelta(days=10)
            self.envio_data_prev_var.set(data_prevista.strftime("%Y-%m-%d"))
            
        except ValueError:
            # Se a data de envio for inválida, mostra mensagem e não preenche
            messagebox.showwarning("Atenção", "Data de envio inválida. Use o formato AAAA-MM-DD.")
            self.envio_data_prev_var.set("")  # Limpa o campo de previsão

    def auto_calcular_previsao(self, *args):
        """Preenche automaticamente a previsão quando a data de envio é alterada"""
        # Só preenche se o campo de previsão estiver vazio
        if self.envio_data_prev_var.get().strip():
            return
            
        try:
            data_envio_str = self.envio_data_envio_var.get().strip()
            if data_envio_str and len(data_envio_str) == 10:  # Só calcula se a data estiver completa
                data_envio = datetime.strptime(data_envio_str, "%Y-%m-%d").date()
                data_prevista = data_envio + timedelta(days=10)
                self.envio_data_prev_var.set(data_prevista.strftime("%Y-%m-%d"))
        except ValueError:
            pass  # Ignora erros de parsing durante a digitação

    def refresh_envio_combos(self):
        # Obter equipamentos disponíveis
        equip_list = equipamentos_disponiveis()
        
        # Limpar mapeamento anterior
        self.equip_map = {}
        equip_values = []
        
        # Preencher lista de equipamentos disponíveis
        for e in equip_list:
            equip_text = f"{e['tipo']} - {e['modelo']} ({e['numero_serie'] or e['numero_os'] or 'N/A'})"
            self.equip_map[equip_text] = e['id_equipamento']
            equip_values.append(equip_text)
        
        # Atualizar o combobox de equipamentos
        self.envio_equip_combo.configure(values=equip_values)
        
        if not equip_values:
            # FORÇAR a exibição de "Nenhum equipamento identificado"
            self.envio_equip_combo.set("Nenhum equipamento identificado")
            self.envio_equip_var.set("Nenhum equipamento identificado")
        else:
            # Selecionar o primeiro equipamento disponível
            self.envio_equip_combo.set(equip_values[0])
            self.envio_equip_var.set(equip_values[0])

        # Obter locais disponíveis
        locais_list = listar_locais()
        self.local_map = {}
        local_values = []
        
        for l in locais_list:
            self.local_map[l['nome_local']] = l['id_local']
            local_values.append(l['nome_local'])
        
        # Atualizar combobox de locais
        self.envio_local_combo.configure(values=local_values)
        
        if not local_values:
            self.envio_local_combo.set("Nenhum local disponível")
            self.envio_local_var.set("Nenhum local disponível")
        else:
            self.envio_local_combo.set(local_values[0])
            self.envio_local_var.set(local_values[0])

    def refresh_mov_list(self):
        for i in self.tree_mov.get_children():
            self.tree_mov.delete(i)
        for idx, m in enumerate(listar_mov_abertas()):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            equip_info = f"{m['tipo']} - {m['modelo'] or '-'} ({m['numero_serie'] or m['numero_os'] or 'N/A'})"
            
            # Garantir que o status seja exibido corretamente
            status = m["status"]
            if status == "Em terceiro":
                status = "Em terceiro"
            elif status == "Atrasado":
                status = "Atrasado"
            
            self.tree_mov.insert("", "end", iid=str(m["id_mov"]), 
                            values=(equip_info, m["nome_local"], m["data_envio"], 
                                    m["data_prevista_retorno"] or "-", status, 
                                    m["observacoes"] or "-"), 
                            tags=(tag,))
            
    def on_registrar_envio(self):
        try:
            selected_equip_text = self.envio_equip_var.get()
            selected_local_text = self.envio_local_var.get()

            # print(f"DEBUG: Equipamento selecionado: {selected_equip_text}")
            # print(f"DEBUG: Local selecionado: {selected_local_text}")

            # Verificar se não há equipamentos ou locais disponíveis
            if (selected_equip_text == "Nenhum equipamento identificado" or 
                selected_local_text == "Nenhum local disponível"):
                messagebox.showwarning("Atenção", "Não há equipamentos ou locais disponíveis.")
                return

            # Verificar se os textos selecionados existem no mapeamento
            # print(f"DEBUG: Equipamentos no mapa: {list(self.equip_map.keys())}")
            # print(f"DEBUG: Locais no mapa: {list(self.local_map.keys())}")

            if (selected_equip_text not in self.equip_map or 
                selected_local_text not in self.local_map):
                # print("DEBUG: Seleção não encontrada nos mapas")
                messagebox.showerror("Erro", "Seleção inválida. Atualizando lista...")
                self.refresh_envio_combos()
                return

            id_equip = self.equip_map[selected_equip_text]
            id_local = self.local_map[selected_local_text]

            # print(f"DEBUG: ID Equipamento: {id_equip}, ID Local: {id_local}")
            # print(f"DEBUG: Data envio: {self.envio_data_envio_var.get()}")
            # print(f"DEBUG: Data prevista: {self.envio_data_prev_var.get()}")

            # Testar a função registrar_envio diretamente
            registrar_envio(id_equip, id_local, self.envio_data_envio_var.get(), 
                        self.envio_data_prev_var.get(), self.envio_obs_var.get())
            
            # print("DEBUG: Envio registrado no banco")
            messagebox.showinfo("Sucesso", "Envio registrado!")
            
            # Limpar campos
            self.envio_obs_var.set("")
            self.envio_data_prev_var.set("")
            
            # FORÇAR atualização completa dos comboboxes
            self.refresh_envio_combos()
            
            # Atualizar outras abas
            self.refresh_mov_list()
            self.refresh_devolucao()
            self.refresh_stats()
            
        # except Exception as ex:
        #     print(f"DEBUG: Erro ao registrar envio: {ex}")
        #     messagebox.showerror("Erro", str(ex))
                
        except Exception as ex:
                messagebox.showerror("Erro", str(ex))
    # -----------------------------
    # Tab Devolução
    # -----------------------------
    def build_tab_devolucao(self):
        tab = self.tabs.tab("Devolução")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        frm = ctk.CTkFrame(frame, fg_color="#071025")
        frm.pack(fill="x", pady=(0,8))

        self._create_ctk_button(frm, text="Registrar Retorno", command=self.on_registrar_retorno).pack(side="right", padx=6)
        self.tree_devol = ttk.Treeview(frame, columns=("Equipamento", "SN", "OS", "Local", "Envio", "Prev. Retorno", "Status"), show="headings", height=14)
        for h in ("Equipamento", "SN", "OS", "Local", "Envio", "Prev. Retorno", "Status"):
            self.tree_devol.heading(h, text=h)
            self.tree_devol.column(h, width=160, anchor="center")
        self.tree_devol.pack(fill="both", expand=True, pady=(8,6))

        self.tree_devol.tag_configure("oddrow", background="#0f1724")
        self.tree_devol.tag_configure("evenrow", background="#071025")

    def refresh_devolucao(self):
        for i in self.tree_devol.get_children():
            self.tree_devol.delete(i)
        for idx, m in enumerate(listar_mov_abertas()):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            equip_info = f"{m['tipo']} — {m['modelo'] or '-'}"
            
            # Garantir que o status seja exibido corretamente
            status = m["status"]
            if status == "Em terceiro":
                status = "Em terceiro"
            elif status == "Atrasado":
                status = "Atrasado"
            
            self.tree_devol.insert("", "end", iid=str(m["id_mov"]), 
                                values=(equip_info, m["numero_serie"] or "-", m["numero_os"] or "-", 
                                        m["nome_local"], m["data_envio"], m["data_prevista_retorno"] or "-", 
                                        status), 
                                tags=(tag,))

    def on_registrar_retorno(self):
        sel = self.tree_devol.selection()
        if not sel:
            messagebox.showerror("Erro","Selecione um movimento.")
            return
        try:
            id_mov = int(sel[0])
            registrar_retorno(id_mov)
            messagebox.showinfo("Sucesso","Retorno registrado!")
            self.refresh_envio_combos()
            self.refresh_devolucao()
            self.refresh_all()
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    # -----------------------------
    # Tab Retornados
    # -----------------------------            

    def build_tab_retornados(self):
        """Nova aba para gerenciar equipamentos retornados"""
        tab = self.tabs.tab("Retornados")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)
        
        btn_frame = ctk.CTkFrame(frame, fg_color="#071025")
        btn_frame.pack(fill="x", pady=(0, 8))
        
        self._create_ctk_button(btn_frame, text="Marcar como Devolvido", 
                            command=self.on_marcar_devolvido, width=150).pack(side="left", padx=6)
        self._create_ctk_button(btn_frame, text="Atualizar Lista", 
                            command=self.refresh_retornados, width=100).pack(side="left", padx=6)
        self.tree_retornados = ttk.Treeview(frame, 
                                        columns=("Equipamento", "Local", "Envio", "Retorno", "Status", "Obs"), 
                                        show="headings", height=12)
        
        headings = {
            "Equipamento": "Equipamento",
            "Local": "Local", 
            "Envio": "Data Envio",
            "Retorno": "Data Retorno",
            "Status": "Status",
            "Obs": "Observações"
        }
        
        for h in headings:
            self.tree_retornados.heading(h, text=headings[h])
            self.tree_retornados.column(h, width=140, anchor="center")
        
        self.tree_retornados.pack(fill="both", expand=True)
        self.tree_retornados.tag_configure("oddrow", background="#0f1724")
        self.tree_retornados.tag_configure("evenrow", background="#071025")

    def refresh_retornados(self):
        """Atualiza a lista de equipamentos retornados"""
        for i in self.tree_retornados.get_children():
            self.tree_retornados.delete(i)
        
        retornados = listar_retornados()
        for idx, m in enumerate(retornados):
            tag = 'evenrow' if idx % 2 == 0 else 'oddrow'
            equip_info = f"{m['tipo']} - {m['modelo'] or '-'} ({m['numero_serie'] or m['numero_os'] or 'N/A'})"
            
            self.tree_retornados.insert("", "end", iid=str(m["id_mov"]), 
                                    values=(equip_info, m["nome_local"], m["data_envio"], 
                                            m["data_retorno"] or "-", m["status"], 
                                            m["observacoes"] or "-"), 
                                    tags=(tag,))

    def on_marcar_devolvido(self):
        """Marca um equipamento retornado como devolvido"""
        sel = self.tree_retornados.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione um equipamento retornado.")
            return
        
        try:
            id_mov = int(sel[0])
            marcar_como_devolvido(id_mov)
            messagebox.showinfo("Sucesso", "Equipamento marcado como devolvido!")
            self.refresh_retornados()
            self.refresh_stats()
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))

    # -----------------------------
    # Tab Relatórios
    # -----------------------------
    def build_tab_relatorios(self):
        tab = self.tabs.tab("Relatórios")
        frame = ctk.CTkFrame(tab, fg_color="#071025")
        frame.pack(fill="both", expand=True, padx=12, pady=12)

        frm = ctk.CTkFrame(frame, fg_color="#071025")
        frm.pack(fill="x", pady=(0,8))

        self.rel_local_var = ctk.StringVar()
        self.rel_tipo_var = ctk.StringVar()
        self.rel_status_var = ctk.StringVar()

        ctk.CTkLabel(frm, text="Local:", text_color="#e6eef8").grid(row=0, column=0, padx=6, pady=6, sticky="w")
        self.rel_local_combo = ctk.CTkOptionMenu(frm, values=[], variable=self.rel_local_var, width=220)
        self.rel_local_combo.grid(row=0, column=1, padx=6, pady=6)
        
        ctk.CTkLabel(frm, text="Tipo:", text_color="#e6eef8").grid(row=0, column=2, padx=6, pady=6, sticky="w")
        self.rel_tipo_combo = ctk.CTkOptionMenu(frm, values=["",] + TIPOS_PADRAO, variable=self.rel_tipo_var, width=180)
        self.rel_tipo_combo.grid(row=0, column=3, padx=6, pady=6)
        
        ctk.CTkLabel(frm, text="Status:", text_color="#e6eef8").grid(row=0, column=4, padx=6, pady=6, sticky="w")
        self.rel_status_combo = ctk.CTkOptionMenu(frm, 
                                                values=["", "Em terceiro", "Atrasado", "Retornado", "Devolvido"], 
                                                variable=self.rel_status_var, width=140)
        self.rel_status_combo.grid(row=0, column=5, padx=6, pady=6)

        ctk.CTkButton(frm, text="Filtrar", command=self.on_filtrar).grid(row=0, column=6, padx=6)
        ctk.CTkButton(frm, text="Exportar CSV", command=self.on_exportar).grid(row=0, column=7, padx=6)

        self.tree_rel = ttk.Treeview(frame, columns=("ID","Tipo","Modelo","SN","OS","Local","Envio","Prev","Retorno","Status","Obs"), show="headings", height=14)
        for h in ("ID","Tipo","Modelo","SN","OS","Local","Envio","Prev","Retorno","Status","Obs"):
            self.tree_rel.heading(h, text=h)
            self.tree_rel.column(h, width=120, anchor="center")
        self.tree_rel.pack(fill="both", expand=True, pady=(8,6))

    def refresh_relatorios_combos(self):
        locais_list = [""] + [f"{l['nome_local']}" for l in listar_locais()]
        try:
            self.rel_local_combo.configure(values=locais_list)
            self.rel_local_combo.set(locais_list[0])
        except Exception:
            pass

    def on_filtrar(self):
        id_local = None
        if self.rel_local_var.get():
            for l in listar_locais():
                if l["nome_local"] == self.rel_local_var.get():
                    id_local = l["id_local"]
                    break
        tipo = self.rel_tipo_var.get() or None
        status = self.rel_status_var.get() or None
        rows = filtro_relatorio(id_local=id_local, tipo=tipo, status=status)
        for i in self.tree_rel.get_children():
            self.tree_rel.delete(i)
        for r in rows:
            self.tree_rel.insert("", "end", values=(r["id_mov"], r["tipo"], r["modelo"], r["numero_serie"], r["numero_os"], r["nome_local"], r["data_envio"], r["data_prevista_retorno"], r["data_retorno"], r["status"], r["observacoes"]))

    def on_exportar(self):
        try:
            rows = []
            for item in self.tree_rel.get_children():
                # treeview.item(item)['values'] retorna uma tupla, que é o que precisamos
                rows.append(self.tree_rel.item(item)["values"])
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filepath = os.path.join(EXPORT_DIR, f"relatorio_{ts}.csv")
            exportar_csv(rows, filepath)
            messagebox.showinfo("Sucesso", f"Relatório exportado em:\n{filepath}")
        except Exception as ex:
            messagebox.showerror("Erro", str(ex))


# -----------------------------
# Inicialização
# -----------------------------
if __name__ == "__main__":
    # print(f"DEBUG: Trabalhando no diretório: {os.getcwd()}")
    # print(f"DEBUG: Arquivo do banco: {os.path.abspath(DB_NAME)}")
    
    # if os.path.exists(DB_NAME):
    #     print("DEBUG: Banco de dados já existe - usando existente")
    # else:
    #     print("DEBUG: Criando novo banco de dados")
    
    init_db()
    corrigir_status_existentes()  
    app = App()
    
    def verificar_dados_completos():
        # verificar_dados()
        verificar_movimentacoes_detalhadas()
    
    app.after(2000, verificar_dados_completos)
    app.mainloop()