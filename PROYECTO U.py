import json
import secrets
import os
import calendar
import sys
import time
import textwrap
import hashlib
import sqlite3
from datetime import datetime

# --- INTERFAZ GRÁFICA ---
import tkinter as tk
from tkinter import ttk # <--- IMPORTANTE: De aquí sale el PanedWindow y Treeview
from tkinter import messagebox, filedialog, colorchooser, CENTER, LEFT, RIGHT, TOP, BOTTOM, BOTH, X, Y, VERTICAL, HORIZONTAL, W, E, END, NW, NO

import ttkbootstrap as tb 
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import ToolTip, DateEntry

# --- IMÁGENES ---
from PIL import Image, ImageTk

# --- EXCEL ---
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side as ExcelSide

# --- GRÁFICAS Y PDF (MATPLOTLIB) ---
import matplotlib
matplotlib.use("TkAgg") # Configura el backend para interfaz gráfica

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import numpy as np

# --- ESTOS SON LOS QUE FALTABAN PARA EL PDF ---
import matplotlib.image as mpimg       # Para leer el logo en el PDF
from matplotlib.patches import Rectangle # Para dibujar los cuadros azules en el PDF



# --- RUTAS DEL SISTEMA ---
# Detectar carpeta segura para guardar la Base de Datos y Configuración
appdata = os.getenv('APPDATA')
carpeta_configuracion = os.path.join(appdata, "SistemaInventarioUNINDETEC_SQL")

if not os.path.exists(carpeta_configuracion):
    try:
        os.makedirs(carpeta_configuracion)
    except:
        carpeta_configuracion = os.path.join(os.getenv('TEMP'), "SistemaInventarioUNINDETEC_SQL")
        os.makedirs(carpeta_configuracion, exist_ok=True)

RUTA_DB = os.path.join(carpeta_configuracion, "inventario_master.db")
ARCHIVO_UI_CONFIG = os.path.join(carpeta_configuracion, "config_ui.json")

class GestorBaseDatos:
    def __init__(self, ruta_db):
        """
        Inicializa el gestor con la ruta específica (Local o NAS).
        Si la carpeta no existe, intenta crearla.
        """
        self.ruta_db = ruta_db
        
        # Asegurar que el directorio exista
        directorio = os.path.dirname(ruta_db)
        if directorio and not os.path.exists(directorio):
            try:
                os.makedirs(directorio)
            except Exception as e:
                print(f"Error creando directorio DB: {e}")

        # Ejecutar creación/actualización de tablas al iniciar
        self.crear_tablas()

    def conectar(self):
        conn = sqlite3.connect(self.ruta_db, timeout=10)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")   # Evita bloqueos en escrituras simultáneas
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def crear_tablas(self):
        """Crea las tablas y aplica migraciones"""
        sql_script = """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            rol TEXT DEFAULT 'OPERADOR',
            nombre_completo TEXT DEFAULT 'Usuario del Sistema',
            email TEXT DEFAULT 'sin.email@institucion.gob.mx',
            foto_path TEXT DEFAULT '',
            permisos TEXT DEFAULT '{}'
        );
        CREATE TABLE IF NOT EXISTS inventario (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            partida TEXT,
            material TEXT,
            factura TEXT DEFAULT 'S/F',
            stock REAL DEFAULT 0,
            ultimo_movimiento TEXT,
            UNIQUE(partida, material)
        );
        CREATE TABLE IF NOT EXISTS historial (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha_hora TEXT,
            tipo TEXT,
            partida TEXT,
            material TEXT,
            cantidad REAL,
            factura TEXT,
            destino TEXT,
            responsable TEXT,
            entrego TEXT
        );
        CREATE TABLE IF NOT EXISTS catalogos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo TEXT,
            valor TEXT
        );
        -- NUEVA TABLA PARA NOMBRES LARGOS DE PARTIDAS
        CREATE TABLE IF NOT EXISTS partidas_desc (
            codigo TEXT PRIMARY KEY,
            descripcion TEXT
        );
        CREATE TABLE IF NOT EXISTS config_sistema (
            clave TEXT PRIMARY KEY,
            valor TEXT
        );
        """
        try:
            with self.conectar() as conn:
                conn.executescript(sql_script)
                # Migraciones previas...
                try: conn.execute("ALTER TABLE inventario ADD COLUMN factura TEXT DEFAULT 'S/F'")
                except: pass
                try: conn.execute("ALTER TABLE historial ADD COLUMN factura TEXT DEFAULT ''")
                except: pass
                # Lógica Admin...
                conn.execute("UPDATE usuarios SET rol='ADMIN' WHERE usuario='ADMIN'")
                cursor = conn.cursor()
                cursor.execute("SELECT COUNT(*) FROM usuarios")
                if cursor.fetchone()[0] == 0:
                    pass_hash = hashlib.sha256("admin".encode()).hexdigest()
                    p_full = '{"crear":1, "entrada":1, "salida":1, "editar":1, "eliminar":1, "catalogos":1, "historico":1, "ajustes":1}'
                    conn.execute("INSERT INTO usuarios (usuario, password, rol, permisos) VALUES (?, ?, ?, ?)", 
                                 ("ADMIN", pass_hash, "ADMIN", p_full))
        except Exception as e:
            print(f"Error tablas: {e}")


    def validar_login(self, usuario, password):
        try:
            with self.conectar() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT * FROM usuarios WHERE usuario = ?", (usuario,))
                user = cursor.fetchone()
            
                if not user:
                    return None
            
                stored_hash = user['password']
            
            # ¿Es formato NUEVO? (contiene ":" separando salt:hash)
                if ":" in stored_hash:
                    salt, hash_guardado = stored_hash.split(":", 1)
                    import hashlib
                    hash_intento = hashlib.pbkdf2_hmac(
                        'sha256', 
                        password.encode(), 
                        salt.encode(), 
                        100000
                    ).hex()
                    valido = (hash_intento == hash_guardado)
                else:
                # Formato VIEJO (SHA256 simple) - compatibilidad temporal
                    import hashlib
                    hash_viejo = hashlib.sha256(password.encode()).hexdigest()
                    valido = (hash_viejo == stored_hash)
                
                # ✅ MIGRACIÓN AUTOMÁTICA: Si la contraseña es correcta,
                # aprovechamos para actualizar al nuevo formato en ese momento
                    if valido:
                        import secrets
                        salt = secrets.token_hex(16)
                        hash_nuevo = hashlib.pbkdf2_hmac(
                            'sha256', password.encode(), salt.encode(), 100000
                        ).hex()
                        nuevo_stored = f"{salt}:{hash_nuevo}"
                        conn.execute(
                            "UPDATE usuarios SET password = ? WHERE usuario = ?",
                            (nuevo_stored, usuario)
                        )
                        conn.commit()
                        print(f"✅ Contraseña de {usuario} migrada al nuevo formato")
            
                return dict(user) if valido else None
            
        except Exception as e:
            print(f"Error en login: {e}")
            return None

    def get_config(self, clave):
        """Recupera configuraciones (Ruta de Logo, Título App, etc.)"""
        try:
            with self.conectar() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT valor FROM config_sistema WHERE clave = ?", (clave,))
                res = cursor.fetchone()
                return res['valor'] if res else None
        except:
            return None

    def set_config(self, clave, valor):
        """Guarda o actualiza una configuración"""
        try:
            with self.conectar() as conn:
                conn.execute("REPLACE INTO config_sistema (clave, valor) VALUES (?, ?)", (clave, valor))
                conn.commit()
        except Exception as e:
            print(f"Error guardando config {clave}: {e}")

    def consultar(self, query, params=()):
        """Helper para consultas SELECT"""
        try:
            with self.conectar() as conn:
                cursor = conn.cursor()
                cursor.execute(query, params)
                return cursor.fetchall()
        except Exception as e:
            print(f"Error SQL (SELECT): {e}")
            return []

    def ejecutar(self, query, params=()):
        """Helper para INSERT, UPDATE, DELETE"""
        try:
            with self.conectar() as conn:
                conn.execute(query, params)
                conn.commit()
        except Exception as e:
            print(f"Error SQL (EXEC): {e}")
            raise e # Lanzamos el error para que la interfaz muestre el mensaje
class GestorTemas:
    """Clase para gestionar temas personalizados"""
    
    # Temas Predefinidos
    TEMAS_PREDEFINIDOS = {
        "Azul Profesional": {
            "color_primario": "#1F4E79",
            "color_secundario": "#4472C4",
            "color_acento": "#00B0F0",
            "color_fondo": "#FFFFFF",
            "color_texto": "#000000",
            "tema_bootstrap": "flatly"
        },
        "Verde Moderno": {
            "color_primario": "#2E7D32",
            "color_secundario": "#66BB6A",
            "color_acento": "#4CAF50",
            "color_fondo": "#FAFAFA",
            "color_texto": "#212121",
            "tema_bootstrap": "minty"
        },
        "Oscuro": {
            "color_primario": "#212121",
            "color_secundario": "#424242",
            "color_acento": "#FFC107",
            "color_fondo": "#303030",
            "color_texto": "#FFFFFF",
            "tema_bootstrap": "darkly"
        },
        "Rojo Corporativo": {
            "color_primario": "#C00000",
            "color_secundario": "#E74C3C",
            "color_acento": "#FF6B6B",
            "color_fondo": "#FFFFFF",
            "color_texto": "#2C3E50",
            "tema_bootstrap": "journal"
        },
        "Morado Creativo": {
            "color_primario": "#6A1B9A",
            "color_secundario": "#9C27B0",
            "color_acento": "#BA68C8",
            "color_fondo": "#F5F5F5",
            "color_texto": "#212121",
            "tema_bootstrap": "pulse"
        },
        "Naranja Energético": {
            "color_primario": "#E65100",
            "color_secundario": "#FF6F00",
            "color_acento": "#FFB300",
            "color_fondo": "#FFFFFF",
            "color_texto": "#212121",
            "tema_bootstrap": "sandstone"
        }
    }
    
    @staticmethod
    def get_tema_actual(db_manager):
        """Obtener tema actual de la base de datos"""
        tema = {
            "color_primario": db_manager.get_config("TEMA_COLOR_PRIMARIO") or "#1F4E79",
            "color_secundario": db_manager.get_config("TEMA_COLOR_SECUNDARIO") or "#4472C4",
            "color_acento": db_manager.get_config("TEMA_COLOR_ACENTO") or "#00B0F0",
            "color_fondo": db_manager.get_config("TEMA_COLOR_FONDO") or "#FFFFFF",
            "color_texto": db_manager.get_config("TEMA_COLOR_TEXTO") or "#000000",
            "tema_bootstrap": db_manager.get_config("TEMA_BOOTSTRAP") or "flatly"
        }
        return tema
    
    @staticmethod
    def guardar_tema(db_manager, tema):
        """Guardar tema en la base de datos"""
        db_manager.set_config("TEMA_COLOR_PRIMARIO", tema["color_primario"])
        db_manager.set_config("TEMA_COLOR_SECUNDARIO", tema["color_secundario"])
        db_manager.set_config("TEMA_COLOR_ACENTO", tema["color_acento"])
        db_manager.set_config("TEMA_COLOR_FONDO", tema["color_fondo"])
        db_manager.set_config("TEMA_COLOR_TEXTO", tema["color_texto"])
        db_manager.set_config("TEMA_BOOTSTRAP", tema["tema_bootstrap"])

class LoginWindow(tk.Toplevel):
    def __init__(self, parent, db_manager, on_success):
        super().__init__(parent)
        self.title("Acceso Seguro")
        
        # 1. Ocultamos la ventana mientras la construimos para que no se vea "parpadear" al ajustarse
        self.withdraw()
        
        self.db = db_manager
        self.callback = on_success
        self.protocol("WM_DELETE_WINDOW", self.cancelar_login)
        
        # --- CARGAR CONFIGURACIÓN ---
        titulo = self.db.get_config("TITULO_APP")
        if not titulo: titulo = "SOTFWARE DE USO LIBRE" 
        
        subtitulo = self.db.get_config("SUBTITULO_APP")
        if not subtitulo: subtitulo = "CONTROL DE PARTIDAS"
        
        logo_path = self.db.get_config("LOGO_APP")
        if not logo_path or not os.path.exists(logo_path):
            logo_path = "logo.png" if os.path.exists("logo.png") else None

        # --- INTERFAZ ---
        # Logo
        if logo_path:
            try:
                img = Image.open(logo_path)
                img = img.resize((110, 110), Image.LANCZOS)
                self.img_login = ImageTk.PhotoImage(img)
                ttk.Label(self, image=self.img_login).pack(pady=(25, 10))
            except Exception as e:
                print(f"Error imagen login: {e}")
                ttk.Label(self, text="🔒", font=("Segoe UI Emoji", 60)).pack(pady=(20, 10))
        else:
            ttk.Label(self, text="🔒", font=("Segoe UI Emoji", 60)).pack(pady=(20, 10))

        # Textos
        ttk.Label(self, text=titulo, font=("Arial Black", 16), bootstyle="primary", justify=CENTER).pack(pady=(0, 5), padx=10)
        ttk.Label(self, text=subtitulo, font=("Segoe UI", 10, "bold"), justify=CENTER, foreground="gray").pack(pady=(0, 20), padx=10)
        
        # Formulario
        fr = ttk.Frame(self, padding=30)
        fr.pack(fill=BOTH, expand=True)
        
        ttk.Label(fr, text="Usuario:").pack(anchor=W)
        self.entry_user = ttk.Entry(fr)
        self.entry_user.pack(fill=X, pady=(0, 15))
        self.entry_user.focus()
        
        ttk.Label(fr, text="Contraseña:").pack(anchor=W)
        self.entry_pass = ttk.Entry(fr, show="*")
        self.entry_pass.pack(fill=X, pady=(0, 20))
        self.entry_pass.bind("<Return>", lambda e: self.entrar())
        
        ttk.Button(fr, text="INICIAR SESIÓN", bootstyle="primary", command=self.entrar).pack(fill=X, pady=10)
        
        self.lbl_msg = ttk.Label(fr, text="", foreground="red", justify=CENTER)
        self.lbl_msg.pack()

        # --- AUTOAJUSTE Y CENTRADO ---
        self.update_idletasks() # IMPORTANTE: Calcula el tamaño real de los elementos
        
        width = 400 # Ancho fijo que queremos
        height = self.winfo_reqheight() # Altura automática según contenido
        
        # Obtener dimensiones pantalla
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        
        x = (ws // 2) - (width // 2)
        y = (hs // 2) - (height // 2)
        
        # Aplicar geometría calculada
        self.geometry(f'{width}x{height}+{int(x)}+{int(y)}')
        self.resizable(False, False)
        
        # Mostrar ventana ya lista
        self.deiconify()
    
    def cancelar_login(self):
        if messagebox.askyesno("Salir", "¿Deseas salir del sistema?"):
            sys.exit()

    def entrar(self):
        u = self.entry_user.get().strip().upper()
        p = self.entry_pass.get().strip()
        
        user_data = self.db.validar_login(u, p)
        if user_data:
            self.lbl_msg.config(text="Acceso Correcto...", foreground="green")
            self.after(500, lambda: [self.destroy(), self.callback(user_data)])
        else:
            self.lbl_msg.config(text="Usuario o contraseña incorrectos", foreground="red")


            
class SistemaInventario:
    def __init__(self, root, db_manager, usuario_actual):
        self.root = root
        self.db = db_manager
        self.usuario = usuario_actual
        # Cargar tema actual
        self.tema_actual = GestorTemas.get_tema_actual(self.db)
        
        # LEER TITULO DE LA BASE DE DATOS
        titulo_ventana = self.db.get_config("TITULO_APP")
        if not titulo_ventana: 
            titulo_ventana = "Sistema de Inventario"
        
        self.root.title(f"{titulo_ventana} - [Usuario: {self.usuario['usuario']}]")
        self.animacion_actual = None
        
        # --- CAMBIO: INICIAR MAXIMIZADO ---
        try:
            self.root.state('zoomed') # Para Windows
        except:
            self.root.attributes('-zoomed', True) # Para Linux/Otros
            
        # Si por alguna razón falla el zoomed, usamos un tamaño base grande
        if self.root.state() != 'zoomed':
            self.root.geometry("1280x850")
        
        # Estilos visuales
        style = ttk.Style()
        style.configure("Treeview", rowheight=30)
        style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))
        
        # CREAR INTERFAZ
        self.notebook = ttk.Notebook(self.root, bootstyle="primary")
        
        # Inicializar variable para el logo
        self.tk_logo = None
        
        # Llamamos al header
        self.setup_header()
        
        # Empaquetamos el notebook después del header
        self.notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Crear pestañas en el nuevo orden
        self.tab_inv = ttk.Frame(self.notebook, padding=10)
        self.tab_audit = ttk.Frame(self.notebook, padding=10)
        self.tab_consumo = ttk.Frame(self.notebook, padding=10)
        self.tab_hist = ttk.Frame(self.notebook, padding=10)
        
        # NUEVO ORDEN DE PESTAÑAS PRINCIPALES
        self.notebook.add(self.tab_inv, text="📦 INVENTARIO")
        self.notebook.add(self.tab_audit, text="📑 DATOS")
        self.notebook.add(self.tab_consumo, text="📈 CONSUMO")
        self.notebook.add(self.tab_hist, text="🕒 HISTORIAL")
        
        # Construir contenido de pestañas
        self.setup_tab_inventario()
        self.setup_tab_auditoria()
        self.setup_tab_consumo() 
        self.setup_tab_historial()
        
        # Carga inicial de datos
        self.actualizar_combos()
        self.cargar_tabla_inventario()
        self.cargar_tabla_historial()
        self.datos_kardex_procesados = []

   

    def limpiar_filtros(self):
        """
        Resetea la barra de búsqueda y el filtro de categorías,
        y vuelve a cargar todo el inventario.
        """
        try:
            # 1. Limpiar caja de búsqueda (Si existe la variable)
            if hasattr(self, 'var_busqueda'):
                self.var_busqueda.set("")
            
            # 2. Resetear ComboBox de Categorías a "TODAS" (Índice 0)
            if hasattr(self, 'combo_categoria'):
                try:
                    self.combo_categoria.current(0)
                except: pass
            
            # 3. Recargar la tabla de inventario
            # (El sistema busca la función de carga, ya sea cargar_inventario o cargar_datos)
            if hasattr(self, 'cargar_inventario'):
                self.cargar_inventario()
            elif hasattr(self, 'cargar_datos'):
                self.cargar_datos()
                
        except Exception as e:
            print(f"Advertencia al limpiar filtros: {e}")


    def cargar_config_ui(self):
        # Valores por defecto
        self.ui_titulo = "UNINDETEC"
        self.ui_subtitulo = "CONTROL DE INVENTARIO"
        self.ui_tema = "primary"
        self.ui_logo = ""
        
        if os.path.exists(ARCHIVO_UI_CONFIG):
            try:
                with open(ARCHIVO_UI_CONFIG, 'r') as f:
                    data = json.load(f)
                    self.ui_titulo = data.get("titulo", self.ui_titulo)
                    self.ui_subtitulo = data.get("subtitulo", self.ui_subtitulo)
                    self.ui_tema = data.get("tema", self.ui_tema)
                    self.ui_logo = data.get("logo", "")
            except: pass

    # --- EN LA CLASE SistemaInventario ---
    def setup_header(self):
        # 1. Limpiar encabezado anterior
        for widget in self.root.pack_slaves():
            if isinstance(widget, ttk.Frame) and widget != self.notebook:
                widget.destroy()

        # 2. Frame Principal del Header (Más alto y espacioso)
        fr = ttk.Frame(self.root, padding=(20, 15)) # Aumentamos padding vertical
        try: fr.pack(side=TOP, fill=X, before=self.notebook)
        except: fr.pack(side=TOP, fill=X)
        
        # ==========================================
        # SECCIÓN IZQUIERDA: LOGO Y TÍTULO DEL SISTEMA
        # ==========================================
        fr_left = ttk.Frame(fr)
        fr_left.pack(side=LEFT, fill=Y, anchor=W)

        # Logo del Sistema
        logo_path = self.db.get_config("LOGO_APP")
        if not logo_path or not os.path.exists(logo_path):
            logo_path = "logo.png" if os.path.exists("logo.png") else None

        if logo_path:
            try:
                # Logo un poco más grande también (100px)
                img = Image.open(logo_path).resize((100, 100), Image.LANCZOS)
                self.tk_logo = ImageTk.PhotoImage(img, master=fr)
                ttk.Label(fr_left, image=self.tk_logo).pack(side=LEFT, padx=(0, 20))
            except: 
                ttk.Label(fr_left, text="⚓", font=("Arial", 40)).pack(side=LEFT, padx=15)
        else:
            ttk.Label(fr_left, text="⚓", font=("Arial", 40)).pack(side=LEFT, padx=15)

        # Títulos
        txt_fr = ttk.Frame(fr_left)
        txt_fr.pack(side=LEFT, fill=Y, anchor=CENTER)
        
        t_principal = self.db.get_config("TITULO_APP") or "SOTFWARE DE USO LIBRE"
        t_sub = self.db.get_config("SUBTITULO_APP") or "CONTROL DE PARTIDAS"

        ttk.Label(txt_fr, text=t_principal, font=("Arial Black", 28), bootstyle="primary").pack(anchor=W)
        ttk.Label(txt_fr, text=t_sub, font=("Segoe UI", 14, "bold"), bootstyle="secondary").pack(anchor=W)
        
        # ==========================================
        # SECCIÓN DERECHA: FICHA DE USUARIO (AMPLIA)
        # ==========================================
        
        # Contenedor derecho
        fr_right = ttk.Frame(fr)
        fr_right.pack(side=RIGHT, fill=Y, anchor=E)

        # Datos del Usuario
        nombre_user = self.usuario.get('nombre_completo', self.usuario['usuario']).upper()
        email_user = self.usuario.get('email', 'Usuario del Sistema')
        foto_user_path = self.usuario.get('foto_path', '')
        rol_user = self.usuario.get('rol', 'OPERADOR')

        # 1. TEXTOS (BIENVENIDA Y NOMBRE)
        fr_textos_user = ttk.Frame(fr_right)
        fr_textos_user.pack(side=LEFT, padx=(0, 20), anchor=E)

        # Etiqueta de Bienvenida
        ttk.Label(fr_textos_user, text="¡BIENVENIDO(A)!", 
                 font=("Segoe UI", 10, "bold"), bootstyle="success", anchor=E).pack(anchor=E)
        
        # Nombre del Usuario (GRANDE)
        ttk.Label(fr_textos_user, text=nombre_user, 
                 font=("Segoe UI", 18, "bold"), bootstyle="primary", anchor=E).pack(anchor=E)
        
        # Rol y Email
        texto_rol = f"{rol_user} | {email_user}"
        ttk.Label(fr_textos_user, text=texto_rol, 
                 font=("Segoe UI", 10), bootstyle="secondary", anchor=E).pack(anchor=E)

        # 2. FOTO DE PERFIL (GRANDE)
        fr_foto_marco = ttk.Frame(fr_right, padding=2, bootstyle="secondary") # Un marquito fino
        fr_foto_marco.pack(side=LEFT, padx=(0, 15))

        if foto_user_path and os.path.exists(foto_user_path):
            try:
                # FOTO MUCHO MÁS GRANDE (75x75)
                img_u = Image.open(foto_user_path).resize((75, 75), Image.LANCZOS)
                self.tk_foto_user = ImageTk.PhotoImage(img_u, master=fr)
                lbl_foto = ttk.Label(fr_foto_marco, image=self.tk_foto_user)
                lbl_foto.pack()
            except:
                 ttk.Label(fr_foto_marco, text="👤", font=("Segoe UI Emoji", 45)).pack()
        else:
            icono_def = "👨‍✈️" if rol_user == 'ADMIN' else "👤"
            ttk.Label(fr_foto_marco, text=icono_def, font=("Segoe UI Emoji", 45)).pack()

        # 3. BOTÓN DE CONFIGURACIÓN (ENGRANAJE)
        # Separador vertical
        ttk.Separator(fr_right, orient=VERTICAL).pack(side=LEFT, fill=Y, padx=10)
        
        btn_conf = ttk.Button(fr_right, text="⚙️", bootstyle="link", command=self.abrir_menu_admin)
        # Agrandamos el engrane
        for child in btn_conf.winfo_children(): child.configure(font=("Segoe UI Emoji", 30))
        btn_conf.pack(side=LEFT, padx=5)
        btn_conf.pack(side=LEFT)

    def setup_autocomplete(self, combo, lista_completa):
        """Filtra la lista internamente sin bloquear ni interrumpir la escritura"""
        
        def al_escribir(event):
            # Si son teclas especiales, no hacemos nada
            if event.keysym in ['Up', 'Down', 'Return', 'Tab', 'Left', 'Right']:
                return

            texto = combo.get().strip().upper()

            if texto == "":
                # Si está vacío, restauramos toda la lista
                combo['values'] = lista_completa
            else:
                # Filtrar lista (Silenciosamente)
                filtrados = [x for x in lista_completa if texto in str(x).upper()]
                combo['values'] = filtrados
                # NO abrimos la lista automáticamente (evita el bloqueo)

        # Vinculamos al evento de soltar tecla
        combo.bind("<KeyRelease>", al_escribir)
        
        # Carga inicial
        combo['values'] = lista_completa
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA (RESTAURACIÓN DE INTERFAZ) ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA (VERSIÓN ANTI-BLOQUEO) ---
    # ------------------------------------------------------------------
    #  INTERFAZ DE INVENTARIO REDISEÑADA (FÁCIL PARA EL USUARIO)
    # ------------------------------------------------------------------
    def setup_tab_inventario(self):
        # 1. DIVIDIR LA PANTALLA
        self.p_izq = ttk.Frame(self.tab_inv, width=550)
        self.p_izq.pack(side=LEFT, fill=Y, padx=(0, 10))
        self.p_izq.pack_propagate(False) 
        
        p_der = ttk.Frame(self.tab_inv)
        p_der.pack(side=RIGHT, fill=BOTH, expand=True)

        # === PANEL IZQUIERDO: INDICADOR DE SELECCIÓN ===
        fr_info = ttk.LabelFrame(self.p_izq, text=" 📦 Material Seleccionado ", padding=10, bootstyle="secondary")
        fr_info.pack(fill=X, pady=(0, 10))
        
        self.lbl_seleccionado = ttk.Label(fr_info, text="Ninguno (Selecciona en la tabla ➡)", 
                                          font=("Segoe UI", 11, "bold"), foreground="#E67E22", wraplength=350)
        self.lbl_seleccionado.pack(anchor=CENTER)
        self.lbl_stock_actual = ttk.Label(fr_info, text="Stock: --", font=("Segoe UI", 10))
        self.lbl_stock_actual.pack(anchor=CENTER)

        # === PANEL IZQUIERDO: PESTAÑAS DE ACCIÓN ===
        self.nb_acciones = ttk.Notebook(self.p_izq, bootstyle="primary")
        self.nb_acciones.pack(fill=BOTH, expand=True)

        # --- PESTAÑA 1: ENTRADAS (REUBICADA AL PRINCIPIO) ---
        self.tab_entradas = ttk.Frame(self.nb_acciones, padding=15)
        self.nb_acciones.add(self.tab_entradas, text="⬇️ ENTRADAS")

        ttk.Label(self.tab_entradas, text="ENTRADA DE NUEVO MATERIAL", font=("Segoe UI", 12, "bold"), foreground="#27ae60").pack(pady=(0, 15))

        ttk.Label(self.tab_entradas, text="Cantidad a Ingresar:").pack(anchor=W)
        self.ent_cant_ent = ttk.Entry(self.tab_entradas, font=("Segoe UI", 14, "bold"), justify=CENTER)
        self.ent_cant_ent.pack(fill=X, pady=5)

        ttk.Label(self.tab_entradas, text="Factura / Referencia:").pack(anchor=W, pady=(10, 0))
        self.ent_factura_ent = ttk.Entry(self.tab_entradas)
        self.ent_factura_ent.pack(fill=X, pady=2)

        # PERMISO ENTRADA
        estado_ent = "normal" if self.tiene_permiso('entrada') else "disabled"
        btn_ent = ttk.Button(self.tab_entradas, text="✅ REGISTRAR ENTRADA", bootstyle="success", state=estado_ent,
                           command=lambda: self.procesar_movimiento("ENTRADA"))
        btn_ent.pack(fill=X, pady=30, ipady=5)
        
        if estado_ent == "disabled": ToolTip(btn_ent, text="No tienes permiso para registrar entradas")

        # --- PESTAÑA 2: SALIDAS (REUBICADA AL MEDIO) ---
        self.tab_salidas = ttk.Frame(self.nb_acciones, padding=15)
        self.nb_acciones.add(self.tab_salidas, text="⬆️ SALIDAS")

        ttk.Label(self.tab_salidas, text="Registrar Salida", font=("Segoe UI", 12, "bold"), foreground="#c0392b").pack(pady=(0, 15))

        ttk.Label(self.tab_salidas, text="Cantidad a Retirar:").pack(anchor=W)
        self.ent_cant_sal = ttk.Entry(self.tab_salidas, font=("Segoe UI", 14, "bold"), justify=CENTER)
        self.ent_cant_sal.pack(fill=X, pady=5)

        ttk.Label(self.tab_salidas, text="Destino / Área:").pack(anchor=W, pady=(10, 0))
        self.cb_area_sal = ttk.Combobox(self.tab_salidas)
        self.cb_area_sal.pack(fill=X, pady=2)

        ttk.Label(self.tab_salidas, text="Solicita (Nombre):").pack(anchor=W, pady=(10, 0))
        self.ent_resp_sal = ttk.Entry(self.tab_salidas)
        self.ent_resp_sal.pack(fill=X, pady=2)

        ttk.Label(self.tab_salidas, text="Autoriza / Entrega:").pack(anchor=W, pady=(10, 0))
        self.cb_jefe_sal = ttk.Combobox(self.tab_salidas)
        self.cb_jefe_sal.pack(fill=X, pady=2)
        
        # PERMISO SALIDA
        estado_sal = "normal" if self.tiene_permiso('salida') else "disabled"
        btn_sal = ttk.Button(self.tab_salidas, text="🔥 REGISTRAR SALIDA", bootstyle="danger", state=estado_sal,
                           command=lambda: self.procesar_movimiento("SALIDA"))
        btn_sal.pack(fill=X, pady=30, ipady=5)
        
        if estado_sal == "disabled": ToolTip(btn_sal, text="No tienes permiso para registrar salidas")

        # --- PESTAÑA 3: NUEVO (SE QUEDA AL FINAL) ---
        self.tab_nuevo = ttk.Frame(self.nb_acciones, padding=15)
        self.nb_acciones.add(self.tab_nuevo, text="➕ CREAR NUEVO MATERIAL")

        ttk.Label(self.tab_nuevo, text="Crear Producto", font=("Segoe UI", 12, "bold"), foreground="#2980b9").pack(pady=(0, 15))
        
        ttk.Label(self.tab_nuevo, text="Partida:").pack(anchor=W)
        self.cb_partida = ttk.Combobox(self.tab_nuevo, state="readonly")
        self.cb_partida.pack(fill=X, pady=2)
        
        ttk.Label(self.tab_nuevo, text="Descripción:").pack(anchor=W, pady=(10, 0))
        self.txt_desc = tk.Text(self.tab_nuevo, height=4, font=("Segoe UI", 10), wrap="word")
        self.txt_desc.pack(fill=X, pady=2, padx=1) 
        
        # --- NUEVO CAMPO: STOCK INICIAL ---
        ttk.Label(self.tab_nuevo, text="Cantidad Inicial (Stock):").pack(anchor=W, pady=(10, 0))
        self.ent_stock_inicial = ttk.Entry(self.tab_nuevo, justify=CENTER, font=("Segoe UI", 10, "bold"))
        self.ent_stock_inicial.insert(0, "0") # Valor por defecto
        self.ent_stock_inicial.pack(fill=X, pady=2)
        # ----------------------------------

        ttk.Label(self.tab_nuevo, text="Factura Inicial:").pack(anchor=W, pady=(10, 0))
        self.txt_factura_alta = ttk.Entry(self.tab_nuevo)
        self.txt_factura_alta.pack(fill=X, pady=2)
        
        # PERMISO CREAR
        estado_crear = "normal" if self.tiene_permiso('crear') else "disabled"
        btn_crear = ttk.Button(self.tab_nuevo, text="💾 GUARDAR", bootstyle="info", state=estado_crear,
                             command=self.agregar_material)
        btn_crear.pack(fill=X, pady=30, ipady=5)
        
        if estado_crear == "disabled": ToolTip(btn_crear, text="No tienes permiso para crear materiales")

        # === PANEL DERECHO: TABLA ===
        fr_top = ttk.Frame(p_der, padding=(0, 5)); fr_top.pack(fill=X)
        
        ttk.Label(fr_top, text="🔍 Buscar:", font=("Segoe UI", 9, "bold")).pack(side=LEFT)
        self.cb_busqueda_material = ttk.Combobox(fr_top, width=30)
        self.cb_busqueda_material.pack(side=LEFT, padx=5)
        self.cb_busqueda_material.bind("<KeyRelease>", lambda e: self.cargar_tabla_inventario())
        self.cb_busqueda_material.bind("<<ComboboxSelected>>", self.cargar_tabla_inventario)

        ttk.Label(fr_top, text="📂 Filtrar Partida:", font=("Segoe UI", 9, "bold")).pack(side=LEFT, padx=(15, 5))
        self.cb_filtro_partida = ttk.Combobox(fr_top, state="readonly", width=15)
        self.cb_filtro_partida.pack(side=LEFT, padx=5)
        self.cb_filtro_partida.bind("<<ComboboxSelected>>", self.cargar_tabla_inventario)
        
        ttk.Button(fr_top, text="🔄 Ver Todo", bootstyle="link", command=self.limpiar_filtros).pack(side=LEFT)

        cols = ("ID", "PARTIDA", "MATERIAL", "STOCK")
        self.tree_inv = ttk.Treeview(p_der, columns=cols, show="headings", bootstyle="info")
        self.tree_inv.heading("ID", text="ID"); self.tree_inv.column("ID", width=40, anchor=CENTER)
        self.tree_inv.heading("PARTIDA", text="PARTIDA"); self.tree_inv.column("PARTIDA", width=80, anchor=CENTER)
        self.tree_inv.heading("MATERIAL", text="DESCRIPCIÓN"); self.tree_inv.column("MATERIAL", width=400)
        self.tree_inv.heading("STOCK", text="STOCK"); self.tree_inv.column("STOCK", width=80, anchor=CENTER)
        
        sc = ttk.Scrollbar(p_der, orient=VERTICAL, command=self.tree_inv.yview)
        self.tree_inv.configure(yscrollcommand=sc.set)
        self.tree_inv.pack(side=LEFT, fill=BOTH, expand=True)
        sc.pack(side=RIGHT, fill=Y)
        self.tree_inv.tag_configure("BAJO", background="#ffcccc", foreground="#8a1f1f")

        self.tree_inv.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        # MENÚ CONTEXTUAL
        self.menu_inv = tk.Menu(self.root, tearoff=0)
        
        # PERMISOS PARA MENÚ CONTEXTUAL
        if self.tiene_permiso('editar'):
            self.menu_inv.add_command(label="✏️ Corregir/Editar Material", command=self.editar_material_seleccionado)
        
        if self.tiene_permiso('eliminar') or self.usuario.get('rol') == 'ADMIN':
            self.menu_inv.add_separator()
            self.menu_inv.add_command(label="🗑️ Eliminar Material (Solo Admin)", command=self.eliminar_material_seleccionado)
        
        def mostrar_menu_inv(event):
            item = self.tree_inv.identify_row(event.y)
            if item:
                self.tree_inv.selection_set(item)
                estado_borrar = "normal" if self.usuario.get('rol') == 'ADMIN' else "disabled"
                try:
                    self.menu_inv.entryconfig("🗑️ Eliminar Material (Solo Admin)", state=estado_borrar)
                except: pass
                self.menu_inv.post(event.x_root, event.y_root)
        
        if self.tiene_permiso('editar') or self.tiene_permiso('eliminar') or self.usuario.get('rol') == 'ADMIN':
            self.tree_inv.bind("<Button-3>", mostrar_menu_inv)


    def on_tree_select(self, event):
        """Actualiza el letrero naranja de la izquierda al seleccionar un producto"""
        sel = self.tree_inv.selection()
        if sel:
            item = self.tree_inv.item(sel[0])
            vals = item['values']
            # vals[2] es Nombre, vals[3] es Stock
            self.lbl_seleccionado.config(text=vals[2]) 
            self.lbl_stock_actual.config(text=f"Stock Actual: {vals[3]}")
            
            # Efecto visual: Rojo si es poco stock
            try:
                stock = float(vals[3])
                if stock <= 2:
                    self.lbl_seleccionado.config(foreground="red")
                else:
                    self.lbl_seleccionado.config(foreground="#E67E22")
            except: pass
        else:
            self.lbl_seleccionado.config(text="Ninguno (Selecciona en la tabla ➡)", foreground="gray")
            self.lbl_stock_actual.config(text="Stock: --")

        def mostrar_menu_inv(event):
            item = self.tree_inv.identify_row(event.y)
            if item:
                self.tree_inv.selection_set(item)
                estado_borrar = "normal" if self.usuario.get('rol') == 'ADMIN' else "disabled"
                self.menu_inv.entryconfig("🗑️ Eliminar Material (Solo Admin)", state=estado_borrar)
                self.menu_inv.post(event.x_root, event.y_root)
        self.tree_inv.bind("<Button-3>", mostrar_menu_inv)


    # ------------------------------------------------------------------
    #  LÓGICA NUEVA PARA PROCESAR MOVIMIENTOS SEPARADOS
    # ------------------------------------------------------------------
    def procesar_movimiento(self, tipo):
        sel = self.tree_inv.selection()
        if not sel:
            messagebox.showwarning("Atención", "Selecciona un material de la tabla derecha.")
            return

        item = self.tree_inv.item(sel[0])
        valores = item['values']
        id_mat, partida, nombre_mat, stock_actual = valores[0], valores[1], valores[2], float(valores[3])

        cantidad = 0; factura = "S/F"; destino = "S/N"; responsable = "S/N"; entrego = "S/N"

        try:
            if tipo == "SALIDA":
                cantidad = float(self.ent_cant_sal.get())
                destino = self.cb_area_sal.get().strip().upper() or "S/N"
                responsable = self.ent_resp_sal.get().strip().upper() or "S/N"
                entrego = self.cb_jefe_sal.get().strip().upper() or "S/N"
                if cantidad > stock_actual:
                    messagebox.showerror("Error", f"Stock insuficiente ({stock_actual}).")
                    return
            elif tipo == "ENTRADA":
                cantidad = float(self.ent_cant_ent.get())
                factura = self.ent_factura_ent.get().strip().upper() or "S/F"

            if cantidad <= 0: raise ValueError
        except:
            messagebox.showerror("Error", "Cantidad inválida.")
            return

        nuevo_stock = stock_actual + cantidad if tipo == "ENTRADA" else stock_actual - cantidad
        fecha = datetime.now().strftime("%d/%m/%Y")
        fecha_full = datetime.now().strftime("%d/%m/%Y %H:%M")

    # --- TRANSACCIÓN ATÓMICA ---
        try:
            with self.db.conectar() as conn:
                conn.execute("BEGIN")
                try:
                # 1. Actualizar stock
                    if tipo == "ENTRADA" and factura != "S/F":
                        conn.execute(
                            "UPDATE inventario SET stock=?, ultimo_movimiento=?, factura=? WHERE id=?",
                            (nuevo_stock, fecha, factura, id_mat)
                        )
                    else:
                        conn.execute(
                            "UPDATE inventario SET stock=?, ultimo_movimiento=? WHERE id=?",
                            (nuevo_stock, fecha, id_mat)
                        )

                # 2. Guardar historial
                    conn.execute(
                        "INSERT INTO historial (fecha_hora, tipo, partida, material, cantidad, destino, responsable, entrego, factura) VALUES (?,?,?,?,?,?,?,?,?)",
                        (fecha_full, tipo, partida, nombre_mat, cantidad, destino, responsable, entrego, factura)
                    )

                    conn.commit()

                except Exception as e:
                    conn.rollback()
                    messagebox.showerror("Error crítico", f"Movimiento cancelado. La BD no fue modificada.\n\nDetalle: {e}")
                    return

        # --- DESPUÉS de confirmar la transacción, generamos el PDF (no es parte de la BD) ---
            if tipo == "SALIDA":
                folio = self.generar_folio()
                self.generar_pdf_vale(nombre_mat, cantidad, destino, responsable, entrego, folio)

            messagebox.showinfo("Éxito", f"Movimiento registrado. Nuevo stock: {nuevo_stock}")

            if tipo == "SALIDA":
                self.ent_cant_sal.delete(0, END)
            else:
                self.ent_cant_ent.delete(0, END)

            self.cargar_tabla_inventario()

            try:
                self.tree_inv.selection_set(sel[0])
                self.on_tree_select(None)
            except:
                pass

        except Exception as e:
            messagebox.showerror("Error de conexión", f"No se pudo conectar a la base de datos.\n\nDetalle: {e}")

    def cargar_tabla_inventario(self, event=None):
        for i in self.tree_inv.get_children(): self.tree_inv.delete(i)
        
        # 1. Obtener valores de los filtros
        partida_sel = self.cb_filtro_partida.get()
        texto_busqueda = self.cb_busqueda_material.get().strip().upper()
        
        # 2. Construir Consulta SQL
        sql = "SELECT * FROM inventario WHERE 1=1"
        params = []
        
        # Filtro 1: Partida
        if partida_sel and partida_sel != "TODAS":
            sql += " AND partida = ?"
            params.append(partida_sel)
            
        # Filtro 2: Material (Búsqueda parcial)
        if texto_busqueda:
            sql += " AND material LIKE ?"
            params.append(f"%{texto_busqueda}%")
            
        sql += " ORDER BY id DESC"
        
        # 3. Ejecutar y Llenar
        filas = self.db.consultar(sql, tuple(params))
        for f in filas:
            tag = "BAJO" if f['stock'] <= 2 else ""
            self.tree_inv.insert("", END, values=(f['id'], f['partida'], f['material'], f['stock']), tags=(tag,))

    def agregar_material(self):
        partida = self.cb_partida.get()
        mat = self.txt_desc.get("1.0", "end-1c").strip().upper()
        fact = self.txt_factura_alta.get().strip().upper() or "S/F"
        
        # Capturamos el Stock Inicial
        try:
            stock_ini = float(self.ent_stock_inicial.get())
            if stock_ini < 0: raise ValueError
        except:
            stock_ini = 0.0
        
        if not partida or not mat:
            messagebox.showwarning("Faltan datos", "Indica Partida y Descripción")
            return
            
        try:
            # 1. Insertar en INVENTARIO con el stock inicial
            self.db.ejecutar("INSERT INTO inventario (partida, material, factura, stock, ultimo_movimiento) VALUES (?, ?, ?, ?, 'ALTA')", 
                             (partida, mat, fact, stock_ini))
            
            # 2. Si hay stock inicial, registrarlo en el HISTORIAL para que cuadre el Kardex
            if stock_ini > 0:
                fecha_full = datetime.now().strftime("%d/%m/%Y %H:%M")
                usuario_act = self.usuario.get('usuario', 'SISTEMA')
                
                # Insertamos como 'ALTA INICIAL' para diferenciarlo
                self.db.ejecutar("""
                    INSERT INTO historial (fecha_hora, tipo, partida, material, cantidad, destino, responsable, entrego, factura)
                    VALUES (?, 'ALTA INICIAL', ?, ?, ?, 'ALMACEN', ?, 'SISTEMA', ?)
                """, (fecha_full, partida, mat, stock_ini, usuario_act, fact))

            messagebox.showinfo("Éxito", f"Material creado correctamente.\nStock inicial: {stock_ini}")
            
            # Limpiar campos
            self.txt_desc.delete("1.0", END)
            self.txt_factura_alta.delete(0, END)
            self.ent_stock_inicial.delete(0, END)
            self.ent_stock_inicial.insert(0, "0")
            
            # Recargar tablas
            self.cargar_tabla_inventario()
            self.actualizar_combos() 

        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicado", "Este material ya existe en esa partida")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {e}")


    def setup_tab_historial(self):
        cols = ("ID", "FECHA", "TIPO", "PARTIDA", "MATERIAL", "CANT", "DESTINO", "RESP", "ENTREGO")
        # Agregamos ID al principio para poder identificar el registro en la BD
        self.tree_hist = ttk.Treeview(self.tab_hist, columns=cols, show="headings", bootstyle="primary")
        
        anchos = [40, 120, 80, 60, 250, 50, 100, 100, 100]
        for c, w in zip(cols, anchos):
            self.tree_hist.heading(c, text=c)
            self.tree_hist.column(c, width=w, anchor=CENTER if c in ["ID", "CANT", "TIPO"] else W)
            
        sc = ttk.Scrollbar(self.tab_hist, orient=VERTICAL, command=self.tree_hist.yview)
        self.tree_hist.configure(yscrollcommand=sc.set)
        self.tree_hist.pack(side=LEFT, fill=BOTH, expand=True)
        sc.pack(side=RIGHT, fill=Y)
        
        self.tree_hist.tag_configure("ENTRADA", foreground="green")
        self.tree_hist.tag_configure("SALIDA", foreground="blue")
        self.tree_hist.tag_configure("SISTEMA", foreground="gray")
        self.tree_hist.tag_configure("ELIMINADO", foreground="red") # Para auditoría de borrados

        # ========================================================
        # MENÚ CONTEXTUAL HISTORIAL (SOLO ADMIN)
        # ========================================================
        self.menu_hist = tk.Menu(self.root, tearoff=0)
        self.menu_hist.add_command(label="⚠️ Revertir y Eliminar Registro (Admin)", command=self.revertir_historial_admin)

        def mostrar_menu_hist(event):
            # Solo mostrar si es ADMIN
            if self.usuario.get('rol') == 'ADMIN':
                item = self.tree_hist.identify_row(event.y)
                if item:
                    self.tree_hist.selection_set(item)
                    self.menu_hist.post(event.x_root, event.y_root)

        self.tree_hist.bind("<Button-3>", mostrar_menu_hist)

    def cargar_tabla_historial(self):
        for i in self.tree_hist.get_children(): self.tree_hist.delete(i)
        # Traemos también el ID
        filas = self.db.consultar("SELECT id, fecha_hora, tipo, partida, material, cantidad, destino, responsable, entrego FROM historial ORDER BY id DESC LIMIT 500")
        for f in filas:
            color = "ENTRADA" if "ENTRADA" in f['tipo'] else ("SALIDA" if "SALIDA" in f['tipo'] else "SISTEMA")
            self.tree_hist.insert("", END, values=(f['id'], f['fecha_hora'], f['tipo'], f['partida'], f['material'], 
                                                   f['cantidad'], f['destino'], f['responsable'], f['entrego']), tags=(color,))

    # --- PESTAÑA AUDITORIA (KARDEX) ---
    def setup_tab_auditoria(self):
        fr_top = ttk.Frame(self.tab_audit, padding=10)
        fr_top.pack(fill=X)
        
        ttk.Label(fr_top, text="Buscar Material:").pack(side=LEFT)
        self.cb_kardex_mat = ttk.Combobox(fr_top, width=40)
        self.cb_kardex_mat.pack(side=LEFT, padx=5)
        
        ttk.Button(fr_top, text="Generar Kardex", command=self.generar_kardex).pack(side=LEFT)
        ttk.Button(fr_top, text="💾 Exportar Excel", bootstyle="success-outline", command=self.exportar_excel_kardex).pack(side=RIGHT)
        
        self.tree_kardex = ttk.Treeview(self.tab_audit, columns=("FECHA", "MOVIMIENTO", "CANT", "SALDO"), show="headings")
        self.tree_kardex.heading("FECHA", text="Fecha"); self.tree_kardex.heading("MOVIMIENTO", text="Movimiento")
        self.tree_kardex.heading("CANT", text="Cant"); self.tree_kardex.heading("SALDO", text="Saldo")
        self.tree_kardex.pack(fill=BOTH, expand=True, pady=10)

    def generar_kardex(self):
        mat = self.cb_kardex_mat.get()
        if not mat: return
        
        # Limpiar tabla visual
        for i in self.tree_kardex.get_children(): self.tree_kardex.delete(i)
        self.datos_kardex_procesados = [] # Limpiar datos para Excel
        
        # Consultar DB
        movs = self.db.consultar("SELECT * FROM historial WHERE material = ? ORDER BY id ASC", (mat,))
        
        saldo = 0
        for m in movs:
            tipo = m['tipo']
            cant = m['cantidad']
            
            # Calcular saldo
            if "ENTRADA" in tipo or "HISTORICO (+)" in tipo:
                saldo += cant
            elif "SALIDA" in tipo or "HISTORICO (-)" in tipo:
                saldo -= cant
                
            # 1. Insertar en la tabla VISUAL
            self.tree_kardex.insert("", END, values=(m['fecha_hora'], tipo, cant, saldo))
            
            # 2. Guardar datos PROCESADOS para Excel (Diccionario limpio)
            self.datos_kardex_procesados.append({
                "FECHA": m['fecha_hora'],
                "MOVIMIENTO": tipo,
                "CANTIDAD": cant,
                "SALDO": saldo,
                "DESTINO": m['destino'],
                "RESPONSABLE": m['responsable']
            })

    def exportar_excel_kardex(self):
        # Verifica si hay datos procesados (calculados en generar_kardex)
        if not hasattr(self, 'datos_kardex_procesados') or not self.datos_kardex_procesados:
            messagebox.showwarning("Alerta", "Primero genera el Kardex en pantalla.")
            return
        
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if ruta:
            try:
                # Usamos la lista procesada que TIENE EL SALDO
                df = pd.DataFrame(self.datos_kardex_procesados)
                
                # Reordenar columnas para que se vea bien
                cols_ordenadas = ["FECHA", "MOVIMIENTO", "CANTIDAD", "SALDO", "DESTINO", "RESPONSABLE"]
                # Asegurarnos de que solo usamos columnas que existen (por si acaso)
                cols_final = [c for c in cols_ordenadas if c in df.columns]
                df = df[cols_final]
                
                df.to_excel(ruta, index=False)
                messagebox.showinfo("Exportado", "Archivo Excel generado correctamente.")
                os.startfile(ruta) # Abrir automáticamente
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo exportar: {e}")
    # --- MENÚ ADMIN ---
    # --- EN LA CLASE SistemaInventario ---
    def abrir_menu_admin(self):
        top = tk.Toplevel(self.root)
        top.title("Administración del Sistema")
        # Aumentamos un poco la altura para el nuevo botón
        self.centrar_ventana_emergente(top, 420, 580) 
        
        fr = ttk.Frame(top, padding=20)
        fr.pack(fill=BOTH, expand=True)
        
        # Encabezado del menú con info de rol
        rol_actual = self.usuario.get('rol', 'OPERADOR')
        es_admin_rol = (rol_actual == 'ADMIN')

        ttk.Label(fr, text="Menú de Configuración", font=("Segoe UI", 14, "bold"), justify=CENTER).pack(pady=(0, 5))
        
        if not es_admin_rol:
            ttk.Label(fr, text="Modo Restringido: Según tus permisos.", 
                     bootstyle="warning", font=("Segoe UI", 9)).pack(pady=(0, 15))
        else:
             ttk.Label(fr, text="Modo Administrador: Acceso total.", 
                     bootstyle="success", font=("Segoe UI", 9)).pack(pady=(0, 15))
    
        # --- BOTONES ---

        # 1. Gestión de Usuarios (SOLO ADMIN ROL)
        # Esto es demasiado sensible para delegarlo por permiso simple.
        estado_users = "normal" if es_admin_rol else "disabled"
        ttk.Button(fr, text="👥  Gestionar Usuarios y Permisos", bootstyle="primary", state=estado_users,
                   command=self.abrir_gestion_usuarios).pack(fill=X, pady=8, ipady=8)
        
        ttk.Separator(fr).pack(fill=X, pady=10)

        # 2. Temas (Configuración visual básica permitida a todos)
        ttk.Button(fr, text="🎨  Personalizar Temas y Colores", bootstyle="info",
                command=self.abrir_editor_temas).pack(fill=X, pady=8, ipady=8)
                
        # 3. Catálogos (Permiso 'catalogos')
        estado_cat = "normal" if self.tiene_permiso('catalogos') else "disabled"
        btn_cat = ttk.Button(fr, text="📋  Gestión de Catálogos (Partidas)", bootstyle="warning",
                             state=estado_cat, command=self.abrir_gestor_catalogos)
        btn_cat.pack(fill=X, pady=8, ipady=8)
        if estado_cat == "disabled": ToolTip(btn_cat, text="No tienes permiso para editar catálogos")
                
        # 4. Histórico (Permiso 'historico')
        estado_hist = "normal" if self.tiene_permiso('historico') else "disabled"
        btn_hist = ttk.Button(fr, text="📅  Modificar Histórico (Pasado)", bootstyle="danger",
                           state=estado_hist, command=self.abrir_registro_pasado)
        btn_hist.pack(fill=X, pady=8, ipady=8)
        if estado_hist == "disabled": ToolTip(btn_hist, text="No tienes permiso para modificar el historial")
                
        # 5. Ajustes Visuales (Permiso 'ajustes')
        estado_conf = "normal" if self.tiene_permiso('ajustes') else "disabled"
        btn_ajustes = ttk.Button(fr, text="⚙️  Ajustes del Sistema (Logos)", bootstyle="secondary",
                              state=estado_conf, command=self.abrir_ajustes_visuales)
        btn_ajustes.pack(fill=X, pady=8, ipady=8)
        if estado_conf == "disabled": ToolTip(btn_ajustes, text="No tienes permiso para cambiar logos")
        
        ttk.Button(fr, text="Cerrar Menú", bootstyle="outline", command=top.destroy).pack(side=BOTTOM, fill=X, pady=(20, 0))

    def abrir_gestor_catalogos(self):
        # 1. Configurar Ventana
        top = tk.Toplevel(self.root)
        top.title("Gestión de Catálogos")
        # Hacemos la ventana redimensionable y con un tamaño mínimo
        top.geometry("700x600")
        top.minsize(600, 500) 
        
        # Contenedor principal que se expande
        main_frame = ttk.Frame(top)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        tabs = ttk.Notebook(main_frame)
        tabs.pack(fill=BOTH, expand=True)
        
        # --- PESTAÑA ESPECIAL PARA PARTIDAS (Con Descripción) ---
        fr_part = ttk.Frame(tabs, padding=15)
        tabs.add(fr_part, text="📂 Partidas (Códigos)")
        
        ttk.Label(fr_part, text="Catálogo de Partidas Presupuestales:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        ttk.Label(fr_part, text="* Selecciona una fila para editar su descripción", font=("Segoe UI", 8), foreground="gray").pack(anchor=W, pady=(0, 5))
        
        # Frame para la tabla (para que el scrollbar se pegue bien)
        fr_tree_container = ttk.Frame(fr_part)
        fr_tree_container.pack(fill=BOTH, expand=True, pady=5)

        # Lista con columnas
        cols_p = ("CODIGO", "DESCRIPCION")
        tree_part = ttk.Treeview(fr_tree_container, columns=cols_p, show="headings")
        
        tree_part.heading("CODIGO", text="Código"); 
        tree_part.column("CODIGO", width=100, anchor=CENTER, stretch=False) # Código no se estira tanto
        tree_part.heading("DESCRIPCION", text="Descripción"); 
        tree_part.column("DESCRIPCION", width=400, stretch=True) # Descripción ocupa el resto
        
        sc_p = ttk.Scrollbar(fr_tree_container, orient=VERTICAL, command=tree_part.yview)
        tree_part.configure(yscrollcommand=sc_p.set)
        
        tree_part.pack(side=LEFT, fill=BOTH, expand=True)
        sc_p.pack(side=RIGHT, fill=Y)
        
        def cargar_partidas():
            # Guardar selección actual si existe
            sel_id = tree_part.selection()
            prev_cod = tree_part.item(sel_id[0])['values'][0] if sel_id else None

            for i in tree_part.get_children(): tree_part.delete(i)
            
            # Consultamos catalogos y unimos con descripciones
            filas = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='PARTIDA' ORDER BY valor ASC")
            for f in filas:
                cod = f['valor']
                # Buscar descripción
                desc_res = self.db.consultar("SELECT descripcion FROM partidas_desc WHERE codigo=?", (cod,))
                nombre = desc_res[0]['descripcion'] if desc_res else "(Sin descripción)"
                item = tree_part.insert("", END, values=(cod, nombre))
                
                # Restaurar selección
                if prev_cod and str(cod) == str(prev_cod):
                    tree_part.selection_set(item)
                    tree_part.see(item)

        cargar_partidas()

        # --- ÁREA DE EDICIÓN (Inferior) ---
        fr_add_p = ttk.LabelFrame(fr_part, text=" Editar / Agregar ", padding=10, bootstyle="info")
        fr_add_p.pack(fill=X, pady=10)
        
        # Grid para alinear mejor
        fr_add_p.columnconfigure(1, weight=1) # La columna de descripción se estira

        ttk.Label(fr_add_p, text="Código:").grid(row=0, column=0, padx=5, sticky=W)
        e_cod = ttk.Entry(fr_add_p, width=15, font=("Segoe UI", 10, "bold"))
        e_cod.grid(row=1, column=0, padx=5, sticky=EW)
        
        ttk.Label(fr_add_p, text="Descripción (Nombre Largo):").grid(row=0, column=1, padx=5, sticky=W)
        e_desc = ttk.Entry(fr_add_p, width=40)
        e_desc.grid(row=1, column=1, padx=5, sticky=EW)

        # FUNCIONALIDAD DE SELECCIÓN (CLICK EN TABLA)
        def al_seleccionar(event):
            sel = tree_part.selection()
            if not sel: return
            item = tree_part.item(sel[0])
            valores = item['values']
            
            # Llenar campos
            e_cod.delete(0, END); e_cod.insert(0, valores[0])
            e_desc.delete(0, END); e_desc.insert(0, valores[1])
            
        tree_part.bind("<<TreeviewSelect>>", al_seleccionar)

        def guardar_partida():
            c = e_cod.get().strip().upper()
            d = e_desc.get().strip().upper()
            if not c: 
                messagebox.showwarning("Error", "El código es obligatorio", parent=top)
                return
            
            # 1. Guardar en Catalogos (Lista simple) - INSERT OR IGNORE para no duplicar error
            existe = self.db.consultar("SELECT * FROM catalogos WHERE tipo='PARTIDA' AND valor=?", (c,))
            if not existe:
                self.db.ejecutar("INSERT INTO catalogos (tipo, valor) VALUES ('PARTIDA', ?)", (c,))
            
            # 2. Guardar Descripción (REPLACE actualiza si ya existe, inserta si es nuevo)
            self.db.ejecutar("REPLACE INTO partidas_desc (codigo, descripcion) VALUES (?, ?)", (c, d))
            
            # Limpiar y recargar
            e_cod.delete(0, END); e_desc.delete(0, END)
            cargar_partidas()
            self.actualizar_combos()
            
            # --- CAMBIO IMPORTANTE AQUÍ: parent=top y top.lift() ---
            messagebox.showinfo("Guardado", f"Partida {c} guardada/actualizada correctamente.", parent=top)
            top.lift() # Fuerza a la ventana a ponerse al frente de nuevo
            # -------------------------------------------------------

        # Botones grandes y claros
        btn_frame = ttk.Frame(fr_add_p)
        btn_frame.grid(row=1, column=2, padx=10)
        
        ttk.Button(btn_frame, text="💾 Guardar / Actualizar", bootstyle="success", command=guardar_partida).pack(fill=X)
        
        def eliminar_partida():
            sel = tree_part.selection()
            if not sel: 
                messagebox.showwarning("Atención", "Selecciona una partida de la lista para eliminar.", parent=top)
                return
            item = tree_part.item(sel[0])
            cod = item['values'][0]
            
            # --- CAMBIO IMPORTANTE AQUÍ TAMBIÉN ---
            if messagebox.askyesno("Confirmar", f"¿Borrar partida {cod}?\nSe eliminará del catálogo (no del historial).", parent=top):
                self.db.ejecutar("DELETE FROM catalogos WHERE tipo='PARTIDA' AND valor=?", (cod,))
                self.db.ejecutar("DELETE FROM partidas_desc WHERE codigo=?", (cod,))
                cargar_partidas()
                self.actualizar_combos()
                e_cod.delete(0, END); e_desc.delete(0, END)
                top.lift() # Asegurar que no se esconda al borrar

        ttk.Button(fr_part, text="🗑️ Eliminar Seleccionada", bootstyle="danger", command=eliminar_partida).pack(fill=X, pady=(0,5))

        # --- PESTAÑAS SIMPLES (Area, Jefe) ---
        # Se mantienen igual pero con layout mejorado
        def crear_tab_lista_simple(tipo_cat, titulo):
            fr = ttk.Frame(tabs, padding=15)
            tabs.add(fr, text=titulo)
            
            fr_list_cont = ttk.Frame(fr)
            fr_list_cont.pack(fill=BOTH, expand=True, pady=5)
            
            lst = tk.Listbox(fr_list_cont, height=10)
            lst.pack(side=LEFT, fill=BOTH, expand=True)
            
            sb = ttk.Scrollbar(fr_list_cont, orient=VERTICAL, command=lst.yview)
            sb.pack(side=RIGHT, fill=Y)
            lst.config(yscrollcommand=sb.set)

            def cargar():
                lst.delete(0, END)
                fs = self.db.consultar("SELECT valor FROM catalogos WHERE tipo=? ORDER BY valor ASC", (tipo_cat,))
                for x in fs: lst.insert(END, x['valor'])
            cargar()
            
            fr_controls = ttk.Frame(fr)
            fr_controls.pack(fill=X, pady=5)
            
            e_val = ttk.Entry(fr_controls)
            e_val.pack(side=LEFT, fill=X, expand=True, padx=(0,5))
            
            def add():
                v = e_val.get().strip().upper()
                if v: 
                    self.db.ejecutar("INSERT INTO catalogos (tipo, valor) VALUES (?, ?)", (tipo_cat, v))
                    e_val.delete(0, END); cargar(); self.actualizar_combos()
            
            def delete():
                s = lst.curselection()
                if s: 
                    v = lst.get(s[0])
                    self.db.ejecutar("DELETE FROM catalogos WHERE tipo=? AND valor=?", (tipo_cat, v))
                    cargar(); self.actualizar_combos()
            
            ttk.Button(fr_controls, text="➕ Agregar", bootstyle="success", command=add).pack(side=LEFT)
            ttk.Button(fr, text="🗑️ Eliminar Seleccionado", bootstyle="danger", command=delete).pack(fill=X)

        crear_tab_lista_simple("AREA", "🏢 Áreas")
        crear_tab_lista_simple("JEFE", "👤 Jefes")

        # ==========================================================
        # --- NUEVA PESTAÑA: CÓNSTAME (FIRMA) ---
        # ==========================================================
        fr_const = ttk.Frame(tabs, padding=15)
        tabs.add(fr_const, text="✍️ Cónstame")
        
        ttk.Label(fr_const, text="Nombre / Firma de Autoridad (Cónstame):", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        ttk.Label(fr_const, text="Este nombre aparecerá en la parte inferior de los Vales de Salida PDF.", font=("Segoe UI", 8), foreground="gray").pack(anchor=W, pady=(0, 10))
        
        e_const = ttk.Entry(fr_const, font=("Segoe UI", 11))
        e_const.pack(fill=X, pady=5)
        
        # Cargar el valor actual de la base de datos
        res_firma = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='FIRMA'")
        if res_firma:
            e_const.insert(0, res_firma[0]['valor'])
            
        def guardar_constame():
            nueva_firma = e_const.get().strip().upper()
            if nueva_firma:
                # Borramos el anterior y guardamos el nuevo para que solo exista uno
                self.db.ejecutar("DELETE FROM catalogos WHERE tipo='FIRMA'")
                self.db.ejecutar("INSERT INTO catalogos (tipo, valor) VALUES ('FIRMA', ?)", (nueva_firma,))
                messagebox.showinfo("Guardado", "Firma 'Cónstame' actualizada correctamente para los próximos PDFs.", parent=top)
                top.lift()
            else:
                messagebox.showwarning("Atención", "El campo no puede estar vacío.", parent=top)
                top.lift()
                
        ttk.Button(fr_const, text="💾 Guardar Firma", bootstyle="success", command=guardar_constame).pack(pady=15, fill=X)


    def actualizar_combos(self):
        # 1. PARTIDAS (Para Alta, Filtro Inventario y NUEVO FILTRO KARDEX)
        rows = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='PARTIDA' ORDER BY valor ASC")
        lista_p = [r['valor'] for r in rows]
        
        # A. Combo de Alta de Material (Pestaña 1)
        if hasattr(self, 'cb_partida'): 
            self.cb_partida['values'] = lista_p
            self.setup_autocomplete(self.cb_partida, lista_p)
            
        # B. Filtro de Inventario (Pestaña 1 Derecha)
        if hasattr(self, 'cb_filtro_partida'):
            self.cb_filtro_partida['values'] = ["TODAS"] + lista_p
            if not self.cb_filtro_partida.get(): self.cb_filtro_partida.current(0)
        
        # C. --- CORRECCIÓN: Filtro de Kardex (Pestaña 3) ---
        if hasattr(self, 'cb_partida_k'):
            # Le agregamos "TODAS" al principio
            self.cb_partida_k['values'] = ["TODAS"] + lista_p
            # Si está vacío, seleccionamos "TODAS" por defecto para que no se vea aplastado
            if not self.cb_partida_k.get(): 
                self.cb_partida_k.current(0)
        
        # 2. AREAS (CORREGIDO EL NOMBRE DEL COMBOBOX)
        rows = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='AREA' ORDER BY valor ASC")
        lista_a = [r['valor'] for r in rows]
        if hasattr(self, 'cb_area_sal'):  # <-- Corrección aquí
            self.cb_area_sal['values'] = lista_a
            self.setup_autocomplete(self.cb_area_sal, lista_a)
        
        # 3. JEFES (CORREGIDO EL NOMBRE DEL COMBOBOX)
        rows = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='JEFE' ORDER BY valor ASC")
        lista_j = [r['valor'] for r in rows]
        if hasattr(self, 'cb_jefe_sal'):  # <-- Corrección aquí
            self.cb_jefe_sal['values'] = lista_j
            self.setup_autocomplete(self.cb_jefe_sal, lista_j)

        # 4. MATERIALES (Para Buscadores)
        rows_mat = self.db.consultar("SELECT DISTINCT material FROM inventario ORDER BY material ASC")
        lista_materiales = [r['material'] for r in rows_mat]
        
        # Buscador del INVENTARIO
        if hasattr(self, 'cb_busqueda_material'):
            self.setup_autocomplete(self.cb_busqueda_material, lista_materiales)

    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA ---
    def abrir_registro_pasado(self):
        """
        Registro Histórico Manual CON DISEÑO RESPONSIVO
        """
        top = tb.Toplevel(self.root) 
        top.title("Registro Histórico Manual")
        # Tamaño inicial más grande y un mínimo para que no se corte
        top.geometry("500x700")
        top.minsize(450, 600)
        
        ttk.Label(top, text="⚠️ CUIDADO: Esto afecta el stock actual.", bootstyle="warning").pack(pady=10)
        
        # Frame principal con scroll por si la pantalla es muy pequeña (opcional, pero seguro)
        # Usamos pack fill=BOTH para que ocupe todo
        fr = ttk.Frame(top, padding=20)
        fr.pack(fill=BOTH, expand=True)
        
        # --- CAMPOS ---
        # Usamos fill=X en todos los packs para que se estiren horizontalmente
        
        # 1. FECHA
        ttk.Label(fr, text="Fecha del Movimiento:").pack(anchor=W)
        e_fecha = tb.DateEntry(fr, bootstyle="info", dateformat="%d/%m/%Y")
        e_fecha.pack(fill=X, pady=(0, 10))
        
        # 2. TIPO
        ttk.Label(fr, text="Tipo de Movimiento:").pack(anchor=W)
        c_tipo = ttk.Combobox(fr, values=["HISTORICO (+) Entrada/Saldo Inicial", "HISTORICO (-) Salida/Ajuste"], state="readonly")
        c_tipo.current(0)
        c_tipo.pack(fill=X, pady=(0, 10))
        
        # 3. PARTIDA (OBLIGATORIA)
        ttk.Label(fr, text="Partida (Obligatorio):").pack(anchor=W)
        vals_partidas = []
        if hasattr(self, 'cb_partida'): vals_partidas = self.cb_partida['values']
        
        c_partida_hist = ttk.Combobox(fr, values=vals_partidas, state="readonly")
        c_partida_hist.pack(fill=X, pady=(0, 10))
        self.setup_autocomplete(c_partida_hist, list(vals_partidas))

        # 4. MATERIAL
        ttk.Label(fr, text="Material Exacto (Busca):").pack(anchor=W)
        vals_mat = []
        if hasattr(self, 'cb_busqueda_material'):
             vals_mat = self.cb_busqueda_material['values']
             
        c_mat = ttk.Combobox(fr, values=vals_mat)
        self.setup_autocomplete(c_mat, list(vals_mat))
        c_mat.pack(fill=X, pady=(0, 10))
        
        def al_elegir_material(event):
            mat_name = c_mat.get()
            res = self.db.consultar("SELECT partida FROM inventario WHERE material=?", (mat_name,))
            if res:
                c_partida_hist.set(res[0]['partida'])
        c_mat.bind("<<ComboboxSelected>>", al_elegir_material)

        # 5. CANTIDAD
        ttk.Label(fr, text="Cantidad:").pack(anchor=W)
        e_cant = ttk.Entry(fr)
        e_cant.pack(fill=X, pady=(0, 10))

        # 6. FACTURA
        ttk.Label(fr, text="Factura / Documento (Opcional):").pack(anchor=W)
        e_factura_hist = ttk.Entry(fr)
        e_factura_hist.pack(fill=X, pady=(0, 10))
        
        # 7. OBSERVACIÓN
        ttk.Label(fr, text="Observación / Responsable:").pack(anchor=W)
        e_obs = ttk.Entry(fr)
        e_obs.pack(fill=X, pady=(0, 20)) # Un poco más de espacio antes del botón
        
        # --- Lógica de Guardado ---
        def guardar_historico():
            mat = c_mat.get().strip().upper()
            part = c_partida_hist.get().strip()
            tipo_sel = c_tipo.get()
            fecha = e_fecha.entry.get()
            obs = e_obs.get().strip().upper() or "AJUSTE MANUAL"
            fact = e_factura_hist.get().strip().upper() or "S/F"
            
            if not mat or not fecha or not part: 
                messagebox.showwarning("Faltan datos", "Material, Partida y Fecha son obligatorios")
                return

            try:
                cant = float(e_cant.get())
                if cant <= 0: raise ValueError
            except: 
                messagebox.showerror("Error", "Cantidad inválida")
                return

            existe = self.db.consultar("SELECT id FROM inventario WHERE material = ? AND partida = ?", (mat, part))
            
            if "(+)" in tipo_sel:
                tipo_db = "HISTORICO (+)"
                if existe:
                    sql_stock = "UPDATE inventario SET stock = stock + ? WHERE material = ? AND partida = ?"
                    self.db.ejecutar(sql_stock, (cant, mat, part))
                else:
                    if messagebox.askyesno("Nuevo Material", "Este material no existe en esa partida. ¿Crearlo con este stock inicial?"):
                        self.db.ejecutar("INSERT INTO inventario (partida, material, stock, ultimo_movimiento, factura) VALUES (?, ?, ?, ?, ?)",
                                         (part, mat, cant, fecha, fact))
                    else:
                        return
            else:
                tipo_db = "HISTORICO (-)"
                if existe:
                    sql_stock = "UPDATE inventario SET stock = stock - ? WHERE material = ? AND partida = ?"
                    self.db.ejecutar(sql_stock, (cant, mat, part))
                else:
                    messagebox.showerror("Error", "No puedes restar stock de un material que no existe.")
                    return

            try:
                self.db.ejecutar("""
                    INSERT INTO historial (fecha_hora, tipo, partida, material, cantidad, responsable, entrego, factura)
                    VALUES (?, ?, ?, ?, ?, ?, 'AJUSTE HISTORICO', ?)
                """, (fecha, tipo_db, part, mat, cant, obs, fact))
                
                messagebox.showinfo("Éxito", "Registro histórico aplicado correctamente.")
                self.cargar_tabla_inventario()
                self.cargar_tabla_historial()
                top.destroy()
                
            except Exception as e:
                 messagebox.showerror("Error DB", f"{e}")

        # Botón grande al fondo, pegado abajo con espacio
        ttk.Button(fr, text="💾 APLICAR MOVIMIENTO", bootstyle="success", command=guardar_historico).pack(fill=X, side=BOTTOM, pady=10, ipady=5)

    def generar_folio(self):
        rows = self.db.consultar("SELECT COUNT(*) as total FROM historial WHERE tipo='SALIDA'")
        consecutivo = rows[0]['total'] + 1
        return f"{consecutivo:03d}-{datetime.now().year}"

    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA PARA RECUPERAR EL FORMATO DE TABLA ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA PARA TENER EL FORMATO TABLA EXACTO ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA (AJUSTE DE FIRMAS) ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA (CORRECCIÓN DE ANCHO DE TEXTO) ---
    # --- REEMPLAZA ESTA FUNCIÓN COMPLETA (AJUSTE FINAL DE POSICIÓN DEL LOGO) ---
    def generar_pdf_vale(self, material, cantidad, area, resp, jefe, folio):
        """
        Genera el PDF usando el TÍTULO PRINCIPAL (Empresa) y SUBTÍTULO (Depto)
        configurados en los Ajustes del Sistema.
        """
        try:
            # 1. Definir Ruta (Escritorio)
            escritorio = os.path.join(os.environ['USERPROFILE'], 'Desktop')
            nombre_limpio = f"VALE_{folio.replace('/', '-')}.pdf"
            ruta_pdf = os.path.join(escritorio, nombre_limpio)

            # 2. Obtener Datos de la BD
            # AQUI ESTÁ EL CAMBIO: Usamos TITULO_APP y SUBTITULO_APP
            # Esto te permite poner "UNINDETEC" arriba y "DEPARTAMENTO DE CÓMPUTO" abajo
            # editando solo los cuadros de texto de "Títulos de la Ventana" en Ajustes.
            
            empresa_titulo = self.db.get_config("TITULO_APP") 
            if not empresa_titulo: empresa_titulo = "NOMBRE EMPRESA" # Default
            
            depto_subtitulo = self.db.get_config("SUBTITULO_APP")
            if not depto_subtitulo: depto_subtitulo = "DEPARTAMENTO" # Default
            
            # Recuperar firma configurada
            res_firma = self.db.consultar("SELECT valor FROM catalogos WHERE tipo='FIRMA'")
            firma_constame = res_firma[0]['valor'] if res_firma else "AUTORIDAD"
            
            # 3. Configuración Gráfica
            plt.switch_backend('Agg') 
            fig = plt.figure(figsize=(8.5, 11)) 
            ax = fig.add_subplot(111)
            ax.set_xlim(0, 8.5)
            ax.set_ylim(0, 11)
            ax.axis('off')
            
            AZUL_OSCURO = "#1F4E79"
            ROJO_FOLIO = "#C00000"
            
            # --- A. LOGO ---
            logo_path = self.db.get_config("LOGO_PDF")
            if not logo_path or not os.path.exists(logo_path):
                 logo_path = self.db.get_config("LOGO_APP")

            if logo_path and os.path.exists(logo_path):
                try:
                    img = mpimg.imread(logo_path)
                    ax_logo = fig.add_axes([0.15, 0.88, 0.12, 0.10], anchor='NW', zorder=1)
                    ax_logo.imshow(img)
                    ax_logo.axis('off')
                except: pass

            # --- B. ENCABEZADO (EMPRESA - DEPTO) ---
            # Título Grande (Empresa)
            ax.text(4.25, 10.6, empresa_titulo, fontsize=22, weight='bold', color=AZUL_OSCURO, ha='center')
            # Subtítulo (Departamento)
            ax.text(4.25, 10.30, depto_subtitulo, fontsize=11, weight='bold', color='gray', ha='center')

            # Folio y Fecha
            ax.text(8.0, 10.0, f"FOLIO: {folio}", color=ROJO_FOLIO, weight='bold', ha='right', fontsize=11)
            ax.text(8.0, 9.8, f"FECHA: {datetime.now().strftime('%d-%b-%Y').upper()}", color=AZUL_OSCURO, weight='bold', ha='right', fontsize=10)

            # --- C. TÍTULO DEL DOCUMENTO ---
            rect_titulo = Rectangle((2.5, 9.3), 3.5, 0.4, facecolor=AZUL_OSCURO)
            ax.add_patch(rect_titulo)
            ax.text(4.25, 9.45, "VALE DE SALIDA", color='white', weight='bold', fontsize=12, ha='center')

            # --- D. DESTINO ---
            y_dest = 8.8
            ax.text(0.5, y_dest, "DESTINO / EDIFICIO:", weight='bold', color=AZUL_OSCURO, fontsize=11)
            ax.text(2.8, y_dest, area, fontsize=11)
            ax.plot([2.8, 8.0], [y_dest - 0.05, y_dest - 0.05], color='black', linewidth=1)

            # --- E. TABLA PRINCIPAL ---
            y_header = 8.0
            h_row = 0.4 
            x_start = 0.5; x_col1 = 1.5; x_col2 = 2.5; x_end = 8.0
            
            rect_head = Rectangle((x_start, y_header), x_end - x_start, h_row, facecolor=AZUL_OSCURO)
            ax.add_patch(rect_head)
            
            ax.text((x_start + x_col1)/2, y_header + 0.1, "CANT.", color='white', weight='bold', ha='center')
            ax.text((x_col1 + x_col2)/2, y_header + 0.1, "UNIDAD", color='white', weight='bold', ha='center')
            ax.text(x_col2 + 0.2, y_header + 0.1, "DESCRIPCIÓN DEL MATERIAL", color='white', weight='bold', ha='left')

            num_filas = 8
            y_curr = y_header
            for i in range(num_filas + 1): 
                color_linea = AZUL_OSCURO if i == 0 else 'gray'
                grosor = 1 if i == 0 or i == num_filas else 0.5
                ax.plot([x_start, x_end], [y_curr, y_curr], color=color_linea, linewidth=grosor)
                y_curr -= h_row
            
            y_bottom = y_header - (h_row * num_filas)
            ax.plot([x_start, x_start], [y_bottom, y_header + h_row], color=AZUL_OSCURO, linewidth=1)
            ax.plot([x_end, x_end], [y_bottom, y_header + h_row], color=AZUL_OSCURO, linewidth=1)
            ax.plot([x_col1, x_col1], [y_bottom, y_header], color='gray', linewidth=0.5)
            ax.plot([x_col2, x_col2], [y_bottom, y_header], color='gray', linewidth=0.5)

            y_data = y_header - 0.25
            ax.text((x_start + x_col1)/2, y_data, str(cantidad), ha='center', fontsize=11)
            ax.text((x_col1 + x_col2)/2, y_data, "PZA", ha='center', fontsize=11) 
            desc_ajustada = textwrap.fill(material, 65) 
            ax.text(x_col2 + 0.2, y_data, desc_ajustada, ha='left', va='center', fontsize=10)

            # --- F. FIRMAS ---
            y_firmas = 1.5
            
            def dibujar_firma(x_center, titulo, nombre):
                ax.plot([x_center - 1.0, x_center + 1.0], [y_firmas, y_firmas], color='black', linewidth=1)
                ax.text(x_center, y_firmas - 0.2, titulo, ha='center', weight='bold', fontsize=10)
                nombre_wrap = "\n".join(textwrap.wrap(nombre, width=25)) 
                ax.text(x_center, y_firmas - 0.35, nombre_wrap, ha='center', fontsize=7, va='top', color='#1F4E79')

            dibujar_firma(1.8, "ENTREGÓ", jefe)
            dibujar_firma(4.25, "RECIBIÓ / SOLICITÓ", resp)
            dibujar_firma(6.7, "CONSTAME:", firma_constame) 

            # 4. GUARDAR
            fig.savefig(ruta_pdf, dpi=300, bbox_inches='tight')
            plt.close(fig)
            
            intentos = 0
            while not os.path.exists(ruta_pdf) and intentos < 20:
                time.sleep(0.1); intentos += 1
            
            if os.path.exists(ruta_pdf): os.startfile(ruta_pdf)
            else: messagebox.showerror("Error", "No se pudo generar el archivo PDF.")

        except Exception as e:
            messagebox.showerror("Error", f"{e}")
            plt.close()

    
    def abrir_ajustes_visuales(self):
        # 1. Verificar Permisos
        if self.usuario.get('rol') != 'ADMIN' and not self.tiene_permiso('ajustes'):
            messagebox.showerror("Acceso Denegado", "No tienes permiso para modificar la configuración del sistema.")
            return

        top = tk.Toplevel(self.root)
        top.title("Configuración Visual y de Reportes")
        # Ventana un poco más grande para que quepa la nueva sección
        self.centrar_ventana_emergente(top, 900, 750) 

        # Contenedor con Scroll
        canvas = tk.Canvas(top, highlightthickness=0)
        scrollbar = ttk.Scrollbar(top, orient=VERTICAL, command=canvas.yview)
        fr = ttk.Frame(canvas, padding=30) 

        scrollable_window = canvas.create_window((0, 0), window=fr, anchor="nw")

        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(scrollable_window, width=canvas.winfo_width())
        
        canvas.bind("<Configure>", configure_scroll_region)
        fr.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        scrollbar.pack(side=RIGHT, fill=Y)

        # --- CONTENIDO ---
        ttk.Label(fr, text="Personalización del Sistema", font=("Segoe UI", 18, "bold"), bootstyle="primary").pack(pady=(0, 20), anchor=CENTER)

        # ========================================================
        # SECCIÓN 1: LOGOTIPOS (DISTRIBUCIÓN HORIZONTAL)
        # ========================================================
        fr_imgs = ttk.LabelFrame(fr, text=" 🖼️ Logotipos del Sistema ", padding=15, bootstyle="info")
        fr_imgs.pack(fill=X, pady=10)

        f_izq = ttk.Frame(fr_imgs); f_izq.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        f_der = ttk.Frame(fr_imgs); f_der.pack(side=LEFT, fill=BOTH, expand=True, padx=(10, 0))

        ttk.Label(f_izq, text="Logo Interfaz (Pantalla):", font=("Segoe UI", 9, "bold")).pack(anchor=W)
        cont_l1 = ttk.Frame(f_izq)
        cont_l1.pack(fill=X, pady=5)
        self.e_logo_app = ttk.Entry(cont_l1)
        self.e_logo_app.pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        def b_app():
            r = filedialog.askopenfilename(parent=top, filetypes=[("Imágenes", "*.png;*.jpg;*.ico")])
            if r: self.e_logo_app.delete(0, END); self.e_logo_app.insert(0, r); top.lift()
        ttk.Button(cont_l1, text="📂 Buscar", command=b_app).pack(side=LEFT)

        ttk.Label(f_der, text="Logo Reportes (PDF/Excel):", font=("Segoe UI", 9, "bold")).pack(anchor=W)
        cont_l2 = ttk.Frame(f_der)
        cont_l2.pack(fill=X, pady=5)
        self.e_logo_pdf = ttk.Entry(cont_l2)
        self.e_logo_pdf.pack(side=LEFT, fill=X, expand=True, padx=(0,5))
        def b_pdf():
            r = filedialog.askopenfilename(parent=top, filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
            if r: self.e_logo_pdf.delete(0, END); self.e_logo_pdf.insert(0, r); top.lift()
        ttk.Button(cont_l2, text="📂 Buscar", command=b_pdf).pack(side=LEFT)

        # ========================================================
        # SECCIÓN 2: TÍTULOS (DISTRIBUCIÓN HORIZONTAL)
        # ========================================================
        fr_txt = ttk.LabelFrame(fr, text=" 🏷️ Títulos de la Ventana ", padding=15, bootstyle="secondary")
        fr_txt.pack(fill=X, pady=10)

        f_t1 = ttk.Frame(fr_txt); f_t1.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
        f_t2 = ttk.Frame(fr_txt); f_t2.pack(side=LEFT, fill=BOTH, expand=True, padx=(10, 0))

        ttk.Label(f_t1, text="Título Principal (Barra Superior):", font=("Segoe UI", 9, "bold")).pack(anchor=W)
        self.e_titulo = ttk.Entry(f_t1); self.e_titulo.pack(fill=X, pady=5)

        ttk.Label(f_t2, text="Subtítulo (Descripción):", font=("Segoe UI", 9, "bold")).pack(anchor=W)
        self.e_subtitulo = ttk.Entry(f_t2); self.e_subtitulo.pack(fill=X, pady=5)

        # ========================================================
        # SECCIÓN 3: ENCABEZADOS DE REPORTES (ANCHO COMPLETO)
        # ========================================================
        fr_rep = ttk.LabelFrame(fr, text=" 📄 Membrete / Encabezados (Excel y PDF) ", padding=15, bootstyle="warning")
        fr_rep.pack(fill=X, pady=10)
        
        fr_rep.columnconfigure(1, weight=1)

        ttk.Label(fr_rep, text="Línea 1 (Institución):").grid(row=0, column=0, sticky=W, pady=5)
        self.e_h1 = ttk.Entry(fr_rep); self.e_h1.grid(row=0, column=1, sticky=EW, padx=10, pady=5)
        
        ttk.Label(fr_rep, text="Línea 2 (Sub-Institución):").grid(row=1, column=0, sticky=W, pady=5)
        self.e_h2 = ttk.Entry(fr_rep); self.e_h2.grid(row=1, column=1, sticky=EW, padx=10, pady=5)
        
        ttk.Label(fr_rep, text="Línea 3 (Dirección General):").grid(row=2, column=0, sticky=W, pady=5)
        self.e_h3 = ttk.Entry(fr_rep); self.e_h3.grid(row=2, column=1, sticky=EW, padx=10, pady=5)
        
        ttk.Label(fr_rep, text="Línea 4 (Unidad/Depto):").grid(row=3, column=0, sticky=W, pady=5)
        self.e_h4 = ttk.Entry(fr_rep); self.e_h4.grid(row=3, column=1, sticky=EW, padx=10, pady=5)

        # ========================================================
        # SECCIÓN 4: INFORMACIÓN DE LA BASE DE DATOS (SOLO ADMIN)
        # ========================================================
        if self.usuario.get('rol') == 'ADMIN':
            fr_db = ttk.LabelFrame(fr, text=" 🗄️ Ruta de la Base de Datos (Modo Admin) ", padding=15, bootstyle="danger")
            fr_db.pack(fill=X, pady=10)

            ttk.Label(fr_db, text="El sistema está conectado actualmente al siguiente archivo:", font=("Segoe UI", 9)).pack(anchor=W, pady=(0,5))
            
            e_ruta_db = ttk.Entry(fr_db, font=("Consolas", 10, "bold"))
            e_ruta_db.pack(fill=X, pady=5)
            # Insertamos la ruta real que está usando el gestor
            e_ruta_db.insert(0, self.db.ruta_db) 
            e_ruta_db.configure(state="readonly") # Para que puedan copiarlo pero no borrarlo

            def abrir_carpeta_db():
                directorio = os.path.dirname(self.db.ruta_db)
                if not directorio: directorio = os.getcwd()
                if os.path.exists(directorio):
                    os.startfile(directorio)
                else:
                    messagebox.showwarning("Aviso", "La carpeta no se puede abrir directamente.")

            ttk.Button(fr_db, text="📂 Abrir ubicación del archivo", bootstyle="outline-danger", command=abrir_carpeta_db).pack(anchor=E, pady=(5,0))

        # --- CARGA DE DATOS ---
        self.e_logo_app.insert(0, self.db.get_config("LOGO_APP") or "")
        self.e_logo_pdf.insert(0, self.db.get_config("LOGO_PDF") or "")
        self.e_titulo.insert(0, self.db.get_config("TITULO_APP") or "SISTEMA INVENTARIO")
        self.e_subtitulo.insert(0, self.db.get_config("SUBTITULO_APP") or "CONTROL DE STOCK")
        
        self.e_h1.insert(0, self.db.get_config("HEADER_L1") or "SECRETARÍA DE MARINA")
        self.e_h2.insert(0, self.db.get_config("HEADER_L2") or "SUBSECRETARÍA DE MARINA")
        self.e_h3.insert(0, self.db.get_config("HEADER_L3") or "DIRECCIÓN GENERAL DE INDUSTRIA NAVAL")
        self.e_h4.insert(0, self.db.get_config("HEADER_L4") or "UNIDAD DE INVESTIGACIÓN Y DESARROLLO TECNOLÓGICO")

        # --- GUARDAR ---
        def guardar_cambios():
            self.db.set_config("LOGO_APP", self.e_logo_app.get().strip())
            self.db.set_config("LOGO_PDF", self.e_logo_pdf.get().strip())
            self.db.set_config("TITULO_APP", self.e_titulo.get().strip())
            self.db.set_config("SUBTITULO_APP", self.e_subtitulo.get().strip())
            
            self.db.set_config("HEADER_L1", self.e_h1.get().strip())
            self.db.set_config("HEADER_L2", self.e_h2.get().strip())
            self.db.set_config("HEADER_L3", self.e_h3.get().strip())
            self.db.set_config("HEADER_L4", self.e_h4.get().strip())

            if messagebox.askyesno("Reiniciar", "Configuración guardada.\n¿Reiniciar sistema ahora para ver cambios?"):
                import sys, subprocess
                top.destroy(); self.root.destroy()
                script = f'"{sys.argv[0]}"' if " " in sys.argv[0] else sys.argv[0]
                subprocess.Popen(f"{sys.executable} {script}", shell=True)
                sys.exit()
            else:
                top.destroy()

        ttk.Button(fr, text="💾 GUARDAR TODA LA CONFIGURACIÓN", bootstyle="success", command=guardar_cambios).pack(fill=X, pady=20)

        

        # --- GUARDAR ---
        
    def setup_tab_auditoria(self):
        # 1. BARRA DE FILTROS SUPERIOR
        fr_top = ttk.Frame(self.tab_audit, padding=10)
        fr_top.pack(fill=X)
        
        ttk.Label(fr_top, text="Mes:").pack(side=LEFT)
        self.cb_mes_k = ttk.Combobox(fr_top, values=[str(i) for i in range(1, 13)], width=3, state="readonly")
        self.cb_mes_k.current(datetime.now().month - 1)
        self.cb_mes_k.pack(side=LEFT, padx=5)

        ttk.Label(fr_top, text="Año:").pack(side=LEFT)
        self.ent_anio_k = ttk.Entry(fr_top, width=6)
        self.ent_anio_k.insert(0, str(datetime.now().year))
        self.ent_anio_k.pack(side=LEFT, padx=5)

        ttk.Label(fr_top, text="Filtrar Partida:").pack(side=LEFT, padx=(15, 5))
        self.cb_partida_k = ttk.Combobox(fr_top, state="readonly", width=15)
        self.cb_partida_k.pack(side=LEFT)
        # Se llena con self.actualizar_combos()

        ttk.Button(fr_top, text="🔍 Generar Vista Previa", bootstyle="primary", command=self.generar_vista_anexo_c).pack(side=LEFT, padx=15)
        ttk.Button(fr_top, text="💾 Exportar Excel (Anexo C)", bootstyle="success", command=self.exportar_excel_anexo_c).pack(side=RIGHT)

        # 2. TABLA TIPO EXCEL (Treeview Complejo)
        fr_tabla = ttk.Frame(self.tab_audit)
        fr_tabla.pack(fill=BOTH, expand=True, pady=5)

        # Scrollbars dobles
        sc_y = ttk.Scrollbar(fr_tabla, orient=VERTICAL)
        sc_x = ttk.Scrollbar(fr_tabla, orient=HORIZONTAL)

        # Columnas: NP, UNIDAD, DESC, FACTURA, EX_ANT, RECIBIDOS, 1..31, TOTAL, EX_ACT
        dias = [str(d) for d in range(1, 32)]
        cols = ["NP", "UNIDAD", "DESC", "FACTURA", "EX_ANT", "RECIBIDOS"] + dias + ["TOTAL_SAL", "EX_ACT"]
        
        self.tree_kardex = ttk.Treeview(fr_tabla, columns=cols, show="headings", 
                                        yscrollcommand=sc_y.set, xscrollcommand=sc_x.set, selectmode="browse")
        
        sc_y.config(command=self.tree_kardex.yview); sc_y.pack(side=RIGHT, fill=Y)
        sc_x.config(command=self.tree_kardex.xview); sc_x.pack(side=BOTTOM, fill=X)
        self.tree_kardex.pack(side=LEFT, fill=BOTH, expand=True)

        # Configurar Encabezados
        self.tree_kardex.heading("NP", text="N.P."); self.tree_kardex.column("NP", width=35, stretch=NO)
        self.tree_kardex.heading("UNIDAD", text="U."); self.tree_kardex.column("UNIDAD", width=40, stretch=NO)
        self.tree_kardex.heading("DESC", text="DESCRIPCIÓN"); self.tree_kardex.column("DESC", width=200, minwidth=150)
        self.tree_kardex.heading("FACTURA", text="FACTURA"); self.tree_kardex.column("FACTURA", width=80)
        self.tree_kardex.heading("EX_ANT", text="E.ANT"); self.tree_kardex.column("EX_ANT", width=50, anchor=CENTER)
        self.tree_kardex.heading("RECIBIDOS", text="ENT."); self.tree_kardex.column("RECIBIDOS", width=50, anchor=CENTER)
        
        for d in dias:
            self.tree_kardex.heading(d, text=d)
            self.tree_kardex.column(d, width=25, stretch=NO, anchor=CENTER)
            
        self.tree_kardex.heading("TOTAL_SAL", text="T.SAL"); self.tree_kardex.column("TOTAL_SAL", width=50, anchor=CENTER)
        self.tree_kardex.heading("EX_ACT", text="ACT."); self.tree_kardex.column("EX_ACT", width=50, anchor=CENTER)

    def abrir_editor_temas(self):
        top = tk.Toplevel(self.root)
        top.title("🎨 Personalización de Temas")
        
        # Usamos tu función de centrado
        self.centrar_ventana_emergente(top, 750, 700)
        
        tabs = ttk.Notebook(top)
        tabs.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # --- PESTAÑA 1: TEMAS PREDEFINIDOS ---
        tab_predef = ttk.Frame(tabs, padding=20)
        tabs.add(tab_predef, text="🎨 Temas Predefinidos")
        
        ttk.Label(tab_predef, text="Selecciona un tema:", font=("Segoe UI", 12, "bold")).pack(anchor=W, pady=(0, 15))
        
        # Canvas con scroll para los temas
        canvas_temas = tk.Canvas(tab_predef, height=400)
        scroll_temas = ttk.Scrollbar(tab_predef, orient=VERTICAL, command=canvas_temas.yview)
        frame_temas = ttk.Frame(canvas_temas)
        
        canvas_temas.configure(yscrollcommand=scroll_temas.set)
        canvas_temas.pack(side=LEFT, fill=BOTH, expand=True)
        scroll_temas.pack(side=RIGHT, fill=Y)
        canvas_temas.create_window((0, 0), window=frame_temas, anchor=NW)
        
        def seleccionar_tema_predef(nombre_tema):
            tema = GestorTemas.TEMAS_PREDEFINIDOS[nombre_tema]
            GestorTemas.guardar_tema(self.db, tema)
            top.destroy()
            # IMPORTANTE: Llama a la función de reinicio que creamos antes
            self.solicitar_reinicio()
        
        # Generar lista de temas
        for nombre, tema_data in GestorTemas.TEMAS_PREDEFINIDOS.items():
            fr_tema = ttk.Frame(frame_temas, padding=10, relief="solid", borderwidth=1)
            fr_tema.pack(fill=X, pady=5, padx=5)
            
            # Muestra de colores
            fr_colores = ttk.Frame(fr_tema)
            fr_colores.pack(side=LEFT, padx=(0, 15))
            for color in [tema_data["color_primario"], tema_data["color_secundario"], tema_data["color_acento"]]:
                lbl_color = tk.Label(fr_colores, bg=color, width=3, height=1, relief="solid", borderwidth=1)
                lbl_color.pack(side=LEFT, padx=2)
            
            ttk.Label(fr_tema, text=nombre, font=("Segoe UI", 11, "bold")).pack(side=LEFT)
            ttk.Button(fr_tema, text="✓ Aplicar", bootstyle="success",
                    command=lambda n=nombre: seleccionar_tema_predef(n)).pack(side=RIGHT)
        
        frame_temas.update_idletasks()
        canvas_temas.configure(scrollregion=canvas_temas.bbox("all"))
        
        # --- PESTAÑA 2: PERSONALIZADO ---
        tab_custom = ttk.Frame(tabs, padding=20)
        tabs.add(tab_custom, text="🖌️ Personalizado")
        
        colores_personalizados = {
            "color_primario": tk.StringVar(value=self.tema_actual["color_primario"]),
            "color_secundario": tk.StringVar(value=self.tema_actual["color_secundario"]),
            "color_acento": tk.StringVar(value=self.tema_actual["color_acento"]),
            "color_fondo": tk.StringVar(value=self.tema_actual["color_fondo"]),
            "color_texto": tk.StringVar(value=self.tema_actual["color_texto"])
        }
        
        def elegir_color(clave, var):
            color = colorchooser.askcolor(initialcolor=var.get(), title=f"Seleccionar {clave}")
            if color[1]: var.set(color[1])
        
        opciones_color = [
            ("Color Primario", "color_primario"), ("Color Secundario", "color_secundario"),
            ("Color Acento", "color_acento"), ("Color Fondo", "color_fondo"), ("Color Texto", "color_texto")
        ]
        
        for texto, clave in opciones_color:
            fr_color = ttk.Frame(tab_custom, padding=5); fr_color.pack(fill=X, pady=8)
            ttk.Label(fr_color, text=texto, width=20).pack(side=LEFT)
            ttk.Button(fr_color, text="🎨", bootstyle="info", command=lambda c=clave, v=colores_personalizados[clave]: elegir_color(c, v)).pack(side=LEFT)
        
        ttk.Separator(tab_custom).pack(fill=X, pady=20)
        ttk.Label(tab_custom, text="Tema Base:", font=("Segoe UI", 10, "bold")).pack(anchor=W)
        
        # Lista de temas bootstrap disponibles
        temas_bootstrap = ["flatly", "cosmo", "litera", "minty", "pulse", "sandstone", "united", "yeti", "darkly", "superhero", "solar", "cyborg"]
        tema_bootstrap_var = tk.StringVar(value=self.tema_actual["tema_bootstrap"])
        ttk.Combobox(tab_custom, textvariable=tema_bootstrap_var, values=temas_bootstrap, state="readonly").pack(fill=X)
        
        def guardar_tema_personalizado():
            tema_nuevo = {k: v.get() for k, v in colores_personalizados.items()}
            tema_nuevo["tema_bootstrap"] = tema_bootstrap_var.get()
            GestorTemas.guardar_tema(self.db, tema_nuevo)
            top.destroy()
            self.solicitar_reinicio()
        
        ttk.Button(tab_custom, text="💾 GUARDAR Y REINICIAR", bootstyle="success", command=guardar_tema_personalizado).pack(fill=X, pady=20)

    def calcular_datos_kardex(self):
        try:
            mes = int(self.cb_mes_k.get())
            anio = int(self.ent_anio_k.get())
            partida_sel = self.cb_partida_k.get()
        except: messagebox.showerror("Error", "Verifica Mes y Año"); return None

        inicio_mes = datetime(anio, mes, 1)
        ultimo_dia = calendar.monthrange(anio, mes)[1]
        fin_mes = datetime(anio, mes, ultimo_dia, 23, 59, 59)

        # Seleccionar materiales (Filtrado o todos)
        sql = "SELECT id, partida, material FROM inventario"
        params = []
        if partida_sel and partida_sel != "TODAS":
            sql += " WHERE partida = ?"
            params.append(partida_sel)
        sql += " ORDER BY partida, material"
        
        materiales = self.db.consultar(sql, tuple(params))
        datos_procesados = []

        for idx, mat in enumerate(materiales, 1):
            mat_nom = mat['material']
            
            # Obtener TODO el historial de este material
            historial = self.db.consultar("SELECT fecha_hora, tipo, cantidad, factura FROM historial WHERE material = ?", (mat_nom,))
            
            ex_ant = 0
            entradas_mes = 0
            salidas_dias = {d: 0 for d in range(1, 32)}
            facturas_mes = set() 
            
            for h in historial:
                try: 
                    # Intentar leer fecha con y sin hora
                    try: f_obj = datetime.strptime(h['fecha_hora'], "%d/%m/%Y %H:%M")
                    except: f_obj = datetime.strptime(h['fecha_hora'], "%d/%m/%Y")
                except: continue

                cant = h['cantidad']
                tipo = h['tipo'].upper() # Convertimos a mayúsculas para comparar mejor
                
                # --- CORRECCIÓN FUERTE: Detectar cualquier tipo de entrada ---
                # Si dice ENTRADA, HISTORICO (+) o ALTA, es una suma.
                es_entrada = ("ENTRADA" in tipo or "(+)" in tipo or "ALTA" in tipo)
                
                # 1. Movimientos ANTERIORES al mes (Para Saldo Anterior)
                if f_obj < inicio_mes:
                    if es_entrada: ex_ant += cant
                    else: ex_ant -= cant
                
                # 2. Movimientos DURANTE el mes
                elif inicio_mes <= f_obj <= fin_mes:
                    if es_entrada:
                        entradas_mes += cant
                        # Capturar facturas si existen
                        if h['factura'] and h['factura'] != "S/F": facturas_mes.add(h['factura'])
                    else:
                        # Todo lo que no sea entrada se considera SALIDA o AJUSTE (-)
                        dia = f_obj.day
                        salidas_dias[dia] += cant
            
            # Formatear facturas
            str_facturas = ", ".join(facturas_mes) if facturas_mes else ""

            total_sal = sum(salidas_dias.values())
            
            # CALCULO FINAL: Anterior + Entradas - Salidas
            ex_act = (ex_ant + entradas_mes) - total_sal
            
            row = {
                "NP": idx, "UNIDAD": "PZA", "DESC": mat_nom, "FACTURA": str_facturas,
                "EX_ANT": ex_ant, "RECIBIDOS": entradas_mes, "SALIDAS_DIAS": salidas_dias,
                "TOTAL_SAL": total_sal, "EX_ACT": ex_act, "PARTIDA": mat['partida']
            }
            datos_procesados.append(row)
            
        return datos_procesados, mes, anio, partida_sel
    
    def generar_vista_anexo_c(self):
        res = self.calcular_datos_kardex()
        if not res: return
        datos, _, _, _ = res
        
        # Limpiar y llenar
        for i in self.tree_kardex.get_children(): self.tree_kardex.delete(i)
        
        for d in datos:
            # Lista de valores para las columnas 1-31
            vals_dias = [d['SALIDAS_DIAS'][dia] if d['SALIDAS_DIAS'][dia] > 0 else "" for dia in range(1, 32)]
            
            valores = [d['NP'], d['UNIDAD'], d['DESC'], d['FACTURA'], d['EX_ANT'], d['RECIBIDOS']] + vals_dias + [d['TOTAL_SAL'], d['EX_ACT']]
            self.tree_kardex.insert("", END, values=valores)
    
    # ------------------------------------------------------------------
  
    # ------------------------------------------------------------------
  


    def exportar_excel_anexo_c(self):
        # Importación segura
        from openpyxl.styles import Font, Alignment, Border, Side as ExcelSide
        
        res = self.calcular_datos_kardex()
        if not res: return
        datos, mes, anio, partida_sel = res
        
        ruta = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not ruta: return

        wb = Workbook()
        ws = wb.active
        
        nombres_meses = ["", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", 
                         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]
        nombre_mes = nombres_meses[mes] 
        
        ws.title = f"ANEXO C {nombre_mes}-{anio}"
        
        # ESTILOS
        thin = ExcelSide(border_style="thin", color="000000") 
        borde_todo = Border(top=thin, left=thin, right=thin, bottom=thin)
        centrado = Alignment(horizontal='center', vertical='center', wrap_text=True)
        negrita = Font(bold=True, size=10, name='Arial')
        
        # --- ENCABEZADOS DINÁMICOS (DESDE DB) ---
        ws.merge_cells('A1:AK1'); ws['A1'] = "ANEXO C"
        
        # Recuperar configuración o usar default si está vacío
        h1 = self.db.get_config("HEADER_L1") or "INSTITUCIÓN"
        h2 = self.db.get_config("HEADER_L2") or "SUBDIRECCIÓN"
        h3 = self.db.get_config("HEADER_L3") or "DIRECCIÓN GENERAL"
        h4 = self.db.get_config("HEADER_L4") or "DEPARTAMENTO"
        
        ws.merge_cells('A2:AK2'); ws['A2'] = h1
        ws.merge_cells('A3:AK3'); ws['A3'] = h2
        ws.merge_cells('A4:AK4'); ws['A4'] = h3
        ws.merge_cells('A5:AK5'); ws['A5'] = h4
        
        # LÓGICA DE NOMBRE LARGO
        txt_partida = "TODAS LAS PARTIDAS"
        if partida_sel and partida_sel != "TODAS":
            res_desc = self.db.consultar("SELECT descripcion FROM partidas_desc WHERE codigo=?", (partida_sel,))
            nombre_largo = res_desc[0]['descripcion'] if res_desc else ""
            txt_partida = f"PARTIDA {partida_sel} {nombre_largo}"
            
        ws.merge_cells('A7:AK7')
        ws['A7'] = f"CONTROL DE LA {txt_partida} CORRESPONDIENTE AL MES DE {nombre_mes} {anio}"
        
        for row in ws.iter_rows(min_row=1, max_row=7):
            for cell in row:
                cell.alignment = centrado
                cell.font = negrita

        # --- CABECERA DE TABLA ---
        headers_fijos = [("A", "N.P."), ("B", "UNIDAD"), ("C", "DESCRIPCIÓN"), 
                         ("D", "FACTURA"), ("E", "EXIST.\nANT."), ("F", "EFECTOS\nRECIBIDOS")]
        
        for col, texto in headers_fijos:
            ws.merge_cells(f'{col}9:{col}10')
            cell = ws[f'{col}9']
            cell.value = texto
            cell.alignment = centrado
            cell.font = Font(bold=True, size=8)
            cell.border = borde_todo
            ws[f'{col}10'].border = borde_todo 

        ws.merge_cells('G9:AK9')
        ws['G9'] = "SALIDAS (DÍAS DEL MES)"
        ws['G9'].alignment = centrado; ws['G9'].font = Font(bold=True, size=8); ws['G9'].border = borde_todo
        
        col_idx = 7
        for dia in range(1, 32):
            cell = ws.cell(row=10, column=col_idx, value=dia)
            cell.alignment = centrado; cell.font = Font(size=8); cell.border = borde_todo
            ws.column_dimensions[chr(64+col_idx) if col_idx <= 26 else f"A{chr(64+col_idx-26)}"].width = 3.5
            col_idx += 1

        ws.merge_cells('AL9:AL10'); ws['AL9'] = "TOTAL\nSALIDA"
        ws.merge_cells('AM9:AM10'); ws['AM9'] = "EXIST.\nACT."
        for col in ['AL', 'AM']:
            cell = ws[f'{col}9']
            cell.alignment = centrado; cell.font = Font(bold=True, size=8); cell.border = borde_todo
            ws[f'{col}10'].border = borde_todo

        # --- DATOS ---
        fila_act = 11
        for d in datos:
            ws.cell(row=fila_act, column=1, value=d['NP']).border = borde_todo
            ws.cell(row=fila_act, column=2, value=d['UNIDAD']).border = borde_todo
            ws.cell(row=fila_act, column=3, value=d['DESC']).border = borde_todo
            ws.cell(row=fila_act, column=4, value=d['FACTURA']).border = borde_todo
            ws.cell(row=fila_act, column=5, value=d['EX_ANT']).border = borde_todo
            ws.cell(row=fila_act, column=6, value=d['RECIBIDOS']).border = borde_todo
            
            col_d = 7
            for dia in range(1, 32):
                val = d['SALIDAS_DIAS'][dia]
                ws.cell(row=fila_act, column=col_d, value=val if val > 0 else "").border = borde_todo
                col_d += 1
            
            ws.cell(row=fila_act, column=38, value=d['TOTAL_SAL']).border = borde_todo
            ws.cell(row=fila_act, column=39, value=d['EX_ACT']).border = borde_todo
            fila_act += 1
            
        ws.column_dimensions['C'].width = 40 
        ws.column_dimensions['D'].width = 15 

        # --- PIE DE PÁGINA (FIRMAS) ---
        fila_firmas = fila_act + 4
        firmas = [
            ("ELABORÓ", "B", "E"), 
            ("SUPERVISÓ", "M", "P"), 
            ("Vo. Bo.", "S", "V"), 
            ("CONSTAME", "AA", "AE")
        ]
        
        top_line = ExcelSide(border_style="medium", color="000000")
        for titulo, col_inicio, col_fin in firmas:
            ws.merge_cells(f'{col_inicio}{fila_firmas}:{col_fin}{fila_firmas}')
            cell = ws[f'{col_inicio}{fila_firmas}']
            cell.value = titulo
            cell.alignment = centrado; cell.font = negrita
            for c_idx in range(ws[f'{col_inicio}1'].column, ws[f'{col_fin}1'].column + 1):
                ws.cell(row=fila_firmas, column=c_idx).border = Border(top=top_line)

        try:
            wb.save(ruta)
            messagebox.showinfo("Éxito", "Reporte Anexo C generado correctamente.")
            os.startfile(ruta)
        except PermissionError:
            messagebox.showwarning("Archivo Abierto", 
                                   f"No se pudo guardar el archivo.\n\n"
                                   f"El archivo '{os.path.basename(ruta)}' está abierto.\n"
                                   "Por favor, CIÉRRALO y vuelve a intentar.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el Excel: {e}")

    def centrar_ventana_emergente(self, ventana, ancho, alto):
        """Centra una ventana Toplevel en la pantalla y define su tamaño mínimo"""
        ventana.update_idletasks() # Necesario para cálculos precisos
        
        # Obtener dimensiones de la pantalla
        ws = ventana.winfo_screenwidth()
        hs = ventana.winfo_screenheight()
        
        # Calcular posición X e Y
        x = (ws // 2) - (ancho // 2)
        y = (hs // 2) - (alto // 2)
        
        # Aplicar geometría y tamaño mínimo
        ventana.geometry(f"{ancho}x{alto}+{int(x)}+{int(y)}")
        ventana.minsize(ancho, alto) # Evita que se haga demasiado pequeña y corte botones
        
        # Asegurar que se pueda maximizar (True, True es el defecto, pero confirmamos)
        ventana.resizable(True, True) 
        
        # Ponerla al frente
        ventana.lift()
        ventana.focus_force()


    def solicitar_reinicio(self):
        """Pregunta y ejecuta el reinicio automático (CORREGIDO PARA RUTAS CON ESPACIOS)"""
        respuesta = messagebox.askyesno(
            "Configuración Guardada", 
            "Para que los cambios visuales se apliquen correctamente, es necesario reiniciar el sistema.\n\n"
            "¿Deseas reiniciar AHORA?",
            icon='warning'
        )
        if respuesta:
            # 1. Cerrar ventana actual
            self.root.destroy()
            
            # 2. Preparar el reinicio manejando ESPACIOS en la ruta
            python = sys.executable
            script = sys.argv[0]
            
            # Si el nombre del archivo o la carpeta tiene espacios, le ponemos comillas
            if " " in script:
                script = f'"{script}"'
            
            # 3. Reiniciar
            # os.execl reemplaza el proceso actual por uno nuevo
            try:
                os.execl(python, python, script, *sys.argv[1:])
            except Exception as e:
                # Si falla el reinicio automático, avisamos
                print(f"Error al reiniciar: {e}")
                sys.exit()
        
    # --- EN LA CLASE SistemaInventario (NUEVA FUNCIÓN) ---
    def abrir_gestion_usuarios(self):
        # 1. Seguridad: Solo ADMIN puede entrar aquí
        if self.usuario.get('rol') != 'ADMIN':
            messagebox.showerror("Acceso Denegado", "Se requieren permisos de administrador.")
            return

        top = tk.Toplevel(self.root)
        top.title("Gestión de Usuarios y Permisos")
        top.state('zoomed') # Pantalla completa
        top.minsize(900, 600)
        
        # --- VARIABLES DE CONTROL ---
        var_id = tk.StringVar()
        var_user = tk.StringVar()
        var_pass = tk.StringVar()
        var_nombre = tk.StringVar()
        var_email = tk.StringVar() # <--- VARIABLE PARA EL EMAIL
        var_rol = tk.StringVar(value="OPERADOR")
        var_foto_path = tk.StringVar()
        
        # Variables de Permisos
        var_p_crear = tk.IntVar(); var_p_ent = tk.IntVar(); var_p_sal = tk.IntVar()
        var_p_edit = tk.IntVar(); var_p_del = tk.IntVar()
        var_p_cat = tk.IntVar(); var_p_hist = tk.IntVar(); var_p_conf = tk.IntVar()
        
        self.usuario_seleccionado_id = None 
        self.password_actual_hash = "" 

        try:
            import tkinter.ttk as original_ttk
            # Frame principal con padding
            main_container = ttk.Frame(top, padding=10)
            main_container.pack(fill=BOTH, expand=True)

            paned = original_ttk.PanedWindow(main_container, orient=HORIZONTAL)
            paned.pack(fill=BOTH, expand=True)

            # ==========================================
            # PANEL IZQUIERDO: LISTA DE USUARIOS
            # ==========================================
            fr_lista = ttk.Labelframe(paned, text=" Usuarios Registrados ", padding=10, bootstyle="info")
            paned.add(fr_lista, weight=1) 

            cols_u = ("ID", "USUARIO", "ROL", "NOMBRE")
            tree_users = ttk.Treeview(fr_lista, columns=cols_u, show="headings", selectmode="browse")
            
            tree_users.heading("ID", text="ID"); tree_users.column("ID", width=40, anchor=CENTER)
            tree_users.heading("USUARIO", text="Usuario"); tree_users.column("USUARIO", width=100)
            tree_users.heading("ROL", text="Rol"); tree_users.column("ROL", width=80, anchor=CENTER)
            tree_users.heading("NOMBRE", text="Nombre"); tree_users.column("NOMBRE", width=150)
            
            sc_u = ttk.Scrollbar(fr_lista, orient=VERTICAL, command=tree_users.yview)
            tree_users.configure(yscrollcommand=sc_u.set)
            
            tree_users.pack(side=LEFT, fill=BOTH, expand=True)
            sc_u.pack(side=RIGHT, fill=Y)

            ttk.Label(fr_lista, text="💡 Selecciona un usuario para editarlo", 
                      font=("Segoe UI", 8), bootstyle="secondary").pack(side=BOTTOM, fill=X)

            # ==========================================
            # PANEL DERECHO: FORMULARIO DE EDICIÓN
            # ==========================================
            fr_form = ttk.Labelframe(paned, text=" Ficha de Usuario y Permisos ", padding=15, bootstyle="primary")
            paned.add(fr_form, weight=3) 

            # Canvas para scroll si la pantalla es chica
            canvas_form = tk.Canvas(fr_form, highlightthickness=0)
            scrollbar_form = ttk.Scrollbar(fr_form, orient=VERTICAL, command=canvas_form.yview)
            scrollable_frame = ttk.Frame(canvas_form)

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas_form.configure(scrollregion=canvas_form.bbox("all"))
            )

            canvas_form.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas_form.configure(yscrollcommand=scrollbar_form.set)

            canvas_form.pack(side=LEFT, fill=BOTH, expand=True)
            scrollbar_form.pack(side=RIGHT, fill=Y)

            # --- SECCIÓN FOTO ---
            fr_foto = ttk.Frame(scrollable_frame)
            fr_foto.pack(fill=X, pady=(0, 10))
            self.lbl_preview = ttk.Label(fr_foto, text="👤", font=("Segoe UI Emoji", 40), anchor=CENTER)
            self.lbl_preview.pack(side=LEFT, padx=10)
            
            def seleccionar_foto():
                ruta = filedialog.askopenfilename(parent=top, filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
                if ruta:
                    var_foto_path.set(ruta)
                    try:
                        img = Image.open(ruta).resize((70, 70), Image.LANCZOS)
                        self.tk_foto_temp = ImageTk.PhotoImage(img, master=top)
                        self.lbl_preview.configure(image=self.tk_foto_temp, text="")
                    except: pass
                top.lift()
            
            ttk.Button(fr_foto, text="📂 Cambiar Foto", bootstyle="secondary-outline", command=seleccionar_foto).pack(side=LEFT, padx=10)

            # --- CAMPOS DE TEXTO ---
            fr_campos = ttk.Frame(scrollable_frame)
            fr_campos.pack(fill=X)
            fr_campos.columnconfigure(1, weight=1)
            
            # Fila 0: Usuario
            ttk.Label(fr_campos, text="Usuario (Login):").grid(row=0, column=0, sticky=W, pady=5)
            ttk.Entry(fr_campos, textvariable=var_user, font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky=EW, pady=5, padx=5)
            
            # Fila 1: Contraseña
            ttk.Label(fr_campos, text="Contraseña:").grid(row=1, column=0, sticky=W, pady=5)
            fr_pass = ttk.Frame(fr_campos)
            fr_pass.grid(row=1, column=1, sticky=EW, pady=5, padx=5)
            e_pass = ttk.Entry(fr_pass, textvariable=var_pass, show="*")
            e_pass.pack(side=LEFT, fill=X, expand=True)
            ver_pass = tk.BooleanVar()
            ttk.Checkbutton(fr_pass, text="👁️", variable=ver_pass, bootstyle="toolbutton", command=lambda: e_pass.config(show="" if ver_pass.get() else "*")).pack(side=LEFT)
            
            # Fila 2: Nota Password
            self.lbl_help_pass = ttk.Label(fr_campos, text="* Obligatoria para nuevos", font=("Segoe UI", 8), bootstyle="secondary")
            self.lbl_help_pass.grid(row=2, column=1, sticky=W)

            # Fila 3: Nombre Completo
            ttk.Label(fr_campos, text="Nombre Completo:").grid(row=3, column=0, sticky=W, pady=5)
            ttk.Entry(fr_campos, textvariable=var_nombre).grid(row=3, column=1, sticky=EW, pady=5, padx=5)

            # --- Fila 4: CORREO ELECTRÓNICO (NUEVO) ---
            ttk.Label(fr_campos, text="Correo Electrónico:").grid(row=4, column=0, sticky=W, pady=5)
            ttk.Entry(fr_campos, textvariable=var_email).grid(row=4, column=1, sticky=EW, pady=5, padx=5)

            # Fila 5: Rol
            ttk.Label(fr_campos, text="Rol (Etiqueta):").grid(row=5, column=0, sticky=W, pady=5)
            cbox_rol = ttk.Combobox(fr_campos, textvariable=var_rol, values=["OPERADOR", "ADMIN", "SOLO LECTURA"], state="readonly")
            cbox_rol.grid(row=5, column=1, sticky=EW, pady=5, padx=5)

            # --- SECCIÓN PERMISOS ---
            fr_permisos = ttk.LabelFrame(scrollable_frame, text=" Configuración de Accesos ", padding=10, bootstyle="warning")
            fr_permisos.pack(fill=X, pady=15)
            
            col1 = ttk.Frame(fr_permisos); col1.pack(side=LEFT, fill=Y, expand=True, padx=5)
            col2 = ttk.Frame(fr_permisos); col2.pack(side=LEFT, fill=Y, expand=True, padx=5)

            ttk.Label(col1, text="Operativos", font=("Segoe UI", 8, "bold"), foreground="gray").pack(anchor=W)
            ttk.Checkbutton(col1, text="Crear Materiales", variable=var_p_crear, bootstyle="round-toggle").pack(anchor=W, pady=2)
            ttk.Checkbutton(col1, text="Registrar ENTRADAS", variable=var_p_ent, bootstyle="round-toggle").pack(anchor=W, pady=2)
            ttk.Checkbutton(col1, text="Registrar SALIDAS", variable=var_p_sal, bootstyle="round-toggle").pack(anchor=W, pady=2)
            ttk.Checkbutton(col1, text="Editar Maestros", variable=var_p_edit, bootstyle="round-toggle").pack(anchor=W, pady=2)
            
            ttk.Label(col2, text="Módulos Admin", font=("Segoe UI", 8, "bold"), foreground="gray").pack(anchor=W)
            ttk.Checkbutton(col2, text="Gestión Catálogos", variable=var_p_cat, bootstyle="round-toggle").pack(anchor=W, pady=2)
            ttk.Checkbutton(col2, text="Modificar Histórico", variable=var_p_hist, bootstyle="round-toggle").pack(anchor=W, pady=2)
            ttk.Checkbutton(col2, text="Ajustes (Logos)", variable=var_p_conf, bootstyle="round-toggle").pack(anchor=W, pady=2)
            
            # --- Lógica de Roles ---
            def al_cambiar_rol(event):
                rol = var_rol.get()
                if rol == "ADMIN":
                    var_p_crear.set(1); var_p_ent.set(1); var_p_sal.set(1); var_p_edit.set(1); var_p_del.set(1)
                    var_p_cat.set(1); var_p_hist.set(1); var_p_conf.set(1)
                elif rol == "OPERADOR":
                    var_p_crear.set(1); var_p_ent.set(1); var_p_sal.set(1); var_p_edit.set(0); var_p_del.set(0)
                    var_p_cat.set(0); var_p_hist.set(0); var_p_conf.set(0)
                elif rol == "SOLO LECTURA":
                    var_p_crear.set(0); var_p_ent.set(0); var_p_sal.set(0); var_p_edit.set(0); var_p_del.set(0)
                    var_p_cat.set(0); var_p_hist.set(0); var_p_conf.set(0)
            cbox_rol.bind("<<ComboboxSelected>>", al_cambiar_rol)

            # ==========================================
            # FUNCIONES INTERNAS (CRUD USUARIOS)
            # ==========================================
            def limpiar_form():
                self.usuario_seleccionado_id = None; self.password_actual_hash = ""
                var_user.set(""); var_pass.set(""); var_nombre.set(""); var_email.set(""); var_foto_path.set("")
                var_rol.set("OPERADOR"); self.lbl_preview.configure(image="", text="👤")
                tree_users.selection_remove(tree_users.selection())
                al_cambiar_rol(None)
                btn_guardar.configure(text="💾 CREAR NUEVO", bootstyle="success")

            def cargar_usuarios_en_lista():
                for item in tree_users.get_children(): tree_users.delete(item)
                filas = self.db.consultar("SELECT id, usuario, rol, nombre_completo FROM usuarios ORDER BY usuario ASC")
                for f in filas:
                    tree_users.insert("", END, values=(f['id'], f['usuario'], f['rol'], f['nombre_completo']))

            def llenar_formulario(event):
                sel = tree_users.selection()
                if not sel: return
                id_u = tree_users.item(sel[0])['values'][0]
                self.usuario_seleccionado_id = id_u
                
                # Consultamos datos completos
                datos = self.db.consultar("SELECT * FROM usuarios WHERE id=?", (id_u,))
                if datos:
                    u = dict(datos[0])
                    var_user.set(u['usuario']); self.password_actual_hash = u['password']; var_pass.set("")
                    var_nombre.set(u.get('nombre_completo',''))
                    var_email.set(u.get('email','')) # <--- CARGAR EMAIL
                    var_rol.set(u['rol'])
                    var_foto_path.set(u.get('foto_path',''))
                    
                    import json
                    try:
                        p = json.loads(u.get('permisos','{}'))
                        var_p_crear.set(p.get('crear',0)); var_p_ent.set(p.get('entrada',0)); var_p_sal.set(p.get('salida',0))
                        var_p_edit.set(p.get('editar',0)); var_p_del.set(p.get('eliminar',0))
                        var_p_cat.set(p.get('catalogos',0)); var_p_hist.set(p.get('historico',0)); var_p_conf.set(p.get('ajustes',0))
                    except: al_cambiar_rol(None)
                    
                    if u.get('foto_path') and os.path.exists(u['foto_path']):
                         try:
                            img = Image.open(u['foto_path']).resize((70, 70), Image.LANCZOS)
                            self.tk_foto_temp = ImageTk.PhotoImage(img, master=top)
                            self.lbl_preview.configure(image=self.tk_foto_temp, text="")
                         except: pass
                    else: self.lbl_preview.configure(image="", text="👤")
                    btn_guardar.configure(text="💾 ACTUALIZAR", bootstyle="warning")

            tree_users.bind("<<TreeviewSelect>>", llenar_formulario)

            def guardar():
                u = var_user.get().strip().upper()
                p = var_pass.get().strip()
                r = var_rol.get()
                n = var_nombre.get().strip().upper() or u
                e = var_email.get().strip() # <--- OBTENER EMAIL
                f = var_foto_path.get()
                
                import json
                permisos = json.dumps({"crear":var_p_crear.get(), "entrada":var_p_ent.get(), "salida":var_p_sal.get(), 
                                       "editar":var_p_edit.get(), "eliminar":var_p_del.get(), "catalogos":var_p_cat.get(),
                                       "historico":var_p_hist.get(), "ajustes":var_p_conf.get()})
                
                if not u: return
                p_fin = self.password_actual_hash if (self.usuario_seleccionado_id and not p) else hashlib.sha256(p.encode()).hexdigest()
                
                if self.usuario_seleccionado_id:
                    self.db.ejecutar("UPDATE usuarios SET usuario=?, password=?, rol=?, nombre_completo=?, email=?, foto_path=?, permisos=? WHERE id=?", 
                                     (u, p_fin, r, n, e, f, permisos, self.usuario_seleccionado_id))
                else:
                    self.db.ejecutar("INSERT INTO usuarios (usuario, password, rol, nombre_completo, email, foto_path, permisos) VALUES (?,?,?,?,?,?,?)",
                                     (u, p_fin, r, n, e, f, permisos))
                
                limpiar_form(); cargar_usuarios_en_lista()
                messagebox.showinfo("Éxito", "Usuario guardado correctamente.")

            def eliminar():
                if self.usuario_seleccionado_id and messagebox.askyesno("Borrar", "¿Eliminar usuario?"):
                    self.db.ejecutar("DELETE FROM usuarios WHERE id=?", (self.usuario_seleccionado_id,))
                    limpiar_form(); cargar_usuarios_en_lista()

            # --- BOTONERA INFERIOR ---
            fr_btns = ttk.Frame(fr_form, padding=(0, 10))
            fr_btns.pack(side=BOTTOM, fill=X)

            ttk.Button(fr_btns, text="🧹 Limpiar", bootstyle="secondary-outline", command=limpiar_form).pack(side=LEFT, expand=True, fill=X, padx=2)
            btn_guardar = ttk.Button(fr_btns, text="💾 CREAR NUEVO", bootstyle="success", command=guardar)
            btn_guardar.pack(side=LEFT, expand=True, fill=X, padx=2)
            ttk.Button(fr_btns, text="🗑️ Eliminar", bootstyle="danger", command=eliminar).pack(side=LEFT, expand=True, fill=X, padx=2)

            # Carga inicial
            cargar_usuarios_en_lista()

        except Exception as e:
            print(f"Error: {e}")
            top.destroy()


    def registrar_accion(self, tipo, partida, material, cantidad, destino, detalles=""):
        """Registra auditoría en el historial (quién hizo qué)."""
        try:
            from datetime import datetime
            fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
            usuario = self.usuario.get('usuario', 'SISTEMA') if hasattr(self, 'usuario') else 'SISTEMA'
            
            self.db.ejecutar("""
                INSERT INTO historial (fecha_hora, tipo, partida, material, cantidad, destino, responsable, entrego, factura)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (fecha, tipo, partida, material, cantidad, destino, usuario, "AUDITORIA", detalles))
        except Exception as e:
            print(f"Error log: {e}")

    # ------------------------------------------------------------------
    # FUNCIONES DE EDICIÓN Y BORRADO DE MATERIALES
    # ------------------------------------------------------------------
    def editar_material_seleccionado(self):
        """
        Permite corregir datos. 
        Sincroniza AUTOMÁTICAMENTE el historial si cambias el nombre o la factura.
        """
        sel = self.tree_inv.selection()
        if not sel: return
        
        # Obtener datos actuales del renglón seleccionado
        item = self.tree_inv.item(sel[0])
        valores = item['values'] 
        id_mat = valores[0]
        # Recuperamos datos frescos de la BD para no fallar
        db_data = self.db.consultar("SELECT * FROM inventario WHERE id=?", (id_mat,))
        if not db_data: return
        
        data = db_data[0]
        partida_actual = data['partida']
        nombre_actual = data['material']
        factura_actual = data['factura']

        # Ventana de Edición
        top = tk.Toplevel(self.root)
        top.title("Editar y Sincronizar")
        self.centrar_ventana_emergente(top, 450, 400)

        ttk.Label(top, text="✏️ Editar Material", font=("Segoe UI", 12, "bold"), bootstyle="primary").pack(pady=10)

        # Campos
        ttk.Label(top, text="Partida:").pack(anchor=W, padx=20)
        e_partida = ttk.Entry(top); e_partida.pack(fill=X, padx=20)
        e_partida.insert(0, partida_actual)

        ttk.Label(top, text="Descripción (Nombre):").pack(anchor=W, padx=20)
        e_mat = ttk.Entry(top); e_mat.pack(fill=X, padx=20)
        e_mat.insert(0, nombre_actual)
        
        ttk.Label(top, text="Factura (Referencia):").pack(anchor=W, padx=20)
        e_fac = ttk.Entry(top); e_fac.pack(fill=X, padx=20)
        e_fac.insert(0, factura_actual)

        def guardar_cambios():
            p_new = e_partida.get().strip()
            m_new = e_mat.get().strip().upper()
            f_new = e_fac.get().strip().upper()

            if not p_new or not m_new:
                messagebox.showwarning("Error", "El nombre y partida son obligatorios.")
                return

            if messagebox.askyesno("Confirmar", "¿Guardar cambios?\nSe actualizará el inventario y el historial."):
                try:
                    # 1. ACTUALIZAR TABLA INVENTARIO
                    self.db.ejecutar("""
                        UPDATE inventario SET partida=?, material=?, factura=? WHERE id=?
                    """, (p_new, m_new, f_new, id_mat))
                    
                    # 2. SINCRONIZAR HISTORIAL (MAGIA AQUÍ)
                    
                    # A) Si cambió el NOMBRE, actualizamos todo el historial viejo
                    if m_new != nombre_actual: 
                        self.db.ejecutar("""
                            UPDATE historial SET material=? WHERE material=?
                        """, (m_new, nombre_actual))
                        print(f"Historial renombrado: {nombre_actual} -> {m_new}")

                    # B) Si cambió la FACTURA, preguntamos si actualizar entradas viejas
                    if f_new != factura_actual:
                        if messagebox.askyesno("Actualizar Facturas", 
                                               "Has cambiado la factura.\n"
                                               "¿Quieres aplicar esta factura a las ENTRADAS pasadas en el historial?"):
                            self.db.ejecutar("""
                                UPDATE historial SET factura=? 
                                WHERE material=? AND (tipo LIKE '%ENTRADA%' OR tipo LIKE '%ALTA%')
                            """, (f_new, m_new))

                    # 3. LOG DE AUDITORÍA
                    self.registrar_accion("EDICION", p_new, m_new, 0, "SISTEMA", "Actualización de datos maestros")
                    
                    messagebox.showinfo("Éxito", "Material y Historial sincronizados correctamente.")
                    top.destroy()
                    
                    # Recargar todo
                    self.cargar_tabla_inventario()
                    self.cargar_tabla_historial()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Fallo al guardar: {e}")

        ttk.Button(top, text="💾 GUARDAR Y SINCRONIZAR", bootstyle="success", command=guardar_cambios).pack(fill=X, padx=20, pady=20)
    def eliminar_material_seleccionado(self):
        """SOLO ADMIN: Elimina un material completo"""
        if self.usuario.get('rol') != 'ADMIN':
            messagebox.showerror("Acceso Denegado", "Solo el Administrador puede eliminar materiales.")
            return

        sel = self.tree_inv.selection()
        if not sel: return
        
        item = self.tree_inv.item(sel[0])
        id_mat = item['values'][0]
        nombre_mat = item['values'][2]

        if messagebox.askyesno("⚠️ ELIMINAR MATERIAL", 
                               f"¿Estás seguro de ELIMINAR PERMANENTEMENTE:\n\n'{nombre_mat}'?\n\n"
                               "Se perderá el stock actual. Esta acción es irreversible.", icon='warning'):
            try:
                self.db.ejecutar("DELETE FROM inventario WHERE id=?", (id_mat,))
                self.registrar_accion("ELIMINADO", "N/A", nombre_mat, 0, "PAPELERA", "Admin eliminó material")
                messagebox.showinfo("Listo", "Material eliminado.")
                self.cargar_tabla_inventario()
            except Exception as e:
                messagebox.showerror("Error", f"Fallo: {e}")

    def revertir_historial_admin(self):
        """
        Permite al admin borrar un registro del historial y opcionalmente
        REVERTIR el efecto que tuvo en el stock. (SIN DEJAR LOG DE AUDITORÍA)
        """
        sel = self.tree_hist.selection()
        if not sel: return
        
        item = self.tree_hist.item(sel[0])
        vals = item['values']
        
        # OBTENER DATOS DEL RENGLÓN SELECCIONADO
        # Indices: 0=ID, 2=TIPO, 3=PARTIDA, 4=MATERIAL, 5=CANTIDAD
        id_hist = vals[0]
        tipo = vals[2]
        partida = vals[3]
        material = vals[4]
        
        try: 
            cantidad = float(vals[5])
        except: 
            cantidad = 0
        
        # 1. CONFIRMACIÓN DE BORRADO
        if messagebox.askyesno("Eliminar Historial", 
                               f"¿Eliminar registro de '{tipo}' (ID: {id_hist})?\n\n"
                               "Esta acción será permanente."):
            
            revertir = False
            # Solo preguntamos si queremos afectar stock si fue un movimiento real
            if "ENTRADA" in tipo or "SALIDA" in tipo or "HISTORICO" in tipo:
                revertir = messagebox.askyesno("Revertir Stock", 
                                               f"Este registro movió {cantidad} piezas.\n\n"
                                               "¿Deseas REVERTIR ese movimiento en el inventario actual?\n"
                                               "(Si dices SÍ, el stock se ajustará automáticamente).")

            try:
                if revertir:
                    # BUSCAR EL MATERIAL EN EL INVENTARIO ACTUAL
                    res = self.db.consultar("SELECT stock FROM inventario WHERE partida=? AND material=?", (partida, material))
                    if res:
                        stock_actual = res[0]['stock']
                        nuevo_stock = stock_actual
                        
                        # LÓGICA INVERSA (DESHACER EL CAMBIO)
                        if "ENTRADA" in tipo or "HISTORICO (+)" in tipo:
                            # Si metimos cosas y borramos el registro -> RESTAMOS (quitamos lo que metimos por error)
                            nuevo_stock = stock_actual - cantidad
                        elif "SALIDA" in tipo or "HISTORICO (-)" in tipo:
                            # Si sacamos cosas y borramos el registro -> DEVOLVEMOS (regresamos lo que sacamos)
                            nuevo_stock = stock_actual + cantidad
                        
                        # Actualizar en Base de Datos
                        self.db.ejecutar("UPDATE inventario SET stock=? WHERE partida=? AND material=?", 
                                         (nuevo_stock, partida, material))
                        print(f"Stock revertido. Nuevo stock: {nuevo_stock}")
                    else:
                        messagebox.showwarning("Aviso", "El material ya no existe en el inventario, solo se borrará el historial.")

                # 2. ELIMINAR EL RENGLÓN DEL HISTORIAL
                self.db.ejecutar("DELETE FROM historial WHERE id=?", (id_hist,))
                
                # --- AQUÍ QUITAMOS LA LÍNEA DE AUDITORÍA "LOG_BORRADO" ---
                # self.registrar_accion(...)  <-- ELIMINADO

                messagebox.showinfo("Listo", "Registro eliminado correctamente.")
                
                # ACTUALIZAR TABLAS
                self.cargar_tabla_historial()
                self.cargar_tabla_inventario()

            except Exception as e:
                messagebox.showerror("Error", f"No se pudo completar: {e}")
    
    # ------------------------------------------------------------------
    #  NUEVA SECCIÓN: MÓDULO DE CONSUMO Y ESTADÍSTICAS (BI)
    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    #  MÓDULO DE CONSUMO (CORREGIDO: GRÁFICA + LISTA COMPLETA)
    # ------------------------------------------------------------------
    def setup_tab_consumo(self):
        """Construye la interfaz: Gráfica arriba, Lista detallada abajo"""
        
        # Usamos un PanedWindow para dividir la pantalla en 2 (Arriba/Abajo) ajustables
        self.paned_consumo = ttk.PanedWindow(self.tab_consumo, orient=VERTICAL)
        self.paned_consumo.pack(fill=BOTH, expand=True)

        # --- SECCIÓN SUPERIOR: FILTROS Y GRÁFICA ---
        self.frame_sup = ttk.Frame(self.paned_consumo)
        self.paned_consumo.add(self.frame_sup, weight=3) # La gráfica ocupa más espacio

        # 1. Controles (Filtros)
        fr_control = ttk.Frame(self.frame_sup, padding=10)
        fr_control.pack(fill=X)

        ttk.Label(fr_control, text="📅 Periodo:", font=("Segoe UI", 10, "bold")).pack(side=LEFT)
        
        meses = ["01-Enero", "02-Febrero", "03-Marzo", "04-Abril", "05-Mayo", "06-Junio",
                 "07-Julio", "08-Agosto", "09-Septiembre", "10-Octubre", "11-Noviembre", "12-Diciembre"]
        self.cb_mes_graf = ttk.Combobox(fr_control, values=meses, state="readonly", width=12)
        self.cb_mes_graf.current(datetime.now().month - 1)
        self.cb_mes_graf.pack(side=LEFT, padx=5)

        self.ent_anio_graf = ttk.Entry(fr_control, width=6)
        self.ent_anio_graf.insert(0, str(datetime.now().year))
        self.ent_anio_graf.pack(side=LEFT, padx=5)

        ttk.Button(fr_control, text="🔄 Generar Reporte", bootstyle="primary", command=self.generar_grafica_consumo).pack(side=LEFT, padx=15)

        # Etiquetas de resumen rápido
        self.lbl_resumen_total = ttk.Label(fr_control, text="", font=("Segoe UI", 10, "bold"), foreground="#2980b9")
        self.lbl_resumen_total.pack(side=RIGHT, padx=10)

        # 2. Contenedor de la Gráfica
        self.fr_grafica_container = ttk.Frame(self.frame_sup, padding=5, relief="solid", borderwidth=1)
        self.fr_grafica_container.pack(fill=BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(self.fr_grafica_container, text="📊 La gráfica aparecerá aquí", foreground="gray").place(relx=0.5, rely=0.5, anchor=CENTER)

        # --- SECCIÓN INFERIOR: LISTA DETALLADA (TOP) ---
        self.frame_inf = ttk.Frame(self.paned_consumo, padding=10)
        self.paned_consumo.add(self.frame_inf, weight=2) # La lista ocupa menos espacio

        ttk.Label(self.frame_inf, text="📋 Detalle de Consumo (Ranking Completo)", font=("Segoe UI", 11, "bold"), bootstyle="secondary").pack(anchor=W, pady=(0, 5))

        # Tabla (Treeview)
        cols = ("RANK", "MATERIAL", "CANTIDAD", "PORCENTAJE")
        self.tree_consumo = ttk.Treeview(self.frame_inf, columns=cols, show="headings", height=8, bootstyle="info")
        
        self.tree_consumo.heading("RANK", text="N°"); self.tree_consumo.column("RANK", width=40, anchor=CENTER)
        self.tree_consumo.heading("MATERIAL", text="MATERIAL / PRODUCTO"); self.tree_consumo.column("MATERIAL", width=400)
        self.tree_consumo.heading("CANTIDAD", text="CONSUMO (Pzas)"); self.tree_consumo.column("CANTIDAD", width=120, anchor=CENTER)
        self.tree_consumo.heading("PORCENTAJE", text="% TOTAL"); self.tree_consumo.column("PORCENTAJE", width=100, anchor=CENTER)

        sc_y = ttk.Scrollbar(self.frame_inf, orient=VERTICAL, command=self.tree_consumo.yview)
        self.tree_consumo.configure(yscrollcommand=sc_y.set)
        
        self.tree_consumo.pack(side=LEFT, fill=BOTH, expand=True)
        sc_y.pack(side=RIGHT, fill=Y)


    def generar_grafica_consumo(self):
        """
        Gráfica Premium: Animación de entrada fluida + Interacción Hover
        """
        import matplotlib.pyplot as plt 
        from matplotlib.figure import Figure
        from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        import matplotlib.animation as animation 
        import numpy as np

        # 1. Limpiezas previas
        for widget in self.fr_grafica_container.winfo_children(): widget.destroy()
        for i in self.tree_consumo.get_children(): self.tree_consumo.delete(i)
        
        # DETENER ANIMACIÓN ANTERIOR (Vital para que no choque)
        if hasattr(self, 'mi_animacion') and self.mi_animacion:
            try: self.mi_animacion.event_source.stop()
            except: pass

        # 2. Validar Fechas
        try:
            mes_idx = self.cb_mes_graf.current() + 1
            anio = int(self.ent_anio_graf.get())
            nombre_mes = self.cb_mes_graf.get().split("-")[1]
        except:
            messagebox.showerror("Error", "Año inválido.")
            return

        # 3. Obtener Datos
        sql = "SELECT fecha_hora, material, cantidad FROM historial WHERE tipo LIKE '%SALIDA%' OR tipo LIKE '%HISTORICO (-)%'"
        datos = self.db.consultar(sql)
        
        consumo = {}
        total_mes = 0

        for d in datos:
            try:
                fecha_raw = d['fecha_hora']
                if "/" in fecha_raw:
                    partes = fecha_raw.split("/")
                    m = int(partes[1])
                    y = int(partes[2].split(" ")[0])
                    
                    if m == mes_idx and y == anio:
                        mat = d['material']
                        cant = d['cantidad']
                        consumo[mat] = consumo.get(mat, 0) + cant
                        total_mes += cant
            except: continue

        if not consumo:
            ttk.Label(self.fr_grafica_container, text=f"Sin movimientos en {nombre_mes} {anio}", 
                      font=("Segoe UI", 16), foreground="red").place(relx=0.5, rely=0.5, anchor=CENTER)
            self.lbl_resumen_total.config(text="Total: 0 pzas")
            return

        # 4. Ordenar datos (Top 10)
        items_ordenados = sorted(consumo.items(), key=lambda x: x[1], reverse=True)
        
        # Llenar Tabla
        for idx, (nom, cant) in enumerate(items_ordenados, 1):
            porcentaje = (cant / total_mes) * 100
            self.tree_consumo.insert("", END, values=(f"{idx}°", nom, f"{int(cant)}", f"{porcentaje:.1f}%"))

        self.lbl_resumen_total.config(text=f"Total Mes: {int(total_mes)} piezas")

        # Datos Gráfica
        LIMITE_GRAFICA = 10 
        nombres_graf = [x[0] for x in items_ordenados[:LIMITE_GRAFICA]]
        cantidades_graf = [x[1] for x in items_ordenados[:LIMITE_GRAFICA]]
        
        # Invertir para visualización (el mayor arriba)
        nombres_graf.reverse()
        cantidades_graf.reverse()

        # 5. CONFIGURACIÓN DE LA FIGURA
        fig = Figure(figsize=(5, 4), dpi=100)
        ax = fig.add_subplot(111)
        
        COLOR_BARRA = '#3498db' # Azul
        COLOR_HOVER = '#e74c3c' # Rojo
        
        # Inicializar barras CON ANCHO 0 (para animar después)
        barras = ax.barh(nombres_graf, [0]*len(nombres_graf), color=COLOR_BARRA, height=0.6, alpha=0.9)

        # Estilo
        ax.set_title(f"Top Consumo - {nombre_mes} {anio}", fontsize=12, fontweight='bold', color='#2c3e50')
        ax.set_xlabel("Cantidad (Piezas)", fontsize=9)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_color('#bdc3c7')
        ax.xaxis.grid(True, color='#ecf0f1', linestyle='--')
        ax.set_axisbelow(True)
        
        # Ajuste de escala
        if len(nombres_graf) < 5:
            ax.set_ylim(-1, len(nombres_graf))
        
        # Fijar límite X para que no "baile" la gráfica al crecer
        max_val = max(cantidades_graf) if cantidades_graf else 10
        ax.set_xlim(0, max_val * 1.15) 

        # --- A. DEFINICIÓN DE LA ANIMACIÓN ---
        def animar(frame):
            # frame va de 0 a frames_total
            # Usamos una función suave (no lineal) para que se vea elegante
            progreso = frame / 30  # 30 frames total
            
            for bar, target in zip(barras, cantidades_graf):
                # Efecto de desaceleración simple
                current_width = target * progreso
                bar.set_width(current_width)
                
            return barras

        # --- B. INTERACTIVIDAD (HOVER) ---
        annot = ax.annotate("", xy=(0,0), xytext=(10, 10), textcoords="offset points",
                            bbox=dict(boxstyle="round", fc="white", ec="gray", alpha=0.9),
                            fontsize=9, fontweight='bold', color='#2c3e50')
        annot.set_visible(False)

        def actualizar_hover(bar, valor, nombre):
            y = bar.get_y() + bar.get_height() / 2
            x = bar.get_width()
            annot.xy = (x, y)
            text = f"{nombre}\n{int(valor)} pzas"
            annot.set_text(text)

        def al_mover_mouse(event):
            vis = annot.get_visible()
            if event.inaxes == ax:
                encontro = False
                for bar, val, nom in zip(barras, cantidades_graf, nombres_graf):
                    cont, _ = bar.contains(event)
                    if cont:
                        actualizar_hover(bar, val, nom)
                        annot.set_visible(True)
                        bar.set_color(COLOR_HOVER)
                        encontro = True
                    else:
                        bar.set_color(COLOR_BARRA)
                
                if encontro or vis:
                    fig.canvas.draw_idle()

        # Dibujar en Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.fr_grafica_container)
        canvas.draw()
        canvas.get_tk_widget().pack(side=TOP, fill=BOTH, expand=True)
        
        # Conectar eventos
        canvas.mpl_connect("motion_notify_event", al_mover_mouse)
        
        # INICIAR Y GUARDAR ANIMACIÓN (EL SECRETO PARA QUE NO DESAPAREZCA)
        # blit=False es menos eficiente pero más compatible con Tkinter
        self.mi_animacion = animation.FuncAnimation(fig, animar, frames=31, interval=20, blit=False, repeat=False)

    def tiene_permiso(self, accion):
        """
        Verifica si el usuario actual tiene permiso para una acción específica.
        Acciones: 'crear', 'entrada', 'salida', 'editar', 'eliminar', 'ajustes'
        """
        # 1. Si es ADMIN por Rol, tiene acceso a todo siempre
        if self.usuario.get('rol') == 'ADMIN':
            return True
            
        # 2. Leer permisos JSON del usuario
        try:
            permisos_str = self.usuario.get('permisos', '{}')
            if not permisos_str: return False
            
            import json
            permisos_dict = json.loads(permisos_str)
            
            # Retorna True si la clave existe y es 1 (True)
            return permisos_dict.get(accion, 0) == 1
        except:
            return False

class SelectorDB:
    ARCHIVO_CONFIG = "config_conexion.json"

    @staticmethod
    def obtener_ruta_db(root_padre):
        """
        Intenta obtener la ruta. Si falla, abre ventana usando 'root_padre' como base.
        """
        # 1. Intentar leer config
        if os.path.exists(SelectorDB.ARCHIVO_CONFIG):
            try:
                with open(SelectorDB.ARCHIVO_CONFIG, 'r') as f:
                    data = json.load(f)
                    ruta = data.get("ruta_db")
                    if ruta and os.path.exists(ruta):
                        return ruta
            except: pass

        # 2. Si falla, pedir al usuario usando la raíz existente
        return SelectorDB.abrir_ventana_seleccion(root_padre)
    
    

    @staticmethod
    def abrir_ventana_seleccion(root_padre):
        """Muestra el selector como una ventana hija (Toplevel)"""
        
        selector = tk.Toplevel(root_padre)
        selector.title("Configuración Inicial")
        
        # SE ELIMINÓ 'selector.transient(root_padre)' para que tenga su 
        # propia presencia en la barra de tareas de Windows sin depender del cuadro blanco.
        
        selector.grab_set()
        
        w, h = 500, 350
        ws = root_padre.winfo_screenwidth()
        hs = root_padre.winfo_screenheight()
        selector.geometry(f'{w}x{h}+{int((ws/2)-(w/2))}+{int((hs/2)-(h/2))}')
        
        resultado = [""] 

        ttk.Label(selector, text="BIENVENIDO AL SISTEMA", 
                 font=("Arial", 16, "bold"), bootstyle="primary").pack(pady=(30, 10))
        
        ttk.Label(selector, text="Para comenzar, necesitamos una Base de Datos.\n¿Qué deseas hacer?", 
                 justify=CENTER).pack(pady=10)

        def guardar_y_cerrar(ruta):
            try:
                with open(SelectorDB.ARCHIVO_CONFIG, 'w') as f:
                    json.dump({"ruta_db": ruta}, f)
                resultado[0] = ruta
                selector.destroy() 
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar config: {e}")

        def crear_nueva():
            ruta = filedialog.asksaveasfilename(
                parent=selector, 
                title="Crear Nueva Base de Datos",
                defaultextension=".db",
                filetypes=[("Archivos SQLite", "*.db")],
                initialfile="inventario_unindetec.db"
            )
            if ruta: guardar_y_cerrar(ruta)

        def abrir_existente():
            ruta = filedialog.askopenfilename(
                parent=selector,
                title="Seleccionar Base de Datos Existente",
                filetypes=[("Archivos SQLite", "*.db")]
            )
            if ruta: guardar_y_cerrar(ruta)
            
        def al_cerrar():
            if not resultado[0]:
                if messagebox.askyesno("Salir", "¿Deseas salir del sistema?"):
                    root_padre.destroy() 
                    sys.exit()
            else:
                selector.destroy()

        selector.protocol("WM_DELETE_WINDOW", al_cerrar)

        fr_btns = ttk.Frame(selector, padding=20)
        fr_btns.pack(fill=BOTH, expand=True)

        ttk.Button(fr_btns, text="📂 BUSCAR ARCHIVO EXISTENTE", bootstyle="info", 
                   command=abrir_existente).pack(fill=X, pady=5, ipady=8)
        
        ttk.Label(fr_btns, text="- O -", bootstyle="secondary").pack(pady=5)
        
        ttk.Button(fr_btns, text="✨ CREAR BASE DE DATOS NUEVA", bootstyle="success", 
                   command=crear_nueva).pack(fill=X, pady=5, ipady=8)

        root_padre.wait_window(selector)
        
        return resultado[0]
# FUNCIONES GLOBALES (FUERA DE LAS CLASES)
# ==========================================

# ==========================================
# 1. PANTALLA DE CARGA "MODO CIBERSEGURIDAD"


def mostrar_splash_epico(db_manager, callback):
    """
    Splash Screen Profesional: Estilo Minimalista/Corporativo (Blanco)
    CORREGIDO: Sin el error de letterspacing
    """
    splash = tk.Toplevel()
    splash.overrideredirect(True)
    
    # --- CONFIGURACIÓN DE COLORES Y ESTILO ---
    COLOR_FONDO = "#FFFFFF"       # Blanco Puro
    COLOR_TITULO = "#2c3e50"      # Gris Azulado Oscuro
    COLOR_SUBTITULO = "#7f8c8d"   # Gris Medio
    COLOR_ACCENTO = "#3498db"     # Azul Corporativo
    COLOR_BARRA_FONDO = "#ecf0f1" # Gris muy claro
    
    splash.configure(bg=COLOR_FONDO)
    
    # --- DIMENSIONES Y CENTRADO ---
    ancho = 600
    alto = 350
    ws = splash.winfo_screenwidth()
    hs = splash.winfo_screenheight()
    x = (ws/2) - (ancho/2)
    y = (hs/2) - (alto/2)
    splash.geometry(f'{ancho}x{alto}+{int(x)}+{int(y)}')
    
    # --- OBTENER DATOS DE LA BD ---
    tit = db_manager.get_config("TITULO_APP")
    if not tit: tit = "SOTFWARE DE USO LIBRE"
    
    sub = db_manager.get_config("SUBTITULO_APP")
    if not sub: sub = "CONTROL DE PARTIDAS"
    
    logo_path = db_manager.get_config("LOGO_APP")
    
    # --- INTERFAZ GRÁFICA ---
    
    # 1. Contenedor Central
    fr_centro = tk.Frame(splash, bg=COLOR_FONDO)
    fr_centro.pack(expand=True, fill=BOTH, padx=40)
    
    # 2. Logo (Si existe)
    if logo_path and os.path.exists(logo_path):
        try:
            img = Image.open(logo_path)
            img = img.resize((80, 80), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)
            lbl_img = tk.Label(fr_centro, image=img_tk, bg=COLOR_FONDO)
            lbl_img.image = img_tk # Referencia vital
            lbl_img.pack(pady=(30, 15))
        except: pass

    # 3. Títulos
    lbl_titulo = tk.Label(fr_centro, text=tit, font=("Segoe UI", 24, "bold"), 
                          bg=COLOR_FONDO, fg=COLOR_TITULO)
    lbl_titulo.pack(pady=(0, 5))
    
    # TRUCO: Para simular el espaciado elegante, usamos espacios entre letras
    sub_espaciado = " ".join(list(sub.upper()))
    
    lbl_sub = tk.Label(fr_centro, text=sub_espaciado, font=("Segoe UI", 9, "bold"), 
                       bg=COLOR_FONDO, fg=COLOR_SUBTITULO)
    lbl_sub.pack(pady=(0, 40))

    # 4. Barra de Progreso (Canvas)
    canvas_ancho = 500
    canvas_alto = 4
    canvas = tk.Canvas(fr_centro, width=canvas_ancho, height=canvas_alto, 
                       bg=COLOR_BARRA_FONDO, highlightthickness=0)
    canvas.pack(pady=(0, 10))
    
    # Crear el rectángulo de progreso
    barra = canvas.create_rectangle(0, 0, 0, canvas_alto, fill=COLOR_ACCENTO, width=0)
    
    # 5. Texto de estado
    lbl_estado = tk.Label(fr_centro, text="Iniciando...", 
                          font=("Segoe UI", 9), bg=COLOR_FONDO, fg=COLOR_SUBTITULO)
    lbl_estado.pack()

    # --- ANIMACIÓN ---
    mensajes = [
        "Cargando configuración...",
        "Verificando base de datos...",
        "Preparando interfaz...",
        "Listo..."
    ]
    
    progreso_actual = 0
    paso = 2 # Velocidad de avance de la barra
    
    def animar():
        nonlocal progreso_actual
        if progreso_actual < canvas_ancho:
            progreso_actual += paso
            canvas.coords(barra, 0, 0, progreso_actual, canvas_alto)
            
            # Cambiar texto
            porcentaje = (progreso_actual / canvas_ancho) * 100
            idx = int((porcentaje / 100) * (len(mensajes) - 1))
            lbl_estado.config(text=mensajes[idx])
            
            splash.after(15, animar) # 15ms por frame
        else:
            splash.after(400, lambda: [splash.destroy(), callback()])

    splash.after(200, animar)

def iniciar_sistema(root_login, db_manager, usuario_data):
    """Cierra login y abre sistema principal"""
    root_login.withdraw()  # Ocultar en lugar de destruir
    
    root_main = ttk.Window(themename="flatly")
    app = SistemaInventario(root_main, db_manager, usuario_data)
    
    # Cuando se cierre el sistema principal, cerrar todo
    def al_cerrar():
        root_main.destroy()
        root_login.destroy()
    
    root_main.protocol("WM_DELETE_WINDOW", al_cerrar)
    root_main.mainloop()


if __name__ == "__main__":
    print("--- INICIANDO SISTEMA ---")
    
    # 1. Crear ventana principal oculta (Para que los menús emergentes tengan un padre)
    app = tb.Window(themename="flatly")
    app.withdraw() # MANTENER OCULTA, evita el cuadro blanco
    app.title("Sistema de Inventario")

    # 2. PREPARACIÓN Y VALIDACIÓN DE BD
    archivo_config = "config_conexion.json"  
    ruta_db = None
    
    # A) Leer archivo de configuración de conexión
    if os.path.exists(archivo_config):
        try:
            with open(archivo_config, 'r') as f:
                data = json.load(f)
                posible_ruta = data.get("ruta_db")
                if posible_ruta and os.path.exists(posible_ruta):
                    ruta_db = posible_ruta
        except: pass

    # B) Si no hay archivo config o la ruta no existe (NAS desconectado)
    if not ruta_db or not os.path.exists(ruta_db):
        if 'SelectorDB' in globals():
            # NO usamos app.deiconify() aquí, para que no salga el cuadro blanco
            ruta_seleccionada = SelectorDB.abrir_ventana_seleccion(app)
            if ruta_seleccionada:
                ruta_db = ruta_seleccionada
            else:
                print("Operación cancelada por el usuario.")
                sys.exit()
        else:
            ruta_db = "inventario_unindetec.db"

    # 3. INSTANCIAR GESTOR
    db_temp = GestorBaseDatos(ruta_db)
    
    # 4. LEER CONFIGURACIÓN VISUAL DESDE LA BD
    tema_guardado = db_temp.get_config("TEMA_BOOTSTRAP") or "flatly"
    fuente_guardada = db_temp.get_config("FUENTE_SISTEMA") or "Segoe UI"
    tamano_fuente = 10 

    app.style.theme_use(tema_guardado)
    
    # 5. APLICAR FUENTE
    estilo = ttk.Style()
    estilo.configure(".", font=(fuente_guardada, tamano_fuente))
    estilo.configure("Treeview.Heading", font=(fuente_guardada, tamano_fuente, "bold"))
    estilo.configure("Treeview", font=(fuente_guardada, tamano_fuente))
    
    w, h = 500, 400
    ws, hs = app.winfo_screenwidth(), app.winfo_screenheight()
    app.geometry(f"{w}x{h}+{int((ws/2)-(w/2))}+{int((hs/2)-(h/2))}")

    # 6. FUNCIONES DE FLUJO
    def iniciar_app_principal(usuario_data):
        """Cierra Login y MUESTRA el Sistema Principal"""
        for widget in app.winfo_children(): widget.destroy()
        app.withdraw() 
        
        sistema = SistemaInventario(app, db_temp, usuario_data)
        
        try:
            app.state('zoomed') 
        except:
            app.attributes('-zoomed', True)

        app.update_idletasks()
        app.deiconify()

    def mostrar_login():
        """Muestra la pantalla de Login"""
        for widget in app.winfo_children(): widget.destroy()
        try:
            LoginWindow(app, db_temp, iniciar_app_principal)
        except Exception as e:
            messagebox.showerror("Error", f"Error iniciando Login: {e}")

    # 7. EJECUCIÓN DEL SPLASH -> LOGIN
    if 'mostrar_splash_epico' in globals():
        try:
            mostrar_splash_epico(db_temp, mostrar_login)
        except Exception as e:
            print(f"Error en splash: {e}")
            mostrar_login()
    else:
        mostrar_login()

    app.mainloop()