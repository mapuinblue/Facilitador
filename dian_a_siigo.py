"""
DIAN a Siigo - Aplicación de Escritorio v2.6
Mejorado: Formato correcto de VALOR_BASE y nueva paleta de colores pasteles
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from pathlib import Path
import re
from datetime import datetime
import os
import sys
import traceback


class ProcesadorContableDIAN:
    """Procesa archivos DIAN con detección automática de estructura"""
    
    def __init__(self):
        self.IVA_RATE = 0.19
        self.CUENTAS_COMPRAS = {
            'gasto': '14, 51, 61',
            'iva_descontable': '24080103'
        }
        self.CUENTAS_VENTAS = {
            'ingresos': '41',
            'iva_generado': '24080101',
            'iva_credito': '13050501'
        }
    
    def formato_pesos(self, valor):
        """Formatea número al estilo colombiano: 200.000,00"""
        try:
            if pd.isna(valor) or valor == 0 or valor == '':
                return "0,00"
            valor = round(float(valor), 2)
            entero = int(valor)
            decimal = abs(valor - entero)
            entero_formateado = f"{entero:,}".replace(",", ".")
            decimal_formateado = f"{decimal:.2f}"[1:].replace(".", ",")
            return f"{entero_formateado}{decimal_formateado}"
        except:
            return "0,00"
    
    def formato_base_iva(self, valor):
        """Formatea la base del IVA al estilo colombiano: 200.000,00"""
        try:
            if pd.isna(valor) or valor == 0 or valor == '':
                return "0,00"
            valor = int(round(float(valor)))  # Base IVA es entero redondeado
            return f"{valor:,}".replace(",", ".") + ",00"
        except:
            return "0,00"
    
    def redondear_peso(self, valor):
        """Redondea al peso más cercano"""
        try:
            if pd.isna(valor) or valor == '':
                return 0
            return round(float(valor))
        except:
            return 0
    
    def limpiar_nit(self, nit):
        """Limpia y formatea el NIT"""
        if pd.isna(nit):
            return ""
        nit_str = str(nit).strip().replace('.0', '').replace('.00', '')
        return re.sub(r'[^\d]', '', nit_str)
    
    def leer_archivo_dian(self, ruta_archivo):
        """
        Lee archivo DIAN con mejor detección de estructura:
        - Busca encabezados reales buscando patrones conocidos
        - Maneja diferentes formatos de archivo
        """
        extension = Path(ruta_archivo).suffix.lower()
        
        try:
            # Primero intentar leer sin saltar filas para inspeccionar
            if extension == '.csv':
                df_raw = pd.read_csv(ruta_archivo, encoding='utf-8-sig', 
                                   header=None, dtype=str, nrows=10)
            else:
                xl = pd.ExcelFile(ruta_archivo)
                hoja = xl.sheet_names[0]
                df_raw = pd.read_excel(ruta_archivo, sheet_name=hoja, 
                                     header=None, dtype=str, nrows=10)
            
            print("Primeras filas del archivo crudo:")
            for i in range(min(5, len(df_raw))):
                print(f"Fila {i}: {list(df_raw.iloc[i].dropna().head(15))}")
            
            # Buscar la fila que contiene encabezados clave
            header_row = None
            for idx, row in df_raw.iterrows():
                row_str = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                if any(keyword in row_str.lower() for keyword in ['total', 'iva', 'nit', 'emisor', 'receptor']):
                    header_row = idx
                    print(f"Encontrado encabezado en fila {idx}")
                    break
            
            if header_row is None:
                # Usar la primera fila no vacía como encabezado
                for idx, row in df_raw.iterrows():
                    if not row.isnull().all():
                        header_row = idx
                        break
            
            print(f"Usando fila {header_row} como encabezado")
            
            # Leer el archivo con el encabezado encontrado
            if extension == '.csv':
                df = pd.read_csv(ruta_archivo, encoding='utf-8-sig', 
                               skiprows=header_row, header=0, dtype=str)
            else:
                df = pd.read_excel(ruta_archivo, sheet_name=hoja, 
                                 skiprows=header_row, header=0, dtype=str)
            
            # Limpiar nombres de columnas
            df.columns = [str(col).strip() for col in df.columns]
            
            # Mostrar columnas detectadas
            print(f"\nColumnas detectadas ({len(df.columns)}):")
            for i, col in enumerate(df.columns):
                print(f"  {i:2d}. '{col}'")
            
            print(f"\nTotal filas leídas: {len(df)}")
            
            # Buscar columnas por patrones si no tienen los nombres exactos
            column_mapping = {}
            
            # Buscar Total
            for col in df.columns:
                col_lower = str(col).lower()
                if 'total' in col_lower and 'base' not in col_lower:
                    column_mapping['Total'] = col
                elif 'valor' in col_lower and 'total' in col_lower:
                    column_mapping['Total'] = col
                elif 'monetario' in col_lower:
                    column_mapping['Total'] = col
            
            # Buscar IVA
            for col in df.columns:
                col_lower = str(col).lower()
                if 'iva' in col_lower and 'rete' not in col_lower and 'total' not in col_lower:
                    column_mapping['IVA'] = col
                elif 'impuesto' in col_lower and 'valor' in col_lower:
                    column_mapping['IVA'] = col
            
            # Buscar NIT Emisor
            for col in df.columns:
                col_lower = str(col).lower()
                if 'nit' in col_lower and 'emisor' in col_lower:
                    column_mapping['NIT Emisor'] = col
                elif 'documento' in col_lower and 'emisor' in col_lower:
                    column_mapping['NIT Emisor'] = col
            
            # Buscar Nombre Emisor
            for col in df.columns:
                col_lower = str(col).lower()
                if 'nombre' in col_lower and 'emisor' in col_lower:
                    column_mapping['Nombre Emisor'] = col
                elif 'razón' in col_lower and 'social' in col_lower:
                    column_mapping['Nombre Emisor'] = col
            
            # Buscar NIT Receptor
            for col in df.columns:
                col_lower = str(col).lower()
                if 'nit' in col_lower and 'receptor' in col_lower:
                    column_mapping['NIT Receptor'] = col
                elif 'documento' in col_lower and 'receptor' in col_lower:
                    column_mapping['NIT Receptor'] = col
            
            # Buscar Nombre Receptor
            for col in df.columns:
                col_lower = str(col).lower()
                if 'nombre' in col_lower and 'receptor' in col_lower:
                    column_mapping['Nombre Receptor'] = col
            
            print(f"\nMapeo de columnas encontrado: {column_mapping}")
            
            # Renombrar columnas a nombres estándar
            df = df.rename(columns=column_mapping)
            
            # Verificar que tenemos las columnas mínimas requeridas
            if 'Total' not in df.columns:
                # Intentar encontrar por índice (última columna numérica)
                numeric_cols = []
                for col in df.columns:
                    try:
                        # Verificar si la columna contiene números
                        sample = df[col].dropna().head(10)
                        if len(sample) > 0 and any(str(x).replace(',', '').replace('.', '').isdigit() for x in sample):
                            numeric_cols.append(col)
                    except:
                        pass
                
                if numeric_cols:
                    # Usar la última columna numérica como Total
                    df = df.rename(columns={numeric_cols[-1]: 'Total'})
                    print(f"Usando '{numeric_cols[-1]}' como columna Total")
            
            # Filtrar solo Facturas electrónicas si existe la columna
            if 'Tipo de documento' in df.columns:
                original_count = len(df)
                mask = df['Tipo de documento'].astype(str).str.contains('Factura', case=False, na=False)
                df = df[mask]
                print(f"Facturas filtradas: {len(df)} de {original_count}")
            else:
                print("Advertencia: No se encontró columna 'Tipo de documento'")
            
            # Convertir columnas numéricas
            for col in df.columns:
                if col in ['Total', 'IVA', 'ICA', 'Rete IVA', 'Rete Renta', 'Rete ICA']:
                    try:
                        # Reemplazar comas por puntos y convertir a numérico
                        df[col] = df[col].astype(str).str.replace(',', '.', regex=False)
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        print(f"Convertida columna {col} a numérico")
                    except Exception as e:
                        print(f"Error convirtiendo columna {col}: {e}")
                        df[col] = 0
            
            # Si no se encontró columna IVA, calcularla si es posible
            if 'IVA' not in df.columns and 'Total' in df.columns:
                print("Advertencia: No se encontró columna IVA, se asumirá 0")
                df['IVA'] = 0
            
            return df
            
        except Exception as e:
            raise Exception(f"Error leyendo archivo: {str(e)}\nDetalle: {traceback.format_exc()}")
    
    def procesar_compras(self, df):
        """
        Procesa COMPRAS con manejo de columnas faltantes
        """
        registros = []
        
        print(f"\nProcesando {len(df)} compras...")
        print(f"Columnas disponibles: {list(df.columns)}")
        
        # Verificar columnas requeridas
        required_cols = ['Total', 'IVA']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise Exception(f"Columnas faltantes para compras: {missing_cols}")
        
        for idx, row in df.iterrows():
            try:
                # Obtener valores con manejo de columnas faltantes
                total = float(row.get('Total', 0))
                iva = float(row.get('IVA', 0))
                
                # Obtener NIT Emisor o usar valor por defecto
                if 'NIT Emisor' in df.columns:
                    nit = self.limpiar_nit(row.get('NIT Emisor', ''))
                else:
                    # Intentar encontrar columna con NIT
                    nit_cols = [col for col in df.columns if 'nit' in str(col).lower()]
                    nit = self.limpiar_nit(row[nit_cols[0]]) if nit_cols else ""
                
                # Obtener nombre o usar valor por defecto
                if 'Nombre Emisor' in df.columns:
                    obs = str(row.get('Nombre Emisor', f'Compra {idx+1}'))[:50]
                else:
                    obs = f"Compra {idx+1}"
                
                if total == 0 and iva == 0:
                    continue
                
                # Calcular valores
                valor_sin_iva = total - iva
                base_iva = self.redondear_peso(iva / self.IVA_RATE) if iva > 0 else 0
                
                # Fila 1: Gasto (Débito)
                registros.append({
                    'CUENTA': self.CUENTAS_COMPRAS['gasto'],
                    'CC': '',
                    'OBSERVACIONES': obs,
                    'DEBITO': self.formato_pesos(valor_sin_iva),
                    'CREDITO': '',
                    'VALOR_BASE': '',
                    'TERCERO': nit,
                    'H': ''
                })
                
                # Fila 2: IVA (Débito)
                if iva > 0:
                    registros.append({
                        'CUENTA': self.CUENTAS_COMPRAS['iva_descontable'],
                        'CC': '',
                        'OBSERVACIONES': obs,
                        'DEBITO': self.formato_pesos(iva),
                        'CREDITO': '',
                        'VALOR_BASE': self.formato_base_iva(base_iva),
                        'TERCERO': nit,
                        'H': 1
                    })
                    
            except Exception as e:
                print(f"Error procesando fila {idx}: {e}")
                continue
        
        print(f"Registros generados: {len(registros)}")
        return pd.DataFrame(registros)
    
    def procesar_ventas(self, df):
        """
        Procesa VENTAS con manejo de columnas faltantes
        """
        registros = []
        
        print(f"\nProcesando {len(df)} ventas...")
        print(f"Columnas disponibles: {list(df.columns)}")
        
        # Verificar columnas requeridas
        required_cols = ['Total', 'IVA']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            raise Exception(f"Columnas faltantes para ventas: {missing_cols}")
        
        for idx, row in df.iterrows():
            try:
                # Obtener valores con manejo de columnas faltantes
                total = float(row.get('Total', 0))
                iva = float(row.get('IVA', 0))
                
                # Obtener NIT Receptor o usar valor por defecto
                if 'NIT Receptor' in df.columns:
                    nit = self.limpiar_nit(row.get('NIT Receptor', ''))
                else:
                    # Intentar encontrar columna con NIT
                    nit_cols = [col for col in df.columns if 'nit' in str(col).lower()]
                    nit = self.limpiar_nit(row[nit_cols[0]]) if nit_cols else ""
                
                # Obtener nombre o usar valor por defecto
                if 'Nombre Receptor' in df.columns:
                    obs = str(row.get('Nombre Receptor', f'Venta {idx+1}'))[:50]
                else:
                    obs = f"Venta {idx+1}"
                
                if total == 0 and iva == 0:
                    continue
                
                # Calcular valores
                valor_sin_iva = total - iva
                base_iva = self.redondear_peso(iva / self.IVA_RATE) if iva > 0 else 0
                
                # Fila 1: Ingresos (Crédito) - Cuenta 41
                registros.append({
                    'CUENTA': self.CUENTAS_VENTAS['ingresos'],
                    'CC': '',
                    'OBSERVACIONES': obs,
                    'DEBITO': '',
                    'CREDITO': self.formato_pesos(valor_sin_iva),
                    'VALOR_BASE': '',
                    'TERCERO': nit,
                    'H': ''
                })
                
                # Fila 2: IVA Generado (Crédito) - Cuenta 24080101
                if iva > 0:
                    registros.append({
                        'CUENTA': self.CUENTAS_VENTAS['iva_generado'],
                        'CC': '',
                        'OBSERVACIONES': obs,
                        'DEBITO': '',
                        'CREDITO': self.formato_pesos(iva),
                        'VALOR_BASE': self.formato_base_iva(base_iva),
                        'TERCERO': nit,
                        'H': 1
                    })
                    
                    # Fila 3: IVA Débito - Cuenta 13050501
                    registros.append({
                        'CUENTA': self.CUENTAS_VENTAS['iva_credito'],
                        'CC': '',
                        'OBSERVACIONES': obs,
                        'DEBITO': self.formato_pesos(iva),
                        'CREDITO': '',
                        'VALOR_BASE': self.formato_base_iva(base_iva),
                        'TERCERO': nit,
                        'H': ''
                    })
                    
            except Exception as e:
                print(f"Error procesando fila {idx}: {e}")
                continue
        
        print(f"Registros generados: {len(registros)}")
        return pd.DataFrame(registros)


class AplicacionDIAN:
    """Interfaz gráfica"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("DIAN → Siigo | Conversor Contable v2.6")
        self.root.geometry("1000x800")
        
        # Paleta de colores pasteles en rosas
        self.COLORES = {
            'fondo_principal': '#FFF0F5',  # Lavender blush - rosa muy suave
            'fondo_secundario': '#FFE6F2',  # Rosa más claro
            'fondo_frame': '#FFFFFF',  # Blanco para contraste
            'titulo_principal': '#C71585',  # Medium violet red - rosa oscuro elegante
            'texto_principal': '#8B0058',   # Rosa oscuro para buen contraste
            'texto_secundario': '#A0527A',  # Dusty rose - rosa grisáceo
            'boton_principal': '#FFB6C1',   # Light pink
            'boton_secundario': '#FFC0CB',  # Pink
            'boton_accion': '#FF69B4',      # Hot pink
            'boton_exito': '#FF1493',       # Deep pink
            'boton_peligro': '#DB7093',     # Pale violet red (Mismo color para todos los botones de resultados)
            'barra_progreso': '#FFC0CB',    # Pink
            'log_fondo': '#2c3e50',         # Mantener oscuro para el log
            'log_texto': '#FFB6C1',         # Light pink para el texto del log
            'borde': '#FFB6C1'              # Light pink para bordes
        }
        
        self.root.configure(bg=self.COLORES['fondo_principal'])
        
        self.procesador = ProcesadorContableDIAN()
        self.archivo_actual = None
        self.df_resultado = None
        
        self.crear_widgets()
    
    def crear_widgets(self):
        main_frame = tk.Frame(self.root, bg=self.COLORES['fondo_principal'], padx=30, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        tk.Label(main_frame, text="Conversor DIAN a Siigo", 
                font=('Helvetica', 26, 'bold'), 
                bg=self.COLORES['fondo_principal'], 
                fg=self.COLORES['titulo_principal']).pack(pady=(0, 5))
        
        tk.Label(main_frame, text="Procesa archivos de compras y ventas descargados de la DIAN", 
                font=('Helvetica', 12), 
                bg=self.COLORES['fondo_principal'], 
                fg=self.COLORES['texto_secundario']).pack(pady=(0, 20))
        
        # Frame de selección
        frame_archivo = tk.LabelFrame(main_frame, text=" 1. Seleccionar Archivo de la DIAN ", 
                                     bg=self.COLORES['fondo_frame'], 
                                     fg=self.COLORES['texto_principal'],
                                     font=('Helvetica', 12, 'bold'), 
                                     padx=15, pady=15,
                                     relief=tk.RIDGE,
                                     borderwidth=2)
        frame_archivo.pack(fill=tk.X, pady=10)
        
        self.entry_ruta = tk.Entry(frame_archivo, font=('Helvetica', 11), 
                                  width=65, relief=tk.SOLID, bd=2,
                                  bg='white', fg=self.COLORES['texto_principal'])
        self.entry_ruta.pack(side=tk.LEFT, padx=(0, 10), ipady=5)
        
        tk.Button(frame_archivo, text="📁 Buscar Archivo", 
                 command=self.seleccionar_archivo,
                 bg=self.COLORES['boton_principal'], 
                 fg=self.COLORES['texto_principal'],
                 font=('Helvetica', 11, 'bold'),
                 relief=tk.RAISED, 
                 padx=20, pady=6,
                 cursor='hand2',
                 activebackground=self.COLORES['boton_accion'],
                 activeforeground='white').pack(side=tk.LEFT)
        
        # Frame de tipo
        frame_tipo = tk.LabelFrame(main_frame, text=" 2. Tipo de Documento ", 
                                  bg=self.COLORES['fondo_frame'],
                                  fg=self.COLORES['texto_principal'],
                                  font=('Helvetica', 12, 'bold'), 
                                  padx=15, pady=15,
                                  relief=tk.RIDGE,
                                  borderwidth=2)
        frame_tipo.pack(fill=tk.X, pady=10)
        
        self.tipo_var = tk.StringVar(value="auto")
        
        tipos = [
            ("🔍 Detectar automáticamente (recomendado)", "auto"),
            ("🛒 Compras / Recibidos (usará NIT Emisor)", "compras"),
            ("💰 Ventas / Enviados (usará NIT Receptor)", "ventas")
        ]
        
        for texto, valor in tipos:
            tk.Radiobutton(frame_tipo, text=texto, variable=self.tipo_var, 
                          value=valor, 
                          bg=self.COLORES['fondo_frame'], 
                          fg=self.COLORES['texto_principal'],
                          font=('Helvetica', 11),
                          selectcolor=self.COLORES['boton_principal'],
                          activebackground=self.COLORES['fondo_frame'],
                          activeforeground=self.COLORES['texto_principal']).pack(anchor=tk.W, pady=4)
        
        # Botón procesar
        tk.Button(main_frame, text="⚡ PROCESAR ARCHIVO", 
                 command=self.procesar_archivo,
                 bg=self.COLORES['boton_accion'], 
                 fg='white',
                 font=('Helvetica', 16, 'bold'),
                 relief=tk.RAISED, 
                 padx=40, pady=15,
                 cursor='hand2',
                 activebackground=self.COLORES['boton_exito'],
                 activeforeground='white').pack(pady=20)
        
        # Barra de progreso
        style = ttk.Style()
        style.theme_use('default')
        style.configure("TProgressbar",
                       thickness=20,
                       background=self.COLORES['barra_progreso'],
                       troughcolor=self.COLORES['fondo_secundario'])
        
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, 
                                       length=500, mode='determinate',
                                       style="TProgressbar")
        self.progress.pack(pady=10)
        
        # Frame de resultados
        self.frame_resultados = tk.LabelFrame(main_frame, text=" 3. Resultados y Log ", 
                                             bg=self.COLORES['fondo_frame'],
                                             fg=self.COLORES['texto_principal'],
                                             font=('Helvetica', 12, 'bold'), 
                                             padx=15, pady=15,
                                             relief=tk.RIDGE,
                                             borderwidth=2)
        self.frame_resultados.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.lbl_estado = tk.Label(self.frame_resultados, 
                                  text="Esperando archivo...", 
                                  bg=self.COLORES['fondo_frame'], 
                                  fg=self.COLORES['texto_secundario'], 
                                  font=('Helvetica', 12))
        self.lbl_estado.pack(pady=5)
        
        # Botones
        self.frame_botones = tk.Frame(self.frame_resultados, bg=self.COLORES['fondo_frame'])
        self.frame_botones.pack(pady=5)
        
        self.btn_ver = tk.Button(self.frame_botones, text="👁 Ver Vista Previa", 
                                command=self.ver_preview, state=tk.DISABLED,
                                bg=self.COLORES['boton_peligro'], 
                                fg='white',
                                font=('Helvetica', 11, 'bold'),
                                relief=tk.RAISED, 
                                padx=15, pady=8,
                                cursor='hand2',
                                activebackground='#C71585',
                                activeforeground='white',
                                disabledforeground='white')  # <-- Añadido
        self.btn_ver.pack(side=tk.LEFT, padx=5)
        
        self.btn_excel = tk.Button(self.frame_botones, text="💾 Descargar Excel", 
                                  command=self.guardar_excel, state=tk.DISABLED,
                                  bg=self.COLORES['boton_peligro'], 
                                  fg='white',
                                  font=('Helvetica', 11, 'bold'),
                                  relief=tk.RAISED, 
                                  padx=15, pady=8,
                                  cursor='hand2',
                                  activebackground='#C71585',
                                  activeforeground='white',
                                  disabledforeground='white')  # <-- Añadido
        self.btn_excel.pack(side=tk.LEFT, padx=5)
        
        self.btn_query = tk.Button(self.frame_botones, text="📋 Power Query", 
                                  command=self.mostrar_power_query, state=tk.DISABLED,
                                  bg=self.COLORES['boton_peligro'], 
                                  fg='white',
                                  font=('Helvetica', 11, 'bold'),
                                  relief=tk.RAISED, 
                                  padx=15, pady=8,
                                  cursor='hand2',
                                  activebackground='#C71585',
                                  activeforeground='white',
                                  disabledforeground='white')  # <-- Añadido
        self.btn_query.pack(side=tk.LEFT, padx=5)
        
        # Resumen
        self.lbl_resumen = tk.Label(self.frame_resultados, text="", 
                                   bg=self.COLORES['fondo_frame'], 
                                   fg=self.COLORES['texto_principal'],
                                   font=('Helvetica', 11, 'bold'), 
                                   justify=tk.LEFT)
        self.lbl_resumen.pack(pady=5)
        
        # Log
        self.txt_log = scrolledtext.ScrolledText(self.frame_resultados, 
                                                height=12, width=90,
                                                font=('Consolas', 10),
                                                bg=self.COLORES['log_fondo'], 
                                                fg=self.COLORES['log_texto'],
                                                relief=tk.SUNKEN,
                                                borderwidth=2)
        self.txt_log.pack(fill=tk.BOTH, expand=True, pady=10)
        self.txt_log.insert(tk.END, "Log de procesamiento iniciado...\n")
        self.txt_log.config(state=tk.DISABLED)
    
    def log(self, mensaje):
        """Agrega mensaje al log"""
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} - {mensaje}\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)
        self.root.update()
    
    def seleccionar_archivo(self):
        """Abre diálogo para seleccionar archivo"""
        archivo = filedialog.askopenfilename(
            title="Seleccionar archivo de la DIAN",
            filetypes=[
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Archivos CSV", "*.csv"),
                ("Todos los archivos", "*.*")
            ]
        )
        if archivo:
            self.archivo_actual = archivo
            self.entry_ruta.delete(0, tk.END)
            self.entry_ruta.insert(0, archivo)
            nombre = Path(archivo).name
            self.lbl_estado.config(text=f"Archivo seleccionado: {nombre}", fg=self.COLORES['boton_accion'])
            self.log(f"Archivo seleccionado: {nombre}")
            
            # Detectar tipo por nombre
            nombre_lower = nombre.lower()
            if 'recibido' in nombre_lower:
                self.tipo_var.set('compras')
                self.log("Tipo sugerido: Compras")
            elif 'enviado' in nombre_lower:
                self.tipo_var.set('ventas')
                self.log("Tipo sugerido: Ventas")
    
    def procesar_archivo(self):
        """Procesa el archivo seleccionado"""
        if not self.archivo_actual:
            messagebox.showwarning("Atención", "Por favor selecciona un archivo primero.")
            return
        
        # Limpiar log
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete(1.0, tk.END)
        self.txt_log.config(state=tk.DISABLED)
        
        try:
            self.progress['value'] = 10
            self.root.update()
            
            # Leer archivo
            self.log("Leyendo archivo...")
            df = self.procesador.leer_archivo_dian(self.archivo_actual)
            self.log(f"Filas leídas: {len(df)}")
            
            if len(df) == 0:
                raise Exception("No se encontraron facturas en el archivo.")
            
            self.progress['value'] = 30
            
            # Mostrar columnas detectadas
            self.log(f"Columnas detectadas: {', '.join(list(df.columns))}")
            
            # Determinar tipo
            tipo = self.tipo_var.get()
            if tipo == "auto":
                columnas_str = ' '.join([str(c).upper() for c in df.columns])
                if 'NIT EMISOR' in columnas_str:
                    tipo = "compras"
                    self.log("Tipo detectado: Compras (por NIT Emisor)")
                elif 'NIT RECEPTOR' in columnas_str:
                    tipo = "ventas"
                    self.log("Tipo detectado: Ventas (por NIT Receptor)")
                else:
                    tipo = "compras"
                    self.log("Tipo por defecto: Compras")
            
            self.progress['value'] = 50
            
            # Verificar que existan las columnas necesarias para el tipo seleccionado
            if tipo == "compras":
                if 'NIT Emisor' not in df.columns:
                    # No lanzar error, solo advertir
                    self.log("Advertencia: No se encontró 'NIT Emisor', usando valor por defecto")
                if 'Nombre Emisor' not in df.columns:
                    self.log("Advertencia: No se encontró 'Nombre Emisor', usando valor por defecto")
            else:  # ventas
                if 'NIT Receptor' not in df.columns:
                    self.log("Advertencia: No se encontró 'NIT Receptor', usando valor por defecto")
                if 'Nombre Receptor' not in df.columns:
                    self.log("Advertencia: No se encontró 'Nombre Receptor', usando valor por defecto")
            
            # Verificar columnas comunes
            for col in ['Total', 'IVA']:
                if col not in df.columns:
                    raise Exception(f"No se encontró la columna '{col}' en el archivo")
            
            self.log(f"Columnas verificadas: Total, IVA presentes")
            self.progress['value'] = 70
            
            # Procesar según tipo
            self.log(f"Procesando como {tipo}...")
            if tipo == "compras":
                self.df_resultado = self.procesador.procesar_compras(df)
                tipo_nombre = "Compras/Recibidos"
            else:
                self.df_resultado = self.procesador.procesar_ventas(df)
                tipo_nombre = "Ventas/Enviados"
            
            self.progress['value'] = 100
            
            # Verificar resultado
            if len(self.df_resultado) == 0:
                self.log("⚠️ ERROR: No se generaron registros")
                messagebox.showerror("Error", 
                    "No se generaron registros.\n"
                    "Verifica que las facturas tengan valores en Total e IVA.")
                self.progress['value'] = 0
                return
            
            # Éxito
            self.log(f"✅ ÉXITO: {len(self.df_resultado)} filas generadas")
            self.lbl_estado.config(text=f"✅ Completado: {tipo_nombre}", fg=self.COLORES['boton_exito'])
            self.lbl_resumen.config(text=f"Tipo: {tipo_nombre}\n"
                                        f"Filas generadas: {len(self.df_resultado)}\n"
                                        f"Facturas procesadas: {len(df)}")
            
            # Habilitar botones
            self.btn_ver.config(state=tk.NORMAL, bg=self.COLORES['boton_accion'], fg='white')
            self.btn_excel.config(state=tk.NORMAL)
            self.btn_query.config(state=tk.NORMAL)
            
            messagebox.showinfo("Éxito", 
                f"Procesamiento completado.\n\n"
                f"Tipo: {tipo_nombre}\n"
                f"Facturas: {len(df)}\n"
                f"Registros Siigo: {len(self.df_resultado)}")
            
        except Exception as e:
            self.progress['value'] = 0
            error_msg = str(e)
            self.log(f"❌ ERROR: {error_msg}")
            self.log(traceback.format_exc())
            messagebox.showerror("Error", f"Error al procesar:\n\n{error_msg}")
    
    def ver_preview(self):
        """Muestra ventana con vista previa"""
        if self.df_resultado is None or len(self.df_resultado) == 0:
            return
        
        ventana = tk.Toplevel(self.root)
        ventana.title("Vista Previa - Datos para Siigo")
        ventana.geometry("1100x600")
        ventana.configure(bg=self.COLORES['fondo_principal'])
        
        frame = tk.Frame(ventana, padx=10, pady=10, bg=self.COLORES['fondo_principal'])
        frame.pack(fill=tk.BOTH, expand=True)
        
        columnas = list(self.df_resultado.columns)
        tree = ttk.Treeview(frame, columns=columnas, show='headings', height=20)
        
        # Configurar estilo para el treeview
        style = ttk.Style()
        style.configure("Treeview",
                      background=self.COLORES['fondo_frame'],
                      foreground=self.COLORES['texto_principal'],
                      fieldbackground=self.COLORES['fondo_frame'])
        style.configure("Treeview.Heading",
                      background=self.COLORES['boton_principal'],
                      foreground=self.COLORES['texto_principal'],
                      font=('Helvetica', 10, 'bold'))
        
        for col in columnas:
            tree.heading(col, text=col)
            ancho = 150 if col in ['OBSERVACIONES', 'TERCERO'] else 100
            tree.column(col, width=ancho, anchor='center')
        
        for idx, row in self.df_resultado.head(100).iterrows():
            tree.insert('', tk.END, values=list(row))
        
        scrollbar_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        scrollbar_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        tree.grid(row=0, column=0, sticky='nsew')
        scrollbar_y.grid(row=0, column=1, sticky='ns')
        scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        
        tk.Label(ventana, text=f"Mostrando {min(100, len(self.df_resultado))} de {len(self.df_resultado)} filas", 
                fg=self.COLORES['texto_secundario'], 
                bg=self.COLORES['fondo_principal'],
                font=('Helvetica', 9)).pack(pady=5)
    
    def guardar_excel(self):
        """Guarda el resultado en Excel"""
        if self.df_resultado is None:
            return
        
        tipo = self.tipo_var.get()
        prefijo = "Compras" if tipo == "compras" else "Ventas"
        
        archivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{prefijo}_Siigo_{datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        
        if archivo:
            try:
                with pd.ExcelWriter(archivo, engine='openpyxl') as writer:
                    self.df_resultado.to_excel(writer, index=False, sheet_name='Siigo')
                    
                    worksheet = writer.sheets['Siigo']
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if cell.value:
                                    max_length = max(max_length, len(str(cell.value)))
                            except:
                                pass
                        worksheet.column_dimensions[column_letter].width = min(max_length + 2, 50)
                
                self.log(f"✅ Excel guardado: {archivo}")
                
                if messagebox.askyesno("Éxito", "¿Deseas abrir el archivo ahora?"):
                    os.startfile(archivo)
                    
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")
    
    def mostrar_power_query(self):
        """Muestra código Power Query"""
        if self.df_resultado is None:
            return
        
        # Generar código M
        filas = []
        for idx, row in self.df_resultado.iterrows():
            valores = []
            for v in row.values:
                if pd.isna(v) or v == '':
                    valores.append('null')
                elif isinstance(v, (int, float)):
                    valores.append(str(v))
                else:
                    valores.append(f'"{str(v)}"')
            filas.append(f"    {{ {', '.join(valores)} }}")
        
        datos = ",\n".join(filas)
        headers = ", ".join([f'"{col}"' for col in self.df_resultado.columns])
        
        codigo_m = f"""let
    Origen = #table(
        {{ {headers} }},
        {{
{datos}
        }}
    ),
    TipoCambiado = Table.TransformColumnTypes(Origen,{{
        {{"CUENTA", type text}}, 
        {{"CC", type text}}, 
        {{"OBSERVACIONES", type text}}, 
        {{"DEBITO", type text}}, 
        {{"CREDITO", type text}}, 
        {{"VALOR_BASE", type text}},  <!-- Cambiado a texto por el formato -->
        {{"TERCERO", type text}}, 
        {{"H", Int64.Type}}
    }}),
    Limpieza = Table.ReplaceValue(TipoCambiado,"",null,Replacer.ReplaceValue,
        {{"DEBITO", "CREDITO", "H", "VALOR_BASE"}}),
    Filtrado = Table.SelectRows(Limpieza, each ([CUENTA] <> null))
in
    Filtrado"""
        
        ventana = tk.Toplevel(self.root)
        ventana.title("Código Power Query (M)")
        ventana.geometry("900x700")
        ventana.configure(bg=self.COLORES['fondo_principal'])
        
        tk.Label(ventana, text="Copia este código en Excel (Datos > Obtener datos > Editor avanzado)", 
                fg=self.COLORES['texto_principal'], 
                bg=self.COLORES['fondo_principal'],
                font=('Helvetica', 11, 'bold')).pack(pady=10)
        
        texto = scrolledtext.ScrolledText(ventana, wrap=tk.WORD, 
                                         font=('Consolas', 10), height=30,
                                         bg=self.COLORES['log_fondo'], 
                                         fg=self.COLORES['log_texto'])
        texto.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        texto.insert(tk.END, codigo_m)
        
        def copiar():
            self.root.clipboard_clear()
            self.root.clipboard_append(codigo_m)
            messagebox.showinfo("Copiado", "Código copiado al portapapeles")
        
        tk.Button(ventana, text="📋 Copiar al portapapeles", 
                 command=copiar,
                 bg=self.COLORES['boton_exito'], 
                 fg='white',
                 font=('Helvetica', 12, 'bold'),
                 relief=tk.RAISED, 
                 padx=20, pady=10,
                 cursor='hand2',
                 activebackground='#FF0066',
                 activeforeground='white').pack(pady=10)


if __name__ == "__main__":
    # Instalar dependencias si faltan
    try:
        import pandas
        import openpyxl
    except ImportError:
        print("Instalando dependencias...")
        import subprocess
        subprocess.check_call([sys.executable, "-m", "pip", "install", 
                             "pandas", "openpyxl", "-q"])
        print("Dependencias instaladas. Reiniciando...")
        os.execv(sys.executable, ['python'] + sys.argv)
    
    # Iniciar
    root = tk.Tk()
    app = AplicacionDIAN(root)
    root.mainloop()