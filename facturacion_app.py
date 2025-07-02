# facturacion_app.py (Usando xlwings)
import pandas as pd
import xlwings as xw
from tkinter import Tk, filedialog, messagebox, StringVar, ttk, Label
import os
import datetime
import sys
import subprocess # Importar subprocess para re-abrir el archivo

# Importar Image y ImageTk para manejar imágenes en Tkinter
try:
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showwarning("Advertencia", "La librería Pillow no está instalada. No se podrá mostrar el logo. "
                                         "Por favor, instala Pillow con: pip install Pillow")
    Image = None
    ImageTk = None


# Definiciones de colores para xlwings (para relleno de filas)
COLOR_ELVIRA_RGB = (102, 255, 255)   # #66FFFF (Cian/Azul Claro)
COLOR_CARLOS_RGB = (204, 255, 153)   # #CCFF99 (Verde/Amarillo Claro)

class FacturacionProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Procesador de Pronóstico de Cobranza")
        master.geometry("700x550") # Aumentar un poco la altura para el logo y más espacio
        master.resizable(False, False)

        # Configurar el tema de ttk para una apariencia más moderna
        style = ttk.Style()
        style.theme_use('clam') # 'clam' es a menudo más moderno que 'default' o 'alt'

        # Estilos personalizados
        style.configure('TButton', font=('Arial', 10, 'bold'), padding=10, relief='flat', borderwidth=0,
                        background='#007bff', foreground='white') # Azul para botones
        style.map('TButton', background=[('active', '#0056b3')]) # Efecto hover

        style.configure('TEntry', padding=5, relief='flat', borderwidth=1, fieldbackground='#e9ecef',
                        foreground='#495057', bordercolor='#ced4da') # Estilo para entradas de texto
        style.configure('TLabel', font=('Arial', 10)) # Estilo para etiquetas

        # Variables para las rutas de los archivos
        self.excel_origin_path = StringVar()
        self.excel_origin_path.set("Ningún archivo de origen seleccionado")
        self.excel_template_path = StringVar()
        self.excel_template_path.set("Ningún archivo de plantilla seleccionado")

        self.HOJA_ORIGEN = "TABLA (OK)"
        self.OVL_HOJA = "FACTURACION OVL"
        self.LFOV_HOJA = "FACTURACION LFOV"
        self.FILA_INICIO_ENCABEZADOS_ORIGEN = 7
        self.FILA_INICIO_DATOS_DESTINO = 2

        # Define las columnas a importar y sus nombres en el archivo de destino.
        # La clave es el nombre EXACTO del encabezado en la hoja "TABLA (OK)".
        # El valor es el nombre que tendrá la columna en las hojas de destino.
        # La columna 'SALDO TOTAL' se calculará y añadirá después en VBA.
        self.COLUMNAS_ORIGEN_ORDENADAS = {
            'EMISOR': 'EMISOR',
            'NOMBRE O RAZON SOCIAL': 'NOMBRE O RAZON SOCIAL',
            'TIPO DE DOCUMENTO': 'TIPO DE DOCUMENTO',
            'CONCEPTO': 'CONCEPTO',
            'FOLIO': 'FOLIO',
            'CONTRATO': 'CONTRATO',
            'PERIODO \nDE \nRENTA': 'PERIODO DE RENTA',
            'SALDO \nPENDIENTE': 'SALDO PENDIENTE',
            'FECHA DE PAGO': 'FECHA DE PAGO',
            'AGENTE': 'AGENTE'
        }
        self.TIPOS_DOCUMENTO_INCLUIDOS = ['FACTURA', 'NOTA DE CREDITO', 'SALDO A FAVOR', 'RECIBO DE PAGO']

        # Configuración de la cuadrícula
        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        for i in range(12): # Más filas para acomodar el logo y el espaciado
            self.master.rowconfigure(i, weight=1)

        # --- Sección del Logo y el Ícono de la Aplicación (Actualizada para .ico) ---
        self.logo_image = None
        self.logo_label = None
        # Cambiar la ruta para buscar un archivo .ico
        logo_path = os.path.join(os.path.dirname(__file__), "logo.ico") # Asume logo.ico en la misma carpeta
        if Image and ImageTk and os.path.exists(logo_path):
            try:
                # Establece el ícono de la aplicación usando wm_iconbitmap para archivos .ico
                # Esta es la forma más robusta para el ícono de la barra de tareas en Windows.
                master.wm_iconbitmap(logo_path)

                # Para mostrar el logo dentro de la GUI, aún necesitamos un PhotoImage
                # Pillow puede abrir archivos .ico y extraer la imagen para PhotoImage.
                img = Image.open(logo_path)
                # Redimensionar la imagen para el logo en la GUI (ej. 64x64 píxeles)
                img_display = img.resize((64, 64), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img_display)
                self.logo_label = Label(master, image=self.logo_image)
                self.logo_label.grid(row=0, column=0, columnspan=2, pady=(10, 5), sticky='n')
                
            except Exception as e:
                messagebox.showwarning("Advertencia de Logo/Ícono", f"No se pudo cargar el logo/ícono: {e}. Asegúrese de que 'logo.ico' es un archivo de ícono válido y que Pillow está instalado.")
                self.logo_image = None
                self.logo_label = None
        
        # Ajustar el inicio de las etiquetas si hay logo
        current_row = 1 if self.logo_label else 0

        ttk.Label(master, text="1. Seleccione el archivo de Excel ORIGEN:",
                  font=('Arial', 10, 'bold')).grid(row=current_row, column=0, columnspan=2, pady=(20, 5), sticky='w', padx=20)
        current_row += 1
        self.origin_file_entry = ttk.Entry(master, textvariable=self.excel_origin_path, width=70, state='readonly')
        self.origin_file_entry.grid(row=current_row, column=0, columnspan=1, pady=5, sticky='ew', padx=(20, 10))
        self.browse_origin_button = ttk.Button(master, text="Buscar Archivo Origen", command=self.browse_origin_file)
        self.browse_origin_button.grid(row=current_row, column=1, pady=5, sticky='w', padx=(0, 20))

        current_row += 1
        ttk.Label(master, text="2. Seleccione el archivo de PLANTILLA de Destino (se SOBRESCRIBIRÁ):",
                  font=('Arial', 10, 'bold')).grid(row=current_row, column=0, columnspan=2, pady=(15, 5), sticky='w', padx=20)
        current_row += 1
        self.template_file_entry = ttk.Entry(master, textvariable=self.excel_template_path, width=70, state='readonly')
        self.template_file_entry.grid(row=current_row, column=0, columnspan=1, pady=5, sticky='ew', padx=(20, 10))
        self.browse_template_button = ttk.Button(master, text="Buscar Archivo Plantilla", command=self.browse_template_file)
        self.browse_template_button.grid(row=current_row, column=1, pady=5, sticky='w', padx=(0, 20))

        current_row += 1
        ttk.Label(master, text="3. Haga clic para PROCESAR y ACTUALIZAR la PLANTILLA:",
                  font=('Arial', 10, 'bold')).grid(row=current_row, column=0, columnspan=2, pady=(15, 5), sticky='w', padx=20)
        current_row += 1
        self.process_button = ttk.Button(master, text="Procesar y Actualizar Plantilla de Cobranza", command=self.process_excel)
        self.process_button.grid(row=current_row, column=0, columnspan=2, pady=10)
        
        current_row += 1
        ttk.Label(master, text="La PLANTILLA seleccionada será MODIFICADA directamente con los datos procesados.",
                  font=('Arial', 9), foreground="red").grid(row=current_row, column=0, columnspan=2, pady=(0, 10))

        current_row += 1
        self.status_label = ttk.Label(master, text="Listo para iniciar. Seleccione los archivos.", font=('Arial', 9, 'italic'))
        self.status_label.grid(row=current_row, column=0, columnspan=2, pady=(10, 20))

    def browse_origin_file(self):
        file_selected = filedialog.askopenfilename(
            initialdir=os.path.expanduser("~"),
            title="Seleccionar Archivo Excel de Origen",
            filetypes=(("Archivos Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*"))
        )
        if file_selected:
            self.excel_origin_path.set(file_selected)
            self.status_label.config(text=f"Archivo origen seleccionado: {os.path.basename(file_selected)}")
        else:
            self.excel_origin_path.set("Ningún archivo de origen seleccionado")
            self.status_label.config(text="Selección de archivo de origen cancelada.")

    def browse_template_file(self):
        file_selected = filedialog.askopenfilename(
            initialdir=os.path.expanduser("~"),
            title="Seleccionar Archivo Excel de Plantilla de Destino",
            filetypes=(("Archivos Excel con Macros", "*.xlsm"), ("Todos los archivos", "*.*"))
        )
        if file_selected:
            self.excel_template_path.set(file_selected)
            self.status_label.config(text=f"Archivo plantilla seleccionado: {os.path.basename(file_selected)}")
        else:
            self.excel_template_path.set("Ningún archivo de plantilla seleccionado")
            self.status_label.config(text="Selección de archivo de plantilla cancelada.")

    def _interleave_agents(self, df_input):
        """
        Intercala los grupos de clientes de 'ELVIRA' y 'CARLOS' para su presentación.
        """
        if df_input.empty:
            return df_input

        df_input['AGENTE_UPPER_TEMP'] = df_input['AGENTE'].astype(str).str.strip().str.upper()

        df_elvira_all = df_input[df_input['AGENTE_UPPER_TEMP'] == 'ELVIRA'].copy()
        df_carlos_all = df_input[df_input['AGENTE_UPPER_TEMP'] == 'CARLOS'].copy()

        elvira_client_groups = [group.drop(columns=['AGENTE_UPPER_TEMP']) for _, group in df_elvira_all.groupby('NOMBRE O RAZON SOCIAL', sort=True)]
        carlos_client_groups = [group.drop(columns=['AGENTE_UPPER_TEMP']) for _, group in df_carlos_all.groupby('NOMBRE O RAZON SOCIAL', sort=True)] 

        interleaved_dataframes = []
        
        iter_elvira_groups = iter(elvira_client_groups)
        iter_carlos_groups = iter(carlos_client_groups)

        while True:
            el_group = None
            ca_group = None
            try:
                el_group = next(iter_elvira_groups)
            except StopIteration:
                pass

            try:
                ca_group = next(iter_carlos_groups)
            except StopIteration:
                pass
            
            if el_group is not None:
                interleaved_dataframes.append(el_group)
            if ca_group is not None:
                interleaved_dataframes.append(ca_group)
            
            if el_group is None and ca_group is None:
                break
        
        if interleaved_dataframes:
            return pd.concat(interleaved_dataframes, ignore_index=True)
        else:
            return pd.DataFrame(columns=[col for col in df_input.columns if col != 'AGENTE_UPPER_TEMP'])


    def _process_single_sheet(self, ws_sheet, df_sheet):
        """
        Procesa una única hoja (OVL o LFOV) utilizando xlwings.
        Limpia datos, escribe nuevos datos y aplica formatos.
        """
        try:
            num_output_cols = len(self.COLUMNAS_ORIGEN_ORDENADAS) 
            
            # Limpiar datos existentes en la hoja de destino para escribir los nuevos
            last_row_in_sheet = ws_sheet.range('A' + str(ws_sheet.cells.last_cell.row)).end('up').row
            
            if last_row_in_sheet >= self.FILA_INICIO_DATOS_DESTINO:
                target_range_clear_content = ws_sheet.range((self.FILA_INICIO_DATOS_DESTINO, 1), (last_row_in_sheet, num_output_cols))
                target_range_clear_content.clear_contents()
                
                full_range_for_clear = ws_sheet.range((self.FILA_INICIO_DATOS_DESTINO, 1), (last_row_in_sheet, num_output_cols))
                full_range_for_clear.color = xw.constants.ColorIndex.xlColorIndexNone 
                
            # Escribir los datos procesados (DataFrame a Excel)
            if not df_sheet.empty:
                ws_sheet.range(self.FILA_INICIO_DATOS_DESTINO, 1).value = df_sheet.values
            else:
                pass
            
            last_data_row_written = self.FILA_INICIO_DATOS_DESTINO + df_sheet.shape[0] - 1
            if df_sheet.empty:
                last_data_row_written = self.FILA_INICIO_DATOS_DESTINO - 1 

            # Obtener índices de columna para formatos y lógica
            agente_col_idx = list(df_sheet.columns).index('AGENTE') if 'AGENTE' in df_sheet.columns else -1
            folio_col_idx = list(df_sheet.columns).index('FOLIO') if 'FOLIO' in df_sheet.columns else -1
            saldo_pendiente_col_df_idx = list(df_sheet.columns).index('SALDO PENDIENTE') if 'SALDO PENDIENTE' in df_sheet.columns else -1
            fecha_pago_col_idx = list(df_sheet.columns).index('FECHA DE PAGO') if 'FECHA DE PAGO' in df_sheet.columns else -1
            
            agente_col_excel = agente_col_idx + 1 if agente_col_idx != -1 else -1
            folio_col_excel = folio_col_idx + 1 if folio_col_idx != -1 else -1
            saldo_pendiente_col_excel = saldo_pendiente_col_df_idx + 1 if saldo_pendiente_col_df_idx != -1 else -1
            fecha_pago_col_excel = fecha_pago_col_idx + 1 if fecha_pago_col_idx != -1 else -1

            # Aplica formato de colores (Elvira/Carlos) a todas las columnas importadas.
            if last_data_row_written >= self.FILA_INICIO_DATOS_DESTINO:
                cols_for_agent_coloring = num_output_cols

                for r_idx in range(self.FILA_INICIO_DATOS_DESTINO, last_data_row_written + 1):
                    row_fill_color = None
                    
                    cell_agente_value = str(ws_sheet.cells(r_idx, agente_col_excel).value).strip().upper() if agente_col_excel != -1 else ""
                    cell_folio_value = str(ws_sheet.cells(r_idx, folio_col_excel).value).strip().upper() if folio_col_excel != -1 else ""

                    if cell_agente_value == "ELVIRA":
                        row_fill_color = COLOR_ELVIRA_RGB
                    elif cell_agente_value == "CARLOS":
                        row_fill_color = COLOR_CARLOS_RGB

                    if row_fill_color is None:
                        if cell_folio_value == "ELVIRA": 
                            row_fill_color = COLOR_ELVIRA_RGB
                        elif cell_folio_value == "CARLOS":
                            row_fill_color = COLOR_CARLOS_RGB

                    # Aplicar color de fondo a las columnas importadas
                    if row_fill_color:
                        ws_sheet.range((r_idx, 1), (r_idx, cols_for_agent_coloring)).color = row_fill_color
                    else:
                        ws_sheet.range((r_idx, 1), (r_idx, cols_for_agent_coloring)).color = xw.constants.ColorIndex.xlColorIndexNone 
                    
                    # Formato de Fecha para la columna 'FECHA DE PAGO'
                    if fecha_pago_col_excel != -1:
                        cell_fecha = ws_sheet.cells(r_idx, fecha_pago_col_excel)
                        if isinstance(cell_fecha.value, (datetime.datetime, datetime.date)):
                            cell_fecha.number_format = 'dd/mm/yyyy'
                        
            # Ajustar ancho de columnas automáticamente
            ws_sheet.autofit()

        except Exception as e:
            print(f"ERROR en _process_single_sheet para hoja '{ws_sheet.name}': {e}", file=sys.stderr)
            raise


    def process_excel(self):
        origin_path = self.excel_origin_path.get()
        template_path = self.excel_template_path.get()

        if not os.path.exists(origin_path):
            messagebox.showerror("Error", "El archivo de Excel ORIGEN no existe. Por favor, seleccione un archivo válido.", parent=self.master)
            self.status_label.config(text="Error: Archivo de origen no encontrado.")
            return

        if not os.path.exists(template_path):
            messagebox.showerror("Error", "El archivo de Excel PLANTILLA de Destino no existe. Por favor, seleccione un archivo válido.", parent=self.master)
            self.status_label.config(text="Error: Archivo de plantilla no encontrado.")
            return

        app = None
        wb = None
        processed_successfully = False # Bandera para controlar el cierre de Excel
        
        try:
            self.status_label.config(text="Procesando datos. Por favor, espere...") # Mensaje inicial más conciso
            self.master.update_idletasks() # Forzar actualización de la GUI
            
            # Abrir Excel de forma INVISIBLE
            app = xw.App(visible=False) 
            app.api.DisplayAlerts = False
            app.api.ScreenUpdating = False
            app.api.Calculation = xw.constants.Calculation.xlCalculationManual

            # Parte 1: Lectura y procesamiento del archivo de origen con Pandas
            df_origen_con_headers = pd.read_excel(origin_path, sheet_name=self.HOJA_ORIGEN,
                                                 header=self.FILA_INICIO_ENCABEZADOS_ORIGEN - 1)

            # Validar que las columnas requeridas para el procesamiento existan en el DataFrame de origen.
            required_cols_origin = list(self.COLUMNAS_ORIGEN_ORDENADAS.keys())
            for col in required_cols_origin:
                if col not in df_origen_con_headers.columns:
                    messagebox.showerror("Error de Columna en Origen", f"La columna '{col}' no fue encontrada en la hoja '{self.HOJA_ORIGEN}' del archivo de origen. "
                                                                        f"Asegúrese de que el encabezado en la fila {self.FILA_INICIO_ENCABEZADOS_ORIGEN} sea exactamente '{col}'.", parent=self.master)
                    self.status_label.config(text=f"Error: Columna '{col}' no encontrada en origen.")
                    return 

            df_origen_con_headers['EMISOR_LIMPIO'] = df_origen_con_headers['EMISOR'].astype(str).str.strip().str.upper()

            df_filtrado_emisor = df_origen_con_headers[
                df_origen_con_headers['EMISOR_LIMPIO'].isin(['OVL', 'LFOV'])
            ].copy()

            if 'EMISOR_LIMPIO' in df_filtrado_emisor.columns:
                df_filtrado_emisor = df_filtrado_emisor.drop(columns=['EMISOR_LIMPIO'])
            
            # Inicializar df_filtrado_documento antes de la condicional
            df_filtrado_documento = df_filtrado_emisor.copy()

            # Asegurarse de que 'TIPO DE DOCUMENTO' existe antes de filtrar por ella
            if 'TIPO DE DOCUMENTO' not in df_filtrado_emisor.columns:
                messagebox.showerror("Error de Columna", f"La columna 'TIPO DE DOCUMENTO' no fue encontrada en el DataFrame filtrado por EMISOR. "
                                                        "Verifique el archivo de origen y la configuración de columnas.", parent=self.master)
                self.status_label.config(text=f"Error: Columna 'TIPO DE DOCUMENTO' no encontrada para filtrar.")
                return

            df_filtrado_documento = df_filtrado_emisor[
                df_filtrado_emisor['TIPO DE DOCUMENTO'].astype(str).str.strip().str.upper().isin(self.TIPOS_DOCUMENTO_INCLUIDOS)
            ].copy()

            # Filtro de SALDO PENDIENTE: Incluye celdas vacías (NaN) y cualquier valor numérico (positivo, negativo o cero).
            saldo_pendiente_col_name_original = 'SALDO \nPENDIENTE' 
            if saldo_pendiente_col_name_original in df_filtrado_documento.columns:
                pd.to_numeric(df_filtrado_documento[saldo_pendiente_col_name_original], errors='coerce') 
                df_filtrado_final = df_filtrado_documento.copy() 
            else:
                messagebox.showwarning("Advertencia de Columna", f"La columna '{saldo_pendiente_col_name_original}' no fue encontrada después de los filtros anteriores. "
                                                                 "No se pudo verificar el tipo de dato de SALDO PENDIENTE, pero se procederá con los filtros existentes.", parent=self.master)
                df_filtrado_final = df_filtrado_documento.copy()


            columnas_a_seleccionar = []
            for col_original in self.COLUMNAS_ORIGEN_ORDENADAS.keys():
                if col_original in df_filtrado_final.columns: 
                    columnas_a_seleccionar.append(col_original)
                else:
                    pass

            if not columnas_a_seleccionar:
                messagebox.showwarning("Advertencia", "No se encontraron filas que cumplan los criterios de filtro (EMISOR OVL/LFOV o TIPO DE DOCUMENTO) en el archivo de origen. El archivo de salida estará vacío.", parent=self.master)
                self.status_label.config(text="Advertencia: No se encontraron datos para procesar.")
                df_final_empty_cols = list(self.COLUMNAS_ORIGEN_ORDENADAS.values())
                df_ovl = pd.DataFrame(columns=df_final_empty_cols)
                df_lfov = pd.DataFrame(columns=df_final_empty_cols)
            else:
                df_final_processed = df_filtrado_final[columnas_a_seleccionar].rename(columns=self.COLUMNAS_ORIGEN_ORDENADAS)
                
                df_ovl_raw = df_final_processed[df_final_processed['EMISOR'].astype(str).str.strip().str.upper() == 'OVL'].copy()
                df_ovl = self._interleave_agents(df_ovl_raw)

                df_lfov_raw = df_final_processed[df_final_processed['EMISOR'].astype(str).str.strip().str.upper() == 'LFOV'].copy()
                df_lfov = self._interleave_agents(df_lfov_raw)

            # Parte 2: Carga y preparación de la plantilla de destino con xlwings
            wb = app.books.open(template_path)

            if self.OVL_HOJA not in [s.name for s in wb.sheets]:
                messagebox.showerror("Error en Plantilla", f"La hoja '{self.OVL_HOJA}' no fue encontrada en el archivo de plantilla. Asegúrese de que el nombre sea correcto.", parent=self.master)
                print(f"ERROR: Hoja '{self.OVL_HOJA}' no encontrada.", file=sys.stderr)
                return 

            if self.LFOV_HOJA not in [s.name for s in wb.sheets]:
                messagebox.showerror("Error en Plantilla", f"La hoja '{self.LFOV_HOJA}' no fue encontrada en el archivo de plantilla. Asegúrese de que el nombre sea correcto.", parent=self.master)
                print(f"ERROR: Hoja '{self.LFOV_HOJA}' no encontrada.", file=sys.stderr)
                return 

            ws_ovl = wb.sheets[self.OVL_HOJA]
            ws_lfov = wb.sheets[self.LFOV_HOJA]

            # Parte 3: Escritura y formato en las hojas de la plantilla
            self._process_single_sheet(ws_ovl, df_ovl)
            self._process_single_sheet(ws_lfov, df_lfov)

            # Las llamadas a macros VBA han sido eliminadas para que Excel las gestione.

            # Parte 4: Guardar el archivo de plantilla actualizado
            wb.save()
            
            # Cerrar el libro y la aplicación de xlwings
            wb.close() 
            app.quit()

            # Mostrar el mensaje de éxito
            messagebox.showinfo("Éxito", f"¡Procesamiento completado! El archivo de PLANTILLA ha sido actualizado:\n{template_path}", parent=self.master)
            self.status_label.config(text="¡Procesamiento completado con éxito! Plantilla actualizada.")
            
            processed_successfully = True # Marcar como éxito

        except FileNotFoundError:
            messagebox.showerror("Error", f"Uno de los archivos de Excel no fue encontrado. Verifique las rutas.", parent=self.master)
            self.status_label.config(text="Error: Archivos no encontrados.")
            print(f"ERROR: FileNotFoundError - Uno de los archivos no fue encontrado. Ruta origen: {origin_path}, Ruta plantilla: {template_path}", file=sys.stderr)
        except KeyError as e:
            messagebox.showerror("Error de Columna", f"Una columna esperada no fue encontrada. Asegúrese de que los encabezados sean correctos. Detalle: {e}", parent=self.master)
            self.status_label.config(text=f"Error: Columna faltante ({e}).")
            print(f"ERROR: KeyError - Columna faltante. Detalle: {e}", file=sys.stderr)
        except Exception as e:
            messagebox.showerror("Error Inesperado", f"Ocurrió un error inesperado durante el procesamiento: {e}", parent=self.master)
            print(f"ERROR: Ocurrió un error inesperado durante el procesamiento: {e}", file=sys.stderr)
            self.status_label.config(text="Error inesperado durante el procesamiento.")
        finally:
            # Asegurar que Excel application se cierre solo si hubo un error y aún está abierta
            try:
                if not processed_successfully and app and app.alive:
                    # Si hubo un error y la aplicación de Excel sigue viva, intentar cerrarla.
                    if app.books:
                        for open_wb in app.books:
                            try:
                                open_wb.close(False) # Cerrar sin guardar cambios
                            except Exception as e_close_wb:
                                print(f"WARNING: Error al cerrar el libro en el bloque finally: {e_close_wb}", file=sys.stderr)
                    app.quit()
                    print("DEBUG: Aplicación Excel de xlwings cerrada en finally debido a un error.")
            except Exception as e_quit:
                print(f"ERROR: Error al intentar cerrar la aplicación de Excel en finally: {e_quit}", file=sys.stderr)
            self.master.after(100, lambda: self.master.focus_force())
            
            # Después de todo el procesamiento y cierre, si fue exitoso, re-abrir la plantilla
            if processed_successfully:
                try:
                    # Usar subprocess.Popen para abrir el archivo con la aplicación predeterminada
                    subprocess.Popen(['start', '', template_path], shell=True)
                except Exception as e_reopen:
                    print(f"ERROR: No se pudo re-abrir el archivo procesado: {e_reopen}", file=sys.stderr)
                    messagebox.showwarning("Advertencia", f"El procesamiento se completó, pero no se pudo re-abrir el archivo:\n{template_path}\nError: {e_reopen}", parent=self.master)


if __name__ == "__main__":
    root = Tk()
    app = FacturacionProcessorApp(root)
    root.mainloop()
