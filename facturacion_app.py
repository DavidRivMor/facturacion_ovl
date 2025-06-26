# facturacion_app.py (Usando xlwings)
import pandas as pd
import xlwings as xw
from tkinter import Tk, filedialog, messagebox, StringVar, ttk
import os
import datetime
import sys

# Definiciones de colores para xlwings (RGB)
COLOR_ELVIRA_RGB = (102, 255, 255)   # #66FFFF
COLOR_CARLOS_RGB = (204, 255, 153)   # #CCFF99


class FacturacionProcessorApp:
    def __init__(self, master):
        self.master = master
        master.title("Procesador de Pronóstico de Cobranza")
        master.geometry("700x450")
        master.resizable(False, False)

        self.excel_origin_path = StringVar()
        self.excel_origin_path.set("Ningún archivo de origen seleccionado")
        self.excel_template_path = StringVar()
        self.excel_template_path.set("Ningún archivo de plantilla seleccionado")

        self.HOJA_ORIGEN = "TABLA (OK)"
        self.OVL_HOJA = "FACTURACION OVL"
        self.LFOV_HOJA = "FACTURACION LFOV"
        self.FILA_INICIO_ENCABEZADOS_ORIGEN = 7
        self.FILA_INICIO_DATOS_DESTINO = 2

        # Definición de las columnas de origen y su orden deseado para la hoja de destino
        self.COLUMNAS_ORIGEN_ORDENADAS = {
            'EMISOR': 'EMISOR',
            'NOMBRE O RAZON SOCIAL': 'NOMBRE O RAZON SOCIAL',
            'TIPO DE DOCUMENTO': 'TIPO DE DOCUMENTO',
            'CONCEPTO': 'CONCEPTO',
            'FOLIO': 'FOLIO',
            'CONTRATO': 'CONTRATO',
            'PERIODO \nDE \nRENTA': 'PERIODO DE RENTA',
            'IMPORTE': 'IMPORTE',
            'PAGOS': 'PAGOS',
            'FECHA DE PAGO': 'FECHA DE PAGO',
            'AGENTE': 'AGENTE',
            'PRONOSTICO DE COBRANZA': 'PRONOSTICO DE COBRANZA'
        }
        self.TIPOS_DOCUMENTO_INCLUIDOS = ['FACTURA', 'NOTA DE CREDITO', 'SALDO A FAVOR', 'RECIBO DE PAGO']

        # Componentes de la GUI
        self.master.columnconfigure(0, weight=1)
        self.master.columnconfigure(1, weight=1)
        for i in range(10):
            self.master.rowconfigure(i, weight=1)

        ttk.Label(master, text="1. Seleccione el archivo de Excel ORIGEN:",
                  font=('Arial', 10, 'bold')).grid(row=0, column=0, columnspan=2, pady=(20, 5), sticky='w', padx=20)
        self.origin_file_entry = ttk.Entry(master, textvariable=self.excel_origin_path, width=70, state='readonly')
        self.origin_file_entry.grid(row=1, column=0, columnspan=1, pady=5, sticky='ew', padx=(20, 10))
        self.browse_origin_button = ttk.Button(master, text="Buscar Archivo Origen", command=self.browse_origin_file)
        self.browse_origin_button.grid(row=1, column=1, pady=5, sticky='w', padx=(0, 20))

        ttk.Label(master, text="2. Seleccione el archivo de PLANTILLA de Destino (se SOBRESCRIBIRÁ):",
                  font=('Arial', 10, 'bold')).grid(row=2, column=0, columnspan=2, pady=(15, 5), sticky='w', padx=20)
        self.template_file_entry = ttk.Entry(master, textvariable=self.excel_template_path, width=70, state='readonly')
        self.template_file_entry.grid(row=3, column=0, columnspan=1, pady=5, sticky='ew', padx=(20, 10))
        self.browse_template_button = ttk.Button(master, text="Buscar Archivo Plantilla", command=self.browse_template_file)
        self.browse_template_button.grid(row=3, column=1, pady=5, sticky='w', padx=(0, 20))

        ttk.Label(master, text="3. Haga clic para PROCESAR y ACTUALIZAR la PLANTILLA:",
                  font=('Arial', 10, 'bold')).grid(row=4, column=0, columnspan=2, pady=(15, 5), sticky='w', padx=20)
        self.process_button = ttk.Button(master, text="Procesar y Actualizar Plantilla de Cobranza", command=self.process_excel)
        self.process_button.grid(row=5, column=0, columnspan=2, pady=10)
        
        ttk.Label(master, text="La PLANTILLA seleccionada será MODIFICADA directamente con los datos procesados.",
                  font=('Arial', 9), foreground="red").grid(row=6, column=0, columnspan=2, pady=(0, 10))

        self.status_label = ttk.Label(master, text="Listo para iniciar. Seleccione los archivos.", font=('Arial', 9, 'italic'))
        self.status_label.grid(row=7, column=0, columnspan=2, pady=(10, 20))

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
            print(f"DEBUG: Iniciando _process_single_sheet para hoja: {ws_sheet.name}")

            # num_output_cols es el número de columnas de datos que Python escribe (A-L).
            num_output_cols = len(self.COLUMNAS_ORIGEN_ORDENADAS)
            
            # Limpiar datos existentes en la hoja de destino para escribir los nuevos
            last_row_in_sheet = ws_sheet.range('A' + str(ws_sheet.cells.last_cell.row)).end('up').row
            
            if last_row_in_sheet >= self.FILA_INICIO_DATOS_DESTINO:
                target_range_clear_content = ws_sheet.range((self.FILA_INICIO_DATOS_DESTINO, 1), (last_row_in_sheet, num_output_cols))
                target_range_clear_content.clear_contents()
                
                full_range_for_color_clear = ws_sheet.range((self.FILA_INICIO_DATOS_DESTINO, 1), (last_row_in_sheet, num_output_cols + 2))
                full_range_for_color_clear.color = xw.constants.ColorIndex.xlColorIndexNone 
                
                print(f"DEBUG: Contenido y colores de {ws_sheet.name} limpiados hasta fila {last_row_in_sheet}.")

            # Escribir los datos procesados (DataFrame a Excel)
            if not df_sheet.empty:
                ws_sheet.range(self.FILA_INICIO_DATOS_DESTINO, 1).value = df_sheet.values
                print(f"DEBUG: Datos escritos en {ws_sheet.name} a partir de fila {self.FILA_INICIO_DATOS_DESTINO}.")
            else:
                print(f"DEBUG: DataFrame para {ws_sheet.name} está vacío. No se escribieron datos.")
            
            last_data_row_written = self.FILA_INICIO_DATOS_DESTINO + df_sheet.shape[0] - 1
            if df_sheet.empty:
                last_data_row_written = self.FILA_INICIO_DATOS_DESTINO - 1 

            agente_col_idx = list(df_sheet.columns).index('AGENTE') if 'AGENTE' in df_sheet.columns else -1
            folio_col_idx = list(df_sheet.columns).index('FOLIO') if 'FOLIO' in df_sheet.columns else -1
            
            agente_col_excel = agente_col_idx + 1 if agente_col_idx != -1 else -1
            folio_col_excel = folio_col_idx + 1 if folio_col_idx != -1 else -1
            
            # Aplicar Formato de colores (ELVIRA/CARLOS - prioridad AGENTE, luego FOLIO)
            # Excluyendo la columna L ('PRONOSTICO DE COBRANZA')
            if last_data_row_written >= self.FILA_INICIO_DATOS_DESTINO:
                cols_for_coloring = num_output_cols - 1
                
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

                    if row_fill_color:
                        ws_sheet.range((r_idx, 1), (r_idx, cols_for_coloring)).color = row_fill_color
                    else:
                        ws_sheet.range((r_idx, 1), (r_idx, cols_for_coloring)).color = xw.constants.ColorIndex.xlColorIndexNone
                print(f"DEBUG: Formato de colores (Elvira/Carlos) aplicado en {ws_sheet.name} (excluyendo Columna L).")
            
            current_max_row = last_data_row_written
            
            # La lógica para escribir Fórmulas de Suma y Resta (P3:Q14) ha sido ELIMINADA.
            print(f"DEBUG: Fórmulas de suma y resta no gestionadas por Python en {ws_sheet.name}. Se espera que VBA lo haga.")

            # La validación de datos en M y N no es manejada por Python.
            print(f"DEBUG: La validación de datos en M y N no es manejada por Python en {ws_sheet.name}. Se espera que VBA lo haga.")

            # Ajustar ancho de columnas automáticamente
            ws_sheet.autofit()
            print(f"DEBUG: Ancho de columnas autoajustado en {ws_sheet.name}.")

            print(f"DEBUG: _process_single_sheet finalizado para hoja: {ws_sheet.name}")

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
        
        try:
            self.status_label.config(text="Iniciando lectura y procesamiento de datos del archivo de origen...")
            print("DEBUG: Iniciando process_excel.")
            
            # Abrir Excel de forma oculta
            app = xw.App(visible=False) 
            app.api.DisplayAlerts = False
            app.api.ScreenUpdating = False
            app.api.Calculation = xw.constants.Calculation.xlCalculationManual
            print("DEBUG: Aplicación Excel iniciada en modo oculto y configurada.")

            # Parte 1: Lectura y procesamiento del archivo de origen con Pandas
            df_origen_con_headers = pd.read_excel(origin_path, sheet_name=self.HOJA_ORIGEN,
                                                 header=self.FILA_INICIO_ENCABEZADOS_ORIGEN - 1)
            print(f"DEBUG: DataFrame cargado de origen. Columnas: {list(df_origen_con_headers.columns)}")

            required_cols_origin = ['EMISOR', 'TIPO DE DOCUMENTO', 'NOMBRE O RAZON SOCIAL', 'AGENTE', 'IMPORTE', 'PAGOS', 'FECHA DE PAGO', 'FOLIO', 'CONTRATO', 'PERIODO \nDE \nRENTA', 'PRONOSTICO DE COBRANZA']
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
            
            df_filtrado_documento = df_filtrado_emisor[
                df_filtrado_emisor['TIPO DE DOCUMENTO'].astype(str).str.strip().str.upper().isin(self.TIPOS_DOCUMENTO_INCLUIDOS)
            ].copy()
            print(f"DEBUG: DataFrame filtrado por EMISOR y TIPO DE DOCUMENTO. Shape: {df_filtrado_documento.shape}")

            columnas_a_seleccionar = []
            for col_original in self.COLUMNAS_ORIGEN_ORDENADAS.keys():
                if col_original in df_filtrado_documento.columns: 
                    columnas_a_seleccionar.append(col_original)
                else:
                    print(f"Advertencia: Columna '{col_original}' no encontrada en el DataFrame filtrado de origen. Se omitirá.")

            if not columnas_a_seleccionar:
                messagebox.showwarning("Advertencia", "No se encontraron filas que cumplan los criterios de filtro (EMISOR OVL/LFOV o TIPO DE DOCUMENTO) en el archivo de origen. El archivo de salida estará vacío.", parent=self.master)
                self.status_label.config(text="Advertencia: No se encontraron datos para procesar.")
                df_final_empty_cols = list(self.COLUMNAS_ORIGEN_ORDENADAS.values())
                df_ovl = pd.DataFrame(columns=df_final_empty_cols)
                df_lfov = pd.DataFrame(columns=df_final_empty_cols)
            else:
                df_final_processed = df_filtrado_documento[columnas_a_seleccionar].rename(columns=self.COLUMNAS_ORIGEN_ORDENADAS)
                
                df_ovl_raw = df_final_processed[df_final_processed['EMISOR'].astype(str).str.strip().str.upper() == 'OVL'].copy()
                df_ovl = self._interleave_agents(df_ovl_raw)

                df_lfov_raw = df_final_processed[df_final_processed['EMISOR'].astype(str).str.strip().str.upper() == 'LFOV'].copy()
                df_lfov = self._interleave_agents(df_lfov_raw)
            print("DEBUG: Datos de OVL y LFOV interleavados.")

            self.status_label.config(text="Datos del origen procesados. Abriendo plantilla de destino...")
            print("DEBUG: Abriendo plantilla de destino.")

            # Parte 2: Carga y preparación de la plantilla de destino con xlwings
            wb = app.books.open(template_path)
            print(f"DEBUG: Plantilla '{template_path}' abierta.")

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
            print(f"DEBUG: Hojas '{self.OVL_HOJA}' y '{self.LFOV_HOJA}' obtenidas.")

            self.status_label.config(text="Plantilla cargada. Rellenando hojas...")
            print("DEBUG: Rellenando hojas OVL y LFOV.")

            # Parte 3: Escritura y formato en las hojas de la plantilla
            self._process_single_sheet(ws_ovl, df_ovl)
            self._process_single_sheet(ws_lfov, df_lfov)
            print("DEBUG: Proceso de hojas individuales completado.")

            self.status_label.config(text="Hojas rellenadas y formateadas. Guardando plantilla actualizada...")
            print("DEBUG: Guardando plantilla actualizada.")

            # Parte 4: Guardar el archivo de plantilla actualizado
            wb.save()
            print("DEBUG: Plantilla guardada.")
            
            # Cerrar el libro dentro de la instancia de xlwings
            wb.close()
            print("DEBUG: Workbook cerrado dentro de la aplicación xlwings.")

            # Cerrar la aplicación de Excel iniciada por xlwings
            app.quit()
            print("DEBUG: Aplicación Excel de xlwings cerrada.")

            # Mostrar el mensaje de éxito
            messagebox.showinfo("Éxito", f"¡Procesamiento completado! El archivo de PLANTILLA ha sido actualizado y se abrirá:\n{template_path}", parent=self.master)
            self.status_label.config(text="¡Procesamiento completado con éxito! Plantilla actualizada y se abrirá.")
            print("DEBUG: Mensaje de éxito mostrado.")

            # Abrir el archivo de Excel usando el programa predeterminado del sistema operativo
            # Esto lanzará una nueva instancia de Excel controlada por el usuario
            os.startfile(template_path)
            print(f"DEBUG: Archivo de plantilla '{template_path}' abierto por el sistema operativo.")


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
                if app:
                    # Restaurar las configuraciones de Excel si la aplicación aún existe
                    app.api.DisplayAlerts = True
                    app.api.ScreenUpdating = True
                    app.api.Calculation = xw.constants.Calculation.xlCalculationAutomatic

                    # Si la aplicación de xlwings no fue cerrada exitosamente, intentamos cerrarla aquí.
                    if app.books: # Si hay libros aún abiertos en la app
                        for open_wb in app.books:
                            open_wb.close(False) # Cierra sin guardar cambios
                    app.quit()
                    print("DEBUG: Aplicación Excel de xlwings cerrada en finally debido a un error.")
            except Exception as e_quit:
                print(f"ERROR: Error al intentar cerrar la aplicación de Excel en finally: {e_quit}", file=sys.stderr)
            self.master.after(100, lambda: self.master.focus_force())
            print("DEBUG: Finalizado el bloque finally.")

if __name__ == "__main__":
    root = Tk()
    app = FacturacionProcessorApp(root)
    root.mainloop()
