import os
import json
import configparser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import zipfile
import xml.etree.ElementTree as ET

class SimpleExcelReader:
    def __init__(self, filepath):
        self.filepath = filepath
        self.shared_strings = []
        self.rows = []  # 0-indexed list of rows, where each row is a list of cell values
        self._load()

    def _load(self):
        with zipfile.ZipFile(self.filepath, 'r') as z:
            # Load Shared Strings
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            try:
                with z.open('xl/sharedStrings.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    for si in root.findall('ns:si', ns):
                        t_elem = si.find('ns:t', ns)
                        if t_elem is not None:
                            self.shared_strings.append(t_elem.text or "")
                        else:
                            # Handle rich text strings which contain <r><t>text</t></r>
                            parts = []
                            for r in si.findall('ns:r', ns):
                                t = r.find('ns:t', ns)
                                if t is not None and t.text:
                                    parts.append(t.text)
                            self.shared_strings.append("".join(parts))
            except KeyError:
                self.shared_strings = []

            # Load Sheet1 (default first worksheet)
            try:
                sheet_file = 'xl/worksheets/sheet1.xml'
                with z.open(sheet_file) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
            except KeyError:
                # If sheet1.xml doesn't exist, look for any sheet under worksheets
                worksheets = [name for name in z.namelist() if name.startswith('xl/worksheets/sheet')]
                if not worksheets:
                    raise ValueError("No se encontraron hojas de trabajo en el archivo Excel.")
                with z.open(worksheets[0]) as f:
                    tree = ET.parse(f)
                    root = tree.getroot()

            sheet_data = root.find('ns:sheetData', ns)
            if sheet_data is None:
                return

            # Temporary dictionary to store raw coordinates since rows in XML are not guaranteed to be ordered
            temp_rows = {}
            max_col_idx = 0
            
            for row_elem in sheet_data.findall('ns:row', ns):
                row_num = int(row_elem.get('r'))  # 1-based
                row_cells = {}
                
                for c_elem in row_elem.findall('ns:c', ns):
                    r_ref = c_elem.get('r')  # e.g., "A1", "BC23"
                    if not r_ref:
                        continue
                    
                    # Parse column letters
                    col_letters = "".join([char for char in r_ref if char.isalpha()])
                    col_idx = 0
                    for char in col_letters:
                        col_idx = col_idx * 26 + (ord(char.upper()) - ord('A') + 1)
                    col_idx -= 1  # 0-based
                    
                    max_col_idx = max(max_col_idx, col_idx)
                    
                    t = c_elem.get('t')  # type
                    v_elem = c_elem.find('ns:v', ns)
                    val = None
                    if v_elem is not None and v_elem.text is not None:
                        val_str = v_elem.text
                        if t == 's':  # shared string
                            val = self.shared_strings[int(val_str)]
                        elif t == 'b':  # boolean
                            val = (val_str == '1')
                        elif t == 'inlineStr': # inline string
                            is_elem = c_elem.find('ns:is', ns)
                            if is_elem is not None:
                                t_elem = is_elem.find('ns:t', ns)
                                if t_elem is not None:
                                    val = t_elem.text
                        else:
                            # Try parsing as float or int, fallback to raw string
                            try:
                                if '.' in val_str:
                                    val = float(val_str)
                                else:
                                    val = int(val_str)
                            except ValueError:
                                val = val_str
                    row_cells[col_idx] = val
                temp_rows[row_num] = row_cells

            if not temp_rows:
                return

            max_row_num = max(temp_rows.keys())
            
            # Reconstruct list of rows
            for r in range(1, max_row_num + 1):
                row_data = [None] * (max_col_idx + 1)
                if r in temp_rows:
                    for c, val in temp_rows[r].items():
                        row_data[c] = val
                self.rows.append(row_data)

    @property
    def max_row(self):
        return len(self.rows)

    @property
    def last_non_empty_row(self):
        for r_idx in range(len(self.rows) - 1, -1, -1):
            row_data = self.rows[r_idx]
            if any(val is not None and str(val).strip() != "" for val in row_data):
                return r_idx + 1
        return 0

    @property
    def max_column(self):
        if not self.rows:
            return 0
        return max(len(r) for r in self.rows)

    def cell_value(self, row, column):
        """row and column are 1-based indices (like openpyxl)."""
        r_idx = row - 1
        c_idx = column - 1
        if 0 <= r_idx < len(self.rows):
            row_data = self.rows[r_idx]
            if 0 <= c_idx < len(row_data):
                return row_data[c_idx]
        return None


class ExcelToJsonConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de Excel a LoRaWAN JSON")
        self.root.geometry("800x650")
        self.root.minsize(700, 550)
        
        # Color palette & styling
        self.bg_color = "#f5f6f8"
        self.card_color = "#ffffff"
        self.primary_color = "#4f46e5" # indigo
        self.text_color = "#1f2937"
        
        self.root.configure(bg=self.bg_color)
        
        # Configuration file path
        self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.ini")
        self.config = configparser.ConfigParser()
        self.load_config()
        
        # State variables
        self.excel_file_path = tk.StringVar(value="")
        self.columns_list = []
        self.last_focused_entry = None
        self.all_rows_var = tk.BooleanVar(value=True)
        self.active_reader = None
        
        # Setup UI
        self.create_widgets()
        
        # Save on close
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def load_config(self):
        """Loads configuration from config.ini, creating it with defaults if missing."""
        if not os.path.exists(self.config_path):
            self.config["LORAWAN"] = {
                "dev_eui": "AC1F09FFFE24F82C",
                "join_eui": "6C4EEF66F47986A6",
                "lorawan_version": "MAC_V1_0_2",
                "lorawan_phy_version": "PHY_V1_0_2_REV_B",
                "frequency_plan_id": "AU_915_928_FSB_2",
                "supports_join": "True",
                "app_key": "1F33A170A5F1FDA0AB697AAE2B95916B"
            }
            self.config["COMPOUND_FIELDS"] = {
                "device_id": "Dispositivo-{dev_eui}",
                "name": "My Device",
                "description": "Living room temperature sensor"
            }
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.config.write(f)
        else:
            self.config.read(self.config_path, encoding="utf-8")
            
    def save_config(self):
        """Saves current GUI field values back to config.ini."""
        try:
            self.config["LORAWAN"]["dev_eui"] = self.fields["dev_eui"].get()
            self.config["LORAWAN"]["join_eui"] = self.fields["join_eui"].get()
            self.config["LORAWAN"]["lorawan_version"] = self.fields["lorawan_version"].get()
            self.config["LORAWAN"]["lorawan_phy_version"] = self.fields["lorawan_phy_version"].get()
            self.config["LORAWAN"]["frequency_plan_id"] = self.fields["frequency_plan_id"].get()
            self.config["LORAWAN"]["supports_join"] = str(self.fields["supports_join"].get())
            self.config["LORAWAN"]["app_key"] = self.fields["app_key"].get()
            
            self.config["COMPOUND_FIELDS"]["device_id"] = self.compound_fields["device_id"].get()
            self.config["COMPOUND_FIELDS"]["name"] = self.compound_fields["name"].get()
            self.config["COMPOUND_FIELDS"]["description"] = self.compound_fields["description"].get()
            
            with open(self.config_path, "w", encoding="utf-8") as f:
                self.config.write(f)
        except Exception as e:
            print(f"Error guardando configuración: {e}")
            
    def on_closing(self):
        self.save_config()
        self.root.destroy()
            
    def create_widgets(self):
        # Configure Grid
        self.root.columnconfigure(0, weight=2)
        self.root.columnconfigure(1, weight=3)
        self.root.rowconfigure(1, weight=1)
        
        # 1. Top File Selection Frame
        top_frame = tk.Frame(self.root, bg=self.card_color, bd=1, relief="groove")
        top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=15, pady=15, ipadx=10, ipady=10)
        top_frame.columnconfigure(1, weight=1)
        
        lbl_file = tk.Label(top_frame, text="Archivo Excel:", bg=self.card_color, font=("Segoe UI", 10, "bold"))
        lbl_file.grid(row=0, column=0, sticky="w", padx=5)
        
        entry_file = tk.Entry(top_frame, textvariable=self.excel_file_path, state="readonly", font=("Segoe UI", 10))
        entry_file.grid(row=0, column=1, sticky="ew", padx=10)
        
        btn_browse = tk.Button(top_frame, text="Buscar archivo...", command=self.browse_excel, bg=self.primary_color, fg="white", activebackground="#4338ca", font=("Segoe UI", 9, "bold"), relief="flat", padx=10)
        btn_browse.grid(row=0, column=2, sticky="e", padx=5)
        
        # 2. Left Column Frame: Columns List & Row Filters
        left_frame = tk.Frame(self.root, bg=self.bg_color)
        left_frame.grid(row=1, column=0, sticky="nsew", padx=(15, 7), pady=(0, 15))
        left_frame.rowconfigure(1, weight=1)
        left_frame.columnconfigure(0, weight=1)
        
        # Columns List Box
        lbl_cols = tk.Label(left_frame, text="Columnas en Excel (Doble clic para insertar):", bg=self.bg_color, font=("Segoe UI", 10, "bold"))
        lbl_cols.grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        list_container = tk.Frame(left_frame, bg="white", bd=1, relief="sunken")
        list_container.grid(row=1, column=0, sticky="nsew")
        list_container.rowconfigure(0, weight=1)
        list_container.columnconfigure(0, weight=1)
        
        self.cols_listbox = tk.Listbox(list_container, font=("Segoe UI", 9), selectbackground="#e0e7ff", selectforeground="black", bd=0, highlightthickness=0)
        self.cols_listbox.grid(row=0, column=0, sticky="nsew")
        self.cols_listbox.bind("<Double-Button-1>", self.on_column_double_click)
        
        scrollbar = tk.Scrollbar(list_container, orient="vertical", command=self.cols_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.cols_listbox.config(yscrollcommand=scrollbar.set)
        
        # Row Range Filter Frame
        filter_frame = tk.LabelFrame(left_frame, text=" Rango de Filas ", bg=self.bg_color, font=("Segoe UI", 9, "bold"), pady=10, padx=10)
        filter_frame.grid(row=2, column=0, sticky="ew", pady=(15, 0))
        filter_frame.columnconfigure(1, weight=1)
        filter_frame.columnconfigure(2, weight=4) # Give ample space to previews
        
        chk_all = tk.Checkbutton(filter_frame, text="Todo el archivo", variable=self.all_rows_var, command=self.toggle_range_inputs, bg=self.bg_color, font=("Segoe UI", 9))
        chk_all.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
        
        lbl_from = tk.Label(filter_frame, text="Desde:", bg=self.bg_color, font=("Segoe UI", 9))
        lbl_from.grid(row=1, column=0, sticky="w", pady=3)
        self.entry_from_var = tk.StringVar()
        self.entry_from_var.trace_add("write", lambda *args: self.update_previews())
        self.entry_from = tk.Entry(filter_frame, textvariable=self.entry_from_var, font=("Segoe UI", 9), width=8)
        self.entry_from.grid(row=1, column=1, sticky="w", padx=5, pady=3)
        
        self.lbl_preview_from = tk.Label(filter_frame, text="-> -", bg=self.bg_color, font=("Segoe UI", 9, "italic"), fg="#4b5563", anchor="w", justify="left")
        self.lbl_preview_from.grid(row=1, column=2, sticky="w", padx=10, pady=3)
        
        lbl_to = tk.Label(filter_frame, text="Hasta:", bg=self.bg_color, font=("Segoe UI", 9))
        lbl_to.grid(row=2, column=0, sticky="w", pady=3)
        self.entry_to_var = tk.StringVar()
        self.entry_to_var.trace_add("write", lambda *args: self.update_previews())
        self.entry_to = tk.Entry(filter_frame, textvariable=self.entry_to_var, font=("Segoe UI", 9), width=8)
        self.entry_to.grid(row=2, column=1, sticky="w", padx=5, pady=3)
        
        self.lbl_preview_to = tk.Label(filter_frame, text="-> -", bg=self.bg_color, font=("Segoe UI", 9, "italic"), fg="#4b5563", anchor="w", justify="left")
        self.lbl_preview_to.grid(row=2, column=2, sticky="w", padx=10, pady=3)
        
        # Initialize disabled state for range input
        self.toggle_range_inputs()
        
        # 3. Right Column Frame: Config Fields
        right_frame = tk.Frame(self.root, bg=self.bg_color)
        right_frame.grid(row=1, column=1, sticky="nsew", padx=(7, 15), pady=(0, 15))
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        # Canvas & Scrollbar for configuration fields (in case screen is small)
        config_canvas = tk.Canvas(right_frame, bg=self.bg_color, borderwidth=0, highlightthickness=0)
        config_canvas.grid(row=0, column=0, sticky="nsew")
        
        config_scroll = tk.Scrollbar(right_frame, orient="vertical", command=config_canvas.yview)
        config_scroll.grid(row=0, column=1, sticky="ns")
        config_canvas.configure(yscrollcommand=config_scroll.set)
        
        scrollable_frame = tk.Frame(config_canvas, bg=self.bg_color)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: config_canvas.configure(scrollregion=config_canvas.bbox("all"))
        )
        config_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=420)
        
        # Section A: Configuración General (INI)
        sec_ini = tk.LabelFrame(scrollable_frame, text=" Configuración General ", bg=self.bg_color, font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        sec_ini.pack(fill="x", pady=(0, 15))
        
        self.fields = {}
        
        general_keys = [
            ("dev_eui", "Device EUI:"),
            ("join_eui", "Join EUI:"),
            ("lorawan_version", "LoRaWAN Version:"),
            ("lorawan_phy_version", "LoRaWAN PHY Version:"),
            ("frequency_plan_id", "Frequency Plan ID:"),
            ("app_key", "App Key:")
        ]
        
        for i, (key, label) in enumerate(general_keys):
            lbl = tk.Label(sec_ini, text=label, bg=self.bg_color, font=("Segoe UI", 9))
            lbl.grid(row=i, column=0, sticky="w", pady=3)
            
            val = self.config.get("LORAWAN", key, fallback="")
            ent = tk.Entry(sec_ini, font=("Segoe UI", 9))
            ent.insert(0, val)
            ent.grid(row=i, column=1, sticky="ew", padx=10, pady=3)
            ent.bind("<FocusIn>", lambda e, w=ent: self.set_focused_entry(w))
            self.fields[key] = ent
            
        # supports_join Checkbutton
        lbl_join = tk.Label(sec_ini, text="Supports Join:", bg=self.bg_color, font=("Segoe UI", 9))
        lbl_join.grid(row=len(general_keys), column=0, sticky="w", pady=3)
        
        join_val = self.config.getboolean("LORAWAN", "supports_join", fallback=True)
        self.supports_join_var = tk.BooleanVar(value=join_val)
        chk_join = tk.Checkbutton(sec_ini, text="", variable=self.supports_join_var, bg=self.bg_color)
        chk_join.grid(row=len(general_keys), column=1, sticky="w", padx=10, pady=3)
        self.fields["supports_join"] = self.supports_join_var
        
        sec_ini.columnconfigure(1, weight=1)
        
        # Section B: Campos Compuestos (Soporta doble clic de columna)
        sec_comp = tk.LabelFrame(scrollable_frame, text=" Campos Compuestos (Fórmulas) ", bg=self.bg_color, font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        sec_comp.pack(fill="x")
        
        self.compound_fields = {}
        compound_keys = [
            ("device_id", "Device ID:"),
            ("name", "Name:"),
            ("description", "Description:")
        ]
        
        for i, (key, label) in enumerate(compound_keys):
            lbl = tk.Label(sec_comp, text=label, bg=self.bg_color, font=("Segoe UI", 9))
            lbl.grid(row=i, column=0, sticky="w", pady=3)
            
            val = self.config.get("COMPOUND_FIELDS", key, fallback="")
            ent = tk.Entry(sec_comp, font=("Segoe UI", 9))
            ent.insert(0, val)
            ent.grid(row=i, column=1, sticky="ew", padx=10, pady=3)
            ent.bind("<FocusIn>", lambda e, w=ent: self.set_focused_entry(w))
            self.compound_fields[key] = ent
            
        self.compound_fields["name"].bind("<KeyRelease>", lambda e: self.update_previews())
        sec_comp.columnconfigure(1, weight=1)
        
        # 4. Bottom Action Button
        btn_export = tk.Button(self.root, text="Exportar JSON", command=self.export_json, bg="#10b981", fg="white", activebackground="#059669", font=("Segoe UI", 11, "bold"), relief="flat", pady=8)
        btn_export.grid(row=2, column=0, columnspan=2, sticky="ew", padx=15, pady=15)
        
    def set_focused_entry(self, entry):
        self.last_focused_entry = entry
        
    def toggle_range_inputs(self):
        if self.all_rows_var.get():
            self.entry_from.config(state="disabled")
            self.entry_to.config(state="disabled")
            excel_path = self.excel_file_path.get()
            if excel_path:
                try:
                    if hasattr(self, "active_reader") and self.active_reader:
                        reader = self.active_reader
                    else:
                        self.active_reader = SimpleExcelReader(excel_path)
                        reader = self.active_reader
                    last_ne = reader.last_non_empty_row
                    self.entry_from_var.set("1")
                    self.entry_to_var.set(str(last_ne - 1) if last_ne > 1 else "0")
                except Exception:
                    pass
        else:
            self.entry_from.config(state="normal")
            self.entry_to.config(state="normal")
        self.update_previews()
            
    def browse_excel(self):
        filepath = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if filepath:
            self.excel_file_path.set(filepath)
            self.load_excel_headers(filepath)
            
    def load_excel_headers(self, filepath):
        try:
            reader = SimpleExcelReader(filepath)
            
            # Read first row for headers
            headers = []
            if reader.max_row > 0:
                first_row = reader.rows[0]
                for cell in first_row:
                    if cell is not None and str(cell).strip() != "":
                        headers.append(str(cell).strip())
            
            self.columns_list = headers
            self.cols_listbox.delete(0, tk.END)
            for col in headers:
                self.cols_listbox.insert(tk.END, col)
                
            # Automatically populate Desde/Hasta range
            last_ne = reader.last_non_empty_row
            self.active_reader = reader
            self.entry_from_var.set("1") # 1 corresponds to the first data row (Excel row 2)
            self.entry_to_var.set(str(last_ne - 1) if last_ne > 1 else "0")
            self.update_previews()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo leer el archivo Excel:\n{e}")
            
    def on_column_double_click(self, event):
        selection = self.cols_listbox.curselection()
        if selection and self.last_focused_entry:
            col_name = self.cols_listbox.get(selection[0])
            self.last_focused_entry.insert(tk.INSERT, f"{{{col_name}}}")
            self.update_previews()
            
    def resolve_placeholders(self, template_str, row_dict):
        """Replaces placeholders in the format {COLUMN_NAME} with row values."""
        resolved = template_str
        for col_name, val in row_dict.items():
            placeholder = f"{{{col_name}}}"
            if placeholder in resolved:
                resolved = resolved.replace(placeholder, str(val) if val is not None else "")
        return resolved

    def update_previews(self):
        excel_path = self.excel_file_path.get()
        if not excel_path:
            self.lbl_preview_from.config(text="Desde: -")
            self.lbl_preview_to.config(text="Hasta: -")
            return
            
        try:
            if not hasattr(self, "active_reader") or not self.active_reader or self.active_reader.filepath != excel_path:
                self.active_reader = SimpleExcelReader(excel_path)
            reader = self.active_reader
            max_r = reader.max_row
        except Exception:
            self.lbl_preview_from.config(text="Desde: -")
            self.lbl_preview_to.config(text="Hasta: -")
            return
            
        try:
            start_row = int(self.entry_from_var.get())
        except ValueError:
            start_row = None
            
        try:
            end_row = int(self.entry_to_var.get())
        except ValueError:
            end_row = None
            
        headers = []
        if reader.max_row > 0:
            first_row = reader.rows[0]
            for val in first_row:
                headers.append(str(val).strip() if val is not None else "")
                
        def get_row_name(r_data):
            if r_data is None or r_data < 1 or r_data >= max_r:
                return "-"
            r_excel = r_data + 1
            row_dict = {}
            for idx, col_name in enumerate(headers):
                if col_name:
                    val = reader.cell_value(r_excel, idx + 1)
                    row_dict[col_name] = val
            name_formula = self.compound_fields["name"].get()
            return self.resolve_placeholders(name_formula, row_dict)

        name_from = get_row_name(start_row)
        name_to = get_row_name(end_row)
        
        # Limit preview display size to prevent GUI overflow
        if len(name_from) > 28:
            name_from = name_from[:25] + "..."
        if len(name_to) > 28:
            name_to = name_to[:25] + "..."
            
        self.lbl_preview_from.config(text=f"-> {name_from}")
        self.lbl_preview_to.config(text=f"-> {name_to}")
        
    def export_json(self):
        excel_path = self.excel_file_path.get()
        if not excel_path:
            messagebox.showwarning("Advertencia", "Por favor seleccione un archivo Excel primero.")
            return
            
        # Save config changes
        self.save_config()
        
        # Determine range
        try:
            reader = SimpleExcelReader(excel_path)
            self.active_reader = reader
            max_r = reader.max_row
            last_ne = reader.last_non_empty_row
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo Excel:\n{e}")
            return
            
        start_row = 1
        end_row = last_ne - 1
        
        if not self.all_rows_var.get():
            try:
                start_row = int(self.entry_from.get())
                end_row = int(self.entry_to.get())
                if start_row < 1 or end_row < start_row:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Rango de filas inválido. 'Desde' debe ser >= 1, y 'Hasta' >= 'Desde'.")
                return
                
        # Ask output save file path
        output_path = filedialog.asksaveasfilename(
            title="Guardar archivo JSON",
            defaultextension=".json",
            filetypes=[("JSON Files", "*.json")]
        )
        
        if not output_path:
            return
            
        try:
            # Map column index to names
            headers = []
            if reader.max_row > 0:
                first_row = reader.rows[0]
                for val in first_row:
                    headers.append(str(val).strip() if val is not None else "")
                
            devices_list = []
            
            # Read configuration values from GUI fields
            join_eui = self.fields["join_eui"].get()
            lorawan_version = self.fields["lorawan_version"].get()
            lorawan_phy_version = self.fields["lorawan_phy_version"].get()
            frequency_plan_id = self.fields["frequency_plan_id"].get()
            supports_join = self.supports_join_var.get()
            
            # App key settings
            app_key_val = self.fields["app_key"].get()
            
            # For each row, process and create JSON
            for r_data in range(start_row, min(end_row + 1, max_r)):
                r = r_data + 1
                row_dict = {}
                for idx, col_name in enumerate(headers):
                    if col_name:
                        val = reader.cell_value(r, idx + 1)
                        row_dict[col_name] = val
                        
                # If the row is entirely empty, skip it
                if not any(row_dict.values()):
                    continue
                    
                # Resolve inputs (possibly with placeholders)
                dev_eui_resolved = self.resolve_placeholders(self.fields["dev_eui"].get(), row_dict)
                app_key_resolved = self.resolve_placeholders(app_key_val, row_dict)
                
                device_id = self.resolve_placeholders(self.compound_fields["device_id"].get(), row_dict).lower()
                name = self.resolve_placeholders(self.compound_fields["name"].get(), row_dict)
                description = self.resolve_placeholders(self.compound_fields["description"].get(), row_dict)
                
                # Build device payload conforming to requested format
                device_payload = {
                    "ids": {
                        "device_id": device_id,
                        "dev_eui": dev_eui_resolved,
                        "join_eui": join_eui
                    },
                    "name": name,
                    "description": description,
                    "lorawan_version": lorawan_version,
                    "lorawan_phy_version": lorawan_phy_version,
                    "frequency_plan_id": frequency_plan_id,
                    "supports_join": supports_join,
                    "root_keys": {
                        "app_key": {
                            "key": app_key_resolved
                        }
                    }
                }
                
                devices_list.append(device_payload)
            
            # Write JSON file
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(devices_list, f, indent=2, ensure_ascii=False)
                
            messagebox.showinfo("Éxito", f"Se exportaron {len(devices_list)} dispositivos con éxito a:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error durante la exportación:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelToJsonConverterApp(root)
    root.mainloop()
