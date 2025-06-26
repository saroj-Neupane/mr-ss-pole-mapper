import sys
import logging
import threading
import json
import shutil
import os
import subprocess
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, simpledialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd

# Add missing imports
from core.config_manager import ConfigManager
from core.utils import Utils
from core.geocoder import Geocoder
from core.attachment_data_reader import AttachmentDataReader
from core.pole_data_processor import PoleDataProcessor
from core.route_parser import RouteParser


class PoleMapperApp:
    """Main application class"""
    
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("Pole Mapper - Configuration & Processing Tool")
            self.root.geometry("1400x900")
            
            # Add protocol handler for window close button
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
            
            # Initialize flags FIRST to prevent recursion
            self._is_saving_config = False
            self._is_initializing = True
            
            # Initialize managers and paths
            self.base_dir = Utils.get_base_directory()
            logging.debug(f"Base directory: {self.base_dir}")
            
            self.config_manager = ConfigManager(self.base_dir)
            self.cache_file = self.base_dir / "geocode_cache.csv"
            self.paths_file = self.base_dir / "last_paths.json"
            self.last_paths = self.load_last_paths()
            
            # Configuration management
            self.current_config_name = "Default"
            self.config = self.config_manager.get_default_config()
            self.mapping_data = []
            
            # Load saved config
            self.load_config()
            
            # Create GUI
            self.create_widgets()
            self.geocoder = None
            
            # Initialization complete - allow auto-saving
            self._is_initializing = False
              # Set initial UI state
            self.update_ui_state()
            
            logging.info("Pole Mapper application initialized successfully")
            
        except Exception as e:
            logging.error(f"Error in PoleMapperApp initialization: {str(e)}", exc_info=True)
            raise

    def load_last_paths(self):
        """Load last used file paths from JSON"""
        default_paths = {
            "input_file": "",
            "attachment_file": "",
            "output_file": "",
            "qc_file": "",
            "last_directory": str(Path.home())
        }
        
        try:
            if self.paths_file.exists():
                with open(self.paths_file, 'r') as f:
                    loaded_paths = json.load(f)
                    default_paths.update(loaded_paths)
        except Exception as e:
            logging.error(f"Error loading last paths: {e}")
        
        return default_paths

    def save_last_paths(self):
        """Save current file paths to JSON"""
        try:
            paths = {
                "input_file": self.input_var.get() if hasattr(self, 'input_var') else "",
                "attachment_file": self.attachment_var.get() if hasattr(self, 'attachment_var') else "",
                "output_file": self.output_var.get() if hasattr(self, 'output_var') else "",
                "qc_file": self.qc_var.get() if hasattr(self, 'qc_var') else "",
                "last_directory": getattr(self, 'last_directory', str(Path.home()))
            }
            
            with open(self.paths_file, 'w') as f:
                json.dump(paths, f, indent=2)
            
            logging.debug("Saved last paths to JSON")
        except Exception as e:
            logging.error(f"Error saving last paths: {e}")

    def load_config(self):
        """Load configuration"""
        try:
            self.config = self.config_manager.load_config(self.current_config_name)
            
            # Ensure all required keys exist
            default_config = self.config_manager.get_default_config()
            for key, value in default_config.items():
                if key not in self.config:
                    self.config[key] = value
            
            # Load mappings
            self.mapping_data = self.config.get("column_mappings", [])
            if not self.mapping_data:
                self.load_default_mappings()
                self.config["column_mappings"] = self.mapping_data
                
        except Exception as e:
            logging.error(f"Error loading config: {e}")
            self.config = self.config_manager.get_default_config()
            self.load_default_mappings()

    def load_default_mappings(self):
        """Load default column mappings"""
        self.mapping_data = [
            # Basic pole information
            ("Pole", "SCID", "Pole"),
            ("Pole", "To Pole", "To Pole"),
            ("Pole", "Line No.", "Line No."),
            ("Pole", "Span Distance", "Span Distance"),
            ("Pole", "Pole Height/Class", "Pole Height/Class"),
            ("Pole", "Address", "Address"),
            ("Pole", "Guy Info", "Guy Info"),
            ("Pole", "Existing Risers", "Existing Risers"),
            
            # Power attachments
            ("Power", "Height", "Power Height"),
            ("Power", "Midspan", "Power Midspan"),
            
            # Communication attachments - individual comm fields
            ("comm1", "Height", "comm1"),
            ("comm2", "Height", "comm2"),
            ("comm3", "Height", "comm3"),
            ("comm4", "Height", "comm4"),
            
            # NEW: Comprehensive communication attachment fields
            ("All_Comm_Heights", "Summary", "All Communication Heights"),
            ("Total_Comm_Count", "Count", "Total Communication Count"),
            
            # Streetlight
            ("Streetlight", "Height", "Streetlight (bottom of bracket)"),
            ("Street Light", "Height", "Street Light Height"),
        ]
        
        # Add provider-specific mappings
        providers = self.config.get("telecom_providers", [])
        for provider in providers:
                self.mapping_data.append((provider, "Attachment Ht", f"{provider} Attachment Height"))
        if "Proposed MetroNet" not in providers:
            self.mapping_data.append(("Proposed MetroNet", "Attachment Ht", "Proposed MetroNet Attachment Height"))
        
        logging.info(f"Loaded {len(self.mapping_data)} default mappings including comprehensive communication fields")

    def create_widgets(self):
        """Create main GUI widgets"""
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_info_tab(notebook)
        self.create_config_tab(notebook)
        self.create_process_tab(notebook)

    def create_info_tab(self, notebook):
        """Create info tab"""
        info_frame = ttk.Frame(notebook)
        notebook.add(info_frame, text="ℹ️ Info")
        
        text_frame = ttk.Frame(info_frame)
        text_frame.pack(fill=BOTH, expand=True, padx=20, pady=20)
        
        info_text = """
MAKE READY SPREADSHEET BUILDER APPLICATION GUIDE

QUICK START:
1. Go to the Configuration tab.
2. Set up your telecom providers and column mappings.
3. Go to the Processing tab.
4. Select the main input Excel file (with nodes, connections, and sections sheets).
5. Select the attachment data Excel file (with SCID sheets).
6. Select the output Excel template file.
7. Click Process Files.

CONFIGURATION:
- Supports multiple configurations.
- Telecom Providers: Add/remove utility companies.
- Power Keywords: Define what counts as power equipment.
- Telecom Keywords: Set alternate names for each provider (case insensitive).
- Output Settings: Configure the header row, data start row, and worksheet name.
- Column Mappings: Map processed data to specific Excel columns.
- Reset to Defaults: Option to restore default settings.

PROCESSING:
- Reads the main Excel file and filters pole data from the nodes sheet.
- Processes attachment data only for SCID sheets matching the filtered nodes.
- Extracts guy wire info from pole notes automatically.
- Optionally uses geocoding to retrieve addresses.
- Calculates power and telecom attachment heights.
- Generates a formatted output Excel file using your defined column mappings.
    """
        
        text_widget = ScrolledText(text_frame, wrap='word', font=("Arial", 11))
        text_widget.pack(fill=BOTH, expand=True)
        text_widget.insert(END, info_text)
        text_widget.config(state=DISABLED)

    def create_config_tab(self, notebook):
        """Create configuration tab"""
        config_frame = ttk.Frame(notebook)
        notebook.add(config_frame, text="⚙️ Configuration")
        
        # Main layout
        main_paned = ttk.PanedWindow(config_frame, orient=HORIZONTAL)
        main_paned.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Left panel with scrollbar
        self.create_left_panel(main_paned)
        
        # Right panel (Column Mappings)
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=2)
        
        mappings_frame = ttk.LabelFrame(right_frame, text="Column Mappings", padding=15)
        mappings_frame.pack(fill=BOTH, expand=True)
        self.create_mappings_editor(mappings_frame)

    def create_left_panel(self, main_paned):
        """Create scrollable left panel"""
        # Create main left frame
        left_main_frame = ttk.Frame(main_paned)
        main_paned.add(left_main_frame, weight=1)
        
        # Create canvas and scrollbar for left panel
        canvas = Canvas(left_main_frame)
        scrollbar = ttk.Scrollbar(left_main_frame, orient="vertical", command=canvas.yview)
        self.scrollable_left_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=self.scrollable_left_frame, anchor="nw")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure scrolling
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.scrollable_left_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            if hasattr(event, 'delta') and event.delta:
                delta = event.delta
            elif hasattr(event, 'num') and event.num in (4, 5):
                delta = 120 if event.num == 4 else -120
            else:
                delta = 0
            if delta:
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)
        
        # Create all sections
        self.create_config_management_section()
        self.create_providers_section()
        self.create_keywords_section()
        self.create_output_settings_section()

    def create_config_management_section(self):
        """Create configuration management section"""
        config_mgmt = ttk.LabelFrame(self.scrollable_left_frame, text="Configuration Management", padding=15)
        config_mgmt.pack(fill=X, pady=(0, 10), padx=5)
        
        # Configuration dropdown
        ttk.Label(config_mgmt, text="Configuration:", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=W)
        self.config_var = StringVar(value=self.current_config_name)
        self.config_combo = ttk.Combobox(config_mgmt, textvariable=self.config_var, 
                                        values=self.config_manager.get_available_configs(), 
                                        state="readonly", width=25)
        self.config_combo.grid(row=0, column=1, sticky=W, padx=(10, 0))
        self.config_combo.bind('<<ComboboxSelected>>', self.on_config_change)
        
        # Buttons
        btn_frame = ttk.Frame(config_mgmt)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky=EW, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Save As", command=self.save_config_as).pack(side=LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Delete", command=self.delete_selected_config).pack(side=LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_to_defaults).pack(side=LEFT)

    def create_providers_section(self):
        """Create telecom providers and power keywords sections"""
        # Power Company
        power_frame = ttk.LabelFrame(self.scrollable_left_frame, text="Power Company", padding=15)
        power_frame.pack(fill=X, pady=(0, 10), padx=5)
        
        ttk.Label(power_frame, text="Power Company Name:").pack(side=LEFT, padx=(0, 10))
        self.power_company_var = StringVar(value=self.config["power_company"])
        power_entry = ttk.Entry(power_frame, textvariable=self.power_company_var)
        power_entry.pack(side=LEFT, fill=X, expand=True)
        
        # Add trace with recursion protection
        def on_power_company_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            self.config["power_company"] = self.power_company_var.get()
            self.auto_save_config()
        
        self.power_company_var.trace('w', on_power_company_change)
        
        # Telecom Providers
        telecom_frame = ttk.LabelFrame(self.scrollable_left_frame, text="Telecom Providers", padding=15)
        telecom_frame.pack(fill=X, pady=(0, 10), padx=5)
        self.create_list_editor(telecom_frame, "telecom_providers")
        
        # Power Keywords
        power_kw_frame = ttk.LabelFrame(self.scrollable_left_frame, text="Power Keywords", padding=15)
        power_kw_frame.pack(fill=X, pady=(0, 10), padx=5)
        self.create_list_editor(power_kw_frame, "power_keywords")

    def create_keywords_section(self):
        """Create owner keywords section"""
        owner_kw_frame = ttk.LabelFrame(self.scrollable_left_frame, text="Owner Keywords", padding=15)
        owner_kw_frame.pack(fill=X, pady=(0, 10), padx=5)
        self.create_telecom_keywords_editor(owner_kw_frame)

    def create_output_settings_section(self):
        """Create output settings section"""
        output_frame = ttk.LabelFrame(self.scrollable_left_frame, text="Output Settings", padding=15)
        output_frame.pack(fill=X, pady=(0, 10), padx=5)
        self.create_output_settings(output_frame)

    def create_list_editor(self, parent, config_key):
        """Create list editor for telecom providers or power keywords"""
        # Initialize listboxes dict if it doesn't exist
        self.listboxes = getattr(self, 'listboxes', {})
        
        listbox = Listbox(parent, height=6)
        listbox.pack(fill=BOTH, expand=True, pady=(0, 10))
        self.listboxes[config_key] = listbox
        
        # Populate listbox
        for item in self.config[config_key]:
            listbox.insert(END, item)
        
        # Controls
        controls = ttk.Frame(parent)
        controls.pack(fill=X)
        
        entry = ttk.Entry(controls)
        entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 5))
        
        def add_item():
            item = entry.get().strip()
            if item and item not in self.config[config_key]:
                self.config[config_key].append(item)
                listbox.insert(END, item)
                entry.delete(0, END)
                if config_key == "telecom_providers":
                    self.refresh_ui()
                # Only save if not already saving/initializing to prevent recursion
                if not getattr(self, '_is_saving_config', False) and not getattr(self, '_is_initializing', False):
                    self.auto_save_config()
        
        def remove_item():
            selection = listbox.curselection()
            if selection:
                item = listbox.get(selection[0])
                self.config[config_key].remove(item)
                listbox.delete(selection[0])
                if config_key == "telecom_providers":
                    self.refresh_ui()
                # Only save if not already saving/initializing to prevent recursion
                if not getattr(self, '_is_saving_config', False) and not getattr(self, '_is_initializing', False):
                    self.auto_save_config()
        
        ttk.Button(controls, text="Add", command=add_item).pack(side=LEFT, padx=(0, 5))
        ttk.Button(controls, text="Remove", command=remove_item).pack(side=LEFT)
        
        entry.bind('<Return>', lambda e, func=add_item: func())

    def create_telecom_keywords_editor(self, parent):
        """Create telecom keywords editor with scrolling"""
        canvas = Canvas(parent, height=150)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.telecom_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=self.telecom_frame, anchor="nw")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure scrolling
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.telecom_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            if hasattr(event, 'delta') and event.delta:
                delta = event.delta
            elif hasattr(event, 'num') and event.num in (4, 5):
                delta = 120 if event.num == 4 else -120
            else:
                delta = 0
            if delta:
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)
        
        self.populate_telecom_keywords()

    def populate_telecom_keywords(self):
        """Populate telecom keywords"""
        # Clear existing widgets
        for widget in self.telecom_frame.winfo_children():
            widget.destroy()
        
        # Clear existing trace variables
        if hasattr(self, 'telecom_vars'):
            for var in self.telecom_vars.values():
                try:
                    for trace_type, trace_id in var.trace_info() if hasattr(var, 'trace_info') else []:
                        try:
                            var.trace_vdelete(trace_type, trace_id)
                        except Exception:
                            pass
                except Exception:
                    pass
        
        self.telecom_vars = {}
        
        for i, provider in enumerate(self.config["telecom_providers"]):
            ttk.Label(self.telecom_frame, text=f"{provider}:", width=15).grid(row=i, column=0, sticky=W, pady=2)
            
            keywords = self.config["telecom_keywords"].get(provider, [])
            if not keywords and provider:
                keywords = [provider.lower()]
            
            keywords_str = ", ".join(keywords)
            
            var = StringVar(value=keywords_str)
            self.telecom_vars[provider] = var
            
            entry = ttk.Entry(self.telecom_frame, textvariable=var, width=40)
            entry.grid(row=i, column=1, sticky=EW, padx=(10, 0), pady=2)
            
            def create_update_function(prov):
                def update_keywords(*args):
                    if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                        return
                    try:
                        keywords_str = self.telecom_vars[prov].get()
                        keywords = [k.strip() for k in keywords_str.split(",") if k.strip()]
                        self.config["telecom_keywords"][prov] = keywords
                        self.auto_save_config()
                    except Exception as e:
                        logging.error(f"Error updating keywords for {prov}: {e}")
                return update_keywords
            
            var.trace_add('write', create_update_function(provider))
        
        self.telecom_frame.grid_columnconfigure(1, weight=1)

    def create_output_settings(self, parent):
        """Create output settings"""
        # Header Row
        ttk.Label(parent, text="Header Row:").grid(row=0, column=0, sticky=W, pady=2)
        self.header_row_var = StringVar(value=str(self.config["output_settings"]["header_row"]))
        header_entry = ttk.Entry(parent, textvariable=self.header_row_var, width=10)
        header_entry.grid(row=0, column=1, sticky=W, padx=(10, 0), pady=2)
        
        # Data Start Row
        ttk.Label(parent, text="Data Start Row:").grid(row=1, column=0, sticky=W, pady=2)
        self.data_start_row_var = StringVar(value=str(self.config["output_settings"]["data_start_row"]))
        data_entry = ttk.Entry(parent, textvariable=self.data_start_row_var, width=10)
        data_entry.grid(row=1, column=1, sticky=W, padx=(10, 0), pady=2)
        
        # Worksheet Name
        ttk.Label(parent, text="Worksheet Name:").grid(row=2, column=0, sticky=W, pady=2)
        self.worksheet_name_var = StringVar(value=self.config["output_settings"]["worksheet_name"])
        worksheet_entry = ttk.Entry(parent, textvariable=self.worksheet_name_var, width=20)
        worksheet_entry.grid(row=2, column=1, sticky=W, padx=(10, 0), pady=2)
        
        # Trace functions
        def on_header_row_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            try:
                self.config["output_settings"]["header_row"] = int(self.header_row_var.get())
                self.auto_save_config()
            except ValueError:
                pass
        
        def on_data_start_row_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            try:
                self.config["output_settings"]["data_start_row"] = int(self.data_start_row_var.get())
                self.auto_save_config()
            except ValueError:
                pass
        
        def on_worksheet_name_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            self.config["output_settings"]["worksheet_name"] = self.worksheet_name_var.get()
            self.auto_save_config()
        
        self.header_row_var.trace('w', on_header_row_change)
        self.data_start_row_var.trace('w', on_data_start_row_change)
        self.worksheet_name_var.trace('w', on_worksheet_name_change)

    def create_mappings_editor(self, parent):
        """Create column mappings editor"""
        # Header
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Element", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky=W)
        ttk.Label(header_frame, text="Attribute", font=("Arial", 11, "bold")).grid(row=0, column=1, sticky=W, padx=(20, 0))
        ttk.Label(header_frame, text="Output Column", font=("Arial", 11, "bold")).grid(row=0, column=2, sticky=W, padx=(20, 0))
        
        # Mappings area with scrollbar
        canvas = Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        self.mappings_frame = ttk.Frame(canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=self.mappings_frame, anchor="nw")
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Configure scrolling
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        
        self.mappings_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            if hasattr(event, 'delta') and event.delta:
                delta = event.delta
            elif hasattr(event, 'num') and event.num in (4, 5):
                delta = 120 if event.num == 4 else -120
            else:
                delta = 0
            if delta:
                canvas.yview_scroll(int(-1 * (delta / 120)), "units")
        
        canvas.bind_all("<MouseWheel>", on_mousewheel)
        canvas.bind_all("<Button-4>", on_mousewheel)
        canvas.bind_all("<Button-5>", on_mousewheel)
        
        self.populate_mappings()
        
        # Buttons
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=X, pady=(10, 0))
        
        ttk.Button(btn_frame, text="Add Mapping", command=self.add_mapping).pack(side=LEFT, padx=(0, 10))
        ttk.Button(btn_frame, text="Reset to Defaults", command=self.reset_mappings).pack(side=LEFT)

    def populate_mappings(self):
        """Populate mappings"""
        for widget in self.mappings_frame.winfo_children():
            widget.destroy()
        
        for i, (element, attribute, output) in enumerate(self.mapping_data):
            self.create_mapping_row(i, element, attribute, output)

    def create_mapping_row(self, row_idx, element, attribute, output):
        """Create a mapping row"""
        row_frame = ttk.Frame(self.mappings_frame)
        row_frame.pack(fill=X, pady=2)
        
        # Element dropdown
        element_var = StringVar(value=element)
        element_combo = ttk.Combobox(row_frame, textvariable=element_var, 
                                   values=self.get_element_options(), state="readonly", width=15)
        element_combo.grid(row=0, column=0, sticky=W)
        
        # Attribute dropdown
        attribute_var = StringVar(value=attribute)
        attribute_combo = ttk.Combobox(row_frame, textvariable=attribute_var,
                                     values=self.get_attribute_options(element), state="readonly", width=15)
        attribute_combo.grid(row=0, column=1, sticky=W, padx=(20, 0))
        
        # Output entry
        output_var = StringVar(value=output)
        output_entry = ttk.Entry(row_frame, textvariable=output_var, width=40)
        output_entry.grid(row=0, column=2, sticky=W, padx=(20, 0))
        
        # Delete button
        ttk.Button(row_frame, text="Delete", command=lambda idx=row_idx: self.delete_mapping(idx)).grid(row=0, column=3, padx=(20, 0))
        
        # Trace callbacks
        def on_element_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            try:
                attribute_combo['values'] = self.get_attribute_options(element_var.get())
                if attribute_combo['values']:
                    attribute_var.set(attribute_combo['values'][0])
                self.update_mapping_data()
                self.auto_save_config()
            except Exception as e:
                logging.error(f"Error in element change: {e}")
        
        def on_attribute_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            try:
                self.update_mapping_data()
                self.auto_save_config()
            except Exception as e:
                logging.error(f"Error in attribute change: {e}")
        
        def on_output_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            try:
                self.update_mapping_data()
                self.auto_save_config()
            except Exception as e:
                logging.error(f"Error in output change: {e}")
        
        element_var.trace_add('write', on_element_change)
        attribute_var.trace_add('write', on_attribute_change)
        output_var.trace_add('write', on_output_change)
        
        # Store references
        row_frame.element_var = element_var
        row_frame.attribute_var = attribute_var
        row_frame.output_var = output_var

    def get_element_options(self):
        """Get element options"""
        base = ["Pole", "New Guy", "Power", "Span", "System", "Street Light"]
        comm_options = ["comm1", "comm2", "comm3", "comm4"]
        return base + comm_options + self.config["telecom_providers"]

    def get_attribute_options(self, element):
        """Get attribute options"""
        options = {
            "Pole": ["Number", "Address", "Height & Class", "MR Notes", "To Pole", "Latitude", "Longitude", "Tag", "Number of Existing Risers"],
            "New Guy": ["Size", "Lead", "Direction"],
            "Power": ["Lowest Height", "Lowest Midspan"],
            "Span": ["Length"],
            "System": ["Line Number"],
            "Street Light": ["Lowest Height"]
        }
        
        if element in ["comm1", "comm2", "comm3", "comm4"] or element in self.config["telecom_providers"]:
            return ["Attachment Ht", "Midspan Ht"]
        
        return options.get(element, ["Custom"])

    def update_mapping_data(self):
        """Update mapping data from UI"""
        try:
            new_data = []
            for widget in self.mappings_frame.winfo_children():
                if hasattr(widget, 'element_var'):
                    try:
                        element = widget.element_var.get()
                        attribute = widget.attribute_var.get()
                        output = widget.output_var.get()
                        if element and attribute and output.strip():
                            new_data.append((element, attribute, output))
                    except Exception as e:
                        logging.error(f"Error reading mapping row: {e}")
            
            self.mapping_data = new_data
        except Exception as e:
            logging.error(f"Error updating mapping data: {e}")

    def add_mapping(self):
        """Add new mapping"""
        elements = self.get_element_options()
        if elements:
            element = elements[0]
            attributes = self.get_attribute_options(element)
            attribute = attributes[0] if attributes else "Custom"
            self.mapping_data.append((element, attribute, "New Column"))
            self.populate_mappings()
            self.auto_save_config()

    def delete_mapping(self, idx):
        """Delete mapping"""
        try:
            if 0 <= idx < len(self.mapping_data):
                del self.mapping_data[idx]
                self.populate_mappings()
                self.auto_save_config()
        except Exception as e:
            logging.error(f"Error deleting mapping: {e}")

    def reset_mappings(self):
        """Reset mappings to defaults"""
        self.load_default_mappings()
        self.populate_mappings()
        self.config["column_mappings"] = self.mapping_data
        self.auto_save_config()

    def create_process_tab(self, notebook):
        """Create processing tab"""
        process_frame = ttk.Frame(notebook)
        notebook.add(process_frame, text="▶️ Processing")
        
        # Create main layout with left and right panels
        main_paned = ttk.PanedWindow(process_frame, orient=HORIZONTAL)
        main_paned.pack(fill=BOTH, expand=True, padx=10, pady=10)
        
        # Left panel for controls
        left_frame = ttk.Frame(main_paned)
        main_paned.add(left_frame, weight=1)
        
        # Right panel for log
        right_frame = ttk.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        
        # File selection
        self.create_file_selection(left_frame)
        
        # Manual routes
        self.create_manual_routes_section(left_frame)
        
        # Processing options
        self.create_processing_options(left_frame)
        
        # Process button
        ttk.Button(left_frame, text="Process Files", command=self.process_files,
                  style="Accent.TButton").pack(pady=20)
        
        # Progress
        self.create_progress_section(left_frame)
        
        # Processing log in right panel
        self.create_log_section(right_frame)
        
        # Setup logging
        self.setup_logging()
        
        # Initialize route text state
        self.toggle_route_text()

    def create_file_selection(self, parent):
        """Create file selection section"""
        file_frame = ttk.LabelFrame(parent, text="File Selection", padding=15)
        file_frame.pack(fill=X, pady=(0, 10))
        
        # Main input file
        ttk.Label(file_frame, text="Main Input Excel File:").grid(row=0, column=0, sticky=W)
        self.input_var = StringVar(value=self.last_paths["input_file"])
        ttk.Entry(file_frame, textvariable=self.input_var, width=50).grid(row=0, column=1, sticky=EW, padx=(10, 10))
        ttk.Button(file_frame, text="Browse", command=self.browse_input).grid(row=0, column=2)
        
        # Attachment file
        ttk.Label(file_frame, text="Attachment Data File:").grid(row=1, column=0, sticky=W, pady=(10, 0))
        self.attachment_var = StringVar(value=self.last_paths["attachment_file"])
        ttk.Entry(file_frame, textvariable=self.attachment_var, width=50).grid(row=1, column=1, sticky=EW, padx=(10, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Browse", command=self.browse_attachment).grid(row=1, column=2, pady=(10, 0))
        
        # Output file
        ttk.Label(file_frame, text="Output Template File:").grid(row=2, column=0, sticky=W, pady=(10, 0))
        self.output_var = StringVar(value=self.last_paths["output_file"])
        ttk.Entry(file_frame, textvariable=self.output_var, width=50).grid(row=2, column=1, sticky=EW, padx=(10, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Browse", command=self.browse_output).grid(row=2, column=2, pady=(10, 0))
        
        # QC file (optional)
        ttk.Label(file_frame, text="QC File (Optional):").grid(row=3, column=0, sticky=W, pady=(10, 0))
        self.qc_var = StringVar(value=self.last_paths.get("qc_file", ""))
        ttk.Entry(file_frame, textvariable=self.qc_var, width=50).grid(row=3, column=1, sticky=EW, padx=(10, 10), pady=(10, 0))
        ttk.Button(file_frame, text="Browse", command=self.browse_qc).grid(row=3, column=2, pady=(10, 0))
        
        file_frame.grid_columnconfigure(1, weight=1)
        
        # Set last directory from saved paths
        self.last_directory = self.last_paths["last_directory"]

    def create_manual_routes_section(self, parent):
        """Create manual routes section"""
        route_frame = ttk.LabelFrame(parent, text="Manual SCID Route Definition (Optional)", padding=15)
        route_frame.pack(fill=X, pady=(0, 10))
        
        # Instructions
        instructions = ttk.Label(route_frame, 
            text="Define pole-to-pole routes manually. Example: 1,2,3,4;", 
            font=("Arial", 9), foreground="gray")
        instructions.pack(anchor=W, pady=(0, 10))
          # Text area for routes
        self.route_text = ScrolledText(route_frame, height=4, font=("Consolas", 10))
        self.route_text.pack(fill=BOTH, expand=True)
        
        # Route options
        route_options_frame = ttk.Frame(route_frame)
        route_options_frame.pack(fill=X, pady=(10, 0))
        self.use_manual_routes_var = BooleanVar(value=False)
        ttk.Checkbutton(route_options_frame, text="Use manual routes", 
                       variable=self.use_manual_routes_var, command=self.toggle_route_text).pack(anchor=W)
        
        self.clear_routes_var = BooleanVar(value=False)
        ttk.Checkbutton(route_options_frame, text="Clear existing route data", 
                       variable=self.clear_routes_var).pack(anchor=W, pady=(5, 0))
        
        # Add traces for auto-saving
        def on_manual_routes_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            self.auto_save_config()
        
        def on_clear_routes_change(*args):
            if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
                return
            self.auto_save_config()
        
        self.use_manual_routes_var.trace('w', on_manual_routes_change)
        self.clear_routes_var.trace('w', on_clear_routes_change)

    def create_processing_options(self, parent):
        """Create processing options section"""
        options_frame = ttk.LabelFrame(parent, text="Processing Options", padding=15)
        options_frame.pack(fill=X, pady=(0, 10))
        
        self.geocoding_var = BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Use geocoding for addresses", variable=self.geocoding_var).pack(anchor=W)
        
        self.open_output_var = BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Open output file when complete", variable=self.open_output_var).pack(anchor=W)

    def create_progress_section(self, parent):
        """Create progress section"""
        progress_frame = ttk.LabelFrame(parent, text="Progress", padding=15)
        progress_frame.pack(fill=X, pady=(0, 10))
        
        self.progress_var = StringVar(value="Ready to process files...")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(anchor=W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.pack(fill=X, pady=(10, 0))

    def create_log_section(self, parent):
        """Create log section"""
        log_frame = ttk.LabelFrame(parent, text="Processing Log", padding=15)
        log_frame.pack(fill=BOTH, expand=True)
        
        self.log_text = ScrolledText(log_frame, height=25, font=("Consolas", 9))
        self.log_text.pack(fill=BOTH, expand=True)

    def toggle_route_text(self):
        """Enable/disable route text based on checkbox"""
        if self.use_manual_routes_var.get():
            self.route_text.config(state=NORMAL)
        else:
            self.route_text.config(state=DISABLED)

    def browse_input(self):
        """Browse for input Excel file"""
        filename = filedialog.askopenfilename(
            title="Select Main Input Excel File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.input_var.get()) if self.input_var.get() else ""
        )
        if filename:
            self.input_var.set(filename)
            self.auto_save_config()

    def browse_attachment(self):
        """Browse for attachment data file"""
        filename = filedialog.askopenfilename(
            title="Select Attachment Data File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.attachment_var.get()) if self.attachment_var.get() else ""
        )
        if filename:
            self.attachment_var.set(filename)
            self.auto_save_config()

    def browse_output(self):
        """Browse for output template file"""
        filename = filedialog.askopenfilename(
            title="Select Output Template File",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.output_var.get()) if self.output_var.get() else ""
        )
        if filename:
            self.output_var.set(filename)
            self.auto_save_config()

    def browse_qc(self):
        """Browse for QC file"""
        filename = filedialog.askopenfilename(
            title="Select QC File (Optional)",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
            initialdir=os.path.dirname(self.qc_var.get()) if self.qc_var.get() else ""
        )
        if filename:
            self.qc_var.set(filename)
            self.auto_save_config()
    


    def process_files(self):
        """Process the selected files"""
        try:
            # Validate inputs
            if not self.input_var.get():
                logging.error("Please select the main input Excel file.")
                return
            
            if not self.attachment_var.get():
                logging.error("Please select the attachment data file.")
                return
            
            if not self.output_var.get():
                logging.error("Please select the output template file.")
                return
            
            # Clear log
            self.log_text.delete(1.0, END)
            
            # Setup progress callback
            def progress_callback(percentage, message):
                self.progress_var.set(message)
                self.progress_bar['value'] = percentage
                self.root.update_idletasks()
            
            # Start processing in thread to prevent GUI freeze
            import threading
            thread = threading.Thread(target=self._process_files_worker, args=(progress_callback,))
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            logging.error(f"Error starting file processing: {e}")
            logging.error(f"Failed to start processing: {e}")

    def _process_files_worker(self, progress_callback):
        """Process files in background thread"""
        try:
            progress_callback(0, "Starting processing...")
            
            # Update config from UI
            self.update_config_from_ui()
            
            # Read input file
            progress_callback(10, "Reading main input file...")
            
            input_file = self.input_var.get()
            nodes_df = pd.read_excel(input_file, sheet_name='nodes', dtype=str).fillna("")
            connections_df = pd.read_excel(input_file, sheet_name='connections', dtype=str).fillna("")
            sections_df = pd.read_excel(input_file, sheet_name='sections', dtype=str).fillna("")
            
            logging.info(f"Read {len(nodes_df)} nodes, {len(connections_df)} connections, {len(sections_df)} sections")
            
            # Extract valid SCIDs for attachment reader filtering
            progress_callback(15, "Extracting valid SCIDs...")
            from core.utils import Utils
            nodes_df_copy = nodes_df.copy()
            nodes_df_copy['scid'] = nodes_df_copy['scid'].apply(Utils.normalize_scid)
            valid_nodes = Utils.filter_valid_nodes(nodes_df_copy)
            valid_scids = valid_nodes['scid'].tolist()
            
            logging.info(f"Found {len(valid_scids)} valid SCIDs for attachment processing")
            
            # Setup geocoder if enabled
            geocoder = None
            if self.geocoding_var.get():
                progress_callback(20, "Initializing geocoder...")
                geocoder = Geocoder(self.cache_file)
              # Setup attachment reader with valid SCIDs
            progress_callback(25, "Reading attachment data...")
            attachment_reader = AttachmentDataReader(
                self.attachment_var.get(), 
                config=self.config, 
                valid_scids=valid_scids
            )
            
            # Setup QC reader if QC file is provided
            progress_callback(27, "Processing QC file...")
            qc_reader = None
            qc_file = self.qc_var.get()
            if qc_file and qc_file.strip():
                try:
                    from core.qc_reader import QCReader
                    qc_reader = QCReader(qc_file)
                    if qc_reader.is_active():
                        logging.info(f"QC filtering enabled - {len(qc_reader.get_ordered_connections())} connections loaded")
                    else:
                        logging.info("QC file provided but no connections found - processing all connections")
                except ImportError:
                    logging.warning("QC reader not available - processing without QC filtering")
                    qc_reader = None
            else:
                logging.info("No QC file provided - processing all connections")
            
            # Setup processor
            progress_callback(30, "Initializing data processor...")
            processor = PoleDataProcessor(
                config=self.config,
                geocoder=geocoder,
                mapping_data=self.mapping_data,
                attachment_reader=attachment_reader,
                qc_reader=qc_reader
            )
              # Parse manual routes if enabled
            manual_routes = None
            if self.use_manual_routes_var.get():
                progress_callback(35, "Parsing manual routes...")
                route_text = self.route_text.get(1.0, END).strip()
                if route_text:
                    manual_routes = RouteParser.parse_manual_routes(route_text)
                    logging.info(f"Parsed {len(manual_routes)} manual routes")
              # Process data
            progress_callback(40, "Processing pole data...")
            result_data = processor.process_data(
                nodes_df=nodes_df,
                connections_df=connections_df,
                sections_df=sections_df,
                progress_callback=progress_callback,
                manual_routes=manual_routes,
                clear_existing_routes=self.clear_routes_var.get()
            )
            
            # Extract job name from nodes_df
            progress_callback(85, "Generating output file...")
            job_name = ""
            if "job_name" in nodes_df.columns and not nodes_df["job_name"].empty:
                job_name = str(nodes_df["job_name"].iloc[0]).strip()
            if not job_name:
                job_name = "Output"
                
            # Generate actual output file by copying template with job name
            template_path = self.output_var.get()
            actual_output_file = self.generate_output_file(job_name, template_path)
            if not actual_output_file:
                progress_callback(0, "Failed to generate output file!")
                return
            
            # Write output to the newly created file
            if qc_reader and qc_reader.is_active():
                progress_callback(90, "Writing output file and populating QC sheet...")
                processor.write_output(result_data, str(actual_output_file))
            else:
                progress_callback(90, "Writing output file...")
                processor.write_output(result_data, str(actual_output_file))
            
            progress_callback(100, "Processing complete!")
            logging.info(f"Processing complete. Output written to: {actual_output_file}")
            
            # Save last paths
            self.save_last_paths()
            
            # Open output file if requested
            if self.open_output_var.get():
                self.root.after(1000, lambda: self.open_output_file(str(actual_output_file)))
            
            # Log success message
            logging.info(f"Processing completed successfully! Processed {len(result_data)} poles. Output saved to: {actual_output_file}")
                
        except Exception as e:
            logging.error(f"Error during processing: {e}", exc_info=True)
            logging.error(f"An error occurred during processing: {e}")
            progress_callback(0, "Processing failed!")

    def generate_output_file(self, job_name, output_template):
        """Generate actual output file by copying the template using job_name."""
        import shutil
        from pathlib import Path
        
        template_path = Path(output_template)
        if not template_path.exists():
            logging.error(f"Output template file not found: {output_template}")
            return None
            
        # Preserve the original file extension (.xlsx or .xlsm)
        template_extension = template_path.suffix
        actual_output_file = template_path.parent / f"{job_name} Spread Sheet{template_extension}"
        logging.info(f"Generated output file path: {actual_output_file}")
        
        try:
            shutil.copy(template_path, actual_output_file)
            logging.info(f"Successfully copied template to: {actual_output_file}")
            return actual_output_file
        except Exception as e:
            logging.error(f"Error copying template file: {e}")
            logging.error(f"Failed to copy template file: {e}")
            return None

    def open_output_file(self, filepath):
        """Open the output file"""
        try:
            import subprocess
            import os
            if os.name == 'nt':  # Windows
                os.startfile(filepath)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.call(['open', filepath] if sys.platform == 'darwin' else ['xdg-open', filepath])
        except Exception as e:
            logging.warning(f"Could not open output file: {e}")

    def setup_logging(self):
        """Setup logging to display in GUI"""
        class GuiLogHandler(logging.Handler):
            def __init__(self, text_widget):
                super().__init__()
                self.text_widget = text_widget
            
            def emit(self, record):
                try:
                    msg = self.format(record)
                    self.text_widget.insert(END, msg + '\n')
                    self.text_widget.see(END)
                except Exception:
                    pass
        
        # Create handler
        if hasattr(self, 'log_text'):
            gui_handler = GuiLogHandler(self.log_text)
            gui_handler.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', 
                                        datefmt='%H:%M:%S')
            gui_handler.setFormatter(formatter)
            
            # Add to root logger
            logging.getLogger().addHandler(gui_handler)

    def auto_save_config(self):
        """Automatically save configuration with debouncing"""
        # Cancel any pending save
        if hasattr(self, '_save_timer'):
            self.root.after_cancel(self._save_timer)
        
        # Set flag to prevent recursion
        if getattr(self, '_is_saving_config', False):
            return
            
        # Schedule save after short delay (debouncing)
        self._save_timer = self.root.after(500, self._do_auto_save)

    def _do_auto_save(self):
        """Actually perform the auto save"""
        if getattr(self, '_is_saving_config', False) or getattr(self, '_is_initializing', False):
            return
            
        try:
            self._is_saving_config = True
            self.update_config_from_ui()
            self.save_config()
        except Exception as e:
            logging.error(f"Error in auto save: {e}")
        finally:
            self._is_saving_config = False
    
    def update_config_from_ui(self):
        """Update config from current UI state"""
        try:
            # Update power company
            if hasattr(self, 'power_company_var'):
                self.config["power_company"] = self.power_company_var.get()
            
            # Update output settings
            if hasattr(self, 'header_row_var'):
                try:
                    self.config["output_settings"]["header_row"] = int(self.header_row_var.get())
                except ValueError:
                    pass
            
            if hasattr(self, 'data_start_row_var'):
                try:
                    self.config["output_settings"]["data_start_row"] = int(self.data_start_row_var.get())
                except ValueError:
                    pass
            
            if hasattr(self, 'worksheet_name_var'):
                self.config["output_settings"]["worksheet_name"] = self.worksheet_name_var.get()
            
            # Update processing options
            if not "processing_options" in self.config:
                self.config["processing_options"] = {}
            
            if hasattr(self, 'geocoding_var'):
                self.config["processing_options"]["use_geocoding"] = self.geocoding_var.get()
            
            # Update manual routes options
            if not "manual_routes_options" in self.config:
                self.config["manual_routes_options"] = {}
                
            if hasattr(self, 'use_manual_routes_var'):
                self.config["manual_routes_options"]["use_manual_routes"] = self.use_manual_routes_var.get()
                
            if hasattr(self, 'clear_routes_var'):
                self.config["manual_routes_options"]["clear_existing_routes"] = self.clear_routes_var.get()
            
            # Update column mappings
            self.config["column_mappings"] = self.mapping_data
            
        except Exception as e:
            logging.error(f"Error updating config from UI: {e}")

    def on_config_change(self, event=None):
        """Handle configuration dropdown change"""
        try:
            config_name = self.config_var.get()
            if config_name != self.current_config_name:
                self.current_config_name = config_name
                self.load_config()
                self.update_ui_values()
                self.update_ui_state()
        except Exception as e:
            logging.error(f"Error changing configuration: {e}")
    
    def update_ui_values(self):
        """Update UI with current config values"""
        try:
            self._is_initializing = True
            
            # Update power company
            if hasattr(self, 'power_company_var'):
                self.power_company_var.set(self.config["power_company"])
            
            # Update output settings
            if hasattr(self, 'header_row_var'):
                self.header_row_var.set(str(self.config["output_settings"]["header_row"]))
            if hasattr(self, 'data_start_row_var'):
                self.data_start_row_var.set(str(self.config["output_settings"]["data_start_row"]))
            if hasattr(self, 'worksheet_name_var'):
                self.worksheet_name_var.set(self.config["output_settings"]["worksheet_name"])
            
            # Update processing options
            processing_options = self.config.get("processing_options", {})
            if hasattr(self, 'geocoding_var'):
                self.geocoding_var.set(processing_options.get("use_geocoding", False))
            
            # Update manual routes options
            manual_routes_options = self.config.get("manual_routes_options", {})
            if hasattr(self, 'use_manual_routes_var'):
                self.use_manual_routes_var.set(manual_routes_options.get("use_manual_routes", False))
            if hasattr(self, 'clear_routes_var'):
                self.clear_routes_var.set(manual_routes_options.get("clear_existing_routes", False))
            
            # Update route text state based on checkbox
            if hasattr(self, 'toggle_route_text'):
                self.toggle_route_text()
            
            # Update mapping data
            self.mapping_data = self.config.get("column_mappings", [])
            if not self.mapping_data:
                self.load_default_mappings()
            
        except Exception as e:
            logging.error(f"Error updating UI values: {e}")
        finally:
            self._is_initializing = False

    def update_ui_state(self):
        """Update UI state"""
        try:
            # Refresh listboxes
            for config_key, listbox in getattr(self, 'listboxes', {}).items():
                listbox.delete(0, END)
                for item in self.config[config_key]:
                    listbox.insert(END, item)
            
            # Refresh telecom keywords and mappings
            if hasattr(self, 'populate_telecom_keywords'):
                self.populate_telecom_keywords()
            
            if hasattr(self, 'populate_mappings'):
                self.populate_mappings()
                
        except Exception as e:
            logging.error(f"Error updating UI state: {e}")

    def refresh_ui(self):
        """Refresh UI components that depend on telecom providers"""
        try:
            # Update telecom keywords section
            if hasattr(self, 'populate_telecom_keywords'):
                self.populate_telecom_keywords()
            
            # Update mappings if they're already populated
            if hasattr(self, 'populate_mappings'):
                self.populate_mappings()
                
        except Exception as e:
            logging.error(f"Error refreshing UI: {e}")

    def save_config_as(self):
        """Save current configuration with a new name"""
        try:
            new_name = simpledialog.askstring("Save Configuration", "Enter configuration name:")
            if new_name is None:
                # User cancelled the dialog
                logging.info("Configuration save cancelled by user")
                return
                
            if not new_name or not new_name.strip():
                # Empty name provided
                logging.warning("Cannot save configuration with empty name")
                messagebox.showwarning("Invalid Name", "Configuration name cannot be empty.")
                return
                
            new_name = new_name.strip()
            
            # Check if configuration already exists
            if new_name in self.config_manager.get_available_configs():
                if not messagebox.askyesno("Configuration Exists", 
                                         f"Configuration '{new_name}' already exists. Do you want to overwrite it?"):
                    return
            
            # Update config from UI and save
            self.update_config_from_ui()
            success = self.config_manager.save_config(new_name, self.config)
            
            if success:
                # Refresh the dropdown and switch to the new config
                self.refresh_config_list()
                self.config_var.set(new_name)
                self.current_config_name = new_name
                logging.info(f"Configuration '{new_name}' saved successfully!")
                messagebox.showinfo("Success", f"Configuration '{new_name}' saved successfully!")
            else:
                logging.error(f"Failed to save configuration '{new_name}'")
                messagebox.showerror("Error", f"Failed to save configuration '{new_name}'. Check the logs for details.")
                
        except Exception as e:
            logging.error(f"Failed to save configuration: {e}")
            messagebox.showerror("Error", f"An error occurred while saving the configuration: {e}")

    def delete_selected_config(self):
        """Delete the selected configuration"""
        try:
            if self.current_config_name == "Default":
                logging.warning("Cannot delete the Default configuration.")
                messagebox.showwarning("Cannot Delete", "Cannot delete the Default configuration.")
                return
            
            # Confirm deletion
            if not messagebox.askyesno("Confirm Delete", 
                                     f"Are you sure you want to delete configuration '{self.current_config_name}'?"):
                return
            
            # Delete the configuration
            success = self.config_manager.delete_config(self.current_config_name)
            
            if success:
                self.refresh_config_list()
                
                # Switch to Default configuration
                self.current_config_name = "Default"
                self.config_var.set("Default")
                self.load_config()
                
                logging.info("Configuration deleted successfully!")
                messagebox.showinfo("Success", "Configuration deleted successfully!")
            else:
                logging.error(f"Failed to delete configuration '{self.current_config_name}'")
                messagebox.showerror("Error", f"Failed to delete configuration. Check the logs for details.")
                
        except Exception as e:
            logging.error(f"Failed to delete configuration: {e}")
            messagebox.showerror("Error", f"An error occurred while deleting the configuration: {e}")

    def reset_to_defaults(self):
        """Reset current configuration to defaults"""
        try:
            # No confirmation dialog - just reset
            self.config = self.config_manager.get_default_config()
            self.update_ui_values()
            self.auto_save_config()
            logging.info("Configuration reset to defaults!")
        except Exception as e:
            logging.error(f"Failed to reset configuration: {e}")

    def save_config(self):
        """Save current configuration"""
        try:
            self.config_manager.save_config(self.current_config_name, self.config)
        except Exception as e:
            logging.error(f"Error saving config: {e}")

    def on_closing(self):
        """Handle application closing"""
        try:
            # Save last paths
            self.save_last_paths()
            
            # Save current config
            self.update_config_from_ui()
            self.save_config()
            
            logging.info("Application closing")
            self.root.destroy()
        except Exception as e:
            logging.error(f"Error during application close: {e}")
            self.root.destroy()
    
    def global_exception_handler(self, exc_type, exc_value, exc_traceback):
        """Handle uncaught exceptions"""
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        
        if issubclass(exc_type, RecursionError):
            logging.error("Recursion error detected. Application will exit.")
        else:
            logging.error(f"An unexpected error occurred: {exc_value}")
            
        # Continue execution without showing message box

    def refresh_config_list(self):
        """Refresh the configuration dropdown with available configurations"""
        try:
            available_configs = self.config_manager.get_available_configs()
            self.config_combo['values'] = available_configs
            logging.debug(f"Configuration list refreshed: {available_configs}")
        except Exception as e:
            logging.error(f"Error refreshing configuration list: {e}")
