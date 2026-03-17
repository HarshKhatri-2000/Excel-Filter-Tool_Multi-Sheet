"""
Excel Filter Tool v11.0
Developed by: HARSH KHATRI
"""

import os
import sys
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.cell.cell import MergedCell
from copy import copy
import gc

if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))


class FormulaAdjuster:
    """Adjusts Excel formulas when rows change"""
    CELL_PATTERN = re.compile(r'(\$?)([A-Z]+)(\$?)(\d+)', re.IGNORECASE)
    
    @staticmethod
    def adjust(formula, src_row, tgt_row):
        if not formula or not isinstance(formula, str) or not formula.startswith('='):
            return formula
        diff = tgt_row - src_row
        def replace(m):
            col_abs, col, row_abs, row = m.groups()
            if row_abs == '$':
                return f"{col_abs}{col}{row_abs}{row}"
            return f"{col_abs}{col}{row_abs}{max(1, int(row) + diff)}"
        try:
            return FormulaAdjuster.CELL_PATTERN.sub(replace, formula)
        except:
            return formula
    
    @staticmethod
    def is_complex(formula):
        if not formula:
            return False
        patterns = [r'!', r'\[', r'INDIRECT\(', r'OFFSET\(', r'VLOOKUP\(', r'HLOOKUP\(', r'XLOOKUP\(', r'INDEX\(', r'MATCH\(']
        return any(re.search(p, formula, re.IGNORECASE) for p in patterns)


class ExcelFilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Filter Tool v11.0 - By HARSH KHATRI")
        self.root.configure(bg='#f0f0f0')
        
        # Window setup
        self.root.geometry("1300x800")
        self.root.minsize(1100, 700)
        
        # Variables
        self.source_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.manager_column = tk.StringVar(value="E")
        self.selected_sheet = tk.StringVar(value="ALL")
        self.formula_mode = tk.StringVar(value="smart")
        self.include_header = tk.BooleanVar(value=True)
        self.preserve_formatting = tk.BooleanVar(value=True)
        self.skip_empty = tk.BooleanVar(value=True)
        self.create_summary = tk.BooleanVar(value=False)
        
        self.is_processing = False
        self.cancel_flag = False
        self.available_managers = []
        self.available_sheets = []
        self.manager_rows = {}
        self.start_time = None
        self.stats = {'files': 0, 'rows': 0, 'formulas_kept': 0, 'formulas_conv': 0}
        
        try:
            self.output_folder.set(os.path.join(os.path.expanduser("~"), "Desktop", "Filtered_Output"))
        except:
            pass
        
        self.build_ui()
    
    def build_ui(self):
        # ============================================
        # HEADER - WITH DEVELOPER NAME
        # ============================================
        header_frame = tk.Frame(self.root, bg='#1565c0', height=70)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)
        
        # Left side - Title
        title_frame = tk.Frame(header_frame, bg='#1565c0')
        title_frame.pack(side='left', padx=20, pady=10)
        
        tk.Label(title_frame, text="📊 Excel Filter Tool", font=('Arial', 22, 'bold'),
                bg='#1565c0', fg='white').pack(anchor='w')
        tk.Label(title_frame, text="Filter & Split Excel Data by Column Values", 
                font=('Arial', 10), bg='#1565c0', fg='#bbdefb').pack(anchor='w')
        
        # Right side - DEVELOPER NAME (PROMINENT)
        dev_frame = tk.Frame(header_frame, bg='#0d47a1', padx=20, pady=8)
        dev_frame.pack(side='right', padx=20, pady=10)
        
        tk.Label(dev_frame, text="✦ DEVELOPED BY ✦", font=('Arial', 9, 'bold'),
                bg='#0d47a1', fg='#90caf9').pack()
        tk.Label(dev_frame, text="HARSH KHATRI", font=('Arial', 16, 'bold'),
                bg='#0d47a1', fg='white').pack()
        
        # ============================================
        # MAIN SCROLLABLE AREA
        # ============================================
        container = tk.Frame(self.root, bg='#f0f0f0')
        container.pack(fill='both', expand=True, side='top')
        
        # Canvas with scrollbars
        self.canvas = tk.Canvas(container, bg='#f0f0f0', highlightthickness=0)
        v_scroll = ttk.Scrollbar(container, orient='vertical', command=self.canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient='horizontal', command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        v_scroll.pack(side='right', fill='y')
        h_scroll.pack(side='bottom', fill='x')
        self.canvas.pack(side='left', fill='both', expand=True)
        
        # Main content frame inside canvas
        self.main_frame = tk.Frame(self.canvas, bg='#f0f0f0')
        self.canvas_window = self.canvas.create_window((0, 0), window=self.main_frame, anchor='nw')
        
        # Bind resize events
        self.main_frame.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox('all')))
        self.canvas.bind('<Configure>', self._on_canvas_resize)
        
        # Mouse wheel for main canvas
        self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        
        # ============================================
        # TWO COLUMN LAYOUT
        # ============================================
        content = tk.Frame(self.main_frame, bg='#f0f0f0')
        content.pack(fill='both', expand=True, padx=15, pady=15)
        
        # LEFT COLUMN
        left_col = tk.Frame(content, bg='#f0f0f0')
        left_col.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        # RIGHT COLUMN - NO pack_propagate(False) so buttons show!
        right_col = tk.Frame(content, bg='#f0f0f0')
        right_col.pack(side='right', fill='y', padx=(10, 0))
        
        # ============================================
        # LEFT COLUMN - FILE SELECTION
        # ============================================
        file_card = self._create_card(left_col, "📁 File Selection")
        
        # Source file
        src_frame = tk.Frame(file_card, bg='white')
        src_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(src_frame, text="Source File:", font=('Arial', 10, 'bold'), bg='white').pack(anchor='w')
        
        src_row = tk.Frame(src_frame, bg='white')
        src_row.pack(fill='x', pady=5)
        
        self.src_entry = tk.Entry(src_row, textvariable=self.source_file, font=('Arial', 10))
        self.src_entry.pack(side='left', fill='x', expand=True, ipady=5)
        
        tk.Button(src_row, text="📂 Browse", command=self.browse_source,
                 bg='#2196f3', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=15, cursor='hand2').pack(side='left', padx=(10, 5))
        
        tk.Button(src_row, text="🔍 Load Data", command=self.load_data,
                 bg='#9c27b0', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=15, cursor='hand2').pack(side='left')
        
        # Output folder
        out_frame = tk.Frame(file_card, bg='white')
        out_frame.pack(fill='x')
        
        tk.Label(out_frame, text="Output Folder:", font=('Arial', 10, 'bold'), bg='white').pack(anchor='w')
        
        out_row = tk.Frame(out_frame, bg='white')
        out_row.pack(fill='x', pady=5)
        
        self.out_entry = tk.Entry(out_row, textvariable=self.output_folder, font=('Arial', 10))
        self.out_entry.pack(side='left', fill='x', expand=True, ipady=5)
        
        tk.Button(out_row, text="📂 Browse", command=self.browse_output,
                 bg='#2196f3', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=15, cursor='hand2').pack(side='left', padx=(10, 0))
        
        # ============================================
        # LEFT COLUMN - OPTIONS
        # ============================================
        opt_card = self._create_card(left_col, "⚙️ Options")
        
        opt_row1 = tk.Frame(opt_card, bg='white')
        opt_row1.pack(fill='x', pady=(0, 10))
        
        # Sheet selection
        sheet_frame = tk.Frame(opt_row1, bg='white')
        sheet_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(sheet_frame, text="Sheet:", font=('Arial', 10, 'bold'), bg='white').pack(anchor='w')
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.selected_sheet,
                                        values=["ALL"], width=20, state='readonly')
        self.sheet_combo.pack(pady=3)
        
        # Fix combobox scroll - unbind main scroll completely when over combobox
        def on_sheet_combo_enter(event):
            self.canvas.unbind_all('<MouseWheel>')
        def on_sheet_combo_leave(event):
            self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        def on_sheet_combo_open(event):
            self.canvas.unbind_all('<MouseWheel>')
        def on_sheet_combo_close(event):
            self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        self.sheet_combo.bind('<Enter>', on_sheet_combo_enter)
        self.sheet_combo.bind('<Leave>', on_sheet_combo_leave)
        self.sheet_combo.bind('<<ComboboxSelected>>', on_sheet_combo_close)
        self.sheet_combo.bind('<Button-1>', on_sheet_combo_open)
        self.sheet_combo.bind('<MouseWheel>', lambda e: "break")
        self.sheet_combo.bind('<FocusOut>', on_sheet_combo_close)
        
        # Column selection
        col_frame = tk.Frame(opt_row1, bg='white')
        col_frame.pack(side='left', padx=(0, 20))
        
        tk.Label(col_frame, text="Column:", font=('Arial', 10, 'bold'), bg='white').pack(anchor='w')
        self.col_combo = ttk.Combobox(col_frame, textvariable=self.manager_column,
                                      values=list('ABCDEFGHIJKLMNOPQRSTUVWXYZ'), width=8, state='readonly')
        self.col_combo.pack(pady=3)
        
        # Fix column combobox scroll - unbind main scroll completely when over combobox
        def on_col_combo_enter(event):
            self.canvas.unbind_all('<MouseWheel>')
        def on_col_combo_leave(event):
            self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        def on_col_combo_open(event):
            self.canvas.unbind_all('<MouseWheel>')
        def on_col_combo_close(event):
            self.canvas.bind_all('<MouseWheel>', self._on_mousewheel)
        self.col_combo.bind('<Enter>', on_col_combo_enter)
        self.col_combo.bind('<Leave>', on_col_combo_leave)
        self.col_combo.bind('<<ComboboxSelected>>', on_col_combo_close)
        self.col_combo.bind('<Button-1>', on_col_combo_open)
        self.col_combo.bind('<MouseWheel>', lambda e: "break")
        self.col_combo.bind('<FocusOut>', on_col_combo_close)
        
        # Formula mode
        formula_frame = tk.Frame(opt_row1, bg='white')
        formula_frame.pack(side='left')
        
        tk.Label(formula_frame, text="Formula Mode:", font=('Arial', 10, 'bold'), bg='white').pack(anchor='w')
        fm_row = tk.Frame(formula_frame, bg='white')
        fm_row.pack()
        
        for text, val in [("Smart", "smart"), ("Keep All", "keep"), ("Values Only", "values")]:
            tk.Radiobutton(fm_row, text=text, variable=self.formula_mode, value=val,
                          bg='white', font=('Arial', 9)).pack(side='left', padx=5)
        
        # Checkboxes
        opt_row2 = tk.Frame(opt_card, bg='white')
        opt_row2.pack(fill='x')
        
        for text, var in [("Include Headers", self.include_header),
                          ("Preserve Formatting", self.preserve_formatting),
                          ("Skip Empty Values", self.skip_empty),
                          ("Create Summary", self.create_summary)]:
            tk.Checkbutton(opt_row2, text=text, variable=var, bg='white',
                          font=('Arial', 9)).pack(side='left', padx=(0, 15))
        
        # ============================================
        # LEFT COLUMN - VALUE SELECTION
        # ============================================
        val_card = self._create_card(left_col, "👥 Select Values to Filter")
        
        # Search row
        search_row = tk.Frame(val_card, bg='white')
        search_row.pack(fill='x', pady=(0, 10))
        
        tk.Label(search_row, text="🔍", font=('Arial', 12), bg='white').pack(side='left')
        
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *a: self.filter_list())
        search_entry = tk.Entry(search_row, textvariable=self.search_var, font=('Arial', 10), width=25)
        search_entry.pack(side='left', padx=5, ipady=3)
        
        tk.Button(search_row, text="✅ All", command=self.select_all,
                 bg='#4caf50', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=10, cursor='hand2').pack(side='left', padx=3)
        
        tk.Button(search_row, text="❌ None", command=self.deselect_all,
                 bg='#f44336', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=10, cursor='hand2').pack(side='left', padx=3)
        
        tk.Button(search_row, text="🔄 Invert", command=self.invert_selection,
                 bg='#ff9800', fg='white', font=('Arial', 9, 'bold'),
                 relief='flat', padx=10, cursor='hand2').pack(side='left', padx=3)
        
        self.count_label = tk.Label(search_row, text="0 / 0 selected", font=('Arial', 10, 'bold'),
                                    bg='white', fg='#1565c0')
        self.count_label.pack(side='right')
        
        # Listbox
        list_frame = tk.Frame(val_card, bg='white')
        list_frame.pack(fill='both', expand=True)
        
        list_scroll = ttk.Scrollbar(list_frame, orient='vertical')
        list_scroll.pack(side='right', fill='y')
        
        self.listbox = tk.Listbox(list_frame, selectmode='multiple', font=('Arial', 10),
                                  bg='#fafafa', selectbackground='#1565c0',
                                  height=10, yscrollcommand=list_scroll.set,
                                  relief='solid', bd=1, exportselection=False)
        self.listbox.pack(side='left', fill='both', expand=True)
        list_scroll.config(command=self.listbox.yview)
        self.listbox.bind('<<ListboxSelect>>', lambda e: self.update_count())
        
        # Bind mouse wheel for listbox ONLY
        self.listbox.bind('<Enter>', lambda e: self.listbox.bind_all('<MouseWheel>', self._on_listbox_scroll))
        self.listbox.bind('<Leave>', lambda e: self.canvas.bind_all('<MouseWheel>', self._on_mousewheel))
        
        # ============================================
        # RIGHT COLUMN - STATISTICS
        # ============================================
        stat_card = self._create_card(right_col, "📈 Statistics")
        
        self.stat_labels = {}
        stats_info = [
            ('sheets', '📑 Sheets:', '#2196f3'),
            ('values', '👥 Unique Values:', '#9c27b0'),
            ('files', '📁 Files Created:', '#4caf50'),
            ('rows', '📋 Total Rows:', '#00bcd4'),
            ('formulas', '📝 Formulas Kept:', '#ff9800'),
            ('converted', '🔄 Converted:', '#e91e63'),
            ('time', '⏱️ Elapsed Time:', '#607d8b')
        ]
        
        for key, label, color in stats_info:
            row = tk.Frame(stat_card, bg='white')
            row.pack(fill='x', pady=3)
            tk.Label(row, text=label, font=('Arial', 10), bg='white').pack(side='left')
            self.stat_labels[key] = tk.Label(row, text='0' if key != 'time' else '--',
                                             font=('Arial', 11, 'bold'), bg='white', fg=color)
            self.stat_labels[key].pack(side='right')
        
        # ============================================
        # RIGHT COLUMN - PROGRESS
        # ============================================
        prog_card = self._create_card(right_col, "📊 Progress")
        
        self.progress = ttk.Progressbar(prog_card, mode='determinate', length=320)
        self.progress.pack(fill='x', pady=(0, 10))
        
        self.progress_label = tk.Label(prog_card, text="Ready", font=('Arial', 11, 'bold'),
                                       bg='white', fg='#4caf50')
        self.progress_label.pack()
        
        self.detail_label = tk.Label(prog_card, text="", font=('Arial', 9), bg='white', fg='#757575')
        self.detail_label.pack()
        
        # ============================================
        # RIGHT COLUMN - LOG
        # ============================================
        log_card = self._create_card(right_col, "📝 Activity Log")
        
        # Log toolbar
        log_toolbar = tk.Frame(log_card, bg='white')
        log_toolbar.pack(fill='x', pady=(0, 5))
        
        tk.Button(log_toolbar, text="🗑️ Clear", command=self.clear_log,
                 bg='#607d8b', fg='white', font=('Arial', 8),
                 relief='flat', padx=8, cursor='hand2').pack(side='left', padx=2)
        
        tk.Button(log_toolbar, text="💾 Export", command=self.export_log,
                 bg='#9c27b0', fg='white', font=('Arial', 8),
                 relief='flat', padx=8, cursor='hand2').pack(side='left', padx=2)
        
        self.log_count = tk.Label(log_toolbar, text="0 entries", font=('Arial', 8), bg='white', fg='#757575')
        self.log_count.pack(side='right')
        
        # Log text
        log_frame = tk.Frame(log_card, bg='#1a1a2e')
        log_frame.pack(fill='both', expand=True)
        
        log_scroll = ttk.Scrollbar(log_frame, orient='vertical')
        log_scroll.pack(side='right', fill='y')
        
        self.log_text = tk.Text(log_frame, font=('Consolas', 9), bg='#1a1a2e', fg='#4caf50',
                               height=8, wrap='word', state='disabled')
        self.log_text.pack(side='left', fill='both', expand=True)
        log_scroll.config(command=self.log_text.yview)
        self.log_text.config(yscrollcommand=log_scroll.set)
        
        # Log tags
        self.log_text.tag_configure('success', foreground='#4caf50')
        self.log_text.tag_configure('error', foreground='#f44336')
        self.log_text.tag_configure('info', foreground='#2196f3')
        self.log_text.tag_configure('warning', foreground='#ff9800')
        
        # Bind mouse wheel for log ONLY
        self.log_text.bind('<Enter>', lambda e: self.log_text.bind_all('<MouseWheel>', self._on_log_scroll))
        self.log_text.bind('<Leave>', lambda e: self.canvas.bind_all('<MouseWheel>', self._on_mousewheel))
        
        # ============================================
        # RIGHT COLUMN - ACTION BUTTONS (NO CARD - DIRECT)
        # ============================================
        action_frame = tk.Frame(right_col, bg='#f0f0f0')
        action_frame.pack(fill='x', pady=(10, 0))
        
        # PROCESS BUTTON
        self.process_btn = tk.Button(action_frame, text="🚀 PROCESS FILES", command=self.start_process,
                                     bg='#4caf50', fg='white', font=('Arial', 11, 'bold'),
                                     relief='flat', pady=8, cursor='hand2')
        self.process_btn.pack(fill='x', pady=(0, 8))
        
        # Other buttons row 1
        btn_row1 = tk.Frame(action_frame, bg='#f0f0f0')
        btn_row1.pack(fill='x', pady=(0, 5))
        
        self.cancel_btn = tk.Button(btn_row1, text="⏹️ Cancel", command=self.cancel_process,
                                    bg='#9e9e9e', fg='white', font=('Arial', 9),
                                    relief='flat', padx=10, pady=5, cursor='hand2', state='disabled')
        self.cancel_btn.pack(side='left', fill='x', expand=True, padx=(0, 3))
        
        tk.Button(btn_row1, text="📁 Open Folder", command=self.open_folder,
                 bg='#9c27b0', fg='white', font=('Arial', 9),
                 relief='flat', padx=10, pady=5, cursor='hand2').pack(side='left', fill='x', expand=True)
        
        # Other buttons row 2
        btn_row2 = tk.Frame(action_frame, bg='#f0f0f0')
        btn_row2.pack(fill='x')
        
        tk.Button(btn_row2, text="🔄 Reset", command=self.reset_app,
                 bg='#607d8b', fg='white', font=('Arial', 9),
                 relief='flat', padx=10, pady=5, cursor='hand2').pack(side='left', fill='x', expand=True, padx=(0, 3))
        
        tk.Button(btn_row2, text="❓ Help", command=self.show_help,
                 bg='#2196f3', fg='white', font=('Arial', 9),
                 relief='flat', padx=10, pady=5, cursor='hand2').pack(side='left', fill='x', expand=True, padx=(0, 3))
        
        tk.Button(btn_row2, text="❌ Exit", command=self.on_exit,
                 bg='#f44336', fg='white', font=('Arial', 9),
                 relief='flat', padx=10, pady=5, cursor='hand2').pack(side='left', fill='x', expand=True)
        
        # ============================================
        # FOOTER
        # ============================================
        footer = tk.Frame(self.root, bg='#1565c0', height=30)
        footer.pack(fill='x', side='bottom')
        footer.pack_propagate(False)
        
        tk.Label(footer, text="Excel Filter Tool v11.0 | Developed by HARSH KHATRI | © 2024",
                font=('Arial', 9), bg='#1565c0', fg='white').pack(pady=5)
        
        # Log welcome message
        self.log_msg("Excel Filter Tool v11.0", 'info')
        self.log_msg("Developed by HARSH KHATRI", 'info')
        self.log_msg("─" * 40, None)
        self.log_msg("Ready! Load an Excel file to begin.", 'success')
    
    def _create_card(self, parent, title):
        """Create a card-style frame"""
        card = tk.Frame(parent, bg='white', padx=15, pady=12, relief='solid', bd=1)
        card.pack(fill='x', pady=(0, 10))
        
        if title:
            tk.Label(card, text=title, font=('Arial', 11, 'bold'), bg='white').pack(anchor='w', pady=(0, 10))
        
        return card
    
    def _on_canvas_resize(self, event):
        self.canvas.itemconfig(self.canvas_window, width=event.width)
    
    def _on_mousewheel(self, event):
        # Check if mouse is over a combobox popup (dropdown list)
        widget = event.widget
        widget_name = str(widget)
        
        # If it's a combobox popdown, don't scroll main UI
        if 'popdown' in widget_name.lower() or 'combobox' in widget_name.lower():
            return "break"
        
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
    
    def _on_listbox_scroll(self, event):
        self.listbox.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        return 'break'
    
    def _on_log_scroll(self, event):
        self.log_text.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        return 'break'
    
    # ============================================
    # LOGGING
    # ============================================
    def log_msg(self, msg, tag=None):
        self.log_text.config(state='normal')
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert('end', f"[{ts}] {msg}\n", tag)
        self.log_text.see('end')
        self.log_text.config(state='disabled')
        lines = int(self.log_text.index('end-1c').split('.')[0])
        self.log_count.config(text=f"{lines} entries")
    
    def clear_log(self):
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, 'end')
        self.log_text.config(state='disabled')
        self.log_count.config(text="0 entries")
    
    def export_log(self):
        path = filedialog.asksaveasfilename(defaultextension=".txt",
                                            filetypes=[("Text files", "*.txt")])
        if path:
            self.log_text.config(state='normal')
            with open(path, 'w') as f:
                f.write(self.log_text.get(1.0, 'end'))
            self.log_text.config(state='disabled')
            self.log_msg(f"Log exported to {path}", 'success')
    
    # ============================================
    # FILE OPERATIONS
    # ============================================
    def browse_source(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if f:
            self.source_file.set(f)
            self.log_msg(f"Selected: {os.path.basename(f)}", 'info')
            self.listbox.delete(0, 'end')
            self.available_managers = []
            self.sheet_combo['values'] = ["ALL"]
            self.selected_sheet.set("ALL")
    
    def browse_output(self):
        f = filedialog.askdirectory()
        if f:
            self.output_folder.set(f)
            self.log_msg(f"Output folder: {f}", 'info')
    
    def open_folder(self):
        f = self.output_folder.get()
        if os.path.exists(f):
            os.startfile(f)
        else:
            messagebox.showwarning("Warning", "Folder doesn't exist yet")
    
    # ============================================
    # DATA LOADING
    # ============================================
    def load_data(self):
        if not self.source_file.get() or not os.path.exists(self.source_file.get()):
            messagebox.showerror("Error", "Please select a valid Excel file")
            return
        
        threading.Thread(target=self._load_thread, daemon=True).start()
    
    def _load_thread(self):
        try:
            self.root.after(0, lambda: self.log_msg("Loading data...", 'info'))
            self.root.after(0, lambda: self.progress_label.config(text="Loading...", fg='#2196f3'))
            
            start = datetime.now()
            col_idx = column_index_from_string(self.manager_column.get())
            skip_empty = self.skip_empty.get()
            
            wb = load_workbook(self.source_file.get(), read_only=True, data_only=True)
            self.available_sheets = wb.sheetnames[:]
            
            selected = self.selected_sheet.get()
            sheets = self.available_sheets if selected == "ALL" else [selected]
            
            unique = set()
            self.manager_rows = {}
            
            for sn in sheets:
                if sn not in wb.sheetnames:
                    continue
                ws = wb[sn]
                row_num = 1
                for row in ws.iter_rows(min_row=2, values_only=True):
                    row_num += 1
                    if row and len(row) >= col_idx:
                        val = row[col_idx - 1]
                        if val:
                            val_str = str(val).strip()
                            if skip_empty and not val_str:
                                continue
                            if val_str.lower() not in ['nan', 'none', '']:
                                if val_str.lower() not in ['manager', 'name', 'employee', 'id']:
                                    unique.add(val_str)
                                    if val_str not in self.manager_rows:
                                        self.manager_rows[val_str] = {}
                                    if sn not in self.manager_rows[val_str]:
                                        self.manager_rows[val_str][sn] = []
                                    self.manager_rows[val_str][sn].append(row_num)
            
            wb.close()
            self.available_managers = sorted(unique, key=lambda x: x.lower())
            elapsed = (datetime.now() - start).total_seconds()
            
            def update_ui():
                self.sheet_combo['values'] = ["ALL"] + self.available_sheets
                self.listbox.delete(0, 'end')
                for m in self.available_managers:
                    cnt = sum(len(r) for r in self.manager_rows.get(m, {}).values())
                    self.listbox.insert('end', f"{m}  ({cnt} rows)")
                
                self.stat_labels['sheets'].config(text=str(len(self.available_sheets)))
                self.stat_labels['values'].config(text=str(len(self.available_managers)))
                self.progress_label.config(text="Ready", fg='#4caf50')
                self.log_msg(f"Found {len(self.available_managers)} unique values in {elapsed:.1f}s", 'success')
                self.update_count()
            
            self.root.after(0, update_ui)
            
        except Exception as e:
            self.root.after(0, lambda: self.log_msg(f"Error: {e}", 'error'))
            self.root.after(0, lambda: self.progress_label.config(text="Error", fg='#f44336'))
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
    
    # ============================================
    # LIST OPERATIONS
    # ============================================
    def filter_list(self):
        search = self.search_var.get().lower()
        self.listbox.delete(0, 'end')
        for m in self.available_managers:
            if search in m.lower():
                cnt = sum(len(r) for r in self.manager_rows.get(m, {}).values())
                self.listbox.insert('end', f"{m}  ({cnt} rows)")
        self.update_count()
    
    def select_all(self):
        self.listbox.select_set(0, 'end')
        self.update_count()
    
    def deselect_all(self):
        self.listbox.selection_clear(0, 'end')
        self.update_count()
    
    def invert_selection(self):
        for i in range(self.listbox.size()):
            if self.listbox.selection_includes(i):
                self.listbox.selection_clear(i)
            else:
                self.listbox.select_set(i)
        self.update_count()
    
    def update_count(self):
        sel = len(self.listbox.curselection())
        tot = self.listbox.size()
        self.count_label.config(text=f"{sel} / {tot} selected")
    
    # ============================================
    # PROCESSING
    # ============================================
    def start_process(self):
        if self.is_processing:
            return
        
        if not self.source_file.get() or not os.path.exists(self.source_file.get()):
            messagebox.showerror("Error", "Please select a valid source file")
            return
        
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return
        
        if not self.available_managers:
            messagebox.showerror("Error", "Please load data first")
            return
        
        indices = self.listbox.curselection()
        if not indices:
            selected = self.available_managers[:]
        else:
            selected = [self.listbox.get(i).rsplit('  (', 1)[0] for i in indices]
        
        if not messagebox.askyesno("Confirm", f"Process {len(selected)} value(s)?"):
            return
        
        self.is_processing = True
        self.cancel_flag = False
        self.start_time = datetime.now()
        self.stats = {'files': 0, 'rows': 0, 'formulas_kept': 0, 'formulas_conv': 0}
        
        self.process_btn.config(state='disabled', bg='#9e9e9e')
        self.cancel_btn.config(state='normal', bg='#f44336')
        
        threading.Thread(target=self._process_thread, args=(selected,), daemon=True).start()
    
    def cancel_process(self):
        self.cancel_flag = True
        self.cancel_btn.config(state='disabled')
        self.log_msg("Cancelling...", 'warning')
    
    def _process_thread(self, selected):
        try:
            output = self.output_folder.get()
            col_idx = column_index_from_string(self.manager_column.get())
            formula_mode = self.formula_mode.get()
            include_header = self.include_header.get()
            preserve_fmt = self.preserve_formatting.get()
            create_summary = self.create_summary.get()
            
            os.makedirs(output, exist_ok=True)
            
            self.root.after(0, lambda: self.log_msg("─" * 40, None))
            self.root.after(0, lambda: self.log_msg(f"Processing {len(selected)} values...", 'info'))
            
            src_wb = load_workbook(self.source_file.get(), data_only=False)
            src_wb_val = load_workbook(self.source_file.get(), data_only=True)
            
            sheet_sel = self.selected_sheet.get()
            sheets = src_wb.sheetnames if sheet_sel == "ALL" else [sheet_sel]
            
            # Cache sheet info
            cache = {}
            for sn in sheets:
                if sn in src_wb.sheetnames:
                    ws = src_wb[sn]
                    cache[sn] = {
                        'ws': ws, 'ws_val': src_wb_val[sn],
                        'max_col': ws.max_column,
                        'widths': {c: ws.column_dimensions[c].width for c in ws.column_dimensions if ws.column_dimensions[c].width}
                    }
            
            total = len(selected)
            summary = []
            
            for idx, val in enumerate(selected):
                if self.cancel_flag:
                    self.root.after(0, lambda: self.log_msg("Cancelled!", 'warning'))
                    break
                
                pct = int((idx + 1) / total * 100)
                self.root.after(0, lambda p=pct, v=val[:25], i=idx, t=total: (
                    self.progress.configure(value=p),
                    self.progress_label.config(text=f"Processing {i+1}/{t}", fg='#2196f3'),
                    self.detail_label.config(text=v),
                    self._update_stats()
                ))
                
                self.root.after(0, lambda v=val: self.log_msg(f"▶ {v}", 'info'))
                
                new_wb = Workbook()
                del new_wb['Sheet']
                
                val_rows = 0
                sheets_added = 0
                
                for sn in sheets:
                    rows = self.manager_rows.get(val, {}).get(sn, [])
                    if not rows:
                        continue
                    
                    c = cache.get(sn)
                    if not c:
                        continue
                    
                    ws_src, ws_val = c['ws'], c['ws_val']
                    max_col = c['max_col']
                    
                    clean = self._clean(sn, 31)
                    ws_new = new_wb.create_sheet(title=clean)
                    sheets_added += 1
                    
                    if preserve_fmt:
                        for col, w in c['widths'].items():
                            ws_new.column_dimensions[col].width = w
                    
                    if include_header:
                        self._copy_row(ws_src, ws_val, 1, ws_new, 1, max_col, formula_mode, 1, 1)
                    
                    tgt = 2 if include_header else 1
                    for sr in rows:
                        self._copy_row(ws_src, ws_val, sr, ws_new, tgt, max_col, formula_mode, sr, tgt)
                        tgt += 1
                        val_rows += 1
                    
                    self.root.after(0, lambda s=clean, r=len(rows): self.log_msg(f"   ✓ {s}: {r} rows", 'success'))
                
                if sheets_added > 0:
                    fname = self._clean(val, 100) or f"Value_{idx+1}"
                    new_wb.save(os.path.join(output, f"{fname}.xlsx"))
                    self.stats['files'] += 1
                    self.stats['rows'] += val_rows
                    summary.append({'value': val, 'sheets': sheets_added, 'rows': val_rows, 'file': f"{fname}.xlsx"})
                    self.root.after(0, lambda f=fname: self.log_msg(f"   💾 {f}.xlsx", 'success'))
                
                new_wb.close()
                if idx % 10 == 0:
                    gc.collect()
            
            if create_summary and summary and not self.cancel_flag:
                self._make_summary(output, summary)
            
            src_wb.close()
            src_wb_val.close()
            gc.collect()
            
            elapsed = (datetime.now() - self.start_time).total_seconds()
            
            self.root.after(0, lambda: (
                self.progress.configure(value=100),
                self.progress_label.config(text="Complete!", fg='#4caf50'),
                self.detail_label.config(text=""),
                self._update_stats(),
                self.log_msg("─" * 40, None),
                self.log_msg(f"Done! {self.stats['files']} files in {elapsed:.1f}s", 'success')
            ))
            
            if not self.cancel_flag:
                self.root.after(0, lambda: messagebox.showinfo("Success",
                    f"Processing Complete!\n\n"
                    f"Files: {self.stats['files']}\n"
                    f"Rows: {self.stats['rows']}\n"
                    f"Time: {elapsed:.1f}s\n\n"
                    f"Output: {output}"))
        
        except Exception as e:
            self.root.after(0, lambda: self.log_msg(f"Error: {e}", 'error'))
            self.root.after(0, lambda: self.progress_label.config(text="Error", fg='#f44336'))
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        
        finally:
            self.is_processing = False
            self.root.after(0, lambda: self.process_btn.config(state='normal', bg='#4caf50'))
            self.root.after(0, lambda: self.cancel_btn.config(state='disabled', bg='#9e9e9e'))
            gc.collect()
    
    def _copy_row(self, ws_src, ws_val, sr, ws_tgt, tr, max_col, mode, src_row, tgt_row):
        for c in range(1, max_col + 1):
            src = ws_src.cell(row=sr, column=c)
            tgt = ws_tgt.cell(row=tr, column=c)
            
            if isinstance(src, MergedCell):
                continue
            
            val = src.value
            
            if isinstance(val, str) and val.startswith('='):
                calc = ws_val.cell(row=sr, column=c).value
                if mode == "values":
                    tgt.value = calc
                    self.stats['formulas_conv'] += 1
                elif mode == "keep":
                    tgt.value = FormulaAdjuster.adjust(val, src_row, tgt_row)
                    self.stats['formulas_kept'] += 1
                else:  # smart
                    if FormulaAdjuster.is_complex(val):
                        tgt.value = calc
                        self.stats['formulas_conv'] += 1
                    else:
                        tgt.value = FormulaAdjuster.adjust(val, src_row, tgt_row)
                        self.stats['formulas_kept'] += 1
            else:
                tgt.value = val
            
            if src.has_style:
                try:
                    tgt.font = copy(src.font)
                    tgt.fill = copy(src.fill)
                    tgt.border = copy(src.border)
                    tgt.alignment = copy(src.alignment)
                    tgt.number_format = src.number_format
                except:
                    pass
    
    def _update_stats(self):
        self.stat_labels['files'].config(text=str(self.stats['files']))
        self.stat_labels['rows'].config(text=str(self.stats['rows']))
        self.stat_labels['formulas'].config(text=str(self.stats['formulas_kept']))
        self.stat_labels['converted'].config(text=str(self.stats['formulas_conv']))
        if self.start_time:
            self.stat_labels['time'].config(text=f"{(datetime.now() - self.start_time).total_seconds():.1f}s")
    
    def _make_summary(self, folder, data):
        wb = Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws.append(['Value', 'Sheets', 'Rows', 'File'])
        for d in data:
            ws.append([d['value'], d['sheets'], d['rows'], d['file']])
        ws.column_dimensions['A'].width = 40
        ws.column_dimensions['D'].width = 50
        wb.save(os.path.join(folder, "_Summary.xlsx"))
        wb.close()
        self.root.after(0, lambda: self.log_msg("Summary created: _Summary.xlsx", 'success'))
    
    def _clean(self, name, max_len):
        if not name:
            return ""
        clean = str(name).strip()
        for c in '/\\:*?"<>|\n\r':
            clean = clean.replace(c, '_')
        return clean[:max_len].strip(' .')
    
    # ============================================
    # OTHER FUNCTIONS
    # ============================================
    def reset_app(self):
        self.source_file.set("")
        self.available_managers = []
        self.manager_rows = {}
        self.listbox.delete(0, 'end')
        self.search_var.set("")
        self.sheet_combo['values'] = ["ALL"]
        self.selected_sheet.set("ALL")
        self.progress.configure(value=0)
        self.progress_label.config(text="Ready", fg='#4caf50')
        self.detail_label.config(text="")
        self.stats = {'files': 0, 'rows': 0, 'formulas_kept': 0, 'formulas_conv': 0}
        for k in self.stat_labels:
            self.stat_labels[k].config(text='0' if k != 'time' else '--')
        self.clear_log()
        self.log_msg("Application reset", 'info')
        gc.collect()
    
    def show_help(self):
        help_text = """
EXCEL FILTER TOOL v11.0
Developed by: HARSH KHATRI

HOW TO USE:
────────────────────────────────
1. Click 'Browse' to select an Excel file
2. Select Sheet (ALL or specific)
3. Choose Column to filter by
4. Click 'Load Data' to analyze
5. Select values from the list
6. Click 'PROCESS FILES'

FORMULA MODES:
────────────────────────────────
• Smart: Auto-adjusts simple formulas,
  converts complex ones to values
• Keep All: Adjusts all formulas
• Values Only: Converts all to values

OPTIONS:
────────────────────────────────
• Include Headers: Copy header row
• Preserve Formatting: Keep column widths
• Skip Empty Values: Ignore blanks
• Create Summary: Generate summary file
        """
        messagebox.showinfo("Help - Excel Filter Tool", help_text)
    
    def on_exit(self):
        if self.is_processing:
            if messagebox.askyesno("Confirm", "Processing in progress. Exit anyway?"):
                self.cancel_flag = True
                self.root.after(500, self.root.destroy)
        elif messagebox.askokcancel("Exit", "Exit application?"):
            self.root.destroy()


def main():
    root = tk.Tk()
    app = ExcelFilterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
