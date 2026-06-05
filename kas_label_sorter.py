"""
KAS Enterprises Label Sorter
A Windows application to sort shipping labels by order number.
Features drag-and-drop file input, single-line output, and duplicate order reporting.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import os
import sys
import json
import tempfile
import subprocess
from datetime import datetime
from collections import defaultdict

# Try to import PIL for image handling
try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# UPS# for item 5330 - special consolidation logic applies
ITEM_5330_UPS = '3030423'

# Built-in product reference data — shipped with the EXE as defaults.
# User modifications are stored in %APPDATA%/KAS/product_data.json and merged on startup.
DEFAULT_PRODUCT_DATA = {
    '3006683': {'part': 'S150102', 'pack_price': 8.40, 'pack_weight': 0.04},
    '3006708': {'part': 'S150104', 'pack_price': 8.52, 'pack_weight': 0.08},
    '3006716': {'part': 'S150105', 'pack_price': 8.76, 'pack_weight': 0.10},
    '3006724': {'part': 'S150106', 'pack_price': 10.20, 'pack_weight': 0.12},
    '3006732': {'part': 'S150107', 'pack_price': 11.76, 'pack_weight': 0.16},
    '3006740': {'part': 'S150108', 'pack_price': 12.72, 'pack_weight': 0.20},
    '3006758': {'part': 'S150109', 'pack_price': 13.80, 'pack_weight': 0.26},
    '3006766': {'part': 'S150110', 'pack_price': 15.72, 'pack_weight': 0.30},
    '3006782': {'part': 'S150112', 'pack_price': 18.84, 'pack_weight': 0.44},
    '3006790': {'part': 'S150113', 'pack_price': 19.56, 'pack_weight': 0.50},
    '3006807': {'part': 'S150114', 'pack_price': 24.00, 'pack_weight': 0.56},
    '3006815': {'part': 'S150115', 'pack_price': 26.04, 'pack_weight': 0.68},
    '3006823': {'part': 'S150116', 'pack_price': 30.48, 'pack_weight': 0.82},
    '3006873': {'part': 'S150121', 'pack_price': 22.86, 'pack_weight': 0.68},
    '3006881': {'part': 'S150122', 'pack_price': 24.30, 'pack_weight': 0.78},
    '3006899': {'part': 'S150123', 'pack_price': 27.12, 'pack_weight': 0.88},
    '3006956': {'part': 'S150129', 'pack_price': 39.00, 'pack_weight': 1.46},
    '3007201': {'part': 'SD50407', 'pack_price': 14.38, 'pack_weight': 0.34},
    '3007227': {'part': 'S160206', 'pack_price': 23.88, 'pack_weight': 0.54},
    '3011256': {'part': '2850Max', 'pack_price': 687.51, 'pack_weight': 20.00},
    '3030423': {'part': 'SAS5330', 'pack_price': 0.99, 'pack_weight': 0.01},
}


def get_appdata_dir():
    """Get the KAS config directory in AppData"""
    if sys.platform == 'win32':
        appdata = os.environ.get('APPDATA', os.path.expanduser('~'))
    else:
        appdata = os.path.expanduser('~')
    kas_dir = os.path.join(appdata, 'KAS')
    os.makedirs(kas_dir, exist_ok=True)
    return kas_dir


def get_product_data_path():
    """Get the path to the user's product data file in AppData"""
    return os.path.join(get_appdata_dir(), 'product_data.json')


def load_product_data():
    """Load product data with merge logic: user edits are preserved, new items are added."""
    config_path = get_product_data_path()

    if not os.path.exists(config_path):
        # First run — write defaults and return them
        save_product_data(DEFAULT_PRODUCT_DATA)
        return dict(DEFAULT_PRODUCT_DATA)

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            saved = json.load(f)
        # Handle both flat format and nested {"products": {...}} format
        if 'products' in saved and isinstance(saved['products'], dict):
            saved = saved['products']
    except (json.JSONDecodeError, IOError):
        # Corrupt file — reset to defaults
        save_product_data(DEFAULT_PRODUCT_DATA)
        return dict(DEFAULT_PRODUCT_DATA)

    # Merge: add any new items from defaults that aren't in user file
    merged = dict(saved)
    changed = False
    for ups_num, default_item in DEFAULT_PRODUCT_DATA.items():
        if ups_num not in merged:
            merged[ups_num] = default_item
            changed = True

    if changed:
        save_product_data(merged)

    return merged


def save_product_data(data):
    """Save product data to the AppData config file."""
    config_path = get_product_data_path()
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2)


# Load product data on startup (merges user edits with any new built-in items)
PRODUCT_DATA = load_product_data()


def get_resource_path(filename):
    """Get the path to a resource file, works for dev and PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # Running as compiled EXE - images are in images/ subfolder
        return os.path.join(sys._MEIPASS, 'images', filename)
    else:
        # Running as script - images are in images/ subfolder
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'images', filename)


class LabelSorter:
    def __init__(self, root):
        self.root = root
        self.root.title("KAS Enterprises - Label Sorter")
        self.root.geometry("800x750")
        self.root.configure(bg='#e8e4e0')
        
        # Store processed data
        self.sorted_content = None
        self.duplicate_report = None
        self.input_file_path = None
        self.records_data = None
        self.has_duplicates = False
        
        # Track if duplicate tab exists
        self.duplicate_tab_added = False
        
        # Store image references to prevent garbage collection
        self.logo_image = None
        self.icon_image = None
        
        self.setup_icon()
        self.setup_styles()
        self.setup_ui()
        self.setup_drag_drop()
    
    def setup_icon(self):
        """Setup the window icon from favicon.ico"""
        if not HAS_PIL:
            return
        
        try:
            favicon_path = get_resource_path('favicon.ico')
            if os.path.exists(favicon_path):
                icon = Image.open(favicon_path)
                # Convert to RGBA if needed and create PhotoImage
                if icon.mode != 'RGBA':
                    icon = icon.convert('RGBA')
                self.icon_image = ImageTk.PhotoImage(icon)
                self.root.iconphoto(True, self.icon_image)
        except Exception as e:
            print(f"Could not load favicon: {e}")
    
    def setup_styles(self):
        """Configure ttk styles for KAS Enterprises theme"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # KAS Colors: Orange #E85D04, Gray #6B6B6B
        
        # Notebook (tabs) styling
        style.configure('TNotebook', background='#e8e4e0', borderwidth=0)
        style.configure('TNotebook.Tab', 
                       background='#e0e0e0', 
                       foreground='#6B6B6B',
                       padding=[20, 8],
                       font=('Segoe UI', 10, 'bold'))
        style.map('TNotebook.Tab',
                 background=[('selected', '#E85D04')],
                 foreground=[('selected', 'white')],
                 padding=[('selected', [20, 8])])
        
        # Frame styling
        style.configure('Dark.TFrame', background='#e8e4e0')
        style.configure('Card.TFrame', background='#e8e4e0')
    
    def setup_ui(self):
        """Setup the main UI components"""
        # Header frame with white background for logo
        header_frame = tk.Frame(self.root, bg='#e8e4e0')
        header_frame.pack(pady=(20, 10), fill='x')
        
        # Load and display logo
        if HAS_PIL:
            try:
                logo_path = get_resource_path('Logo.png')
                if os.path.exists(logo_path):
                    logo = Image.open(logo_path)
                    # Resize logo to fit nicely (max height ~80px, maintain aspect ratio)
                    max_height = 80
                    ratio = max_height / logo.height
                    new_width = int(logo.width * ratio)
                    logo = logo.resize((new_width, max_height), Image.Resampling.LANCZOS)
                    self.logo_image = ImageTk.PhotoImage(logo)
                    
                    logo_label = tk.Label(header_frame, image=self.logo_image, bg='#e8e4e0', cursor='hand2')
                    logo_label.pack()
                    # Secret menu: triple-click on logo
                    logo_label.bind('<Triple-Button-1>', self.open_secret_menu)
                else:
                    # Fallback to text if no logo
                    self._create_text_header(header_frame)
            except Exception as e:
                print(f"Could not load logo: {e}")
                self._create_text_header(header_frame)
        else:
            self._create_text_header(header_frame)
        
        # Subtitle
        subtitle = tk.Label(
            header_frame,
            text="Shipping Label Sorter",
            font=('Segoe UI', 12),
            bg='#e8e4e0',
            fg='#6B6B6B'
        )
        subtitle.pack(pady=(10, 0))
        
        # Drop zone frame
        self.drop_frame = tk.Frame(
            self.root,
            bg='#ddd8d4',
            highlightbackground='#E85D04',
            highlightthickness=2,
            width=600,
            height=100
        )
        self.drop_frame.pack(pady=20, padx=50, fill='x')
        self.drop_frame.pack_propagate(False)
        
        # Drop zone label
        self.drop_label = tk.Label(
            self.drop_frame,
            text="📁 Drag & Drop Label File Here\nor click to browse",
            font=('Segoe UI', 12),
            bg='#ddd8d4',
            fg='#888888',
            cursor='hand2'
        )
        self.drop_label.pack(expand=True)
        self.drop_label.bind('<Button-1>', self.browse_file)
        self.drop_frame.bind('<Button-1>', self.browse_file)
        
        # Status label
        self.status_label = tk.Label(
            self.root,
            text="",
            font=('Segoe UI', 10),
            bg='#e8e4e0',
            fg='#E85D04'
        )
        self.status_label.pack(pady=5)
        
        # Button frame for Save and Print
        btn_frame = tk.Frame(self.root, bg='#e8e4e0')
        btn_frame.pack(pady=(5, 15))
        
        # Save button
        self.save_btn = tk.Button(
            btn_frame,
            text="💾 Save",
            font=('Segoe UI', 11, 'bold'),
            bg='#6B6B6B',
            fg='white',
            activebackground='#555555',
            activeforeground='white',
            padx=20,
            pady=8,
            cursor='hand2',
            state='disabled',
            command=self.save_sorted_file
        )
        self.save_btn.pack(side='left', padx=5)
        
        # Print button
        self.print_btn = tk.Button(
            btn_frame,
            text="🖨️ Print",
            font=('Segoe UI', 11, 'bold'),
            bg='#E85D04',
            fg='white',
            activebackground='#cc5203',
            activeforeground='white',
            padx=20,
            pady=8,
            cursor='hand2',
            state='disabled',
            command=self.print_sorted_file
        )
        self.print_btn.pack(side='left', padx=5)
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(pady=(0, 20), padx=30, fill='both', expand=True)
        
        # Tab 1: Order Summary - Always visible
        self.preview_frame = tk.Frame(self.notebook, bg='#1a1a1a')
        self.notebook.add(self.preview_frame, text='  📊 Order Summary  ')
        
        # Placeholder for preview
        self.preview_placeholder = tk.Label(
            self.preview_frame,
            text="Upload a label file to see the order summary",
            font=('Segoe UI', 12),
            bg='#1a1a1a',
            fg='#6B6B6B'
        )
        self.preview_placeholder.pack(expand=True)
        
        self.preview_text = None  # Will be created when needed
        
        # Tab 2: Duplicate Report - Created only when needed
        self.report_frame = None
        self.report_text = None
    
    def _create_text_header(self, parent):
        """Fallback text header if logo can't be loaded"""
        title_label = tk.Label(
            parent, 
            text="KAS Enterprises",
            font=('Segoe UI', 24, 'bold'),
            bg='#e8e4e0',
            fg='#E85D04',
            cursor='hand2'
        )
        title_label.pack()
        # Secret menu: triple-click on header
        title_label.bind('<Triple-Button-1>', self.open_secret_menu)
    
    def create_preview_text(self):
        """Create the Treeview grid widget (replaces placeholder)"""
        if self.preview_text:
            return
        
        self.preview_placeholder.destroy()
        
        # Configure Treeview styles for dark theme with grid lines
        style = ttk.Style()
        style.configure('Summary.Treeview',
                        background='#1a1a1a',
                        foreground='#cccccc',
                        fieldbackground='#1a1a1a',
                        rowheight=28,
                        font=('Segoe UI', 10))
        style.configure('Summary.Treeview.Heading',
                        background='#2d2d2d',
                        foreground='#ffd700',
                        font=('Segoe UI', 10, 'bold'),
                        borderwidth=1,
                        relief='solid')
        style.map('Summary.Treeview',
                  background=[('selected', '#E85D04')],
                  foreground=[('selected', 'white')])
        style.layout('Summary.Treeview', [
            ('Summary.Treeview.treearea', {'sticky': 'nswe'})
        ])
        
        # Create container frame
        tree_container = tk.Frame(self.preview_frame, bg='#1a1a1a')
        tree_container.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create Treeview with columns
        columns = ('order', 'range', 'weight', 'cost')
        self.preview_tree = ttk.Treeview(
            tree_container,
            columns=columns,
            show='headings',
            style='Summary.Treeview'
        )
        
        # Define headings
        self.preview_tree.heading('order', text='Order #')
        self.preview_tree.heading('range', text='Range')
        self.preview_tree.heading('weight', text='Total Weight')
        self.preview_tree.heading('cost', text='Total Cost')
        
        # Define column widths and alignment
        self.preview_tree.column('order', width=120, anchor='w')
        self.preview_tree.column('range', width=120, anchor='center')
        self.preview_tree.column('weight', width=140, anchor='e')
        self.preview_tree.column('cost', width=140, anchor='e')
        
        # Add scrollbar
        tree_scroll = ttk.Scrollbar(tree_container, orient='vertical', command=self.preview_tree.yview)
        self.preview_tree.configure(yscrollcommand=tree_scroll.set)
        
        tree_scroll.pack(side='right', fill='y')
        self.preview_tree.pack(side='left', fill='both', expand=True)
        
        # Configure row tags for styling
        self.preview_tree.tag_configure('normal', foreground='#cccccc')
        self.preview_tree.tag_configure('cost_high', foreground='#E85D04')
        self.preview_tree.tag_configure('glasses', foreground='#87CEEB')
        self.preview_tree.tag_configure('total_row', foreground='#ffd700', font=('Segoe UI', 10, 'bold'))
        self.preview_tree.tag_configure('total_high', foreground='#E85D04', font=('Segoe UI', 10, 'bold'))
        
        # Totals label below the tree
        self.totals_frame = tk.Frame(self.preview_frame, bg='#1a1a1a')
        self.totals_frame.pack(fill='x', padx=5, pady=(0, 5))
        
        self.totals_label = tk.Label(
            self.totals_frame,
            text="",
            font=('Segoe UI', 10, 'bold'),
            bg='#1a1a1a',
            fg='#E85D04',
            anchor='w'
        )
        self.totals_label.pack(fill='x', padx=10)
        
        # Mark as created (reuse self.preview_text as a flag for compatibility)
        self.preview_text = True
    
    def create_duplicate_tab(self):
        """Create the duplicate report tab (only when duplicates exist)"""
        if self.duplicate_tab_added:
            return
        
        self.report_frame = tk.Frame(self.notebook, bg='#1a1a1a')
        self.notebook.add(self.report_frame, text='  ⚠️ Duplicate Orders  ')
        
        self.report_text = tk.Text(
            self.report_frame,
            font=('Consolas', 10),
            bg='#1a1a1a',
            fg='#cccccc',
            insertbackground='white',
            wrap='word',
            state='disabled'
        )
        report_scrollbar = ttk.Scrollbar(self.report_frame, command=self.report_text.yview)
        self.report_text.configure(yscrollcommand=report_scrollbar.set)
        
        report_scrollbar.pack(side='right', fill='y')
        self.report_text.pack(side='left', fill='both', expand=True)
        
        # Configure text tags for colored output
        self.report_text.tag_configure('header', foreground='#E85D04', font=('Consolas', 12, 'bold'))
        self.report_text.tag_configure('subheader', foreground='#ffd700', font=('Consolas', 11, 'bold'))
        self.report_text.tag_configure('warning', foreground='#ff6b6b', font=('Consolas', 10, 'bold'))
        self.report_text.tag_configure('info', foreground='#E85D04')
        self.report_text.tag_configure('drill', foreground='#ff6b6b')
        self.report_text.tag_configure('count', foreground='#ffd700', font=('Consolas', 10, 'bold'))
        self.report_text.tag_configure('summary_label', foreground='#a0a0a0')
        self.report_text.tag_configure('summary_value', foreground='#E85D04', font=('Consolas', 11, 'bold'))
        
        self.duplicate_tab_added = True
    
    def remove_duplicate_tab(self):
        """Remove the duplicate report tab if it exists"""
        if self.duplicate_tab_added and self.report_frame:
            self.notebook.forget(self.report_frame)
            self.report_frame.destroy()
            self.report_frame = None
            self.report_text = None
            self.duplicate_tab_added = False
    
    def setup_drag_drop(self):
        """Setup drag and drop functionality"""
        try:
            from tkinterdnd2 import DND_FILES
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind('<<Drop>>', self.on_drop)
        except ImportError:
            pass
    
    def on_drop(self, event):
        """Handle file drop event"""
        file_path = event.data
        file_path = file_path.strip('{}')
        self.process_file(file_path)
    
    def browse_file(self, event=None):
        """Open file browser dialog"""
        file_path = filedialog.askopenfilename(
            title="Select Label File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)
    
    def parse_records(self, content):
        """Parse the file content into individual records"""
        records = []
        lines = content.split('\n')
        current_record = []
        
        for line in lines:
            if line.strip() == '{C|}':
                if current_record:
                    records.append(current_record)
                current_record = [line]
            elif current_record:
                current_record.append(line)
        
        if current_record:
            records.append(current_record)
        
        return records
    
    def extract_order_number(self, record):
        """Extract order number from a record"""
        record_text = ''.join(record)
        match = re.search(r'\{RC25;ORD#\s*(\d+-\d+)\|?\}', record_text)
        if match:
            return match.group(1)
        return "000000-000"
    
    def extract_product_description(self, record):
        """Extract product description from a record"""
        record_text = ''.join(record)
        match = re.search(r'\{RC04;([^|]+)\|?\}', record_text)
        if match:
            return match.group(1).strip()
        return ""
    
    def extract_address(self, record):
        """Extract address (RC01 field) from a record"""
        record_text = ''.join(record)
        match = re.search(r'\{RC01;([^|]+)\|?\}', record_text)
        if match:
            return match.group(1).strip()
        return ""
    
    def extract_ups_number(self, record):
        """Extract UPS# (RC02 field) from a record"""
        record_text = ''.join(record)
        match = re.search(r'\{RC02;([^|]+)\|?\}', record_text)
        if match:
            return match.group(1).strip()
        return ""
    
    def record_to_single_line(self, record):
        """Convert a multi-line record to single line format"""
        combined = ''
        for line in record:
            stripped = line.rstrip()
            if stripped:
                combined += stripped
        return combined
    
    def update_case_qty(self, record, new_qty):
        """Update the Case QTY value in a record (for 5330 consolidation)"""
        # RC19 contains "PINK TAG ITEM *** CASE QTY:     X"
        # We need to update the quantity value
        updated_record = []
        for line in record:
            # Look for the CASE QTY pattern and update it
            if 'CASE QTY:' in line:
                # Replace the number after CASE QTY: with the new quantity
                # Pattern: "CASE QTY:     1" -> "CASE QTY:    10"
                updated_line = re.sub(
                    r'(CASE QTY:\s*)(\d+)',
                    lambda m: f"{m.group(1)}{new_qty:>5}",
                    line
                )
                updated_record.append(updated_line)
            else:
                updated_record.append(line)
        return updated_record
    
    def process_file(self, file_path):
        """Process the input file"""
        self.input_file_path = file_path
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{str(e)}")
            return
        
        # Parse records
        records = self.parse_records(content)
        
        if not records:
            messagebox.showwarning("Warning", "No records found in the file.")
            return
        
        # Sort by order number and extract all fields
        records_with_order = []
        for record in records:
            order_num = self.extract_order_number(record)
            product = self.extract_product_description(record)
            address = self.extract_address(record)
            ups_number = self.extract_ups_number(record)
            records_with_order.append((order_num, product, address, ups_number, record))
        
        records_with_order.sort(key=lambda x: x[0])
        self.records_data = records_with_order
        
        # Separate 5330 items for special consolidation logic
        item_5330_records = defaultdict(list)  # Group by full order number
        other_records = []
        
        for order_num, product, address, ups_number, record in records_with_order:
            if ups_number == ITEM_5330_UPS:
                # 5330 item - group by full order number (including suffix)
                item_5330_records[order_num].append((order_num, product, address, ups_number, record))
            else:
                other_records.append((order_num, product, address, ups_number, record))
        
        # Generate single-line output (stored in memory for printing)
        sorted_lines = []
        
        # Process 5330 items: consolidate duplicates and update Case QTY
        for order_num in sorted(item_5330_records.keys()):
            records_list = item_5330_records[order_num]
            qty = len(records_list)
            # Use the first record as the template, update Case QTY
            first_record = records_list[0][4]  # The record is the 5th element
            updated_record = self.update_case_qty(first_record, qty)
            single_line = self.record_to_single_line(updated_record)
            sorted_lines.append(single_line)
        
        # Process other items normally
        for order_num, product, address, ups_number, record in other_records:
            single_line = self.record_to_single_line(record)
            sorted_lines.append(single_line)
        
        # Re-sort all lines by order number (they're embedded in the content)
        # Extract order number from each line for final sort
        def extract_order_from_line(line):
            match = re.search(r'\{RC25;ORD#\s*(\d+-\d+)\|?\}', line)
            return match.group(1) if match else '000000-000'
        
        sorted_lines.sort(key=extract_order_from_line)
        self.sorted_content = '\n'.join(sorted_lines)
        
        # Calculate order summaries (weight and cost per order group)
        # For 5330 items: group by full order number (including suffix)
        # For other items: group by first 6 digits of order number (before the "-")
        order_summaries = defaultdict(lambda: {'weight': 0.0, 'cost': 0.0, 'sub_orders': set(), 'ups_numbers': set()})
        
        # Process 5330 items separately - group by full order number
        for order_num in item_5330_records.keys():
            qty = len(item_5330_records[order_num])
            
            # Split order number into base and sub-order
            if '-' in order_num:
                order_base, sub_order = order_num.split('-', 1)
            else:
                order_base = order_num
                sub_order = '000'
            
            order_summaries[order_base]['sub_orders'].add(sub_order)
            order_summaries[order_base]['ups_numbers'].add(ITEM_5330_UPS)
            
            # Calculate cost and weight for consolidated 5330 items
            if ITEM_5330_UPS in PRODUCT_DATA:
                product_data = PRODUCT_DATA[ITEM_5330_UPS]
                order_summaries[order_base]['weight'] += product_data['pack_weight'] * qty
                order_summaries[order_base]['cost'] += product_data['pack_price'] * qty
        
        # Process other items normally
        for order_num, product, address, ups_number, record in other_records:
            # Split order number into base (first 6 digits) and sub-order (last 3 digits)
            if '-' in order_num:
                order_base, sub_order = order_num.split('-', 1)
            else:
                order_base = order_num
                sub_order = '000'
            
            # Track sub-order numbers for range display
            order_summaries[order_base]['sub_orders'].add(sub_order)
            order_summaries[order_base]['ups_numbers'].add(ups_number)
            
            # Look up data by UPS# - only include known items in totals
            if ups_number in PRODUCT_DATA:
                product_info = PRODUCT_DATA[ups_number]
                order_summaries[order_base]['weight'] += product_info['pack_weight']
                order_summaries[order_base]['cost'] += product_info['pack_price']
        
        # Convert sub_orders set to order range string
        for order_base in order_summaries:
            sub_orders = sorted(order_summaries[order_base]['sub_orders'])
            if len(sub_orders) == 1:
                order_summaries[order_base]['order_range'] = sub_orders[0]
            else:
                order_summaries[order_base]['order_range'] = f"{sub_orders[0]}-{sub_orders[-1]}"
        
        self.order_summaries = dict(order_summaries)
        
        # Check for duplicate orders
        order_counts = defaultdict(list)
        for order_num, product, address, ups_number, record in records_with_order:
            order_counts[order_num].append({'product': product, 'address': address})
        
        duplicates = {k: v for k, v in order_counts.items() if len(v) > 1}
        
        # Check for duplicate addresses
        address_counts = defaultdict(list)
        for order_num, product, address, ups_number, record in records_with_order:
            if address:  # Only track non-empty addresses
                address_counts[address].append({'order_num': order_num, 'product': product})
        
        duplicate_addresses = {k: v for k, v in address_counts.items() if len(v) > 1}
        
        self.has_duplicates = len(duplicates) > 0 or len(duplicate_addresses) > 0
        
        # Update UI
        filename = os.path.basename(file_path)
        status_text = f"✓ Loaded: {filename} ({len(records)} records sorted by order number)"
        if len(duplicates) > 0 or len(duplicate_addresses) > 0:
            warnings = []
            if len(duplicates) > 0:
                warnings.append(f"{len(duplicates)} duplicate orders")
            if len(duplicate_addresses) > 0:
                warnings.append(f"{len(duplicate_addresses)} duplicate addresses")
            status_text += f" — ⚠️ {', '.join(warnings)} found!"
        
        self.status_label.config(text=status_text, fg='#E85D04')
        self.drop_label.config(
            text=f"📄 {filename}\n{len(records)} records loaded & sorted",
            fg='#E85D04'
        )
        
        # Enable buttons
        self.save_btn.config(state='normal')
        self.print_btn.config(state='normal')
        
        # Update preview tab
        self.create_preview_text()
        self.update_preview()
        
        # Handle duplicate tab
        if self.has_duplicates:
            self.create_duplicate_tab()
            self.update_report(order_counts, duplicates, address_counts, duplicate_addresses)
        else:
            self.remove_duplicate_tab()
        
        # Always default to Sorted Preview tab
        self.notebook.select(0)
    
    def update_preview(self):
        """Update the preview tab with order summary Treeview grid"""
        if not self.preview_text or not hasattr(self, 'preview_tree'):
            return
        
        # Clear existing rows
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
        
        if hasattr(self, 'order_summaries') and self.order_summaries:
            grand_weight = 0.0
            grand_cost = 0.0
            
            for order_num in sorted(self.order_summaries.keys()):
                summary = self.order_summaries[order_num]
                weight = summary['weight']
                cost = summary['cost']
                order_range = summary.get('order_range', '')
                ups_numbers = summary.get('ups_numbers', set())
                grand_weight += weight
                grand_cost += cost
                
                # Determine display prefix — sunglasses emoji for glasses orders
                has_glasses = ITEM_5330_UPS in ups_numbers
                order_display = f"\U0001f576\ufe0f {order_num}" if has_glasses else order_num
                
                weight_str = f"{weight:.2f} lb"
                cost_str = f"${cost:.2f}"
                
                # Choose tag based on cost threshold
                if has_glasses:
                    tag = 'glasses'
                elif cost >= 100:
                    tag = 'cost_high'
                else:
                    tag = 'normal'
                
                self.preview_tree.insert('', 'end',
                    values=(order_display, order_range, weight_str, cost_str),
                    tags=(tag,))
            
            # Add total row
            total_tag = 'total_high' if grand_cost >= 100 else 'total_row'
            self.preview_tree.insert('', 'end',
                values=('TOTAL', '', f"{grand_weight:.2f} lb", f"${grand_cost:.2f}"),
                tags=(total_tag,))
            
            # Update totals label
            self.totals_label.config(
                text=f"  Total Unique Orders: {len(self.order_summaries)}"
            )
    
    def update_report(self, order_counts, duplicates, address_counts, duplicate_addresses):
        """Update the duplicate report tab"""
        if not self.report_text:
            return
        
        self.report_text.config(state='normal')
        self.report_text.delete('1.0', tk.END)
        
        # Header
        self.report_text.insert(tk.END, "═" * 55 + "\n", 'header')
        self.report_text.insert(tk.END, "  DUPLICATE ORDERS REPORT\n", 'header')
        self.report_text.insert(tk.END, f"  Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", 'summary_label')
        self.report_text.insert(tk.END, "═" * 55 + "\n\n", 'header')
        
        # Duplicate Orders Section
        if duplicates:
            self.report_text.insert(tk.END, f"Found ", 'summary_label')
            self.report_text.insert(tk.END, f"{len(duplicates)}", 'count')
            self.report_text.insert(tk.END, " orders with multiple labels:\n\n", 'summary_label')
            
            drill_bit_duplicates = []
            other_duplicates = []
            
            for order_num, items in sorted(duplicates.items()):
                products = [item['product'] for item in items]
                has_drill_bit = any('DRILL' in p.upper() for p in products)
                if has_drill_bit:
                    drill_bit_duplicates.append((order_num, items))
                else:
                    other_duplicates.append((order_num, items))
            
            # Drill bit duplicates (highlighted)
            if drill_bit_duplicates:
                self.report_text.insert(tk.END, "⚠️  DRILL BIT ORDERS WITH DUPLICATES\n", 'warning')
                self.report_text.insert(tk.END, "─" * 45 + "\n", 'summary_label')
                
                for order_num, items in drill_bit_duplicates:
                    self.report_text.insert(tk.END, f"  Order:   ", 'summary_label')
                    self.report_text.insert(tk.END, f"{order_num}\n", 'info')
                    self.report_text.insert(tk.END, f"  Labels:  ", 'summary_label')
                    self.report_text.insert(tk.END, f"{len(items)}\n", 'count')
                    self.report_text.insert(tk.END, f"  Product: ", 'summary_label')
                    self.report_text.insert(tk.END, f"{items[0]['product']}\n", 'drill')
                    self.report_text.insert(tk.END, "\n")
            
            # Other duplicates
            if other_duplicates:
                self.report_text.insert(tk.END, "\n📋  OTHER DUPLICATE ORDERS\n", 'subheader')
                self.report_text.insert(tk.END, "─" * 45 + "\n", 'summary_label')
                
                for order_num, items in other_duplicates:
                    self.report_text.insert(tk.END, f"  Order:   ", 'summary_label')
                    self.report_text.insert(tk.END, f"{order_num}\n", 'info')
                    self.report_text.insert(tk.END, f"  Labels:  ", 'summary_label')
                    self.report_text.insert(tk.END, f"{len(items)}\n", 'count')
                    self.report_text.insert(tk.END, f"  Product: {items[0]['product']}\n")
                    self.report_text.insert(tk.END, "\n")
        
        # Duplicate Addresses Section
        if duplicate_addresses:
            self.report_text.insert(tk.END, "\n" + "═" * 55 + "\n", 'header')
            self.report_text.insert(tk.END, "  DUPLICATE ADDRESSES\n", 'header')
            self.report_text.insert(tk.END, "═" * 55 + "\n\n", 'header')
            
            self.report_text.insert(tk.END, f"Found ", 'summary_label')
            self.report_text.insert(tk.END, f"{len(duplicate_addresses)}", 'count')
            self.report_text.insert(tk.END, " addresses with multiple orders:\n\n", 'summary_label')
            
            for address, orders in sorted(duplicate_addresses.items()):
                has_drill_bit = any('DRILL' in o['product'].upper() for o in orders)
                
                if has_drill_bit:
                    self.report_text.insert(tk.END, "⚠️ ", 'warning')
                else:
                    self.report_text.insert(tk.END, "📍 ", 'info')
                
                self.report_text.insert(tk.END, f"Address: ", 'summary_label')
                self.report_text.insert(tk.END, f"{address}\n", 'info')
                self.report_text.insert(tk.END, f"   Orders:  ", 'summary_label')
                self.report_text.insert(tk.END, f"{len(orders)}\n", 'count')
                
                # List each order going to this address
                for order in orders:
                    self.report_text.insert(tk.END, f"   → ", 'summary_label')
                    self.report_text.insert(tk.END, f"{order['order_num']}", 'info')
                    if 'DRILL' in order['product'].upper():
                        self.report_text.insert(tk.END, f" ({order['product'][:25]}...)\n" if len(order['product']) > 25 else f" ({order['product']})\n", 'drill')
                    else:
                        self.report_text.insert(tk.END, f" ({order['product'][:25]}...)\n" if len(order['product']) > 25 else f" ({order['product']})\n", 'summary_label')
                
                self.report_text.insert(tk.END, "\n")
        
        # Summary section
        self.report_text.insert(tk.END, "\n" + "═" * 55 + "\n", 'header')
        self.report_text.insert(tk.END, "  SUMMARY\n", 'header')
        self.report_text.insert(tk.END, "═" * 55 + "\n\n", 'header')
        
        total_records = len(self.records_data)
        unique_orders = len(order_counts)
        duplicate_count = len(duplicates)
        duplicate_addr_count = len(duplicate_addresses)
        drill_count = sum(1 for items in order_counts.values() 
                        for item in items if 'DRILL' in item['product'].upper())
        
        self.report_text.insert(tk.END, "  Total Records:       ", 'summary_label')
        self.report_text.insert(tk.END, f"{total_records}\n", 'summary_value')
        
        self.report_text.insert(tk.END, "  Unique Orders:       ", 'summary_label')
        self.report_text.insert(tk.END, f"{unique_orders}\n", 'summary_value')
        
        self.report_text.insert(tk.END, "  Duplicate Orders:    ", 'summary_label')
        if duplicate_count > 0:
            self.report_text.insert(tk.END, f"{duplicate_count}\n", 'warning')
        else:
            self.report_text.insert(tk.END, f"{duplicate_count}\n", 'summary_value')
        
        self.report_text.insert(tk.END, "  Duplicate Addresses: ", 'summary_label')
        if duplicate_addr_count > 0:
            self.report_text.insert(tk.END, f"{duplicate_addr_count}\n", 'warning')
        else:
            self.report_text.insert(tk.END, f"{duplicate_addr_count}\n", 'summary_value')
        
        self.report_text.insert(tk.END, "  Drill Bit Labels:    ", 'summary_label')
        self.report_text.insert(tk.END, f"{drill_count}", 'summary_value')
        if total_records > 0:
            pct = (drill_count / total_records) * 100
            self.report_text.insert(tk.END, f" ({pct:.0f}%)\n", 'summary_label')
        else:
            self.report_text.insert(tk.END, "\n")
        
        # Full order breakdown table
        self.report_text.insert(tk.END, "\n" + "═" * 55 + "\n", 'header')
        self.report_text.insert(tk.END, "  ALL ORDERS BREAKDOWN\n", 'header')
        self.report_text.insert(tk.END, "═" * 55 + "\n\n", 'header')
        
        self.report_text.insert(tk.END, "  Order Number   │ Count │ Product\n", 'subheader')
        self.report_text.insert(tk.END, "  ───────────────┼───────┼" + "─" * 25 + "\n", 'summary_label')
        
        for order_num in sorted(order_counts.keys()):
            items = order_counts[order_num]
            count = len(items)
            product = items[0]['product']
            product_display = product[:22] + "..." if len(product) > 25 else product
            
            is_duplicate = count > 1
            is_drill = 'DRILL' in product.upper()
            
            self.report_text.insert(tk.END, f"  {order_num:<14} │ ", 'summary_label')
            
            if is_duplicate:
                self.report_text.insert(tk.END, f"{count:^5}", 'warning')
                self.report_text.insert(tk.END, " │ ", 'summary_label')
            else:
                self.report_text.insert(tk.END, f"{count:^5}", 'info')
                self.report_text.insert(tk.END, " │ ", 'summary_label')
            
            if is_drill and is_duplicate:
                self.report_text.insert(tk.END, f"{product_display}\n", 'drill')
            elif is_drill:
                self.report_text.insert(tk.END, f"{product_display}\n")
            else:
                self.report_text.insert(tk.END, f"{product_display}\n")
        
        self.report_text.config(state='disabled')
    
    def open_secret_menu(self, event=None):
        """Open the secret reference data editor (triggered by triple-click on logo)"""
        global PRODUCT_DATA
        
        # Create the secret menu window
        editor = tk.Toplevel(self.root)
        editor.title("🔧 Reference Data Editor")
        editor.geometry("700x500")
        editor.configure(bg='#1a1a1a')
        editor.transient(self.root)
        
        # Center the window
        editor.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 700) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 500) // 2
        editor.geometry(f"+{x}+{y}")
        
        # Header
        header = tk.Label(
            editor,
            text="🔐 Secret Menu - Reference Data Editor",
            font=('Segoe UI', 14, 'bold'),
            bg='#1a1a1a',
            fg='#E85D04'
        )
        header.pack(pady=(15, 5))
        
        subtitle = tk.Label(
            editor,
            text="Triple-click unlocked! Edit product prices and weights below.",
            font=('Segoe UI', 10),
            bg='#1a1a1a',
            fg='#888888'
        )
        subtitle.pack(pady=(0, 10))
        
        # Create frame for the data list
        list_frame = tk.Frame(editor, bg='#1a1a1a')
        list_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Column headers
        header_frame = tk.Frame(list_frame, bg='#2a2a2a')
        header_frame.pack(fill='x')
        
        tk.Label(header_frame, text="UPS#", font=('Consolas', 10, 'bold'), bg='#2a2a2a', fg='#E85D04', width=12).pack(side='left', padx=5)
        tk.Label(header_frame, text="Part", font=('Consolas', 10, 'bold'), bg='#2a2a2a', fg='#E85D04', width=12).pack(side='left', padx=5)
        tk.Label(header_frame, text="Price", font=('Consolas', 10, 'bold'), bg='#2a2a2a', fg='#E85D04', width=10).pack(side='left', padx=5)
        tk.Label(header_frame, text="Weight", font=('Consolas', 10, 'bold'), bg='#2a2a2a', fg='#E85D04', width=10).pack(side='left', padx=5)
        tk.Label(header_frame, text="Actions", font=('Consolas', 10, 'bold'), bg='#2a2a2a', fg='#E85D04', width=15).pack(side='left', padx=5)
        
        # Scrollable list of items
        canvas = tk.Canvas(list_frame, bg='#1a1a1a', highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='#1a1a1a')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)
        
        # Store entry widgets for editing
        entry_widgets = {}
        
        def refresh_list():
            """Refresh the list of products"""
            for widget in scrollable_frame.winfo_children():
                widget.destroy()
            entry_widgets.clear()
            
            for ups_num in sorted(PRODUCT_DATA.keys()):
                data = PRODUCT_DATA[ups_num]
                row = tk.Frame(scrollable_frame, bg='#1a1a1a')
                row.pack(fill='x', pady=2)
                
                # UPS# (read-only display)
                tk.Label(row, text=ups_num, font=('Consolas', 9), bg='#1a1a1a', fg='#cccccc', width=12).pack(side='left', padx=5)
                
                # Part
                part_entry = tk.Entry(row, font=('Consolas', 9), bg='#2a2a2a', fg='#90EE90', width=12, insertbackground='white')
                part_entry.insert(0, data.get('part', ''))
                part_entry.pack(side='left', padx=5)
                
                # Price
                price_entry = tk.Entry(row, font=('Consolas', 9), bg='#2a2a2a', fg='#87CEEB', width=10, insertbackground='white')
                price_entry.insert(0, f"{data.get('pack_price', 0):.2f}")
                price_entry.pack(side='left', padx=5)
                
                # Weight
                weight_entry = tk.Entry(row, font=('Consolas', 9), bg='#2a2a2a', fg='#ffd700', width=10, insertbackground='white')
                weight_entry.insert(0, f"{data.get('pack_weight', 0):.2f}")
                weight_entry.pack(side='left', padx=5)
                
                entry_widgets[ups_num] = {
                    'part': part_entry,
                    'price': price_entry,
                    'weight': weight_entry
                }
                
                # Delete button
                del_btn = tk.Button(
                    row, text="🗑️ Delete", font=('Segoe UI', 8),
                    bg='#8B0000', fg='white', padx=5,
                    command=lambda u=ups_num: delete_item(u)
                )
                del_btn.pack(side='left', padx=5)
        
        def delete_item(ups_num):
            """Delete an item from reference data"""
            if messagebox.askyesno("Confirm Delete", f"Delete UPS# {ups_num}?"):
                del PRODUCT_DATA[ups_num]
                refresh_list()
        
        def add_new_item():
            """Open dialog to add a new item"""
            add_dialog = tk.Toplevel(editor)
            add_dialog.title("Add New Product")
            add_dialog.geometry("300x200")
            add_dialog.configure(bg='#1a1a1a')
            add_dialog.transient(editor)
            add_dialog.grab_set()
            
            # Center
            add_dialog.update_idletasks()
            x = editor.winfo_x() + (editor.winfo_width() - 300) // 2
            y = editor.winfo_y() + (editor.winfo_height() - 200) // 2
            add_dialog.geometry(f"+{x}+{y}")
            
            tk.Label(add_dialog, text="Add New Product", font=('Segoe UI', 12, 'bold'), bg='#1a1a1a', fg='#E85D04').pack(pady=10)
            
            fields_frame = tk.Frame(add_dialog, bg='#1a1a1a')
            fields_frame.pack(pady=10)
            
            tk.Label(fields_frame, text="UPS#:", bg='#1a1a1a', fg='#cccccc').grid(row=0, column=0, sticky='e', padx=5, pady=3)
            ups_entry = tk.Entry(fields_frame, bg='#2a2a2a', fg='white', insertbackground='white')
            ups_entry.grid(row=0, column=1, padx=5, pady=3)
            
            tk.Label(fields_frame, text="Part:", bg='#1a1a1a', fg='#cccccc').grid(row=1, column=0, sticky='e', padx=5, pady=3)
            part_entry = tk.Entry(fields_frame, bg='#2a2a2a', fg='white', insertbackground='white')
            part_entry.grid(row=1, column=1, padx=5, pady=3)
            
            tk.Label(fields_frame, text="Price:", bg='#1a1a1a', fg='#cccccc').grid(row=2, column=0, sticky='e', padx=5, pady=3)
            price_entry = tk.Entry(fields_frame, bg='#2a2a2a', fg='white', insertbackground='white')
            price_entry.grid(row=2, column=1, padx=5, pady=3)
            
            tk.Label(fields_frame, text="Weight:", bg='#1a1a1a', fg='#cccccc').grid(row=3, column=0, sticky='e', padx=5, pady=3)
            weight_entry = tk.Entry(fields_frame, bg='#2a2a2a', fg='white', insertbackground='white')
            weight_entry.grid(row=3, column=1, padx=5, pady=3)
            
            def do_add():
                ups = ups_entry.get().strip()
                part = part_entry.get().strip()
                try:
                    price = float(price_entry.get().strip())
                    weight = float(weight_entry.get().strip())
                except ValueError:
                    messagebox.showerror("Error", "Price and Weight must be numbers")
                    return
                
                if not ups:
                    messagebox.showerror("Error", "UPS# is required")
                    return
                
                if ups in PRODUCT_DATA:
                    messagebox.showerror("Error", f"UPS# {ups} already exists")
                    return
                
                PRODUCT_DATA[ups] = {
                    'part': part,
                    'pack_price': price,
                    'pack_weight': weight
                }
                add_dialog.destroy()
                refresh_list()
            
            btn_frame = tk.Frame(add_dialog, bg='#1a1a1a')
            btn_frame.pack(pady=10)
            tk.Button(btn_frame, text="Add", command=do_add, bg='#E85D04', fg='white', font=('Segoe UI', 10, 'bold'), padx=15).pack(side='left', padx=5)
            tk.Button(btn_frame, text="Cancel", command=add_dialog.destroy, bg='#6B6B6B', fg='white', font=('Segoe UI', 10), padx=15).pack(side='left', padx=5)
        
        def save_all_changes():
            """Save all changes to the AppData config file"""
            # Update PRODUCT_DATA from entry widgets
            for ups_num, widgets in entry_widgets.items():
                try:
                    PRODUCT_DATA[ups_num] = {
                        'part': widgets['part'].get().strip(),
                        'pack_price': float(widgets['price'].get().strip()),
                        'pack_weight': float(widgets['weight'].get().strip())
                    }
                except ValueError:
                    messagebox.showerror("Error", f"Invalid number format for UPS# {ups_num}")
                    return
            
            # Save to AppData
            try:
                save_product_data(PRODUCT_DATA)
                json_path = get_product_data_path()
                messagebox.showinfo("Saved", f"Reference data saved successfully!\n\n{json_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save: {e}")
        
        # Initial list population
        refresh_list()
        
        # Button frame
        btn_frame = tk.Frame(editor, bg='#1a1a1a')
        btn_frame.pack(pady=15)
        
        tk.Button(
            btn_frame, text="➕ Add New", command=add_new_item,
            bg='#228B22', fg='white', font=('Segoe UI', 10, 'bold'), padx=15, pady=5
        ).pack(side='left', padx=10)
        
        tk.Button(
            btn_frame, text="💾 Save All Changes", command=save_all_changes,
            bg='#E85D04', fg='white', font=('Segoe UI', 10, 'bold'), padx=15, pady=5
        ).pack(side='left', padx=10)
        
        tk.Button(
            btn_frame, text="Close", command=editor.destroy,
            bg='#6B6B6B', fg='white', font=('Segoe UI', 10), padx=15, pady=5
        ).pack(side='left', padx=10)
        
        # Info label at bottom
        info_label = tk.Label(
            editor,
            text=f"\U0001f4c2 Data file: {get_product_data_path()}",
            font=('Segoe UI', 8),
            bg='#1a1a1a',
            fg='#555555'
        )
        info_label.pack(pady=(0, 10))
    
    def save_sorted_file(self):
        """Save the sorted content to a file"""
        if not self.sorted_content:
            messagebox.showwarning("Warning", "No processed content to save.")
            return
        
        if self.input_file_path:
            base_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
            default_name = f"{base_name}_sorted.txt"
        else:
            default_name = "sorted_labels.txt"
        
        file_path = filedialog.asksaveasfilename(
            title="Save Sorted File",
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.sorted_content)
                messagebox.showinfo("Success", f"File saved successfully!\n\n{file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file:\n{str(e)}")
    
    def print_sorted_file(self):
        """Print the sorted content - opens file in Notepad for printing with dialog"""
        if not self.sorted_content:
            messagebox.showwarning("Warning", "No processed content to print.")
            return
        
        try:
            # Create a temporary file with the sorted content
            if self.input_file_path:
                base_name = os.path.splitext(os.path.basename(self.input_file_path))[0]
                temp_name = f"{base_name}_sorted.txt"
            else:
                temp_name = "sorted_labels.txt"
            
            # Create temp file in system temp directory
            temp_path = os.path.join(tempfile.gettempdir(), temp_name)
            
            with open(temp_path, 'w', encoding='utf-8') as f:
                f.write(self.sorted_content)
            
            # Try to use win32api for proper print dialog (if available)
            try:
                import win32api
                import win32print
                
                # Show printer selection dialog
                self.show_printer_dialog(temp_path)
                
            except ImportError:
                # Fallback: Open in Notepad so user can use File > Print
                subprocess.Popen(['notepad.exe', temp_path])
                messagebox.showinfo(
                    "Print", 
                    "File opened in Notepad.\n\nUse File → Print (Ctrl+P) to print with your preferred printer."
                )
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not print:\n{str(e)}")
    
    def show_printer_dialog(self, file_path):
        """Show a printer selection dialog and print to selected printer"""
        try:
            import win32print
            import win32api
            
            # Get list of printers
            printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
            default_printer = win32print.GetDefaultPrinter()
            
            if not printers:
                messagebox.showerror("Error", "No printers found.")
                return
            
            # Create printer selection dialog
            dialog = tk.Toplevel(self.root)
            dialog.title("Select Printer")
            dialog.geometry("350x150")
            dialog.configure(bg='#e8e4e0')
            dialog.transient(self.root)
            dialog.grab_set()
            
            # Center the dialog
            dialog.update_idletasks()
            x = self.root.winfo_x() + (self.root.winfo_width() - 350) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - 150) // 2
            dialog.geometry(f"+{x}+{y}")
            
            tk.Label(dialog, text="Select Printer:", font=('Segoe UI', 11), bg='#e8e4e0').pack(pady=(15, 5))
            
            # Printer dropdown
            selected_printer = tk.StringVar(value=default_printer)
            printer_dropdown = ttk.Combobox(dialog, textvariable=selected_printer, values=printers, width=40, state='readonly')
            printer_dropdown.pack(pady=10, padx=20)
            
            def do_print():
                printer = selected_printer.get()
                dialog.destroy()
                try:
                    win32api.ShellExecute(0, "printto", file_path, f'"{printer}"', ".", 0)
                    messagebox.showinfo("Print", f"Sent to printer:\n{printer}")
                except Exception as e:
                    messagebox.showerror("Error", f"Print failed:\n{str(e)}")
            
            def cancel():
                dialog.destroy()
            
            btn_frame = tk.Frame(dialog, bg='#e8e4e0')
            btn_frame.pack(pady=15)
            
            tk.Button(btn_frame, text="Print", command=do_print, bg='#E85D04', fg='white', 
                     font=('Segoe UI', 10, 'bold'), padx=15, pady=5).pack(side='left', padx=5)
            tk.Button(btn_frame, text="Cancel", command=cancel, bg='#6B6B6B', fg='white',
                     font=('Segoe UI', 10, 'bold'), padx=15, pady=5).pack(side='left', padx=5)
            
        except Exception as e:
            # Fallback to notepad
            subprocess.Popen(['notepad.exe', file_path])
            messagebox.showinfo(
                "Print", 
                "File opened in Notepad.\n\nUse File → Print (Ctrl+P) to print."
            )


def main():
    """Main entry point"""
    try:
        from tkinterdnd2 import TkinterDnD
        root = TkinterDnD.Tk()
    except ImportError:
        root = tk.Tk()
    
    app = LabelSorter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
