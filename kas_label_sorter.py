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

# Drill bit reference data - lookup by UPS# (RC02 field)
# Format: UPS#: {'part': Part, 'single_price': SinglePrice, 'pack_price': PackPrice, 'pack_weight': PackWeight}
DRILL_BIT_DATA = {
    '3006708': {'part': 'S150104', 'single_price': 0.71, 'pack_price': 8.52, 'pack_weight': 0.08},
    '3006716': {'part': 'S150105', 'single_price': 0.73, 'pack_price': 8.76, 'pack_weight': 0.10},
    '3006724': {'part': 'S150106', 'single_price': 0.85, 'pack_price': 10.20, 'pack_weight': 0.12},
    '3006732': {'part': 'S150107', 'single_price': 0.98, 'pack_price': 11.76, 'pack_weight': 0.16},
    '3006740': {'part': 'S150108', 'single_price': 1.06, 'pack_price': 12.72, 'pack_weight': 0.20},
    '3006758': {'part': 'S150109', 'single_price': 1.15, 'pack_price': 13.80, 'pack_weight': 0.26},
    '3006766': {'part': 'S150110', 'single_price': 1.31, 'pack_price': 15.72, 'pack_weight': 0.30},
    '3006782': {'part': 'S150112', 'single_price': 1.57, 'pack_price': 18.84, 'pack_weight': 0.44},
    '3006790': {'part': 'S150113', 'single_price': 1.63, 'pack_price': 19.56, 'pack_weight': 0.50},
    '3006807': {'part': 'S150114', 'single_price': 2.00, 'pack_price': 24.00, 'pack_weight': 0.56},
    '3006815': {'part': 'S150115', 'single_price': 2.17, 'pack_price': 26.04, 'pack_weight': 0.68},
    '3006823': {'part': 'S150116', 'single_price': 2.54, 'pack_price': 30.48, 'pack_weight': 0.82},
    '3006873': {'part': 'S150121', 'single_price': 3.81, 'pack_price': 22.86, 'pack_weight': 0.68},
    '3006881': {'part': 'S150122', 'single_price': 4.05, 'pack_price': 24.30, 'pack_weight': 0.78},
    '3006899': {'part': 'S150123', 'single_price': 4.52, 'pack_price': 27.12, 'pack_weight': 0.88},
    '3006956': {'part': 'S150129', 'single_price': 6.50, 'pack_price': 39.00, 'pack_weight': 1.46},
    '3007201': {'part': 'SD50407', 'single_price': 14.38, 'pack_price': 14.38, 'pack_weight': 0.34},
    '3007227': {'part': 'S160206', 'single_price': 1.99, 'pack_price': 23.88, 'pack_weight': 0.54},
}


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
                    
                    logo_label = tk.Label(header_frame, image=self.logo_image, bg='#e8e4e0')
                    logo_label.pack()
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
            fg='#E85D04'
        )
        title_label.pack()
    
    def create_preview_text(self):
        """Create the preview text widget (replaces placeholder)"""
        if self.preview_text:
            return
        
        self.preview_placeholder.destroy()
        
        self.preview_text = tk.Text(
            self.preview_frame,
            font=('Consolas', 9),
            bg='#1a1a1a',
            fg='#cccccc',
            insertbackground='white',
            wrap='none',
            state='disabled'
        )
        preview_scroll_y = ttk.Scrollbar(self.preview_frame, command=self.preview_text.yview)
        preview_scroll_x = ttk.Scrollbar(self.preview_frame, orient='horizontal', command=self.preview_text.xview)
        self.preview_text.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        
        preview_scroll_y.pack(side='right', fill='y')
        preview_scroll_x.pack(side='bottom', fill='x')
        self.preview_text.pack(side='left', fill='both', expand=True)
        
        # Preview text tags
        self.preview_text.tag_configure('order_num', foreground='#E85D04', font=('Consolas', 10, 'bold'))
        self.preview_text.tag_configure('header', foreground='#E85D04', font=('Consolas', 11, 'bold'))
        self.preview_text.tag_configure('subheader', foreground='#ffd700', font=('Consolas', 10, 'bold'))
        self.preview_text.tag_configure('summary_label', foreground='#a0a0a0')
        self.preview_text.tag_configure('summary_value', foreground='#4ecca3', font=('Consolas', 10))
        self.preview_text.tag_configure('cost', foreground='#90EE90', font=('Consolas', 10))
        self.preview_text.tag_configure('info', foreground='#87CEEB')
    
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
        
        # Generate single-line output (stored in memory for printing)
        sorted_lines = []
        for order_num, product, address, ups_number, record in records_with_order:
            single_line = self.record_to_single_line(record)
            sorted_lines.append(single_line)
        
        self.sorted_content = '\n'.join(sorted_lines)
        
        # Calculate order summaries (weight and cost per unique order)
        order_summaries = defaultdict(lambda: {'weight': 0.0, 'cost': 0.0, 'items': []})
        for order_num, product, address, ups_number, record in records_with_order:
            # Look up drill bit data by UPS#
            if ups_number in DRILL_BIT_DATA:
                drill_data = DRILL_BIT_DATA[ups_number]
                order_summaries[order_num]['weight'] += drill_data['pack_weight']
                order_summaries[order_num]['cost'] += drill_data['pack_price']
                order_summaries[order_num]['items'].append({
                    'part': drill_data['part'],
                    'ups': ups_number,
                    'weight': drill_data['pack_weight'],
                    'cost': drill_data['pack_price']
                })
            else:
                # Unknown UPS# - still track it but with zero values
                order_summaries[order_num]['items'].append({
                    'part': 'UNKNOWN',
                    'ups': ups_number,
                    'weight': 0.0,
                    'cost': 0.0
                })
        
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
        """Update the preview tab with order summary table"""
        if not self.preview_text:
            return
        
        self.preview_text.config(state='normal')
        self.preview_text.delete('1.0', tk.END)
        
        if hasattr(self, 'order_summaries') and self.order_summaries:
            # Header
            self.preview_text.insert(tk.END, "═" * 60 + "\n", 'header')
            self.preview_text.insert(tk.END, "  ORDER SUMMARY\n", 'header')
            self.preview_text.insert(tk.END, "═" * 60 + "\n\n", 'header')
            
            # Column headers
            self.preview_text.insert(tk.END, f"  {'Order Number':<18} {'Total Weight':>14} {'Total Cost':>14}\n", 'subheader')
            self.preview_text.insert(tk.END, "  " + "─" * 50 + "\n", 'summary_label')
            
            # Calculate grand totals
            grand_weight = 0.0
            grand_cost = 0.0
            
            # Display each order summary
            for order_num in sorted(self.order_summaries.keys()):
                summary = self.order_summaries[order_num]
                weight = summary['weight']
                cost = summary['cost']
                grand_weight += weight
                grand_cost += cost
                
                self.preview_text.insert(tk.END, f"  ", 'summary_label')
                self.preview_text.insert(tk.END, f"{order_num:<18}", 'order_num')
                self.preview_text.insert(tk.END, f" {weight:>12.2f} lb", 'summary_value')
                self.preview_text.insert(tk.END, f" {cost:>13.2f}\n", 'cost')
            
            # Grand totals
            self.preview_text.insert(tk.END, "  " + "─" * 50 + "\n", 'summary_label')
            self.preview_text.insert(tk.END, f"  {'GRAND TOTAL':<18}", 'header')
            self.preview_text.insert(tk.END, f" {grand_weight:>12.2f} lb", 'summary_value')
            self.preview_text.insert(tk.END, f" ${grand_cost:>12.2f}\n", 'cost')
            
            self.preview_text.insert(tk.END, "\n" + "═" * 60 + "\n", 'header')
            self.preview_text.insert(tk.END, f"  Total Unique Orders: {len(self.order_summaries)}\n", 'info')
            self.preview_text.insert(tk.END, "═" * 60 + "\n", 'header')
        
        self.preview_text.config(state='disabled')
    
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
