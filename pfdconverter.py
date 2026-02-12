import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import threading
from pdf2docx import Converter
import sys
import subprocess

# Handle PIL import with auto-install
try:
    from PIL import Image, ImageTk
except ImportError:
    print("Installing Pillow...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "Pillow"])
    from PIL import Image, ImageTk

# Handle fitz import with auto-install
try:
    import fitz
except ImportError:
    print("Installing PyMuPDF...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "PyMuPDF"])
    import fitz

# Handle docx import with auto-install
try:
    from docx import Document
except ImportError:
    print("Installing python-docx...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "python-docx"])
    from docx import Document

# Handle reportlab import with auto-install
try:
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
except ImportError:
    print("Installing reportlab...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "reportlab"])
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch

import io


class ConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Document Converter")
        self.window.state('zoomed')
        self.window.configure(bg='#f5f5f7')
        
        self.selected_file = None
        self.current_mode = None
        self.converted_file = None
        self.preview_image = None
        self.setup_fonts()
        self.setup_ui()
        
        # Bind resize event
        self.window.bind('<Configure>', self.on_window_resize)
    
    def setup_fonts(self):
        self.font_regular = ('Helvetica', 10)
        self.font_medium = ('Helvetica', 11)
        self.font_bold = ('Helvetica', 24, 'bold')
        self.font_button = ('Helvetica', 11)
        self.font_card_title = ('Helvetica', 13, 'bold')
        
    def setup_ui(self):
        # Main container with padding
        main_container = tk.Frame(self.window, bg='#f5f5f7')
        main_container.pack(expand=True, fill='both', padx=40, pady=30)
        
        # Configure grid weights
        main_container.grid_columnconfigure(0, weight=35)  # Left panel - 35%
        main_container.grid_columnconfigure(1, weight=65)  # Right panel - 65%
        main_container.grid_rowconfigure(0, weight=1)
        
        # ============ LEFT PANEL - CONVERTER ============
        left_panel = tk.Frame(main_container, bg='#f5f5f7')
        left_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 15))
        
        # Header
        header = tk.Frame(left_panel, bg='#f5f5f7')
        header.pack(fill='x', pady=(0, 20))
        
        title = tk.Label(
            header,
            text='Document Converter',
            font=self.font_bold,
            bg='#f5f5f7',
            fg='#1d1d1f',
            anchor='w'
        )
        title.pack(fill='x')
        
        # Divider
        divider = tk.Frame(left_panel, height=1, bg='#d2d2d7')
        divider.pack(fill='x', pady=(0, 25))
        
        # Cards container
        cards_frame = tk.Frame(left_panel, bg='#f5f5f7')
        cards_frame.pack(fill='x')
        
        # PDF to Word card
        self.create_option_card(
            cards_frame,
            'PDF to Word',
            'Convert PDF files to editable Word documents',
            'pdf',
            0
        )
        
        # Word to PDF card
        self.create_option_card(
            cards_frame,
            'Word to PDF',
            'Convert Word documents to PDF format',
            'docx',
            1
        )
        
        # File info frame
        info_container = tk.Frame(left_panel, bg='#f5f5f7')
        info_container.pack(fill='x', pady=(25, 15))
        
        self.info_frame = tk.Frame(info_container, bg='#ffffff', relief='flat', bd=0, height=50)
        self.info_frame.pack(fill='x')
        self.info_frame.pack_propagate(False)
        
        self.file_label = tk.Label(
            self.info_frame,
            text='No file selected',
            font=self.font_medium,
            bg='#ffffff',
            fg='#86868b',
            anchor='w',
            padx=15,
            pady=12
        )
        self.file_label.pack(fill='both')
        
        # Convert button
        button_container = tk.Frame(left_panel, bg='#f5f5f7')
        button_container.pack(fill='x', pady=(0, 20))
        
        self.convert_btn = tk.Button(
            button_container,
            text='Convert',
            font=self.font_button,
            bg='#86868b',
            fg='#ffffff',
            bd=0,
            padx=25,
            pady=10,
            activebackground='#666666',
            activeforeground='#ffffff',
            state='disabled',
            relief='flat',
            cursor='',
            width=15
        )
        self.convert_btn.pack()
        
        # Status bar
        status_container = tk.Frame(left_panel, bg='#f5f5f7')
        status_container.pack(fill='x', side='bottom', pady=(10, 0))
        
        self.status_frame = tk.Frame(status_container, bg='#ffffff', relief='flat', bd=0, height=44)
        self.status_frame.pack(fill='x')
        self.status_frame.pack_propagate(False)
        
        self.status_label = tk.Label(
            self.status_frame,
            text='Ready',
            font=self.font_regular,
            bg='#ffffff',
            fg='#86868b',
            anchor='w',
            padx=15,
            pady=12
        )
        self.status_label.pack(fill='both')
        
        # ============ RIGHT PANEL - PREVIEW ============
        right_panel = tk.Frame(main_container, bg='#ffffff', relief='flat', bd=0)
        right_panel.grid(row=0, column=1, sticky='nsew', padx=(15, 0))
        
        # Preview header
        preview_header = tk.Frame(right_panel, bg='#ffffff', height=50)
        preview_header.pack(fill='x', padx=25, pady=(20, 10))
        preview_header.pack_propagate(False)
        
        preview_title = tk.Label(
            preview_header,
            text='Preview',
            font=('Helvetica', 16, 'bold'),
            bg='#ffffff',
            fg='#1d1d1f'
        )
        preview_title.pack(side='left')
        
        self.preview_filename = tk.Label(
            preview_header,
            text='',
            font=self.font_regular,
            bg='#ffffff',
            fg='#86868b'
        )
        self.preview_filename.pack(side='right', padx=(10, 0))
        
        # Preview area with scrollbar
        preview_container = tk.Frame(right_panel, bg='#f5f5f7')
        preview_container.pack(expand=True, fill='both', padx=25, pady=(0, 25))
        
        # Create canvas with scrollbar
        self.preview_canvas = tk.Canvas(
            preview_container,
            bg='#ffffff',
            highlightthickness=1,
            highlightbackground='#e6e6e8',
            relief='flat'
        )
        
        scrollbar = tk.Scrollbar(
            preview_container,
            orient='vertical',
            command=self.preview_canvas.yview
        )
        
        self.preview_canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack canvas and scrollbar
        self.preview_canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Create inner frame for content
        self.preview_inner = tk.Frame(self.preview_canvas, bg='#ffffff')
        self.preview_canvas.create_window((0, 0), window=self.preview_inner, anchor='nw')
        
        # Bind events
        self.preview_inner.bind('<Configure>', self.on_inner_configure)
        self.preview_canvas.bind('<Configure>', self.on_canvas_configure)
        
        # Show placeholder
        self.show_preview_placeholder()
    
    def on_window_resize(self, event):
        if event.widget == self.window:
            self.update_preview_layout()
    
    def on_inner_configure(self, event):
        self.preview_canvas.configure(scrollregion=self.preview_canvas.bbox('all'))
    
    def on_canvas_configure(self, event):
        self.preview_canvas.itemconfig('all', width=event.width)
    
    def update_preview_layout(self):
        if hasattr(self, 'preview_canvas'):
            self.preview_canvas.itemconfig('all', width=self.preview_canvas.winfo_width())
    
    def show_preview_placeholder(self):
        # Clear inner frame
        for widget in self.preview_inner.winfo_children():
            widget.destroy()
        
        # Create placeholder
        placeholder = tk.Frame(self.preview_inner, bg='#ffffff', height=400)
        placeholder.pack(expand=True, fill='both', padx=40, pady=80)
        
        icon_label = tk.Label(
            placeholder,
            text='👀',
            font=('Helvetica', 48),
            bg='#ffffff',
            fg='#d2d2d7'
        )
        icon_label.pack(pady=(0, 20))
        
        text_label = tk.Label(
            placeholder,
            text='No preview available',
            font=('Helvetica', 16),
            bg='#ffffff',
            fg='#86868b'
        )
        text_label.pack()
        
        subtext_label = tk.Label(
            placeholder,
            text='Convert a file to see preview',
            font=('Helvetica', 12),
            bg='#ffffff',
            fg='#a1a1a6'
        )
        subtext_label.pack(pady=(10, 0))
        
        self.preview_filename.configure(text='')
    
    def create_option_card(self, parent, title, description, mode, col):
        card = tk.Frame(
            parent,
            bg='#ffffff',
            highlightbackground='#e6e6e8',
            highlightthickness=1,
            bd=0,
            height=200
        )
        card.pack(side='left', padx=(0 if col == 0 else 8, 8 if col == 0 else 0), fill='both', expand=True)
        card.pack_propagate(False)
        
        content = tk.Frame(card, bg='#ffffff')
        content.pack(fill='both', expand=True, padx=18, pady=18)
        
        # Icon and title row
        header_row = tk.Frame(content, bg='#ffffff')
        header_row.pack(anchor='w', fill='x', pady=(0, 12))
        
        icon = tk.Label(
            header_row,
            text='📄' if mode == 'pdf' else '📝',
            font=('Helvetica', 32),
            bg='#ffffff',
            fg='#1d1d1f'
        )
        icon.pack(side='left', padx=(0, 10))
        
        title_label = tk.Label(
            header_row,
            text=title,
            font=self.font_card_title,
            bg='#ffffff',
            fg='#1d1d1f'
        )
        title_label.pack(side='left')
        
        # Description
        desc_label = tk.Label(
            content,
            text=description,
            font=self.font_regular,
            bg='#ffffff',
            fg='#86868b',
            justify='left',
            wraplength=220
        )
        desc_label.pack(anchor='w', pady=(0, 20))
        
        # Select button
        select_btn = tk.Button(
            content,
            text='Select File',
            font=self.font_regular,
            bg='#ffffff',
            fg='#0066cc',
            bd=0,
            padx=0,
            pady=2,
            cursor='hand2',
            activebackground='#ffffff',
            activeforeground='#004999',
            command=lambda m=mode: self.select_file(m),
            relief='flat'
        )
        select_btn.pack(anchor='w')
        
        # Hover effect
        def on_enter(e):
            card.configure(highlightbackground='#0066cc')
            select_btn.configure(fg='#004999')
        
        def on_leave(e):
            card.configure(highlightbackground='#e6e6e8')
            select_btn.configure(fg='#0066cc')
        
        card.bind('<Enter>', on_enter)
        card.bind('<Leave>', on_leave)
        
        return card
    
    def select_file(self, mode):
        if mode == 'pdf':
            filetypes = [('PDF files', '*.pdf')]
            file_type = 'PDF'
        else:
            filetypes = [('Word files', '*.docx')]
            file_type = 'Word'
        
        file_path = filedialog.askopenfilename(
            title=f'Select {file_type} File',
            filetypes=filetypes
        )
        
        if file_path:
            self.selected_file = file_path
            self.current_mode = mode
            self.converted_file = None
            self.show_preview_placeholder()
            
            filename = os.path.basename(file_path)
            self.file_label.configure(
                text=f'Selected: {filename}',
                fg='#1d1d1f'
            )
            
            self.convert_btn.configure(
                bg='#0066cc',
                fg='#ffffff',
                state='normal',
                cursor='hand2',
                activebackground='#004999',
                text='Convert',
                command=self.start_conversion
            )
            
            self.status_label.configure(
                text='Ready to convert',
                fg='#1d1d1f'
            )
    
    def start_conversion(self):
        if not self.selected_file:
            return
        
        self.convert_btn.configure(
            bg='#86868b',
            fg='#ffffff',
            state='disabled',
            cursor='',
            text='Converting...',
            command=None
        )
        
        self.status_label.configure(
            text='Converting...',
            fg='#1d1d1f'
        )
        
        thread = threading.Thread(target=self.convert_file)
        thread.daemon = True
        thread.start()
    
    def convert_file(self):
        error = None
        output_path = None
        
        try:
            if self.current_mode == 'pdf':
                output_path = str(Path(self.selected_file).with_suffix('.docx'))
                cv = Converter(self.selected_file)
                cv.convert(output_path, start=0, end=None)
                cv.close()
            else:
                output_path = str(Path(self.selected_file).with_suffix('.pdf'))
                
                # Clean DOCX to PDF conversion without artifacts
                doc = Document(self.selected_file)
                c = canvas.Canvas(output_path, pagesize=letter)
                width, height = letter
                
                y = height - inch
                line_height = 14
                margin = inch
                
                for paragraph in doc.paragraphs:
                    text = paragraph.text.strip()
                    if text:
                        # Clean text - remove any control characters
                        clean_text = ''.join(char for char in text if ord(char) >= 32 or char in '\n\t')
                        
                        # Draw text with proper spacing
                        c.setFont('Helvetica', 11)
                        c.setFillColorRGB(0, 0, 0)
                        c.drawString(margin, y, clean_text)
                        y -= line_height
                        
                        if y < margin:
                            c.showPage()
                            y = height - margin
                            c.setFont('Helvetica', 11)
                
                c.save()
                
        except Exception as e:
            error = str(e)
        
        if error:
            self.window.after(0, lambda err=error: self.conversion_error(err))
        else:
            self.converted_file = output_path
            self.window.after(0, lambda path=output_path: self.conversion_success(path))
    
    def preview_pdf(self, pdf_path):
        # Clear inner frame
        for widget in self.preview_inner.winfo_children():
            widget.destroy()
        
        try:
            doc = fitz.open(pdf_path)
            
            # Create container for pages
            pages_container = tk.Frame(self.preview_inner, bg='#ffffff')
            pages_container.pack(expand=True, fill='both', padx=30, pady=30)
            
            canvas_width = self.preview_canvas.winfo_width() - 60
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Calculate zoom to fit width
                zoom = canvas_width / page.rect.width
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                
                # Convert to PIL Image
                img_data = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_data))
                
                # Convert to PhotoImage
                self.preview_image = ImageTk.PhotoImage(img)
                
                # Page frame
                page_frame = tk.Frame(pages_container, bg='#ffffff')
                page_frame.pack(pady=(0, 20))
                
                # Page number
                if len(doc) > 1:
                    page_label = tk.Label(
                        page_frame,
                        text=f'Page {page_num + 1} of {len(doc)}',
                        font=('Helvetica', 10),
                        bg='#ffffff',
                        fg='#86868b'
                    )
                    page_label.pack(pady=(0, 5))
                
                # Page image
                image_label = tk.Label(
                    page_frame,
                    image=self.preview_image,
                    bg='#ffffff',
                    bd=0
                )
                image_label.pack()
            
            doc.close()
            
        except Exception as e:
            error_label = tk.Label(
                self.preview_inner,
                text='⚠️ Preview not available',
                font=('Helvetica', 12),
                bg='#ffffff',
                fg='#86868b'
            )
            error_label.pack(expand=True)
    
    def preview_docx(self, docx_path):
        # Clear inner frame
        for widget in self.preview_inner.winfo_children():
            widget.destroy()
        
        try:
            doc = Document(docx_path)
            
            # Create container
            container = tk.Frame(self.preview_inner, bg='#ffffff')
            container.pack(expand=True, fill='both', padx=30, pady=30)
            
            # Header
            header_frame = tk.Frame(container, bg='#ffffff')
            header_frame.pack(fill='x', pady=(0, 20))
            
            tk.Label(
                header_frame,
                text='Document Content',
                font=('Helvetica', 14, 'bold'),
                bg='#ffffff',
                fg='#1d1d1f'
            ).pack(anchor='w')
            
            # Content
            content_frame = tk.Frame(container, bg='#ffffff')
            content_frame.pack(fill='both', expand=True)
            
            text_widget = tk.Text(
                content_frame,
                font=('Helvetica', 11),
                bg='#ffffff',
                fg='#1d1d1f',
                wrap='word',
                bd=0,
                highlightthickness=0,
                relief='flat'
            )
            text_widget.pack(fill='both', expand=True)
            
            # Insert clean text
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    # Clean text - remove any control characters
                    clean_text = ''.join(char for char in paragraph.text if ord(char) >= 32 or char == '\n')
                    text_widget.insert('end', clean_text + '\n\n')
            
            text_widget.configure(state='disabled')
            
        except Exception as e:
            error_label = tk.Label(
                self.preview_inner,
                text='📄 Document preview',
                font=('Helvetica', 12),
                bg='#ffffff',
                fg='#86868b'
            )
            error_label.pack(expand=True)
    
    def conversion_success(self, output_path):
        self.status_label.configure(
            text='Conversion completed',
            fg='#1d1d1f'
        )
        
        self.convert_btn.configure(
            bg='#0066cc',
            fg='#ffffff',
            state='normal',
            cursor='hand2',
            text='Convert',
            command=self.start_conversion
        )
        
        self.file_label.configure(
            text='No file selected',
            fg='#86868b'
        )
        
        # Update preview
        filename = os.path.basename(output_path)
        self.preview_filename.configure(text=filename)
        
        if output_path.endswith('.pdf'):
            self.preview_pdf(output_path)
        else:
            self.preview_docx(output_path)
        
        self.selected_file = None
    
    def conversion_error(self, error_msg):
        self.status_label.configure(
            text='Conversion failed',
            fg='#ff3b30'
        )
        
        self.convert_btn.configure(
            bg='#0066cc',
            fg='#ffffff',
            state='normal',
            cursor='hand2',
            text='Convert',
            command=self.start_conversion
        )
        
        self.show_preview_placeholder()
        messagebox.showerror('Error', f'Failed to convert file.')
    
    def run(self):
        self.window.mainloop()


def main():
    app = ConverterApp()
    app.run()


if __name__ == "__main__":
    main()
