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
import re
import shutil
import tempfile


class ConverterApp:
    def __init__(self):
        self.window = tk.Tk()
        self.window.title("Document Converter")
        self.window.state('zoomed')
        self.window.configure(bg='#f5f5f7')
        
        self.selected_file = None
        self.current_mode = None
        self.converted_file = None
        self.preview_file = None
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
        main_container.grid_columnconfigure(0, weight=35)
        main_container.grid_columnconfigure(1, weight=65)
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
        
        # Download button - ONLY DOWNLOADS WHEN CLICKED
        download_container = tk.Frame(left_panel, bg='#f5f5f7')
        download_container.pack(fill='x', pady=(0, 15))
        
        self.download_btn = tk.Button(
            download_container,
            text='⬇️ Download Converted File',
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
            width=20,
            command=self.download_file
        )
        self.download_btn.pack()
        
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
            text='Select a file to generate preview',
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
            self.preview_file = None
            
            # Clear previous preview
            self.show_preview_placeholder()
            
            filename = os.path.basename(file_path)
            self.file_label.configure(
                text=f'Selected: {filename}',
                fg='#1d1d1f'
            )
            
            # Disable download button until conversion is complete
            self.download_btn.configure(
                state='disabled',
                bg='#86868b',
                fg='#ffffff',
                activebackground='#666666',
                cursor=''
            )
            
            self.status_label.configure(
                text='Generating preview...',
                fg='#1d1d1f'
            )
            
            # Start preview generation thread - NO DOWNLOAD, ONLY PREVIEW
            thread = threading.Thread(target=self.generate_preview)
            thread.daemon = True
            thread.start()
    
    def generate_preview(self):
        """Generate preview only - NO FILE SAVING, NO DOWNLOAD"""
        error = None
        preview_path = None
        
        try:
            # Create temporary file for preview
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf' if self.current_mode == 'docx' else '.docx') as tmp_file:
                preview_path = tmp_file.name
            
            if self.current_mode == 'pdf':
                # PDF to Word preview
                cv = Converter(self.selected_file)
                cv.convert(preview_path, start=0, end=None)
                cv.close()
                
            else:
                # Word to PDF preview
                doc = Document(self.selected_file)
                c = canvas.Canvas(preview_path, pagesize=letter)
                width, height = letter
                
                y = height - inch
                line_height = 14
                margin = inch
                paragraph_spacing = 8
                
                for paragraph in doc.paragraphs:
                    text = paragraph.text
                    
                    # Clean the text thoroughly
                    clean_text = self.clean_text(text)
                    
                    if clean_text:
                        # Draw text with proper spacing
                        c.setFont('Helvetica', 11)
                        c.setFillColorRGB(0, 0, 0)
                        
                        # Handle long text wrapping
                        words = clean_text.split()
                        line = []
                        
                        for word in words:
                            line.append(word)
                            line_width = c.stringWidth(' '.join(line), 'Helvetica', 11)
                            
                            if line_width > (width - 2 * margin):
                                line.pop()
                                if line:
                                    c.drawString(margin, y, ' '.join(line))
                                    y -= line_height
                                    line = [word]
                                
                                if y < margin:
                                    c.showPage()
                                    y = height - margin
                                    c.setFont('Helvetica', 11)
                        
                        # Draw remaining text
                        if line:
                            c.drawString(margin, y, ' '.join(line))
                            y -= line_height
                        
                        # Add paragraph spacing
                        y -= paragraph_spacing
                        
                        if y < margin:
                            c.showPage()
                            y = height - margin
                            c.setFont('Helvetica', 11)
                
                c.save()
                
        except Exception as e:
            error = str(e)
            if preview_path and os.path.exists(preview_path):
                os.unlink(preview_path)
        
        if error:
            self.window.after(0, lambda err=error: self.preview_error(err))
        else:
            self.preview_file = preview_path
            self.window.after(0, lambda path=preview_path: self.preview_success(path))
    
    def clean_text(self, text):
        """Remove all hidden characters and artifacts"""
        if not text:
            return ""
        
        # Remove non-printable characters except spaces and newlines
        cleaned = ''.join(char for char in text if ord(char) >= 32 or char == '\n' or char == '\t')
        
        # Replace multiple spaces with single space
        cleaned = re.sub(r' +', ' ', cleaned)
        
        # Fix spacing after periods
        cleaned = re.sub(r'\.(?=[^\s])', '. ', cleaned)
        cleaned = re.sub(r'\. +', '. ', cleaned)
        
        # Remove any remaining control characters
        cleaned = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', cleaned)
        
        return cleaned.strip()
    
    def download_file(self):
        """Download the converted file - ONLY WHEN DOWNLOAD BUTTON IS CLICKED"""
        if not self.converted_file:
            # If no converted file exists, convert first then download
            self.start_conversion_for_download()
            return
            
        if self.converted_file and os.path.exists(self.converted_file):
            # Ask where to save the file
            save_path = filedialog.asksaveasfilename(
                defaultextension=os.path.splitext(self.converted_file)[1],
                filetypes=[
                    ('PDF files', '*.pdf') if self.converted_file.endswith('.pdf') else ('Word files', '*.docx')
                ],
                initialfile=os.path.basename(self.converted_file)
            )
            
            if save_path:
                try:
                    # Copy file to selected location
                    shutil.copy2(self.converted_file, save_path)
                    messagebox.showinfo('Download Complete', f'File saved to:\n{save_path}')
                except Exception as e:
                    messagebox.showerror('Download Failed', f'Could not save file:\n{str(e)}')
    
    def start_conversion_for_download(self):
        """Convert file for download - ONLY CALLED WHEN DOWNLOAD BUTTON IS CLICKED"""
        if not self.selected_file:
            return
        
        self.download_btn.configure(
            state='disabled',
            bg='#86868b',
            fg='#ffffff',
            activebackground='#666666',
            cursor='',
            text='⏳ Converting...'
        )
        
        self.status_label.configure(
            text='Converting file for download...',
            fg='#1d1d1f'
        )
        
        # Start conversion thread
        thread = threading.Thread(target=self.convert_for_download)
        thread.daemon = True
        thread.start()
    
    def convert_for_download(self):
        """Convert file for permanent storage and download"""
        error = None
        output_path = None
        
        try:
            # Create permanent file in user's temp directory
            output_path = str(Path(self.selected_file).with_suffix('.docx' if self.current_mode == 'pdf' else '.pdf'))
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path) if os.path.dirname(output_path) else '.', exist_ok=True)
            
            if self.current_mode == 'pdf':
                # PDF to Word conversion
                cv = Converter(self.selected_file)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                
            else:
                # Word to PDF conversion
                doc = Document(self.selected_file)
                c = canvas.Canvas(output_path, pagesize=letter)
                width, height = letter
                
                y = height - inch
                line_height = 14
                margin = inch
                paragraph_spacing = 8
                
                for paragraph in doc.paragraphs:
                    text = paragraph.text
                    
                    # Clean the text thoroughly
                    clean_text = self.clean_text(text)
                    
                    if clean_text:
                        c.setFont('Helvetica', 11)
                        c.setFillColorRGB(0, 0, 0)
                        
                        words = clean_text.split()
                        line = []
                        
                        for word in words:
                            line.append(word)
                            line_width = c.stringWidth(' '.join(line), 'Helvetica', 11)
                            
                            if line_width > (width - 2 * margin):
                                line.pop()
                                if line:
                                    c.drawString(margin, y, ' '.join(line))
                                    y -= line_height
                                    line = [word]
                                
                                if y < margin:
                                    c.showPage()
                                    y = height - margin
                                    c.setFont('Helvetica', 11)
                        
                        if line:
                            c.drawString(margin, y, ' '.join(line))
                            y -= line_height
                        
                        y -= paragraph_spacing
                        
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
            self.window.after(0, lambda: self.download_ready())
    
    def download_ready(self):
        """Called when conversion for download is complete"""
        self.status_label.configure(
            text='Conversion complete - Ready to download',
            fg='#1d1d1f'
        )
        
        self.download_btn.configure(
            state='normal',
            bg='#34a853',
            fg='#ffffff',
            activebackground='#2d8c46',
            cursor='hand2',
            text='⬇️ Download Converted File',
            command=self.download_file
        )
        
        # Trigger the download
        self.download_file()
    
    def preview_success(self, preview_path):
        self.status_label.configure(
            text='Preview generated',
            fg='#1d1d1f'
        )
        
        # Enable download button
        self.download_btn.configure(
            state='normal',
            bg='#34a853',
            fg='#ffffff',
            activebackground='#2d8c46',
            cursor='hand2',
            text='⬇️ Download Converted File',
            command=self.download_file
        )
        
        self.file_label.configure(
            text=f'Selected: {os.path.basename(self.selected_file)}',
            fg='#1d1d1f'
        )
        
        # Update preview
        filename = os.path.basename(preview_path)
        self.preview_filename.configure(text=f'Preview: {os.path.basename(self.selected_file)}')
        
        if preview_path.endswith('.pdf'):
            self.preview_pdf(preview_path)
        else:
            self.preview_docx(preview_path)
    
    def preview_error(self, error_msg):
        self.status_label.configure(
            text='Preview generation failed',
            fg='#ff3b30'
        )
        
        # Disable download button on preview error
        self.download_btn.configure(
            state='disabled',
            bg='#86868b',
            fg='#ffffff',
            activebackground='#666666',
            cursor=''
        )
        
        self.show_preview_placeholder()
        messagebox.showerror('Error', f'Failed to generate preview.\n\n{error_msg}')
    
    def conversion_error(self, error_msg):
        self.status_label.configure(
            text='Conversion failed',
            fg='#ff3b30'
        )
        
        # Re-enable download button
        self.download_btn.configure(
            state='normal',
            bg='#34a853',
            fg='#ffffff',
            activebackground='#2d8c46',
            cursor='hand2',
            text='⬇️ Download Converted File',
            command=self.download_file
        )
        
        messagebox.showerror('Error', f'Failed to convert file.\n\n{error_msg}')
    
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
            
            # Insert clean text with preserved paragraph formatting
            for paragraph in doc.paragraphs:
                if paragraph.text:
                    clean_text = self.clean_text(paragraph.text)
                    if clean_text:
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
    
    def run(self):
        self.window.mainloop()


def main():
    app = ConverterApp()
    app.run()


if __name__ == "__main__":
    app = ConverterApp()
    app.run()
