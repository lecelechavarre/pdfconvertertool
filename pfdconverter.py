import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import threading
from pdf2docx import Converter
from docx2pdf import convert
import customtkinter as ctk

# Set appearance mode and color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class ConverterApp:
    def __init__(self):
        self.window = ctk.CTk()
        self.window.title("PDF <-> Word Converter Pro")
        self.window.geometry("800x600")
        self.window.resizable(False, False)
        
        # Center window
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.window.winfo_screenheight() // 2) - (600 // 2)
        self.window.geometry(f'800x600+{x}+{y}')
        
        self.selected_file = None
        self.setup_ui()
    
    def setup_ui(self):
        # Header
        header_frame = ctk.CTkFrame(self.window, height=80, corner_radius=0)
        header_frame.pack(fill="x")
        header_frame.pack_propagate(False)
        
        header_label = ctk.CTkLabel(
            header_frame, 
            text="📄 Document Converter Pro", 
            font=("Helvetica", 28, "bold")
        )
        header_label.pack(expand=True)
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Convert between PDF and Word formats seamlessly",
            font=("Helvetica", 12)
        )
        subtitle.pack()
        
        # Main container
        main_container = ctk.CTkFrame(self.window, fg_color="transparent")
        main_container.pack(expand=True, fill="both", padx=40, pady=30)
        
        # Left panel - PDF to Word
        left_panel = ctk.CTkFrame(main_container, width=350, height=400)
        left_panel.pack(side="left", expand=True, fill="both", padx=(0, 20))
        left_panel.pack_propagate(False)
        
        pdf_word_header = ctk.CTkLabel(
            left_panel,
            text="PDF → Word",
            font=("Helvetica", 20, "bold")
        )
        pdf_word_header.pack(pady=(25, 10))
        
        pdf_word_desc = ctk.CTkLabel(
            left_panel,
            text="Convert PDF documents to editable Word files",
            font=("Helvetica", 11),
            text_color="gray"
        )
        pdf_word_desc.pack(pady=(0, 25))
        
        self.pdf_word_btn = ctk.CTkButton(
            left_panel,
            text="📂 Select PDF File",
            command=lambda: self.select_file("pdf"),
            width=250,
            height=45,
            font=("Helvetica", 13, "bold"),
            corner_radius=8
        )
        self.pdf_word_btn.pack(pady=10)
        
        self.pdf_word_convert = ctk.CTkButton(
            left_panel,
            text="🔄 Convert to Word",
            command=lambda: self.start_conversion("pdf_to_word"),
            width=250,
            height=45,
            font=("Helvetica", 13, "bold"),
            corner_radius=8,
            fg_color="#27ae60",
            hover_color="#229954",
            state="disabled"
        )
        self.pdf_word_convert.pack(pady=10)
        
        # Right panel - Word to PDF
        right_panel = ctk.CTkFrame(main_container, width=350, height=400)
        right_panel.pack(side="right", expand=True, fill="both")
        right_panel.pack_propagate(False)
        
        word_pdf_header = ctk.CTkLabel(
            right_panel,
            text="Word → PDF",
            font=("Helvetica", 20, "bold")
        )
        word_pdf_header.pack(pady=(25, 10))
        
        word_pdf_desc = ctk.CTkLabel(
            right_panel,
            text="Convert Word documents to PDF format",
            font=("Helvetica", 11),
            text_color="gray"
        )
        word_pdf_desc.pack(pady=(0, 25))
        
        self.word_pdf_btn = ctk.CTkButton(
            right_panel,
            text="📂 Select Word File",
            command=lambda: self.select_file("docx"),
            width=250,
            height=45,
            font=("Helvetica", 13, "bold"),
            corner_radius=8
        )
        self.word_pdf_btn.pack(pady=10)
        
        self.word_pdf_convert = ctk.CTkButton(
            right_panel,
            text="🔄 Convert to PDF",
            command=lambda: self.start_conversion("word_to_pdf"),
            width=250,
            height=45,
            font=("Helvetica", 13, "bold"),
            corner_radius=8,
            fg_color="#27ae60",
            hover_color="#229954",
            state="disabled"
        )
        self.word_pdf_convert.pack(pady=10)
        
        # Status panel
        status_frame = ctk.CTkFrame(self.window, height=120, corner_radius=0)
        status_frame.pack(fill="x", side="bottom")
        status_frame.pack_propagate(False)
        
        self.status_label = ctk.CTkLabel(
            status_frame,
            text="💡 Ready to convert files",
            font=("Helvetica", 13),
            anchor="w"
        )
        self.status_label.pack(pady=(20, 5), padx=30, anchor="w")
        
        self.progress_bar = ctk.CTkProgressBar(status_frame, width=700)
        self.progress_bar.pack(pady=10, padx=30)
        self.progress_bar.set(0)
    
    def select_file(self, file_type):
        if file_type == "pdf":
            file_path = filedialog.askopenfilename(
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf")]
            )
            if file_path:
                self.selected_file = file_path
                self.pdf_word_convert.configure(state="normal")
                self.status_label.configure(
                    text=f"📎 Selected: {os.path.basename(file_path)}",
                    text_color="#27ae60"
                )
        else:
            file_path = filedialog.askopenfilename(
                title="Select Word File",
                filetypes=[("Word files", "*.docx")]
            )
            if file_path:
                self.selected_file = file_path
                self.word_pdf_convert.configure(state="normal")
                self.status_label.configure(
                    text=f"📎 Selected: {os.path.basename(file_path)}",
                    text_color="#27ae60"
                )
    
    def start_conversion(self, conversion_type):
        if not self.selected_file:
            messagebox.showerror("Error", "No file selected")
            return
        
        # Disable buttons during conversion
        self.pdf_word_convert.configure(state="disabled")
        self.word_pdf_convert.configure(state="disabled")
        self.pdf_word_btn.configure(state="disabled")
        self.word_pdf_btn.configure(state="disabled")
        
        # Start conversion in separate thread
        thread = threading.Thread(
            target=self.convert_file,
            args=(conversion_type,)
        )
        thread.daemon = True
        thread.start()
    
    def convert_file(self, conversion_type):
        try:
            self.progress_bar.set(0.3)
            self.status_label.configure(
                text="⚙️ Converting...",
                text_color="#f39c12"
            )
            
            if conversion_type == "pdf_to_word":
                output_path = str(Path(self.selected_file).with_suffix('.docx'))
                
                # PDF to Word conversion
                cv = Converter(self.selected_file)
                cv.convert(output_path, start=0, end=None)
                cv.close()
                
                self.progress_bar.set(1.0)
                self.window.after(0, lambda: self.conversion_success(
                    "PDF to Word", output_path
                ))
                
            else:
                output_path = str(Path(self.selected_file).with_suffix('.pdf'))
                
                # Word to PDF conversion
                convert(self.selected_file, output_path)
                
                self.progress_bar.set(1.0)
                self.window.after(0, lambda: self.conversion_success(
                    "Word to PDF", output_path
                ))
                
        except Exception as e:
            self.window.after(0, lambda: self.conversion_error(str(e)))
    
    def conversion_success(self, conversion_type, output_path):
        self.progress_bar.set(0)
        self.status_label.configure(
            text=f"✅ {conversion_type} conversion completed successfully!",
            text_color="#27ae60"
        )
        
        # Re-enable buttons
        self.pdf_word_convert.configure(state="normal")
        self.word_pdf_convert.configure(state="normal")
        self.pdf_word_btn.configure(state="normal")
        self.word_pdf_btn.configure(state="normal")
        
        messagebox.showinfo(
            "Success",
            f"File converted successfully!\nSaved as:\n{output_path}"
        )
    
    def conversion_error(self, error_msg):
        self.progress_bar.set(0)
        self.status_label.configure(
            text="❌ Conversion failed",
            text_color="#e74c3c"
        )
        
        # Re-enable buttons
        self.pdf_word_convert.configure(state="normal")
        self.word_pdf_convert.configure(state="normal")
        self.pdf_word_btn.configure(state="normal")
        self.word_pdf_btn.configure(state="normal")
        
        messagebox.showerror(
            "Conversion Error",
            f"Failed to convert file:\n{error_msg}"
        )
    
    def run(self):
        self.window.mainloop()


def main():
    # Check and install required packages
    required_packages = ['pdf2docx', 'docx2pdf', 'customtkinter']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
        except ImportError:
            missing_packages.append(package)
    
    if missing_packages:
        print("Missing required packages. Please install:")
        for package in missing_packages:
            print(f"pip install {package}")
        print("\nInstalling missing packages...")
        
        import subprocess
        import sys
        
        for package in missing_packages:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print("Packages installed successfully!")
    
    # Run application
    app = ConverterApp()
    app.run()


if __name__ == "__main__":
    main()