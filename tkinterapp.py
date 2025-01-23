import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from pathlib import Path
import queue
import os
from PIL import Image
from PyPDF2 import PdfWriter, PdfReader
import comtypes.client
import win32com.client
import tempfile
from datetime import datetime

class PDFMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Page Modifier & Merger")
        self.root.geometry("600x500")
        
        # Variables
        self.files_queue = queue.Queue()
        self.processing = False
        self.uploaded_files = []
        self.settings = {
            'add_numeration': tk.BooleanVar(value=True),
            'added_pages_numeration': tk.BooleanVar(value=False),
            'show_files_count': tk.BooleanVar(value=False),
            'images_numeration': tk.BooleanVar(value=True)
        }
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF Page Modifier & Merger", 
                              font=('Helvetica', 14, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=10)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=3, pady=5)
        
        # Upload button
        self.upload_btn = ttk.Button(button_frame, text="Upload Files", 
                                   command=self.upload_files)
        self.upload_btn.grid(row=0, column=0, padx=5)
        
        # Settings button
        self.settings_btn = ttk.Button(button_frame, text="Settings", 
                                     command=self.show_settings)
        self.settings_btn.grid(row=0, column=1, padx=5)
        
        # Reset button
        self.reset_btn = ttk.Button(button_frame, text="Reset", 
                                  command=self.reset_app)
        self.reset_btn.grid(row=0, column=2, padx=5)
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Ready")
        self.status_label.grid(row=2, column=0, columnspan=3, pady=5)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, length=540, mode='determinate')
        self.progress_bar.grid(row=3, column=0, columnspan=3, pady=5, sticky='ew')
        
        # Files listbox
        self.files_listbox = tk.Listbox(main_frame, height=10, width=70)
        self.files_listbox.grid(row=4, column=0, columnspan=3, pady=5)
        
        # Scrollbar for listbox
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", 
                                command=self.files_listbox.yview)
        scrollbar.grid(row=4, column=3, sticky='ns')
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Options frame
        options_frame = ttk.Frame(main_frame)
        options_frame.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Print on both sides
        self.both_sides_var = tk.BooleanVar()
        self.both_sides_check = ttk.Checkbutton(options_frame, 
                                               text="Print On Both Sides?",
                                               variable=self.both_sides_var)
        self.both_sides_check.grid(row=0, column=0, padx=10)
        
        # Delete specific pages
        self.delete_pages_var = tk.BooleanVar()
        self.delete_pages_check = ttk.Checkbutton(options_frame, 
                                                 text="Delete Specific Pages",
                                                 variable=self.delete_pages_var)
        self.delete_pages_check.grid(row=0, column=1, padx=10)
        
        self.delete_pages_entry = ttk.Entry(options_frame, width=15)
        self.delete_pages_entry.grid(row=0, column=2, padx=5)
        
        # Slides per page
        ttk.Label(options_frame, text="Slides Per Page:").grid(row=0, column=3, 
                                                             padx=5)
        self.slides_per_page = ttk.Combobox(options_frame, values=[1, 2, 4, 6, 9], 
                                          width=5)
        self.slides_per_page.set(1)
        self.slides_per_page.grid(row=0, column=4, padx=5)
        
        # Process button
        self.process_btn = ttk.Button(main_frame, text="Process & Merge PDFs",
                                    command=self.process_files)
        self.process_btn.grid(row=6, column=0, columnspan=3, pady=10)
        
    def show_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("400x300")
        
        frame = ttk.Frame(settings_window, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Add numeration
        ttk.Checkbutton(frame, text="Add page numbers", 
                       variable=self.settings['add_numeration']).grid(row=0, 
                       column=0, sticky='w', pady=5)
        
        # Added pages numeration
        ttk.Checkbutton(frame, text="Number added blank pages", 
                       variable=self.settings['added_pages_numeration']).grid(row=1,
                       column=0, sticky='w', pady=5)
        
        # Show files count
        ttk.Checkbutton(frame, text="Show files count with page numbers", 
                       variable=self.settings['show_files_count']).grid(row=2, 
                       column=0, sticky='w', pady=5)
        
        # Images numeration
        ttk.Checkbutton(frame, text="Add page numbers to images", 
                       variable=self.settings['images_numeration']).grid(row=3, 
                       column=0, sticky='w', pady=5)
    
    def upload_files(self):
        filetypes = (
            ("All supported files", "*.pdf;*.doc;*.docx;*.ppt;*.pptx;*.jpg;*.jpeg;*.png;*.tiff"),
            ("PDF files", "*.pdf"),
            ("Word documents", "*.doc;*.docx"),
            ("PowerPoint files", "*.ppt;*.pptx"),
            ("Image files", "*.jpg;*.jpeg;*.png;*.tiff")
        )
        
        files = filedialog.askopenfilenames(filetypes=filetypes)
        if files:
            self.uploaded_files.extend(files)
            self.update_files_list()
    
    def update_files_list(self):
        self.files_listbox.delete(0, tk.END)
        for i, file in enumerate(self.uploaded_files, 1):
            filename = Path(file).name
            self.files_listbox.insert(tk.END, f"{i}- {filename}")
    
    def reset_app(self):
        self.uploaded_files.clear()
        self.update_files_list()
        self.progress_bar['value'] = 0
        self.status_label['text'] = "Ready"
        self.process_btn['state'] = 'normal'
    
    def convert_to_pdf(self, input_file, output_file):
        file_ext = Path(input_file).suffix.lower()
        
        if file_ext in ['.doc', '.docx']:
            word = win32com.client.Dispatch('Word.Application')
            doc = word.Documents.Open(input_file)
            doc.SaveAs(output_file, FileFormat=17)
            doc.Close()
            word.Quit()
            
        elif file_ext in ['.ppt', '.pptx']:
            powerpoint = win32com.client.Dispatch('Powerpoint.Application')
            presentation = powerpoint.Presentations.Open(input_file)
            presentation.SaveAs(output_file, 32)  # 32 = PDF format
            presentation.Close()
            powerpoint.Quit()
            
        elif file_ext in ['.jpg', '.jpeg', '.png', '.tiff']:
            image = Image.open(input_file)
            image = image.convert('RGB')
            image.save(output_file, 'PDF')
    
    def process_files(self):
        if not self.uploaded_files:
            messagebox.showwarning("Warning", "Please upload files first!")
            return
        
        self.process_btn['state'] = 'disabled'
        self.status_label['text'] = "Processing files..."
        self.progress_bar['value'] = 0
        
        # Start processing thread
        threading.Thread(target=self.process_files_thread, daemon=True).start()
    
    def process_files_thread(self):
        try:
            merger = PdfWriter()
            temp_dir = tempfile.mkdtemp()
            total_files = len(self.uploaded_files)
            file_path = ""  # Initialize file_path
            
            for i, file in enumerate(self.uploaded_files):
                file_path = file  # Update file_path for the current file
                progress = (i / total_files) * 100
                self.root.after(0, lambda: self.update_progress(progress))
                
                file_ext = Path(file).suffix.lower()
                if file_ext != '.pdf':
                    pdf_path = os.path.join(temp_dir, f"converted_{i}.pdf")
                    self.convert_to_pdf(file, pdf_path)
                    pdf_file = pdf_path
                else:
                    pdf_file = file
                
                # Process PDF pages
                reader = PdfReader(pdf_file)
                if self.delete_pages_var.get():
                    pages_to_delete = self.parse_delete_pages(
                        self.delete_pages_entry.get(),
                        len(reader.pages)
                    )
                    for page_num in range(len(reader.pages)):
                        if page_num + 1 not in pages_to_delete:
                            merger.add_page(reader.pages[page_num])
                else:
                    self.merge_pdfs(merger, [pdf_file])
            
            # Save the merged PDF
            output_path = os.path.join(os.path.dirname(self.uploaded_files[0]),
                                     f"merged_{datetime.now():%Y%m%d_%H%M%S}.pdf")
            
            with open(output_path, 'wb') as output_file:
                merger.write(output_file)
            
            self.root.after(0, lambda: self.process_complete(output_path))
            
        except Exception as e:
            error_message = f"An error occurred: {str(e)}\nFile Path: {file_path}"
            self.root.after(0, lambda: self.show_error(error_message))
        finally:
            # Cleanup
            import shutil
            shutil.rmtree(temp_dir, ignore_errors=True)
    
    def parse_delete_pages(self, delete_str, max_pages):
        pages_to_delete = set()
        if not delete_str:
            return pages_to_delete
        
        parts = delete_str.replace(' ', '').split(',')
        for part in parts:
            if part == 'last':
                pages_to_delete.add(max_pages)
            elif '-' in part:
                start, end = map(int, part.split('-'))
                pages_to_delete.update(range(start, end + 1))
            else:
                pages_to_delete.add(int(part))
        
        return pages_to_delete
    
    def update_progress(self, value):
        self.progress_bar['value'] = value
    
    def process_complete(self, output_path):
        self.progress_bar['value'] = 100
        self.status_label['text'] = "Processing complete!"
        self.process_btn['state'] = 'normal'
        
        if messagebox.askyesno("Success", 
                             "PDF created successfully!\nDo you want to open it?"):
            os.startfile(output_path)
    
    def show_error(self, error_message):
        self.status_label['text'] = "Error occurred!"
        self.process_btn['state'] = 'normal'
        messagebox.showerror("Error", f"An error occurred:\n{error_message}")

    def merge_pdfs(self, pdf_writer, pdf_files):
        for pdf_file in pdf_files:
            pdf_reader = PdfReader(pdf_file)
            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFMergerApp(root)
    root.mainloop()
