import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pypandoc
import zipfile
import threading
import queue
import time
from PIL import Image, ImageTk

class MarkdownPreviewWindow:
    def __init__(self, master, md_content):
        self.window = tk.Toplevel(master)
        self.window.title("Markdown Preview")
        self.window.geometry("700x500")
        
        self.text_area = scrolledtext.ScrolledText(self.window, wrap=tk.WORD, font=("Consolas", 11))
        self.text_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.text_area.insert(tk.END, md_content)
        self.text_area.config(state=tk.DISABLED)

def count_files_and_folders(folder_path, recursive):
    if recursive:
        num_folders = sum([len(dirs) for _, dirs, _ in os.walk(folder_path)])
        num_files = sum([len(files) for _, _, files in os.walk(folder_path)])
    else:
        num_folders = len([f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))])
        num_files = len([f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))])
    return num_folders, num_files

def process_md_file(md_path):
    # Replace {.underline} with <u>...</u>
    with open(md_path, "r", encoding="utf-8") as file:
        content = file.read()
    updated_content = content.replace("{.underline}", "")
    updated_content = updated_content.replace("[", "<u>").replace("]", "</u>")
    with open(md_path, "w", encoding="utf-8") as file:
        file.write(updated_content)
    return updated_content

def convert_docx_to_md(file_path, output_folder, media_extraction, progress_queue):
    try:
        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)
        
        # Get the file name without path
        file_name = os.path.basename(file_path)
        file_name_no_ext = os.path.splitext(file_name)[0]
        
        # Define output paths
        md_path = os.path.join(output_folder, file_name_no_ext + '.md')
        zip_path = os.path.join(output_folder, file_name_no_ext + '.zip')
        media_folder = os.path.join(output_folder, "media_" + file_name_no_ext)
        
        # Update progress
        progress_queue.put(("status", f"Converting {file_name}..."))
        progress_queue.put(("progress", 20))
        
        # Create media folder if needed
        if media_extraction:
            os.makedirs(media_folder, exist_ok=True)
            extra_args = ['--wrap=none', f'--extract-media={media_folder}']
        else:
            extra_args = ['--wrap=none']
        
        # Convert file
        pypandoc.convert_file(file_path, 'md', format='docx', 
                              outputfile=md_path, extra_args=extra_args)
        
        progress_queue.put(("progress", 50))
        progress_queue.put(("status", f"Processing {file_name}..."))
        
        # Process the Markdown file
        md_content = process_md_file(md_path)
        
        progress_queue.put(("progress", 70))
        progress_queue.put(("status", f"Creating ZIP for {file_name}..."))
        
        # Create ZIP
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            zipf.write(md_path, os.path.basename(md_path))
            if os.path.exists(media_folder):
                for root, dirs, files in os.walk(media_folder):
                    for file in files:
                        file_path_to_zip = os.path.join(root, file)
                        arcname = os.path.relpath(file_path_to_zip, output_folder)
                        zipf.write(file_path_to_zip, arcname)
        
        progress_queue.put(("progress", 90))
        progress_queue.put(("md_preview", md_content))
        progress_queue.put(("file_done", file_path))
        
        return md_content
    
    except Exception as e:
        progress_queue.put(("error", f"Error processing {file_path}: {str(e)}"))
        return None

class DocxToMarkdownConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Enhanced DOCX to Markdown Converter")
        self.root.geometry("600x500")
        
        # Create a queue for thread communication
        self.progress_queue = queue.Queue()
        
        # Create frames for organization
        self.input_frame = ttk.LabelFrame(root, text="Input Options")
        self.input_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        self.output_frame = ttk.LabelFrame(root, text="Output Options")
        self.output_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        self.options_frame = ttk.LabelFrame(root, text="Conversion Options")
        self.options_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        self.progress_frame = ttk.LabelFrame(root, text="Progress")
        self.progress_frame.pack(fill="x", expand=False, padx=10, pady=5)
        
        self.button_frame = ttk.Frame(root)
        self.button_frame.pack(fill="x", expand=False, padx=10, pady=10)
        
        self.files_frame = ttk.LabelFrame(root, text="Selected Files")
        self.files_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Input widgets
        ttk.Label(self.input_frame, text="Select DOCX Files:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.input_mode = tk.StringVar(value="files")
        ttk.Radiobutton(self.input_frame, text="Files", variable=self.input_mode, value="files").grid(row=0, column=1, padx=5, pady=5)
        ttk.Radiobutton(self.input_frame, text="Folder", variable=self.input_mode, value="folder").grid(row=0, column=2, padx=5, pady=5)
        
        self.browse_button = ttk.Button(self.input_frame, text="Browse", command=self.browse_input)
        self.browse_button.grid(row=0, column=3, padx=5, pady=5)
        
        # Files list
        self.files_listbox = tk.Listbox(self.files_frame, width=70, height=10)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add scrollbar to listbox
        files_scrollbar = ttk.Scrollbar(self.files_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        files_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.config(yscrollcommand=files_scrollbar.set)
        
        # Output widgets
        ttk.Label(self.output_frame, text="Output Folder:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.output_entry = ttk.Entry(self.output_frame, width=50)
        self.output_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        self.output_button = ttk.Button(self.output_frame, text="Browse", command=self.browse_output)
        self.output_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Options widgets
        self.media_var = tk.BooleanVar(value=True)
        self.media_checkbox = ttk.Checkbutton(self.options_frame, text="Extract media", variable=self.media_var)
        self.media_checkbox.pack(padx=5, pady=5, anchor="w")
        
        self.preview_var = tk.BooleanVar(value=True)
        self.preview_checkbox = ttk.Checkbutton(self.options_frame, text="Show preview after conversion", variable=self.preview_var)
        self.preview_checkbox.pack(padx=5, pady=5, anchor="w")
        
        # Progress widgets
        self.status_label = ttk.Label(self.progress_frame, text="Ready")
        self.status_label.pack(padx=5, pady=2, anchor="w")
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, length=580, mode='determinate')
        self.progress_bar.pack(padx=5, pady=5, fill="x")
        
        # Buttons
        self.convert_button = ttk.Button(self.button_frame, text="Convert", command=self.start_conversion)
        self.convert_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        self.clear_button = ttk.Button(self.button_frame, text="Clear All", command=self.clear_all)
        self.clear_button.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Initialize variables
        self.files_to_convert = []
        self.conversion_in_progress = False
        self.md_content = None
        
        # Configure grid expansions
        self.input_frame.columnconfigure(1, weight=1)
        self.output_frame.columnconfigure(1, weight=1)
        
        # Start checking queue
        self.check_queue()
    
    def browse_input(self):
        if self.input_mode.get() == "files":
            files = filedialog.askopenfilenames(filetypes=[("Word Documents", "*.docx")])
            if files:
                self.files_to_convert = list(files)
                self.update_files_listbox()
        else:  # folder mode
            folder = filedialog.askdirectory()
            if folder:
                self.files_to_convert = [
                    os.path.join(folder, f) for f in os.listdir(folder) 
                    if f.lower().endswith('.docx') and os.path.isfile(os.path.join(folder, f))
                ]
                self.update_files_listbox()
    
    def browse_output(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, folder)
    
    def update_files_listbox(self):
        self.files_listbox.delete(0, tk.END)
        for file in self.files_to_convert:
            self.files_listbox.insert(tk.END, os.path.basename(file))
    
    def clear_all(self):
        self.files_to_convert = []
        self.files_listbox.delete(0, tk.END)
        self.output_entry.delete(0, tk.END)
        self.progress_bar['value'] = 0
        self.status_label.config(text="Ready")
    
    def start_conversion(self):
        if not self.files_to_convert:
            messagebox.showwarning("Warning", "Please select at least one DOCX file.")
            return
        
        output_folder = self.output_entry.get()
        if not output_folder:
            output_folder = os.path.dirname(self.files_to_convert[0])
            self.output_entry.insert(0, output_folder)
        
        # Disable UI during conversion
        self.toggle_ui_state(False)
        
        # Start conversion thread
        self.conversion_thread = threading.Thread(
            target=self.run_conversion,
            args=(self.files_to_convert, output_folder, self.media_var.get())
        )
        self.conversion_thread.daemon = True
        self.conversion_thread.start()
    
    def run_conversion(self, files, output_folder, media_extraction):
        total_files = len(files)
        completed_files = 0
        
        for file_path in files:
            # Reset progress for each file
            self.progress_queue.put(("progress", 0))
            self.progress_queue.put(("status", f"Processing file {completed_files + 1} of {total_files}"))
            
            # Convert the file
            md_content = convert_docx_to_md(file_path, output_folder, media_extraction, self.progress_queue)
            
            # Update completed count
            completed_files += 1
            
            # Calculate overall progress
            overall_progress = int((completed_files / total_files) * 100)
            self.progress_queue.put(("overall_progress", overall_progress))
            
            # Update status
            self.progress_queue.put(("status", f"Completed {completed_files} of {total_files} files"))
        
        # All files processed
        self.progress_queue.put(("complete", f"Converted {completed_files} files successfully!"))
    
    def toggle_ui_state(self, enabled):
        state = "normal" if enabled else "disabled"
        self.browse_button.config(state=state)
        self.output_button.config(state=state)
        self.convert_button.config(state=state)
        self.clear_button.config(state=state)
        self.conversion_in_progress = not enabled
    
    def check_queue(self):
        try:
            while not self.progress_queue.empty():
                message = self.progress_queue.get(0)
                
                # Process message based on its type
                if message[0] == "status":
                    self.status_label.config(text=message[1])
                
                elif message[0] == "progress":
                    self.progress_bar['value'] = message[1]
                
                elif message[0] == "overall_progress":
                    self.progress_bar['value'] = message[1]
                
                elif message[0] == "md_preview":
                    self.md_content = message[1]
                
                elif message[0] == "file_done":
                    # If preview is enabled, show the preview for the last converted file
                    if self.preview_var.get() and self.md_content:
                        MarkdownPreviewWindow(self.root, self.md_content)
                
                elif message[0] == "complete":
                    messagebox.showinfo("Conversion Complete", message[1])
                    self.toggle_ui_state(True)
                
                elif message[0] == "error":
                    messagebox.showerror("Error", message[1])
        
        except queue.Empty:
            pass
        finally:
            # Schedule next check
            self.root.after(100, self.check_queue)

def main():
    root = tk.Tk()
    app = DocxToMarkdownConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()