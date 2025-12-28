"""
Simple GUI for USGS Table Extractor.
"""

import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from pathlib import Path
import sys
import traceback

# Add parent directory to path so we can import from src
sys.path.insert(0, str(Path(__file__).parent))

from src.docx_extractor import DocxTableExtractor


class ExtractorGUI:
    """Simple GUI for table extraction."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("USGS Table Extractor")
        self.root.geometry("700x500")
        self.root.configure(bg='#f0f0f0')
        
        # Header
        header_frame = tk.Frame(root, bg='#2c3e50', height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title = tk.Label(header_frame, 
                        text="USGS Table Extractor", 
                        font=("Arial", 18, "bold"),
                        bg='#2c3e50',
                        fg='white')
        title.pack(pady=15)
        
        subtitle = tk.Label(header_frame,
                           text="PDF/Word → Excel Table Extraction",
                           font=("Arial", 10),
                           bg='#2c3e50',
                           fg='#ecf0f1')
        subtitle.pack()
        
        # Main content
        content_frame = tk.Frame(root, bg='#f0f0f0')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Instructions
        instructions = tk.Label(content_frame,
                               text="Select a single file or an entire folder to process:",
                               font=("Arial", 10),
                               bg='#f0f0f0')
        instructions.pack(pady=(0, 10))
        
        # Buttons
        btn_frame = tk.Frame(content_frame, bg='#f0f0f0')
        btn_frame.pack(pady=10)
        
        self.btn_file = tk.Button(btn_frame, 
                                  text="📄 Select File", 
                                  command=self.select_file,
                                  width=18, 
                                  height=2,
                                  font=("Arial", 10, "bold"),
                                  bg='#3498db',
                                  fg='white',
                                  cursor='hand2')
        self.btn_file.grid(row=0, column=0, padx=10)
        
        self.btn_folder = tk.Button(btn_frame, 
                                    text="📁 Select Folder", 
                                    command=self.select_folder,
                                    width=18, 
                                    height=2,
                                    font=("Arial", 10, "bold"),
                                    bg='#2ecc71',
                                    fg='white',
                                    cursor='hand2')
        self.btn_folder.grid(row=0, column=1, padx=10)
        
        # Log section
        log_frame = tk.Frame(content_frame, bg='#f0f0f0')
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        log_label = tk.Label(log_frame, 
                            text="Processing Log:",
                            font=("Arial", 10, "bold"),
                            bg='#f0f0f0')
        log_label.pack(anchor=tk.W)
        
        self.log = scrolledtext.ScrolledText(log_frame, 
                                             height=15, 
                                             width=80,
                                             font=("Consolas", 9),
                                             bg='#2c3e50',
                                             fg='#ecf0f1')
        self.log.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        
        # Status bar
        self.status_bar = tk.Label(root, 
                                   text="Ready",
                                   font=("Arial", 9),
                                   bg='#34495e',
                                   fg='white',
                                   anchor=tk.W,
                                   padx=10)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Initialize
        self.extractor = None
        self.processing = False
        
        self.write_log("="*70)
        self.write_log("USGS Table Extractor - Ready")
        self.write_log("="*70)
        self.write_log("\nSelect a file or folder to begin processing.\n")
    
    def write_log(self, message):
        """Write to log window."""
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.root.update()
    
    def update_status(self, message):
        """Update status bar."""
        self.status_bar.config(text=message)
        self.root.update()
    
    def disable_buttons(self):
        """Disable buttons during processing."""
        self.btn_file.config(state=tk.DISABLED)
        self.btn_folder.config(state=tk.DISABLED)
        self.processing = True
    
    def enable_buttons(self):
        """Enable buttons after processing."""
        self.btn_file.config(state=tk.NORMAL)
        self.btn_folder.config(state=tk.NORMAL)
        self.processing = False
    
    def initialize_extractor(self):
        """Initialize the extractor."""
        if self.extractor is None:
            self.write_log("Initializing extractor...")
            self.update_status("Initializing...")
            try:
                self.extractor = DocxTableExtractor(
                    clean_data=True,
                    auto_convert_pdf=True
                )
                self.write_log("✅ Extractor initialized\n")
            except Exception as e:
                self.write_log(f"❌ Failed to initialize: {e}\n")
                messagebox.showerror("Error", f"Failed to initialize:\n{e}")
                return False
        return True
    
    def select_file(self):
        """Process single file."""
        if self.processing:
            return
        
        file_path = filedialog.askopenfilename(
            title="Select PDF or Word file",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        if not self.initialize_extractor():
            return
        
        self.disable_buttons()
        
        self.write_log("\n" + "="*70)
        self.write_log(f"Processing: {Path(file_path).name}")
        self.write_log("="*70)
        
        try:
            input_file = Path(file_path)
            output_dir = input_file.parent / "output"
            output_dir.mkdir(exist_ok=True)
            output_file = output_dir / f"{input_file.stem}_output.xlsx"
            
            self.update_status(f"Processing {input_file.name}...")
            self.write_log(f"\nInput:  {input_file}")
            self.write_log(f"Output: {output_file}\n")
            
            self.extractor.process_file(input_file, output_file)
            
            self.write_log(f"\n✅ SUCCESS!")
            self.write_log(f"Saved to: {output_file}")
            self.write_log("="*70 + "\n")
            
            self.update_status("Ready")
            
            result = messagebox.askyesno(
                "Success", 
                f"File processed!\n\nOutput: {output_file.name}\n\nOpen folder?",
                icon='info'
            )
            
            if result:
                import os
                os.startfile(output_dir)
            
        except Exception as e:
            self.write_log(f"\n❌ ERROR: {str(e)}")
            self.write_log(traceback.format_exc())
            self.update_status("Error")
            messagebox.showerror("Error", f"Failed:\n\n{str(e)}")
        
        finally:
            self.enable_buttons()
    
    def select_folder(self):
        """Batch process folder."""
        if self.processing:
            return
        
        folder_path = filedialog.askdirectory(
            title="Select folder with PDF/Word files"
        )
        
        if not folder_path:
            return
        
        if not self.initialize_extractor():
            return
        
        self.disable_buttons()
        
        self.write_log("\n" + "="*70)
        self.write_log(f"Batch Processing: {Path(folder_path).name}")
        self.write_log("="*70)
        
        try:
            input_dir = Path(folder_path)
            output_dir = input_dir / "output"
            output_dir.mkdir(exist_ok=True)
            
            pdf_files = list(input_dir.glob("*.pdf"))
            docx_files = list(input_dir.glob("*.docx"))
            all_files = pdf_files + docx_files
            
            if not all_files:
                self.write_log("\n⚠️ No PDF or Word files found")
                messagebox.showwarning("No Files", "No files found in folder.")
                self.enable_buttons()
                return
            
            self.write_log(f"\nFound {len(all_files)} file(s):\n")
            for f in all_files:
                self.write_log(f"  • {f.name}")
            self.write_log("")
            
            success_count = 0
            error_count = 0
            
            for i, file_path in enumerate(all_files, 1):
                self.update_status(f"Processing {i}/{len(all_files)}: {file_path.name}")
                self.write_log(f"[{i}/{len(all_files)}] {file_path.name}... ")
                
                try:
                    output_file = output_dir / f"{file_path.stem}_output.xlsx"
                    self.extractor.process_file(file_path, output_file)
                    self.write_log("✅")
                    success_count += 1
                except Exception as e:
                    self.write_log(f"❌ {str(e)}")
                    error_count += 1
            
            self.write_log("\n" + "="*70)
            self.write_log("BATCH COMPLETE")
            self.write_log("="*70)
            self.write_log(f"✅ Success: {success_count}")
            self.write_log(f"❌ Failed:  {error_count}")
            self.write_log(f"📁 Output: {output_dir}")
            self.write_log("="*70 + "\n")
            
            self.update_status("Ready")
            
            result = messagebox.askyesno(
                "Batch Complete",
                f"Processing complete!\n\n✅ Success: {success_count}\n❌ Failed: {error_count}\n\nOpen folder?",
                icon='info'
            )
            
            if result:
                import os
                os.startfile(output_dir)
            
        except Exception as e:
            self.write_log(f"\n❌ ERROR: {str(e)}")
            self.write_log(traceback.format_exc())
            messagebox.showerror("Error", f"Failed:\n\n{str(e)}")
        
        finally:
            self.enable_buttons()


def main():
    """Run GUI."""
    root = tk.Tk()
    app = ExtractorGUI(root)
    
    # Center window
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()