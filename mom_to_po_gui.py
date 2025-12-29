"""
MOM to PO Generator - GUI Version
Complete workflow: MOM PDF → Analysis → Extraction → HTML → Word
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import json
from pathlib import Path
from word_generator import WordGenerator


class MOMtoPOApp:
    """GUI Application for MOM to PO conversion"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("MOM to PO Generator v3.0")
        self.root.geometry("800x600")
        
        # Variables
        self.template_path = tk.StringVar()
        self.mom_pdf_path = tk.StringVar()
        self.step3_json_path = tk.StringVar()
        self.output_dir = tk.StringVar(value="./outputs")
        
        # Setup UI
        self._setup_ui()
    
    def _setup_ui(self):
        """Setup GUI layout"""
        
        # Title
        title_label = ttk.Label(
            self.root, 
            text="MOM to PO Generator", 
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=20)
        
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Step 3 JSON input
        ttk.Label(main_frame, text="Step 3 HTML JSON File:").grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.step3_json_path, width=50).grid(
            row=0, column=1, pady=5, padx=5
        )
        ttk.Button(main_frame, text="Browse", command=self._browse_step3_json).grid(
            row=0, column=2, pady=5
        )
        
        # Word template (optional)
        ttk.Label(main_frame, text="Word Template (Optional):").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.template_path, width=50).grid(
            row=1, column=1, pady=5, padx=5
        )
        ttk.Button(main_frame, text="Browse", command=self._browse_template).grid(
            row=1, column=2, pady=5
        )
        
        # Output directory
        ttk.Label(main_frame, text="Output Directory:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.output_dir, width=50).grid(
            row=2, column=1, pady=5, padx=5
        )
        ttk.Button(main_frame, text="Browse", command=self._browse_output_dir).grid(
            row=2, column=2, pady=5
        )
        
        # Generate button
        generate_btn = ttk.Button(
            main_frame, 
            text="Generate Word Document", 
            command=self._generate_word,
            style="Accent.TButton"
        )
        generate_btn.grid(row=3, column=0, columnspan=3, pady=30)
        
        # Status log
        ttk.Label(main_frame, text="Log:").grid(
            row=4, column=0, sticky=tk.NW, pady=5
        )
        
        self.log_text = tk.Text(main_frame, height=15, width=70)
        self.log_text.grid(row=5, column=0, columnspan=3, pady=5)
        
        scrollbar = ttk.Scrollbar(main_frame, command=self.log_text.yview)
        scrollbar.grid(row=5, column=3, sticky='ns')
        self.log_text['yscrollcommand'] = scrollbar.set
    
    def _browse_step3_json(self):
        """Browse for Step 3 JSON file"""
        filename = filedialog.askopenfilename(
            title="Select Step 3 HTML JSON File",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if filename:
            self.step3_json_path.set(filename)
    
    def _browse_template(self):
        """Browse for Word template"""
        filename = filedialog.askopenfilename(
            title="Select Word Template",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if filename:
            self.template_path.set(filename)
    
    def _browse_output_dir(self):
        """Browse for output directory"""
        dirname = filedialog.askdirectory(title="Select Output Directory")
        if dirname:
            self.output_dir.set(dirname)
    
    def _log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def _generate_word(self):
        """Generate Word document from Step 3 JSON"""
        
        # Validate inputs
        if not self.step3_json_path.get():
            messagebox.showerror("Error", "Please select Step 3 HTML JSON file")
            return
        
        if not os.path.exists(self.step3_json_path.get()):
            messagebox.showerror("Error", "Step 3 JSON file not found")
            return
        
        # Create output directory
        os.makedirs(self.output_dir.get(), exist_ok=True)
        
        try:
            self._log("=" * 50)
            self._log("Starting Word document generation...")
            
            # Load JSON to get PO number for filename
            with open(self.step3_json_path.get(), 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            po_no = data.get('html_data', {}).get('PO_NO', 'Unknown')
            output_filename = f"PO_{po_no}.docx"
            output_path = os.path.join(self.output_dir.get(), output_filename)
            
            self._log(f"PO Number: {po_no}")
            self._log(f"Output: {output_path}")
            
            # Generate Word document
            generator = WordGenerator(self.template_path.get() if self.template_path.get() else None)
            
            if self.template_path.get() and os.path.exists(self.template_path.get()):
                self._log("Using template mode...")
                generator.generate_with_template(
                    self.template_path.get(),
                    self.step3_json_path.get(),
                    output_path
                )
            else:
                self._log("Using default mode (no template)...")
                generator.generate_from_html_json(
                    self.step3_json_path.get(),
                    output_path
                )
            
            self._log("✅ Word document generated successfully!")
            self._log(f"📄 File: {output_path}")
            
            # Ask to open file
            if messagebox.askyesno("Success", "Word document generated!\n\nOpen file?"):
                os.startfile(output_path)
        
        except Exception as e:
            self._log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to generate Word document:\n{str(e)}")

    
# Run application
if __name__ == "__main__":
    root = tk.Tk()
    app = MOMtoPOApp(root)
    root.mainloop()