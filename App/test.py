import os
import sys
import tkinter as tk
from tkinter import filedialog, Label, Button, messagebox
import pandas as pd
import numpy as np
import itertools
import xlwings as xw
from PIL import Image, ImageTk

class ExcelIntegrationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Yazaki Data Integration")
        self.root.geometry("400x500")
        
        # Setup logo
        self.setup_logo()
        
        # Create notebook/tabs
        self.create_tabs()
        
        # Initialize file paths
        self.mmsta_file_path = ""
        self.output_dir = ""
        self.ypp_cae_file_path = ""
        self.wire_list_file_path = ""

    def resource_path(self, relative_path):
        try:
            if getattr(sys, 'frozen', False):
                base_path = sys._MEIPASS
            else:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)
        except Exception as e:
            print(f"Error loading resource: {e}")
            return relative_path

    def setup_logo(self):
        try:
            logo_path = self.resource_path("../assets/yazaki_logo.png")
            img = Image.open(logo_path)
            img = img.resize((200, 70), Image.LANCZOS)
            logo = ImageTk.PhotoImage(img)
            self.root.logo = logo
            logo_label = Label(self.root, image=logo)
            logo_label.pack(pady=10)
        except:
            logo_label = Label(self.root, text="Yazaki Data Integration", font=("Arial", 14, "bold"))
            logo_label.pack(pady=10)

    def create_tabs(self):
        # Separator Tab
        separator_frame = tk.Frame(self.root)
        separator_frame.pack(padx=20, pady=10)
        
        Label(separator_frame, text="MMSTA Separator", font=("Arial", 12, "bold")).pack(pady=10)
        
        Button(separator_frame, text="Select MMSTA File", command=self.select_mmsta_file).pack(pady=5)
        self.mmsta_label = Label(separator_frame, text="MMSTA File: Not selected", wraplength=350)
        self.mmsta_label.pack()
        
        Button(separator_frame, text="Select Output Folder", command=self.select_output_dir).pack(pady=5)
        self.output_label = Label(separator_frame, text="Output Folder: Not selected", wraplength=350)
        self.output_label.pack()
        
        Button(separator_frame, text="Separate MMSTA", command=self.main_separator).pack(pady=10)

        # Integrator Tab
        integrator_frame = tk.Frame(self.root)
        integrator_frame.pack(padx=20, pady=10)
        
        Label(integrator_frame, text="Excel Integration", font=("Arial", 12, "bold")).pack(pady=10)
        
        Button(integrator_frame, text="Select YPP CAE File", command=self.select_ypp_cae_file).pack(pady=5)
        self.ypp_label = Label(integrator_frame, text="YPP CAE File: Not selected", wraplength=350)
        self.ypp_label.pack()
        
        Button(integrator_frame, text="Select MMSTA Sep File", command=self.select_mmsta_sep_file).pack(pady=5)
        self.mmsta_sep_label = Label(integrator_frame, text="MMSTA Sep File: Not selected", wraplength=350)
        self.mmsta_sep_label.pack()
        
        Button(integrator_frame, text="Select Wire List File", command=self.select_wire_list_file).pack(pady=5)
        self.wire_list_label = Label(integrator_frame, text="Wire List File: Not selected", wraplength=350)
        self.wire_list_label.pack()
        
        Button(integrator_frame, text="Integrate Files", command=self.main_integrator).pack(pady=10)

    def select_mmsta_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
        if file_path and file_path.lower().endswith((".xlsx", ".xls")):
            self.mmsta_file_path = file_path
            self.mmsta_label.config(text=f"MMSTA File: {os.path.basename(file_path)}")

    def select_output_dir(self):
        output_dir = filedialog.askdirectory()
        if output_dir:
            self.output_dir = output_dir
            self.output_label.config(text=f"Output Folder: {output_dir}")

    def select_ypp_cae_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
        if file_path and file_path.lower().endswith((".xlsx", ".xls")):
            self.ypp_cae_file_path = file_path
            self.ypp_label.config(text=f"YPP CAE File: {os.path.basename(file_path)}")

    def select_mmsta_sep_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
        if file_path and file_path.lower().endswith((".xlsx", ".xls")):
            self.mmsta_sep_file_path = file_path
            self.mmsta_sep_label.config(text=f"MMSTA Sep File: {os.path.basename(file_path)}")

    def select_wire_list_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls;*.XLSX")])
        if file_path and file_path.lower().endswith((".xlsx", ".xls")):
            self.wire_list_file_path = file_path
            self.wire_list_label.config(text=f"Wire List File: {os.path.basename(file_path)}")

    def main_separator(self):
        if not self.mmsta_file_path or not self.output_dir:
            messagebox.showerror("Error", "Please select MMSTA file and output folder")
            return

        try:
            output_file = os.path.join(self.output_dir, "output.xlsx")
            separator(self.mmsta_file_path, output_file)
            messagebox.showinfo("Success", f"Files separated successfully!\nOutput: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error during separation: {str(e)}")

    def main_integrator(self):
        if not all([self.ypp_cae_file_path, self.mmsta_sep_file_path, self.wire_list_file_path, self.output_dir]):
            messagebox.showerror("Error", "Please select all required files and output folder")
            return

        try:
            output_file = os.path.join(self.output_dir, "integrated_output.xlsx")
            integrator(
                self.mmsta_sep_file_path, 
                self.wire_list_file_path, 
                self.ypp_cae_file_path, 
                output_file
            )
            messagebox.showinfo("Success", f"Files integrated successfully!\nOutput: {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Error during integration: {str(e)}")

# Add your existing functions from the original script here
# (separator, integrator, Read_wire_list, processing_function, etc.)

def main():
    root = tk.Tk()
    app = ExcelIntegrationApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()