import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import webbrowser
import threading
import time

def select_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(0, file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, folder_path)

def reset_form():
    input_file_entry.delete(0, tk.END)
    output_folder_entry.delete(0, tk.END)
    base_name_entry.delete(0, tk.END)
    chunk_size_entry.delete(0, tk.END)
    progress_var.set(0)
    progress_label.config(text="")
    success_label.grid_remove()
    open_folder_button.grid_remove()
    split_another_button.grid_remove()
    progress_bar.grid_remove()

def open_output_folder():
    folder_path = output_folder_entry.get()
    if os.path.isdir(folder_path):
        webbrowser.open(folder_path)
    else:
        messagebox.showerror("Error", "Output folder not found.")

def start_split_thread():
    # Hide all elements except progress bar during processing
    for widget in frame.winfo_children():
        if widget != progress_bar and widget != progress_label:
            widget.grid_remove()
    
    # Show only progress bar and percentage
    progress_bar.grid(row=0, column=0, columnspan=3, pady=50)
    progress_label.grid(row=1, column=0, columnspan=3)
    progress_var.set(0)
    progress_label.config(text="0%")
    
    thread = threading.Thread(target=split_excel_file)
    thread.start()

def split_excel_file():
    try:
        input_file = input_file_entry.get()
        output_folder = output_folder_entry.get()
        base_name = base_name_entry.get()
        chunk_size = int(chunk_size_entry.get())

        if not os.path.isfile(input_file) or not os.path.isdir(output_folder):
            raise Exception("Invalid input file or output folder.")

        df = pd.read_excel(input_file, engine="openpyxl")
        total_rows = len(df)
        total_chunks = (total_rows + chunk_size - 1) // chunk_size

        for i in range(total_chunks):
            start_row = i * chunk_size
            end_row = min(start_row + chunk_size, total_rows)
            chunk = df.iloc[start_row:end_row]
            output_file = os.path.join(output_folder, f"{base_name}_P{i+1}.csv")
            chunk.to_csv(output_file, index=False)

            percent_done = int(((i + 1) / total_chunks) * 100)
            progress_var.set(percent_done)
            progress_label.config(text=f"{percent_done}%")
            root.update_idletasks()
            time.sleep(0.1)

        # Hide progress bar and show success message with buttons
        progress_bar.grid_remove()
        progress_label.grid_remove()
        
        success_label.grid(row=0, column=0, columnspan=3, pady=30)
        open_folder_button.grid(row=1, column=0, columnspan=3, pady=10)
        split_another_button.grid(row=2, column=0, columnspan=3, pady=5)
        
    except Exception as e:
        messagebox.showerror("❌ Error", str(e))
        show_main_interface()
    finally:
        pass

def show_main_interface():
    # Hide success elements
    success_label.grid_remove()
    open_folder_button.grid_remove()
    split_another_button.grid_remove()
    progress_bar.grid_remove()
    progress_label.grid_remove()
    
    # Show main interface
    excel_label.grid(row=0, column=0, sticky="w", pady=(0,5))
    input_file_entry.grid(row=1, column=0, columnspan=2, pady=5)
    browse_input_button.grid(row=1, column=2, padx=(5,0))
    
    folder_label.grid(row=2, column=0, sticky="w", pady=(10,5))
    output_folder_entry.grid(row=3, column=0, columnspan=2, pady=5)
    browse_output_button.grid(row=3, column=2, padx=(5,0))
    
    base_name_label.grid(row=4, column=0, sticky="w", pady=(10,5))
    base_name_entry.grid(row=5, column=0, columnspan=3, sticky="w", pady=5)
    
    chunk_size_label.grid(row=6, column=0, sticky="w", pady=(10,5))
    chunk_size_entry.grid(row=7, column=0, sticky="w", pady=5)
    
    split_button.grid(row=8, column=0, columnspan=3, pady=20)

# GUI SETUP
root = tk.Tk()
root.title("📊 Excel Splitter Pro")
root.geometry("550x450")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

# Configure modern style
style = ttk.Style()
style.theme_use('clam')
style.configure('Title.TLabel', font=('Segoe UI', 12, 'bold'))
style.configure('Modern.TButton', font=('Segoe UI', 10))

frame = ttk.Frame(root, padding=30)
frame.pack(fill="both", expand=True)

# Main interface elements
excel_label = ttk.Label(frame, text="📈 Excel File:", style='Title.TLabel')
input_file_entry = ttk.Entry(frame, width=45, font=('Segoe UI', 9))
browse_input_button = ttk.Button(frame, text="📂 Browse", command=select_input_file, style='Modern.TButton')

folder_label = ttk.Label(frame, text="📁 Output Folder:", style='Title.TLabel')
output_folder_entry = ttk.Entry(frame, width=45, font=('Segoe UI', 9))
browse_output_button = ttk.Button(frame, text="📂 Browse", command=select_output_folder, style='Modern.TButton')

base_name_label = ttk.Label(frame, text="🏷️ Output Base Name:", style='Title.TLabel')
base_name_entry = ttk.Entry(frame, width=30, font=('Segoe UI', 9))

chunk_size_label = ttk.Label(frame, text="📊 Chunk Size (rows per file):", style='Title.TLabel')
chunk_size_entry = ttk.Entry(frame, width=20, font=('Segoe UI', 9))

split_button = ttk.Button(frame, text="⚡ Split Excel File", command=start_split_thread, style='Modern.TButton')

# Progress elements (hidden initially)
progress_var = tk.IntVar()
progress_bar = ttk.Progressbar(frame, variable=progress_var, maximum=100, length=400, style='TProgressbar')
progress_label = ttk.Label(frame, text="", font=('Segoe UI', 14, 'bold'))

# Success elements (hidden initially)
success_label = ttk.Label(frame, text="✅ File splited successfully!", font=('Segoe UI', 16, 'bold'), foreground='green')
open_folder_button = ttk.Button(frame, text="📁 Open Output Folder", command=open_output_folder, style='Modern.TButton')
split_another_button = ttk.Button(frame, text="🔄 Split Another File", command=lambda: [reset_form(), show_main_interface()], style='Modern.TButton')

# Show main interface initially
show_main_interface()

root.mainloop()