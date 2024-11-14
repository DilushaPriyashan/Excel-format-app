import tkinter as tk
from tkinter import filedialog, ttk
import shutil
import os
import subprocess
import threading
import time

def upload_excel_file():
    excel_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    if excel_path:
        excel_path_label.config(text=f"Excel File: {excel_path}")
        
        new_excel_path = os.path.join(os.getcwd(), "input.xlsx")
        

        if os.path.exists(new_excel_path):
            os.remove(new_excel_path)

     
        shutil.copy(excel_path, new_excel_path)
        print("Excel file saved as:", new_excel_path)
        check_files_uploaded()

def upload_text_file():
    text_path = filedialog.askopenfilename(
        title="Select Text File",
        filetypes=[("Text Files", "*.txt")]
    )
    if text_path:
      
        text_path_label.config(text=f"Text File: {text_path}")
        
        new_text_path = os.path.join(os.getcwd(), "values.txt")
        
     
        if os.path.exists(new_text_path):
            os.remove(new_text_path)

      
        shutil.copy(text_path, new_text_path)
        print("Text file saved as:", new_text_path)
        check_files_uploaded()

def check_files_uploaded():
   
    if (excel_path_label.cget("text") != "Excel File: None" and 
            text_path_label.cget("text") != "Text File: None"):
        sort_button.config(state=tk.NORMAL)

def run_sorting_script():
   
    progress_bar.start()
    
    
    result = subprocess.run(['python', 'main.py'], capture_output=True, text=True)
    
    
    if result.returncode == 0:
        print("main.py executed successfully.")
    else:
        print("Error executing main.py:")
        print(result.stderr)

def sort_files():
    
    sort_button.config(state=tk.DISABLED)
    
    
    threading.Thread(target=execute_sorting).start()

def execute_sorting():
    run_sorting_script()
    
  
    for _ in range(5):
        time.sleep(1)
        
   
    progress_bar.stop()
    
  
    sorting_status_label.config(text="File sorted! Please select a path to save.")
    
   
    save_button.config(state=tk.NORMAL)

def save_sorted_file():
    
    existing_sorted_file_path = os.path.join(os.getcwd(), "sorted_3.xlsx")
    

    save_path = filedialog.asksaveasfilename(
        title="Save Sorted File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if save_path:
        
        shutil.copy(existing_sorted_file_path, save_path)
        
        print(f"Sorted file copied to: {save_path}")
     
        sorting_status_label.config(text=f"Sorted file copied to: {save_path}")

root = tk.Tk()
root.title("File Uploader")
root.geometry("500x400")  


label = tk.Label(root, text="Upload Excel and Text files:")
label.pack(pady=10)


excel_path_label = tk.Label(root, text="Excel File: None")
excel_path_label.pack(pady=5)


upload_excel_button = tk.Button(root, text="Upload Excel", command=upload_excel_file)
upload_excel_button.pack(pady=5)


text_path_label = tk.Label(root, text="Text File: None")
text_path_label.pack(pady=5)


upload_text_button = tk.Button(root, text="Upload Text", command=upload_text_file)
upload_text_button.pack(pady=5)


sort_button = tk.Button(root, text="Sort", command=sort_files, state=tk.DISABLED)
sort_button.pack(pady=10)


progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="indeterminate")
progress_bar.pack(pady=10)


sorting_status_label = tk.Label(root, text="")
sorting_status_label.pack(pady=5)


save_button = tk.Button(root, text="Save Sorted File", command=save_sorted_file, state=tk.DISABLED)
save_button.pack(pady=10)


close_button = tk.Button(root, text="Close", command=root.quit)
close_button.pack(pady=10)

root.mainloop()
