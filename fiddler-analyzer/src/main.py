import subprocess
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def update_saz_path(saz_path):
    """Update the saz file path in index.py"""
    index_path = os.path.join(os.path.dirname(__file__), 'index.py')
    
    with open(index_path, 'r') as file:
        content = file.read()
        
    updated_content = content.replace(
        'saz_FILE_PATH = r"..."', 
        f'saz_FILE_PATH = r"{saz_path}"'
    )
    
    with open(index_path, 'w') as file:
        file.write(updated_content)

def run_script(script_name, file_path=None):
    """Run a Python script and return success status"""
    print(f"\nRunning {script_name}...")
    cmd = [sys.executable, script_name]
    if file_path:
        cmd.append(file_path)
    result = subprocess.run(
        cmd,
        cwd=os.path.dirname(os.path.abspath(__file__))
    )
    return result.returncode == 0

def main():
    root = tk.Tk()
    root.title("Fiddler Tool")
    root.geometry("400x200")
    
    label = tk.Label(root, text="Please select your .saz file to begin analysis:")
    label.pack(pady=20)
    
    def select_file():
        file_path = filedialog.askopenfilename(
            title="Select SAZ File",
            filetypes=[
                ("saz files", "*.saz"),
                ("All files", "*.*")  
            ],
            defaultextension=".saz"
        )
        if file_path:
            root.destroy()
            try:
                update_saz_path(file_path)
                
                scripts = [
                    'index.py',
                    'ms_docs_search.py',
                    'arin_search.py'
                ]
                
                for script in scripts:
                    if not run_script(script, file_path):
                        messagebox.showerror(
                            "Error",
                            f"Failed to run {script}"
                        )
                        sys.exit(1)
                
                messagebox.showinfo(
                    "Success",
                    "Analysis completed successfully!\nResults saved to Excel file."
                )
                
            except Exception as e:
                messagebox.showerror(
                    "Error",
                    f"An error occurred: {str(e)}"
                )
                sys.exit(1)
    
    select_button = tk.Button(
        root,
        text="Select saz File",
        command=select_file
    )
    select_button.pack(pady=20)
    
    root.mainloop()

if __name__ == "__main__":
    main()