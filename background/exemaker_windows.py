import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import sys
import os

class RLCBuilder(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RLC PyInstaller Builder v2.1")
        self.geometry("750x750")

        self.script_path = tk.StringVar()
        self.exe_name = tk.StringVar(value="RocketLaunchCountdown")
        self.hide_console = tk.BooleanVar(value=True)
        self.additional_files = [] # For folders/assets
        self.supporting_scripts = [] # For secondary .py files

        # --- Step 1: Main Script ---
        step1 = tk.LabelFrame(self, text="Step 1: Main Script (Entry Point)")
        step1.pack(fill="x", padx=10, pady=5)
        tk.Entry(step1, textvariable=self.script_path, width=60).pack(side="left", padx=5, pady=5)
        tk.Button(step1, text="Browse", command=self.browse_script).pack(side="left", padx=5)

        # --- Step 2: Supporting Scripts ---
        # These are scripts like path_handler.py that the main script imports
        step2 = tk.LabelFrame(self, text="Step 2: Supporting Logic Scripts (Imported .py files)")
        step2.pack(fill="x", padx=10, pady=5)
        self.scripts_listbox = tk.Listbox(step2, height=3)
        self.scripts_listbox.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        s_btn_frame = tk.Frame(step2)
        s_btn_frame.pack(side="right")
        tk.Button(s_btn_frame, text="Add Script", command=self.add_script).pack(fill="x", padx=5)
        tk.Button(s_btn_frame, text="Clear", command=self.clear_scripts).pack(fill="x", padx=5)

        # --- Step 3: Include Folders ---
        step3 = tk.LabelFrame(self, text="Step 3: Include External Folders (Popouts, Database)")
        step3.pack(fill="x", padx=10, pady=5)
        self.files_listbox = tk.Listbox(step3, height=3)
        self.files_listbox.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        btn_frame = tk.Frame(step3)
        btn_frame.pack(side="right")
        tk.Button(btn_frame, text="Add Folder", command=self.add_folder).pack(fill="x", padx=5)
        tk.Button(btn_frame, text="Clear", command=self.clear_folders).pack(fill="x", padx=5)

        # --- Step 4: Options ---
        step4 = tk.LabelFrame(self, text="Step 4: Build Options")
        step4.pack(fill="x", padx=10, pady=5)
        tk.Entry(step4, textvariable=self.exe_name, width=30).pack(side="left", padx=5, pady=5)
        tk.Checkbutton(step4, text="Hide Console Window", 
                       variable=self.hide_console).pack(side="left", padx=5, pady=5)

        tk.Button(self, text="BUILD ONE-FILE EXE", bg="#2e7d32", fg="white", font=("Arial", 10, "bold"),
                  command=self.start_build).pack(pady=10)

        self.log = scrolledtext.ScrolledText(self, height=12, bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    def browse_script(self):
        path = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
        if path: self.script_path.set(path)

    def add_script(self):
        paths = filedialog.askopenfilenames(filetypes=[("Python files", "*.py")])
        for path in paths:
            self.supporting_scripts.append(path)
            self.scripts_listbox.insert(tk.END, os.path.basename(path))

    def clear_scripts(self):
        self.supporting_scripts.clear()
        self.scripts_listbox.delete(0, tk.END)

    def add_folder(self):
        path = filedialog.askdirectory(title="Select Folder to Bundle")
        if path:
            folder_name = os.path.basename(path)
            # Use a dot for the destination to keep it relative to root
            self.additional_files.append(f"{path}{os.pathsep}{folder_name}")
            self.files_listbox.insert(tk.END, f"Folder: {folder_name}")

    def clear_folders(self):
        self.additional_files.clear()
        self.files_listbox.delete(0, tk.END)

    def log_write(self, text):
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def start_build(self):
        script = self.script_path.get()
        if not script or not os.path.isfile(script):
            messagebox.showerror("Error", "Please select your main RLC script.")
            return
        self.log.delete("1.0", tk.END)
        threading.Thread(target=self.build_exe, args=(script,), daemon=True).start()

    def build_exe(self, script):
        output_dir = os.path.join(os.path.dirname(script), "dist")
        
        cmd = [
            "py",
            "-3.14",
            "-m",
            "PyInstaller",
            "--onefile",
            "--name", self.exe_name.get(),
            "--distpath", output_dir,
        ]

        # Handle Console
        if self.hide_console.get():
            cmd.append("--noconsole")
        else:
            cmd.append("--console")

        # Add logic scripts via search paths
        for s in self.supporting_scripts:
            # We add the directory of the supporting script to the search path
            script_dir = os.path.dirname(s)
            cmd.extend(["--paths", script_dir])

        # Add asset/popout folders
        for item in self.additional_files:
            cmd.extend(["--add-data", item])

        cmd.append(script)

        self.log_write(f"BUILD STARTING...\nCommand: {' '.join(cmd)}\n\n")

        try:
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.log_write(line)
            process.wait()
            if process.returncode == 0:
                messagebox.showinfo("Success", f"EXE saved in 'dist' folder.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    RLCBuilder().mainloop()