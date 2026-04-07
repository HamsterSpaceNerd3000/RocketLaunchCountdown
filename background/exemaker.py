import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import sys
import os

class RLCBuilder(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RLC PyInstaller Builder v2.0")
        self.geometry("750x650")

        self.script_path = tk.StringVar()
        self.exe_name = tk.StringVar(value="RocketLaunchCountdown")
        self.hide_console = tk.BooleanVar(value=True) # Default to hidden
        self.additional_files = []

        # --- Step 1: Main Script ---
        step1 = tk.LabelFrame(self, text="Step 1: Main Script")
        step1.pack(fill="x", padx=10, pady=5)
        tk.Entry(step1, textvariable=self.script_path, width=60).pack(side="left", padx=5, pady=5)
        tk.Button(step1, text="Browse", command=self.browse_script).pack(side="left", padx=5)

        # --- Step 2: Output Name ---
        step2 = tk.LabelFrame(self, text="Step 2: Output Name")
        step2.pack(fill="x", padx=10, pady=5)
        tk.Entry(step2, textvariable=self.exe_name, width=60).pack(side="left", padx=5, pady=5)

        # --- Step 3: Include Popouts/Assets ---
        step3 = tk.LabelFrame(self, text="Step 3: Include Folders (e.g., Popouts)")
        step3.pack(fill="x", padx=10, pady=5)
        self.files_listbox = tk.Listbox(step3, height=3)
        self.files_listbox.pack(side="left", fill="x", expand=True, padx=5, pady=5)
        
        btn_frame = tk.Frame(step3)
        btn_frame.pack(side="right")
        tk.Button(btn_frame, text="Add Folder", command=self.add_folder).pack(fill="x", padx=5)
        tk.Button(btn_frame, text="Clear", command=self.clear_files).pack(fill="x", padx=5)

        # --- Step 4: Options ---
        step4 = tk.LabelFrame(self, text="Step 4: Build Options")
        step4.pack(fill="x", padx=10, pady=5)
        tk.Checkbutton(step4, text="Hide Console Window (--noconsole)", 
                       variable=self.hide_console).pack(side="left", padx=5, pady=5)

        tk.Button(self, text="BUILD ONE-FILE EXE", bg="#2e7d32", fg="white", font=("Arial", 10, "bold"),
                  command=self.start_build).pack(pady=10)

        # Log window
        self.log = scrolledtext.ScrolledText(self, height=12, bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    def browse_script(self):
        path = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
        if path: self.script_path.set(path)

    def add_folder(self):
        path = filedialog.askdirectory(title="Select Folder to Bundle")
        if path:
            folder_name = os.path.basename(path)
            self.additional_files.append(f"{path}{os.pathsep}{folder_name}")
            self.files_listbox.insert(tk.END, f"Folder: {folder_name}")

    def clear_files(self):
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
        
        # Base Command
        cmd = [
            sys.executable, "-m", "PyInstaller",
            "--onefile",
            "--name", self.exe_name.get(),
            "--distpath", output_dir,
        ]

        # Handle Console Toggle
        if self.hide_console.get():
            cmd.append("--noconsole")
        else:
            cmd.append("--console")

        # Add additional folders
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