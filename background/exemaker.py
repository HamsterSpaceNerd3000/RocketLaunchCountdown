import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import sys
import os


class SimpleBuilder(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Simple PyInstaller Builder")
        self.geometry("700x500")

        self.script_path = tk.StringVar()

        # Top controls
        top_frame = tk.Frame(self)
        top_frame.pack(fill="x", padx=10, pady=10)

        tk.Label(top_frame, text="Python Script:").pack(side="left")

        tk.Entry(top_frame, textvariable=self.script_path, width=60).pack(
            side="left", padx=5
        )

        tk.Button(top_frame, text="Browse", command=self.browse_script).pack(
            side="left"
        )

        tk.Button(self, text="Build EXE", bg="#2e7d32", fg="white",
                  command=self.start_build).pack(pady=5)

        # Log window
        self.log = scrolledtext.ScrolledText(self, height=20)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

    def browse_script(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Python files", "*.py")]
        )
        if file_path:
            self.script_path.set(file_path)

    def log_write(self, text):
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def start_build(self):
        script = self.script_path.get()

        if not script or not os.path.isfile(script):
            messagebox.showerror("Error", "Please select a valid Python script.")
            return

        self.log.delete("1.0", tk.END)

        thread = threading.Thread(
            target=self.build_exe,
            args=(script,),
            daemon=True
        )
        thread.start()

    def build_exe(self, script):
        self.log_write(f"Using Python: {sys.executable}\n\n")

        output_dir = os.path.join(os.path.dirname(script), "dist")
        os.makedirs(output_dir, exist_ok=True)

        cmd = [
            sys.executable,
            "-m",
            "PyInstaller",
            "--onefile",
            "--noconsole",
            "--distpath",
            output_dir,
            script
        ]

        self.log_write("Running command:\n")
        self.log_write(" ".join(cmd) + "\n\n")

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            for line in process.stdout:
                self.log_write(line)

            process.wait()

            if process.returncode == 0:
                self.log_write("\nBuild completed successfully.\n")
                messagebox.showinfo("Success", "EXE built successfully!")
            else:
                self.log_write(f"\nBuild failed (code {process.returncode}).\n")
                messagebox.showerror("Error", "Build failed. Check log.")

        except Exception as e:
            self.log_write(f"\nError starting PyInstaller:\n{e}\n")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = SimpleBuilder()
    app.mainloop()
