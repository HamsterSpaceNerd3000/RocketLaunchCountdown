#!/usr/bin/env python3
"""
Simple Tkinter GUI to run PyInstaller and make an exe from a .py file.

Features:
- choose script
- choose onefile / dir
- choose windowed (noconsole)
- add icon
- add additional data (file or folder; multiple, separated by semicolons)
- set output folder
- show live PyInstaller output
- cancel build
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import subprocess
import threading
import sys
import shlex
import os
from pathlib import Path

class BuilderGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Python → EXE (PyInstaller GUI)")
        self.geometry("800x600")
        self.create_widgets()
        self.proc = None  # subprocess handle
        self.stop_requested = False

    def create_widgets(self):
        frame_top = tk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=8)

        tk.Label(frame_top, text="Script:").grid(row=0, column=0, sticky="e")
        self.script_entry = tk.Entry(frame_top, width=70)
        self.script_entry.grid(row=0, column=1, padx=6)
        tk.Button(frame_top, text="Browse...", command=self.browse_script).grid(row=0, column=2)

        tk.Label(frame_top, text="Output folder:").grid(row=1, column=0, sticky="e")
        self.out_entry = tk.Entry(frame_top, width=70)
        self.out_entry.grid(row=1, column=1, padx=6)
        tk.Button(frame_top, text="Choose...", command=self.choose_output).grid(row=1, column=2)

        opts_frame = tk.LabelFrame(self, text="Options", padx=8, pady=8)
        opts_frame.pack(fill="x", padx=10)

        self.onefile_var = tk.BooleanVar(value=True)
        tk.Checkbutton(opts_frame, text="Onefile (single exe)", variable=self.onefile_var).grid(row=0, column=0, sticky="w", padx=6, pady=2)

        self.windowed_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opts_frame, text="Windowed (no console) / --noconsole", variable=self.windowed_var).grid(row=0, column=1, sticky="w", padx=6, pady=2)

        tk.Label(opts_frame, text="Icon (.ico):").grid(row=1, column=0, sticky="e")
        self.icon_entry = tk.Entry(opts_frame, width=50)
        self.icon_entry.grid(row=1, column=1, sticky="w", padx=6)
        tk.Button(opts_frame, text="Browse", command=self.browse_icon).grid(row=1, column=2)

        tk.Label(opts_frame, text="Additional data (src;dest pairs separated by ';', e.g. resources;resources):").grid(row=2, column=0, columnspan=3, sticky="w", pady=(6,0))
        self.data_entry = tk.Entry(opts_frame, width=110)
        self.data_entry.grid(row=3, column=0, columnspan=3, padx=6, pady=4)

        tk.Label(opts_frame, text="Extra PyInstaller args:").grid(row=4, column=0, sticky="w")
        self.extra_entry = tk.Entry(opts_frame, width=80)
        self.extra_entry.grid(row=4, column=1, columnspan=2, padx=6, pady=4, sticky="w")

        run_frame = tk.Frame(self)
        run_frame.pack(fill="x", padx=10, pady=8)

        self.build_btn = tk.Button(run_frame, text="Build EXE", command=self.start_build, bg="#2b7a78", fg="white")
        self.build_btn.pack(side="left", padx=(0,6))

        self.cancel_btn = tk.Button(run_frame, text="Cancel", command=self.request_cancel, state="disabled", bg="#b00020", fg="white")
        self.cancel_btn.pack(side="left")

        clear_btn = tk.Button(run_frame, text="Clear Log", command=self.clear_log)
        clear_btn.pack(side="left", padx=6)

        open_out_btn = tk.Button(run_frame, text="Open Output Folder", command=self.open_output)
        open_out_btn.pack(side="right")

        self.log = scrolledtext.ScrolledText(self, height=18, font=("Consolas", 10))
        self.log.pack(fill="both", expand=True, padx=10, pady=(0,10))

    def browse_script(self):
        path = filedialog.askopenfilename(filetypes=[("Python files", "*.py")])
        if path:
            self.script_entry.delete(0, tk.END)
            self.script_entry.insert(0, path)
            # default output to script parent /dist
            parent = os.path.dirname(path)
            default_out = os.path.join(parent, "dist")
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, default_out)

    def choose_output(self):
        path = filedialog.askdirectory()
        if path:
            self.out_entry.delete(0, tk.END)
            self.out_entry.insert(0, path)

    def browse_icon(self):
        path = filedialog.askopenfilename(filetypes=[("Icon files", "*.ico")])
        if path:
            self.icon_entry.delete(0, tk.END)
            self.icon_entry.insert(0, path)

    def clear_log(self):
        self.log.delete("1.0", tk.END)

    def open_output(self):
        out = self.out_entry.get().strip()
        if not out:
            messagebox.showinfo("Output folder", "No output folder set.")
            return
        os.startfile(out) if os.name == "nt" else subprocess.run(["xdg-open", out])

    def request_cancel(self):
        if self.proc and self.proc.poll() is None:
            self.stop_requested = True
            # try terminate politely
            try:
                self.proc.terminate()
            except Exception:
                pass
            self.log_insert("\nCancellation requested...\n")
        self.cancel_btn.config(state="disabled")

    def log_insert(self, text):
        self.log.insert(tk.END, text)
        self.log.see(tk.END)

    def start_build(self):
        script = self.script_entry.get().strip()
        if not script or not os.path.isfile(script):
            messagebox.showerror("Error", "Please choose a valid Python script to build.")
            return

        out_dir = self.out_entry.get().strip() or os.path.dirname(script)
        os.makedirs(out_dir, exist_ok=True)

        self.build_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")
        self.stop_requested = False
        self.clear_log()

        # Start build on a thread to keep UI responsive
        thread = threading.Thread(target=self.run_pyinstaller, args=(script, out_dir), daemon=True)
        thread.start()

    def run_pyinstaller(self, script, out_dir):
        # Build the PyInstaller command
        cmd = [sys.executable, "-m", "PyInstaller"]

        if self.onefile_var.get():
            cmd.append("--onefile")
        else:
            cmd.append("--onedir")

        if self.windowed_var.get():
            cmd.append("--noconsole")

        icon = self.icon_entry.get().strip()
        if icon:
            cmd.extend(["--icon", icon])

        # add additional data: user can provide pairs like "data;data" or "assets;assets"
        data_spec = self.data_entry.get().strip()
        if data_spec:
            # support multiple separated by semicolons or vertical bars
            pairs = [p for p in (data_spec.split(";") + data_spec.split("|")) if p.strip()]
            # normalize pairs to PyInstaller format: src;dest (on Windows use ';' in CLI but PyInstaller expects src;dest as single argument)
            for p in pairs:
                # if user typed "src:dest" or "src->dest", replace with semicolon
                p_fixed = p.replace(":", ";").replace("->", ";")
                cmd.extend(["--add-data", p_fixed])

        # user extra args (raw)
        extra = self.extra_entry.get().strip()
        if extra:
            # split carefully
            cmd.extend(shlex.split(extra))

        # ensure output path goes to chosen dir: use --distpath
        cmd.extend(["--distpath", out_dir])

        # entry script
        cmd.append(script)

        self.log_insert("Running PyInstaller with command:\n" + " ".join(shlex.quote(c) for c in cmd) + "\n\n")

        # spawn the process
        try:
            self.proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=1,
                universal_newlines=True,
            )
        except Exception as e:
            self.log_insert(f"Failed to start PyInstaller: {e}\n")
            self.build_btn.config(state="normal")
            self.cancel_btn.config(state="disabled")
            return

        # Stream output line by line
        try:
            for line in self.proc.stdout:
                if line:
                    self.log_insert(line)
                if self.stop_requested:
                    try:
                        self.proc.terminate()
                    except Exception:
                        pass
                    break
            self.proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.log_insert("Process did not exit in time after termination.\n")
        except Exception as e:
            self.log_insert(f"Error while running PyInstaller: {e}\n")

        retcode = self.proc.returncode if self.proc else None
        if self.stop_requested:
            self.log_insert("\nBuild cancelled by user.\n")
        elif retcode == 0:
            self.log_insert("\nBuild finished successfully.\n")
        else:
            self.log_insert(f"\nBuild finished with return code {retcode}.\n")

        # Re-enable buttons
        self.build_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.proc = None
        self.stop_requested = False

if __name__ == "__main__":
    app = BuilderGUI()
    app.mainloop()
