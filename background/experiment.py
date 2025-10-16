import tkinter as tk
import requests
import threading
import time
import json

SETTINGS_FILE = "settings.json"

class CountdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Launch Control - GO/NOGO")

        self.go_nogo_labels = {}
        self.sheet_data = {}
        self.last_data = {}
        self.running = True

        # Load settings
        self.settings = self.load_settings()

        tk.Label(root, text="GO/NOGO STATUS", font=("Arial", 16, "bold")).pack(pady=10)

        # Create display area
        self.frame = tk.Frame(root)
        self.frame.pack(pady=10)

        # Buttons
        tk.Button(root, text="Add Spreadsheet", command=self.add_spreadsheet_window).pack(pady=5)
        tk.Button(root, text="Stop", command=self.stop).pack(pady=5)

        self.start_update_thread()

    def load_settings(self):
        try:
            with open(SETTINGS_FILE, "r") as f:
                return json.load(f)
        except FileNotFoundError:
            return {"spreadsheets": []}

    def save_settings(self):
        with open(SETTINGS_FILE, "w") as f:
            json.dump(self.settings, f, indent=4)

    def add_spreadsheet_window(self):
        win = tk.Toplevel(self.root)
        win.title("Add Spreadsheet")

        tk.Label(win, text="Name:").grid(row=0, column=0)
        name_entry = tk.Entry(win)
        name_entry.grid(row=0, column=1)

        tk.Label(win, text="Link (CSV export or share link):").grid(row=1, column=0)
        link_entry = tk.Entry(win, width=60)
        link_entry.grid(row=1, column=1)

        tk.Label(win, text="Range cell (e.g., L2):").grid(row=2, column=0)
        range_entry = tk.Entry(win)
        range_entry.grid(row=2, column=1)

        def save_sheet():
            name = name_entry.get().strip()
            link = link_entry.get().strip()
            cell = range_entry.get().strip().upper()
            if name and link and cell:
                self.settings["spreadsheets"].append({
                    "name": name,
                    "link": link,
                    "cell": cell
                })
                self.save_settings()
                self.add_go_nogo_label(name)
                win.destroy()

        tk.Button(win, text="Save", command=save_sheet).grid(row=3, column=0, columnspan=2, pady=10)

    def add_go_nogo_label(self, name):
        if name not in self.go_nogo_labels:
            label = tk.Label(self.frame, text=f"{name}: ---", font=("Arial", 14), width=25)
            label.pack(pady=2)
            self.go_nogo_labels[name] = label

    def update_labels(self):
        for sheet in self.settings["spreadsheets"]:
            name = sheet["name"]
            link = sheet["link"]
            cell = sheet["cell"]

            # Convert normal sheet link to CSV export link if needed
            if "/edit" in link and "export" not in link:
                link = link.split("/edit")[0] + "/gviz/tq?tqx=out:csv"

            try:
                r = requests.get(link, timeout=5)
                if r.status_code == 200:
                    content = r.text
                    if name not in self.last_data or self.last_data[name] != content:
                        self.last_data[name] = content
                        # Just read raw content and extract cell text if possible
                        value = self.extract_cell_value(content, cell)
                        self.update_label_color(name, value)
            except Exception as e:
                print(f"Error updating {name}: {e}")

    def extract_cell_value(self, csv_data, cell):
        # Simple CSV parser to get cell data like L2
        try:
            rows = [r.split(",") for r in csv_data.splitlines() if r.strip()]
            col = ord(cell[0]) - 65
            row = int(cell[1:]) - 1
            return rows[row][col].strip().upper()
        except Exception:
            return "ERR"

    def update_label_color(self, name, value):
        label = self.go_nogo_labels.get(name)
        if not label:
            return

        if "GO" in value:
            label.config(text=f"{name}: GO", bg="green", fg="white")
        elif "NO" in value:
            label.config(text=f"{name}: NO GO", bg="red", fg="white")
        else:
            label.config(text=f"{name}: ---", bg="gray", fg="black")

    def start_update_thread(self):
        threading.Thread(target=self.update_loop, daemon=True).start()

    def update_loop(self):
        while self.running:
            self.update_labels()
            time.sleep(0.1)

    def stop(self):
        self.running = False
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = CountdownApp(root)
    root.mainloop()
