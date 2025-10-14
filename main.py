import tkinter as tk
import time
from datetime import datetime

HTML_FILE = "countdown.html"

def write_html(mission_name, timer_text):
    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    margin: 0;
    background-color: black;
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
}}
#mission {{
    color: white;
    font-family: Consolas, monospace;
    font-size: 4vw;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    max-width: 90vw;
    text-align: center;
}}
#timer {{
    color: white;
    font-family: Consolas, monospace;
    font-size: 8vw;
    line-height: 1;
    white-space: nowrap;
    text-align: center;
}}
</style>
<script>
setTimeout(() => location.reload(), 1000);
</script>
</head>
<body>
<div id="mission">{mission_name}</div>
<div id="timer">{timer_text}</div>
</body>
</html>"""
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)


class CountdownApp:
    def __init__(self, root):
        self.root = root
        self.root.title("RocketLaunchCountdown")
        self.root.config(bg="black")
        self.root.attributes("-topmost", True)

        # State
        self.running = False
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.target_time = None
        self.hold_start_time = None
        self.remaining_time = 0
        self.mission_name = "Mission"

        # Display
        self.text = tk.Label(
            root,
            text="T-00:00:00",
            font=("Consolas", 80, "bold"),
            fg="white",
            bg="black"
        )
        self.text.pack(padx=50, pady=20)

        # Mission name input
        frame_top = tk.Frame(root, bg="black")
        frame_top.pack(pady=5)
        tk.Label(frame_top, text="Mission Name:", fg="white", bg="black").pack(side="left")
        self.mission_entry = tk.Entry(frame_top, width=20, font=("Arial", 18))
        self.mission_entry.insert(0, self.mission_name)
        self.mission_entry.pack(side="left")

        # Mode toggle
        frame_mode = tk.Frame(root, bg="black")
        frame_mode.pack(pady=5)
        self.mode_var = tk.StringVar(value="duration")
        tk.Radiobutton(frame_mode, text="Duration", variable=self.mode_var, value="duration",
                       fg="white", bg="black", command=self.update_inputs).pack(side="left", padx=5)
        tk.Radiobutton(frame_mode, text="Clock Time", variable=self.mode_var, value="clock",
                       fg="white", bg="black", command=self.update_inputs).pack(side="left", padx=5)

        # Duration inputs
        frame_duration = tk.Frame(root, bg="black")
        frame_duration.pack(pady=5)
        tk.Label(frame_duration, text="H", fg="white", bg="black").pack(side="left")
        self.hours_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.hours_entry.insert(0, "0")
        self.hours_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="M", fg="white", bg="black").pack(side="left")
        self.minutes_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.minutes_entry.insert(0, "5")
        self.minutes_entry.pack(side="left", padx=2)
        tk.Label(frame_duration, text="S", fg="white", bg="black").pack(side="left")
        self.seconds_entry = tk.Entry(frame_duration, width=3, font=("Arial", 18))
        self.seconds_entry.insert(0, "0")
        self.seconds_entry.pack(side="left", padx=2)

        # Clock time input
        frame_clock = tk.Frame(root, bg="black")
        frame_clock.pack(pady=5)
        tk.Label(frame_clock, text="HH:MM", fg="white", bg="black").pack(side="left")
        self.clock_entry = tk.Entry(frame_clock, width=7, font=("Arial", 18))
        self.clock_entry.insert(0, "14:00")
        self.clock_entry.pack(side="left", padx=5)

        # Control buttons
        frame_buttons = tk.Frame(root, bg="black")
        frame_buttons.pack(pady=10)

        self.start_btn = tk.Button(frame_buttons, text="▶ Start", command=self.start, font=("Arial", 14))
        self.start_btn.grid(row=0, column=0, padx=5)

        # Hold and resume share the same position
        self.hold_btn = tk.Button(frame_buttons, text="⏸ Hold", command=self.hold, font=("Arial", 14))
        self.hold_btn.grid(row=0, column=1, padx=5)

        self.resume_btn = tk.Button(frame_buttons, text="⏵ Resume", command=self.resume, font=("Arial", 14))
        self.resume_btn.grid(row=0, column=1, padx=5)
        self.resume_btn.grid_remove()  # hidden at start

        self.scrub_btn = tk.Button(frame_buttons, text="🚫 Scrub", command=self.scrub, font=("Arial", 14), fg="red")
        self.scrub_btn.grid(row=0, column=2, padx=5)

        self.reset_btn = tk.Button(frame_buttons, text="⟳ Reset", command=self.reset, font=("Arial", 14))
        self.reset_btn.grid(row=0, column=3, padx=5)


        self.update_inputs()
        self.update_clock()

    # ----------------------------
    # Update input visibility based on mode
    # ----------------------------
    def update_inputs(self):
        if self.mode_var.get() == "duration":
            self.hours_entry.config(state="normal")
            self.minutes_entry.config(state="normal")
            self.seconds_entry.config(state="normal")
            self.clock_entry.config(state="disabled")
        else:
            self.hours_entry.config(state="disabled")
            self.minutes_entry.config(state="disabled")
            self.seconds_entry.config(state="disabled")
            self.clock_entry.config(state="normal")

    # ----------------------------
    # Control logic
    # ----------------------------
    def start(self):
        self.mission_name = self.mission_entry.get().strip() or "Mission"
        self.running = True
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.show_hold_button()

        try:
            if self.mode_var.get() == "duration":
                h = int(self.hours_entry.get())
                m = int(self.minutes_entry.get())
                s = int(self.seconds_entry.get())
                total_seconds = h * 3600 + m * 60 + s
            else:
                now = datetime.now()
                parts = [int(p) for p in self.clock_entry.get().split(":")]
                h, m = parts[0], parts[1]
                s = parts[2] if len(parts) == 3 else 0
                target_today = now.replace(hour=h, minute=m, second=s, microsecond=0)
                if target_today < now:
                    target_today = target_today.replace(day=now.day + 1)
                total_seconds = (target_today - now).total_seconds()
        except Exception:
            self.text.config(text="Invalid time")
            write_html(self.mission_name, "Invalid time")
            return

        self.target_time = time.time() + total_seconds
        self.remaining_time = total_seconds

    def hold(self):
        if self.running and not self.on_hold and not self.scrubbed:
            self.on_hold = True
            self.hold_start_time = time.time()
            self.remaining_time = max(0, self.target_time - self.hold_start_time)
            self.show_resume_button()

    def resume(self):
        if self.running and self.on_hold and not self.scrubbed:
            self.on_hold = False
            self.target_time = time.time() + self.remaining_time
            self.show_hold_button()

    def show_hold_button(self):
        self.resume_btn.grid_remove()
        self.hold_btn.grid()

    def show_resume_button(self):
        self.hold_btn.grid_remove()
        self.resume_btn.grid()


    def scrub(self):
        self.scrubbed = True
        self.running = False
        write_html(self.mission_name, "SCRUB")
        self.text.config(text="SCRUB")

    def reset(self):
        self.running = False
        self.on_hold = False
        self.scrubbed = False
        self.counting_up = False
        self.text.config(text="T-00:00:00")
        write_html(self.mission_name, "T-00:00:00")
        self.show_hold_button()

    # ----------------------------
    # Clock updating
    # ----------------------------
    def format_time(self, seconds, prefix="T-"):
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        return f"{prefix}{h:02}:{m:02}:{s:02}"

    def update_clock(self):
        if self.running and not self.scrubbed:
            now = time.time()
            if self.on_hold:
                elapsed = int(now - self.hold_start_time)
                timer_text = self.format_time(elapsed, "H+")
            elif self.target_time:
                diff = int(self.target_time - now)
                if diff <= 0 and not self.counting_up:
                    self.counting_up = True
                    self.target_time = now
                    diff = 0
                if self.counting_up:
                    elapsed = int(now - self.target_time)
                    timer_text = self.format_time(elapsed, "T+")
                else:
                    timer_text = self.format_time(diff, "T-")
            else:
                timer_text = "T-00:00:00"
        else:
            timer_text = self.text.cget("text")

        self.text.config(text=timer_text)
        write_html(self.mission_name, timer_text)
        self.root.after(200, self.update_clock)


if __name__ == "__main__":
    root = tk.Tk()
    app = CountdownApp(root)
    write_html("Mission", "T-00:00:00")
    root.mainloop()
