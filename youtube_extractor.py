import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import yt_dlp
import threading

# Globals
stop_flag = False
existing_titles = set()
existing_file_path = None

def format_date(raw):
    if raw and len(raw) == 8:
        return f"{raw[6:8]}-{raw[4:6]}-{raw[0:4]}"
    return raw

def load_existing_titles(filepath):
    global existing_titles
    try:
        df = pd.read_excel(filepath)
        existing_titles = set(df['Title'].dropna().astype(str).str.strip().tolist())
        return df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file.\n{e}")
        return None

def save_to_excel(filepath, data, append=False):
    df_new = pd.DataFrame(data)
    if append:
        df_existing = pd.read_excel(filepath)
        df_combined = pd.concat([df_existing, df_new], ignore_index=True)
        df_combined.to_excel(filepath, index=False)
    else:
        df_new.to_excel(filepath, index=False)

def extract_data():
    global stop_flag, existing_file_path
    stop_flag = False
    url = url_entry.get().strip()
    if not url:
        messagebox.showwarning("Input Error", "Please enter a playlist URL.")
        return

    try:
        extract_btn.config(state="disabled")
        cancel_btn.config(state="normal")
        progress_label.config(text="Processing...")
        root.update()

        flat_opts = {
            'quiet': True,
            'extract_flat': True,
            'skip_download': True,
            'force_generic_extractor': True,
            'nocheckcertificate': True,
        }
        with yt_dlp.YoutubeDL(flat_opts) as ydl:
            playlist = ydl.extract_info(url, download=False)

        entries = playlist.get("entries", [])
        total = len(entries)
        if not total:
            messagebox.showinfo("No Videos", "No videos found in playlist.")
            return

        progress_bar["maximum"] = total
        data = []

        video_opts = {
            'quiet': True,
            'ignoreerrors': True,
            'skip_download': True,
            'extract_flat': False,
            'nocheckcertificate': True,
            'postprocessors': [],
            'postprocessor_hooks': [],
            'prefer_ffmpeg': False,
            'ffmpeg_location': 'NUL',
            'check_formats': False,
        }

        for idx, flat in enumerate(entries, 1):
            if stop_flag:
                progress_label.config(text="Cancelled.")
                break

            vid_url = f"https://www.youtube.com/watch?v={flat.get('id')}"
            try:
                with yt_dlp.YoutubeDL(video_opts) as ydl:
                    video = ydl.extract_info(vid_url, download=False)

                title = video.get("title", "").strip()
                if mode.get() == "existing" and title in existing_titles:
                    print(f"[SKIP] '{title}' already in file.")
                    continue

                video_data = {
                    "Title": title,
                    "Description": video.get("description", ""),
                    "Channel": video.get("uploader", ""),
                    "Upload Date": format_date(video.get("upload_date", "")),
                    "URL": video.get("webpage_url", vid_url)
                }
                data.append(video_data)
                print(f"[{idx}] {title}")
            except Exception as e:
                print(f"[ERROR] Video failed at index {idx}: {e}")

            progress_bar["value"] = idx
            percent = int((idx / total) * 100)
            progress_label.config(text=f"{idx}/{total} processed ({percent}%)")
            root.update_idletasks()
            root.after(1)

        if stop_flag or not data:
            extract_btn.config(state="normal")
            cancel_btn.config(state="disabled")
            return

        if mode.get() == "new":
            filepath = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Excel File"
            )
            if filepath:
                save_to_excel(filepath, data, append=False)
                messagebox.showinfo("Success", f"Saved to:\n{filepath}")
        else:
            save_to_excel(existing_file_path, data, append=True)
            messagebox.showinfo("Success", f"Appended to:\n{existing_file_path}")

        progress_label.config(text="Done.")
    except Exception as e:
        messagebox.showerror("Error", f"Error during extraction:\n{e}")
    finally:
        extract_btn.config(state="normal")
        cancel_btn.config(state="disabled")

def cancel_extraction():
    global stop_flag
    stop_flag = True
    cancel_btn.config(state="disabled")

def run_thread():
    thread = threading.Thread(target=extract_data)
    thread.start()

def toggle_mode():
    global existing_file_path
    if mode.get() == "new":
        file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file:
            return
        df = load_existing_titles(file)
        if df is not None:
            existing_file_path = file
            mode.set("existing")
            switch_btn.config(text="Switch to New")
    else:
        mode.set("new")
        existing_file_path = None
        existing_titles.clear()
        switch_btn.config(text="Switch to Existing")

# GUI Setup
YOUTUBE_RED = "#FF0000"
WHITE = "#FFFFFF"
YELLOW = "#FFD700"

root = tk.Tk()
root.title("YouTube Extractor")
root.geometry("500x340")
root.configure(bg=YOUTUBE_RED)
mode = tk.StringVar(value="new")  # Default is new file

# Header + Switch
top_frame = tk.Frame(root, bg=YOUTUBE_RED)
top_frame.pack(fill="x", pady=(10, 0))

tk.Label(top_frame, text="YouTube Playlist Extractor",
         font=("Helvetica", 16, "bold"), bg=YOUTUBE_RED, fg=WHITE).pack(side="left", padx=(20, 0))

switch_btn = tk.Button(top_frame, text="Switch to Existing", command=toggle_mode,
                       bg=WHITE, fg=YOUTUBE_RED, font=("Helvetica", 10, "bold"))
switch_btn.pack(side="right", padx=20)

tk.Label(root, text="Enter Playlist URL:", font=("Helvetica", 12),
         bg=YOUTUBE_RED, fg=WHITE).pack(pady=(15, 0))

url_entry = tk.Entry(root, width=50, font=("Helvetica", 11))
url_entry.pack(pady=5)

extract_btn = tk.Button(root, text="Extract", command=run_thread,
                        bg=WHITE, fg=YOUTUBE_RED, font=("Helvetica", 12, "bold"))
extract_btn.pack(pady=(10, 5))

cancel_btn = tk.Button(root, text="Cancel", command=cancel_extraction,
                       bg=WHITE, fg=YOUTUBE_RED, font=("Helvetica", 12, "bold"))
cancel_btn.pack()
cancel_btn.config(state="disabled")

style = ttk.Style()
style.theme_use('default')
style.configure("red.Horizontal.TProgressbar", thickness=2, troughcolor=YELLOW, background=YOUTUBE_RED)

progress_bar = ttk.Progressbar(root, length=400, mode="determinate", style="red.Horizontal.TProgressbar")
progress_bar.pack(pady=(10, 2))

progress_label = tk.Label(root, text="", bg=YOUTUBE_RED, fg=WHITE, font=("Helvetica", 10))
progress_label.pack()

root.mainloop()
