# Copyright (c) 2026 MyCompany LLC
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#     http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

import os
import sys
import shutil
import socket
import webbrowser
import tkinter as tk
from tkinter import messagebox
import win32com.client

# Ensure only one instance of the script is running
try:
    lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    lock_socket.bind(('127.0.0.1', 47291))
except socket.error:
    sys.exit(0)

# Colors Palette (Sleek Dark Theme)
BG_COLOR = "#1E1E2E"
SURFACE_COLOR = "#313244"
ACCENT_COLOR = "#89B4FA"
HOVER_COLOR = "#A6E3A1"
TEXT_COLOR = "#CDD6F4"
WARNING_COLOR = "#F38BA8"
BTN_BG = "#45475A"

# Autocorrect pairs
AUTOCORRECT_PAIRS = [
    {"shortcode": "\\cents", "symbol": "\u00a2"},
    {"shortcode": "\\repeat", "symbol": "\u00af"},
    {"shortcode": "\\repeating", "symbol": "\u00af"},
    {"shortcode": "\\vinculum", "symbol": "\u00af"},
    {"shortcode": "\\infinity", "symbol": "\u221e"},
    {"shortcode": "\\2root", "symbol": "\u221a"},
    {"shortcode": "\\3root", "symbol": "\u221b"},
    {"shortcode": "\\4root", "symbol": "\u221c"},
    {"shortcode": "\\comp", "symbol": "\u2218"},
    {"shortcode": "\\deg", "symbol": "\u00b0"},
    {"shortcode": "\\rad", "symbol": "\u33ad"},
    {"shortcode": "\\join", "symbol": "\u22c8"},
    {"shortcode": "\\qed", "symbol": "\u220e"},
    {"shortcode": "\\endproof", "symbol": "\u220e"},
    {"shortcode": "\\circle", "symbol": "\u25ef"},
    {"shortcode": "\\circledot", "symbol": "\u2299"},
    {"shortcode": "\\line", "symbol": "\u20e1"},
    {"shortcode": "\\seg", "symbol": "\u00af"},
    {"shortcode": "\\measangle", "symbol": "\u2221"},
    {"shortcode": "\\rightangle", "symbol": "\u221f"},
    {"shortcode": "\\triangle", "symbol": "\u25b3"},
    {"shortcode": "\\parallelogram", "symbol": "\u25b1"},
    {"shortcode": "\\notparallel", "symbol": "\u2226"},
    {"shortcode": "\\ray", "symbol": "\u20d7"},
    {"shortcode": "\\arc", "symbol": "\u23dc"},
    {"shortcode": "\\nlt", "symbol": "\u226e"},
    {"shortcode": "\\notlt", "symbol": "\u226e"},
    {"shortcode": "\\ngt", "symbol": "\u226f"},
    {"shortcode": "\\notgt", "symbol": "\u226f"},
    {"shortcode": "\\nleq", "symbol": "\u2270"},
    {"shortcode": "\\notle", "symbol": "\u2270"},
    {"shortcode": "\\nge", "symbol": "\u2271"},
    {"shortcode": "\\notge", "symbol": "\u2271"},
    {"shortcode": "\\ngeq", "symbol": "\u2271"},
    {"shortcode": "\\notgeq", "symbol": "\u2271"},
    {"shortcode": "\\not", "symbol": "\u00ac"},
    {"shortcode": "\\muchgreater", "symbol": "\u226b"},
    {"shortcode": "\\muchless", "symbol": "\u226a"},
    {"shortcode": "\\notapprox", "symbol": "\u2249"},
    {"shortcode": "\\notcong", "symbol": "\u2247"},
    {"shortcode": "\\doubleint", "symbol": "\u222c"},
    {"shortcode": "\\tripleint", "symbol": "\u222d"},
    {"shortcode": "\\dprime", "symbol": "\u2033"},
    {"shortcode": "\\doubleprime", "symbol": "\u2033"},
    {"shortcode": "\\tprime", "symbol": "\u2034"},
    {"shortcode": "\\tripleprime", "symbol": "\u2034"},
    {"shortcode": "\\qprime", "symbol": "\u2057"},
    {"shortcode": "\\quadprime", "symbol": "\u2057"},
    {"shortcode": "\\grad", "symbol": "\u2207"},
    {"shortcode": "\\laplace", "symbol": "\u2206"},
    {"shortcode": "\\union", "symbol": "\u222a"},
    {"shortcode": "\\Union", "symbol": "\u22c3"},
    {"shortcode": "\\intersection", "symbol": "\u2229"},
    {"shortcode": "\\Intersection", "symbol": "\u22c2"},
    {"shortcode": "\\notsubset", "symbol": "\u2284"},
    {"shortcode": "\\notsuperset", "symbol": "\u2285"},
    {"shortcode": "\\notsubseteq", "symbol": "\u2288"},
    {"shortcode": "\\notsuperseteq", "symbol": "\u2289"},
    {"shortcode": "\\subsetnoteq", "symbol": "\u228a"},
    {"shortcode": "\\supersetnoteq", "symbol": "\u228b"},
    {"shortcode": "\\belongs", "symbol": "\u2208"},
    {"shortcode": "\\element", "symbol": "\u2208"},
    {"shortcode": "\\contains", "symbol": "\u220b"},
    {"shortcode": "\\owns", "symbol": "\u220b"},
    {"shortcode": "\\powerset", "symbol": "\u2118"},
    {"shortcode": "\\complement", "symbol": "\u2201"},
    {"shortcode": "\\divide", "symbol": "\u2223"},
    {"shortcode": "\\notdivide", "symbol": "\u2224"},
    {"shortcode": "\\and", "symbol": "\u2227"},
    {"shortcode": "\\land", "symbol": "\u2227"},
    {"shortcode": "\\or", "symbol": "\u2228"},
    {"shortcode": "\\lor", "symbol": "\u2228"},
    {"shortcode": "\\nand", "symbol": "\u22bc"},
    {"shortcode": "\\nor", "symbol": "\u22bd"},
    {"shortcode": "\\xor", "symbol": "\u2295"},
    {"shortcode": "\\xnor", "symbol": "\u2299"},
    {"shortcode": "\\proves", "symbol": "\u22a2"},
    {"shortcode": "\\tautology", "symbol": "\u22a4"},
    {"shortcode": "\\false", "symbol": "\u22a5"},
    {"shortcode": "\\contradiction", "symbol": "\u22a5"},
    {"shortcode": "\\implication", "symbol": "\u2192"},
    {"shortcode": "\\implies", "symbol": "\u2192"},
    {"shortcode": "\\biconditional", "symbol": "\u2194"},
    {"shortcode": "\\Implication", "symbol": "\u21d2"},
    {"shortcode": "\\Implies", "symbol": "\u21d2"},
    {"shortcode": "\\Biconditional", "symbol": "\u21d4"},
    {"shortcode": "\\forces", "symbol": "\u22a9"},
    {"shortcode": "\\entailment", "symbol": "\u22a8"},
    {"shortcode": "\\true", "symbol": "\u22a8"},
    {"shortcode": "\\foreach", "symbol": "\u2200"},
    {"shortcode": "\\forsome", "symbol": "\u2203"},
    {"shortcode": "\\stddev", "symbol": "\u03c3"},
    {"shortcode": "\\mean", "symbol": "\u03bc"},
    {"shortcode": "\\corr", "symbol": "\u03c1"},
    {"shortcode": "\\expect", "symbol": "\U0001d53c"},
    {"shortcode": "\\prob", "symbol": "\u2119"},
    {"shortcode": "\\kron", "symbol": "\u2297"},
    {"shortcode": "\\hadamard", "symbol": "\u2299"},
    {"shortcode": "\\adjoint", "symbol": "\u2020"},
    {"shortcode": "\\identity", "symbol": "\U0001d408"},
    {"shortcode": "\\directsum", "symbol": "\u2295"}
]

def get_acl_paths():
    appdata = os.environ.get('APPDATA', '')
    source_path = os.path.join(appdata, 'Microsoft', 'Office', 'mso0127.acl')
    script_dir = os.path.dirname(os.path.abspath(__file__))
    backup_path = os.path.join(script_dir, 'mathautocorrect.backup')
    return source_path, backup_path

def backup_file():
    source, backup = get_acl_paths()
    if os.path.exists(source):
        try:
            shutil.copy2(source, backup)
            messagebox.showinfo("Backup Success", f"Backup created successfully:\n{backup}")
        except Exception as e:
            messagebox.showerror("Backup Error", f"Failed to back up the file:\n{e}")
    else:
        messagebox.showerror("Backup Error", "The Microsoft Word math AutoCorrect file (mso0127.acl) could not be located on your system.")

def restore_file():
    source, backup = get_acl_paths()
    if os.path.exists(backup):
        try:
            # Ensure target directory exists
            os.makedirs(os.path.dirname(source), exist_ok=True)
            shutil.copy2(backup, source)
            messagebox.showinfo("Restore Success", "Backup restored successfully.\n\nIf Microsoft Word is running, please close and restart it for changes to take effect.")
        except Exception as e:
            messagebox.showerror("Restore Error", f"Failed to restore the file:\n{e}")
    else:
        messagebox.showerror("Restore Error", f"No backup file found at:\n{backup}")

def add_autocorrect_entries():
    confirm = messagebox.askokcancel("Confirmation", "Do you want to continue? This will take just a few seconds.")
    if not confirm:
        return
        
    try:
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open Microsoft Word via COM interface.\nEnsure Word is installed.\nError details: {e}")
        return

    try:
        entries = word.OMathAutoCorrect.Entries
        count = 0
        failed = []
        for pair in AUTOCORRECT_PAIRS:
            shortcode = pair["shortcode"]
            symbol = pair["symbol"]
            try:
                entries.Add(shortcode, symbol)
                count += 1
            except Exception:
                failed.append(shortcode)
                
        word.Quit()
        
        msg = f"Successfully added {count} new math AutoCorrect codes."
        if failed:
            msg += f"\n\nUnable to add {len(failed)} codes:\n" + ", ".join(failed)
        msg += "\n\nIf Word is running, you will need to restart it for changes to take effect."
        messagebox.showinfo("Success", msg)
        
    except Exception as e:
        try:
            word.Quit()
        except Exception:
            pass
        messagebox.showerror("Error", f"An error occurred while adding AutoCorrect entries:\n{e}")

def create_doc_menu(parent):
    doc_win = tk.Toplevel(parent)
    doc_win.title("Docs and Links")
    doc_win.configure(bg=BG_COLOR)
    
    # Center documentation window
    w, h = 320, 260
    sw = doc_win.winfo_screenwidth()
    sh = doc_win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    doc_win.geometry(f"{w}x{h}+{x}+{y}")
    doc_win.resizable(False, False)
    doc_win.transient(parent)
    doc_win.grab_set()

    # Escape key closes sub-window
    doc_win.bind("<Escape>", lambda e: doc_win.destroy())

    # Buttons layout helper
    def open_url(url):
        webbrowser.open(url)

    links = [
        ("Visit math-a11y webpage", "https://daisy.github.io/math-a11y/docs/ms-math/"),
        ("Visit commonly used codes list", "https://daisy.org/msmathcodes"),
        ("Visit proposed new codes wiki", "https://github.com/daisy/math-a11y/wiki/Proposed-New-Math-AutoCorrect-Codes-for-Commonly-Used-Symbols")
    ]

    # Title label
    lbl = tk.Label(doc_win, text="Documentation & Links", bg=BG_COLOR, fg=TEXT_COLOR, font=("Segoe UI", 12, "bold"))
    lbl.pack(pady=10)

    # Style functions
    def make_btn(text, cmd):
        btn = tk.Button(doc_win, text=text, command=cmd, bg=BTN_BG, fg=TEXT_COLOR, 
                        activebackground=SURFACE_COLOR, activeforeground=ACCENT_COLOR,
                        font=("Segoe UI", 9), borderwidth=0, height=2, cursor="hand2")
        btn.pack(fill=tk.X, padx=20, pady=5)
        btn.bind("<Enter>", lambda e: btn.configure(bg=ACCENT_COLOR, fg=BG_COLOR))
        btn.bind("<Leave>", lambda e: btn.configure(bg=BTN_BG, fg=TEXT_COLOR))
        return btn

    for text, url in links:
        make_btn(text, lambda u=url: open_url(u))

    ret_btn = tk.Button(doc_win, text="Return to main menu", command=doc_win.destroy, 
                        bg=SURFACE_COLOR, fg=WARNING_COLOR, activebackground=BG_COLOR,
                        activeforeground=WARNING_COLOR, font=("Segoe UI", 9, "bold"), 
                        borderwidth=0, height=2, cursor="hand2")
    ret_btn.pack(fill=tk.X, padx=20, pady=10)
    ret_btn.bind("<Enter>", lambda e: ret_btn.configure(bg=WARNING_COLOR, fg=BG_COLOR))
    ret_btn.bind("<Leave>", lambda e: ret_btn.configure(bg=SURFACE_COLOR, fg=WARNING_COLOR))

def main():
    root = tk.Tk()
    root.title("DAISY math-a11y working group")
    root.configure(bg=BG_COLOR)

    # Center window
    w, h = 380, 360
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    root.resizable(False, False)

    # Escape key support
    root.bind("<Escape>", lambda e: root.destroy())

    # Warning text frame
    warn_frame = tk.Frame(root, bg=BG_COLOR)
    warn_frame.pack(pady=15, padx=20, fill=tk.X)
    
    warn_lbl = tk.Label(warn_frame, text="This utility is provided in good faith. However, use at your own risk! If you are unsure, exit with no changes.",
                        bg=BG_COLOR, fg=WARNING_COLOR, font=("Segoe UI", 9, "italic"), wraplength=340, justify="center")
    warn_lbl.pack()

    # Style functions
    def make_btn(text, cmd, is_accent=False):
        normal_bg = ACCENT_COLOR if is_accent else BTN_BG
        normal_fg = BG_COLOR if is_accent else TEXT_COLOR
        hover_bg = HOVER_COLOR if is_accent else ACCENT_COLOR
        hover_fg = BG_COLOR
        
        btn = tk.Button(root, text=text, command=cmd, bg=normal_bg, fg=normal_fg, 
                        activebackground=SURFACE_COLOR, activeforeground=TEXT_COLOR,
                        font=("Segoe UI", 9, "bold" if is_accent else "normal"), 
                        borderwidth=0, height=2, cursor="hand2")
        btn.pack(fill=tk.X, padx=30, pady=5)
        
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_bg, fg=hover_fg))
        btn.bind("<Leave>", lambda e: btn.configure(bg=normal_bg, fg=normal_fg))
        return btn

    make_btn("Documentation and links menu", lambda: create_doc_menu(root))
    make_btn("Backup Word autocorrect file", backup_file)
    make_btn("Add new math autocorrect codes", add_autocorrect_entries, is_accent=True)
    make_btn("Restore backup of Word autocorrect file", restore_file)
    
    exit_btn = tk.Button(root, text="Exit", command=root.destroy, 
                         bg=SURFACE_COLOR, fg=WARNING_COLOR, activebackground=BG_COLOR, 
                         activeforeground=WARNING_COLOR, font=("Segoe UI", 9, "bold"), 
                         borderwidth=0, height=2, cursor="hand2")
    exit_btn.pack(fill=tk.X, padx=30, pady=15)
    exit_btn.bind("<Enter>", lambda e: exit_btn.configure(bg=WARNING_COLOR, fg=BG_COLOR))
    exit_btn.bind("<Leave>", lambda e: exit_btn.configure(bg=SURFACE_COLOR, fg=WARNING_COLOR))

    root.mainloop()

if __name__ == "__main__":
    main()
