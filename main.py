"""
GUI.desktop main – Tkinter GUI application for SPIDA ↔ Katapult Comparer.
Launch with:
    python -m gui.main
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import tkintermapview as tkm
from pathlib import Path
import pandas as pd
import json
import traceback
import sys
from PIL import Image, ImageDraw, ImageTk

# Ensure project root is on sys.path so sibling modules (compare.py, spida_writer.py) resolve
if getattr(sys, 'frozen', False):
    # Running as PyInstaller executable
    ROOT_DIR = Path(sys._MEIPASS)
else:
    # Running as Python script
    ROOT_DIR = Path(__file__).resolve().parent.parent

if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

try:
    from .compare import compare, haversine_m
    from .spida_writer import apply_edit
except ImportError:
    from compare import compare, haversine_m
    from spida_writer import apply_edit

# Replace previous import of EditableTree with robust fallback
try:
    from .editable_tree import EditableTree  # when run as module (python -m gui.main)
except ImportError:  # when run as script via path
    from editable_tree import EditableTree


# ------------------------------------------------------------------
# Map icon creation utilities
# ------------------------------------------------------------------

def make_circle_icon(radius_px=4,
                     fill="#00c853",               # inner fill
                     outline="#006644",            # thin rim
                     outline_px=1) -> tk.PhotoImage:
    """Create a circular PhotoImage icon for map markers."""
    size = radius_px * 2 + outline_px * 2
    img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse(
        (outline_px, outline_px, size - outline_px, size - outline_px),
        fill=fill, outline=outline, width=outline_px
    )
    return ImageTk.PhotoImage(img)

# Create circle icons for each match tier (keep global refs to prevent GC)
TIER_CIRCLE_ICONS = {}

def init_circle_icons():
    """Initialize circle icons for each match tier color."""
    global TIER_CIRCLE_ICONS
    tier_colours = {
        "scid":               ("#00c853", "#006644"),  # green - exact SCID match
        "pole_num":           ("#2979ff", "#1a237e"),  # blue - pole number match
        "coord_direct":       ("#ffb300", "#ff6f00"),  # amber - coordinate < 1m
        "coord_spec_verified":("#ff9800", "#e65100"),  # orange - coordinate + spec verified
        "katapult_only":      ("#d500f9", "#6a0080"),  # purple - Katapult only
        "unmatched":          ("#d50000", "#b71c1c"),  # red - unmatched SPIDA
    }
    
    for tier, (fill, outline) in tier_colours.items():
        TIER_CIRCLE_ICONS[tier] = make_circle_icon(
            radius_px=4,
            fill=fill,
            outline=outline,
            outline_px=1
        )


class PoleDetailDialog(ttk.Toplevel):
    """Beautiful custom dialog for displaying pole details with dark mode styling."""
    
    def __init__(self, parent, title, details_text, tier="unmatched"):
        super().__init__(parent)
        
        self.title(title)
        self.geometry("500x650")
        self.resizable(False, False)
        
        # Center the dialog on parent window
        self.transient(parent)
        self.grab_set()
        
        # Dark mode tier colors with enhanced contrast
        tier_colors = {
            "scid": ("#43b581", "#1f2f24"),
            "pole_num": ("#7289da", "#1e2337"), 
            "coord_direct": ("#faa61a", "#2f2518"),
            "coord_spec_verified": ("#fd7e14", "#2f1e13"),
            "katapult_only": ("#ad1aea", "#2b1631"),
            "unmatched": ("#e74c3c", "#2f1a1a")
        }
        
        accent_color, header_bg = tier_colors.get(tier, ("#6c757d", "#22252a"))
        bg_dark = "#2b3035"  # Define the color locally
        
        # Configure dark mode styling
        self.configure(bg=bg_dark)
        
        # Modern header with gradient-like effect
        header_frame = ttk.Frame(self)
        header_frame.pack(fill="x", padx=0, pady=0)
        
        # Colored accent bar with glow effect
        accent_frame = tk.Frame(header_frame, bg=accent_color, height=6)
        accent_frame.pack(fill="x")
        
        # Dark title area with modern typography
        title_frame = tk.Frame(header_frame, bg=header_bg, pady=20)
        title_frame.pack(fill="x")
        
        # Title and subtitle only
        title_label = ttk.Label(
            title_frame, 
            text="QuiC", 
            font=("Segoe UI", 20, "bold"),
            foreground="#ffffff",
            background=header_bg
        )
        title_label.pack(side=LEFT)

        subtitle_label = ttk.Label(
            title_frame, 
            text="SPIDA ↔ Katapult Comparison Tool", 
            font=("Segoe UI", 10),
            foreground="#b9bbbe",
            background=header_bg
        )
        subtitle_label.pack(side=LEFT, padx=(10, 0))
        
        # Subtitle with tier info
        tier_display = tier.replace('_', ' ').title()
        subtitle_label = tk.Label(
            title_frame,
            text=f"Match Type: {tier_display}",
            font=("Segoe UI", 11),
            bg=header_bg,
            fg=accent_color
        )
        subtitle_label.pack(pady=(5, 0))
        
        # Content area with modern dark styling
        content_frame = ttk.Frame(self, padding=25)
        content_frame.pack(fill="both", expand=True)
        
        # Create modern scrollable text widget
        text_frame = ttk.Frame(content_frame)
        text_frame.pack(fill="both", expand=True)
        
        self.text_widget = tk.Text(
            text_frame,
            wrap="word",
            font=("JetBrains Mono", 10),  # Modern monospace font
            bg="#1e2124",
            fg="#dcddde",
            relief="flat",
            borderwidth=0,
            padx=20,
            pady=20,
            cursor="arrow",
            selectbackground=accent_color,
            selectforeground="#ffffff",
            insertbackground=accent_color
        )
        
        # Modern scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.text_widget.yview)
        self.text_widget.configure(yscrollcommand=scrollbar.set)
        
        self.text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Format and insert the text with dark mode colors
        self.format_pole_details(details_text, accent_color)
        
        # Make text widget read-only
        self.text_widget.config(state="disabled")
        
        # Modern button area
        button_frame = ttk.Frame(self, padding=(25, 0, 25, 25))
        button_frame.pack(fill="x")
        
        # Action buttons with modern styling
        button_container = ttk.Frame(button_frame)
        button_container.pack(side="right")
        
        copy_btn = ttk.Button(
            button_container, 
            text="📋 Copy", 
            command=lambda: self.copy_to_clipboard(details_text),
            style="secondary.TButton"
        )
        copy_btn.pack(side="left", padx=(0, 10))
        
        close_btn = ttk.Button(
            button_container, 
            text="✓ Close", 
            command=self.destroy,
            style="success.TButton"
        )
        close_btn.pack(side="left")
        
        # Center on parent
        self.center_on_parent(parent)
        
        # Focus and bind keyboard shortcuts
        self.focus_set()
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Control-c>", lambda e: self.copy_to_clipboard(details_text))
    
    def copy_to_clipboard(self, text):
        """Copy pole details to clipboard."""
        self.clipboard_clear()
        self.clipboard_append(text)
        # Brief visual feedback
        self.status_flash("📋 Copied to clipboard!")
    
    def status_flash(self, message):
        """Show a brief status message."""
        # This could be enhanced with a temporary status label
        self.title(f"Pole Details - {message}")
        self.after(2000, lambda: self.title("Pole Details"))
    
    def format_pole_details(self, details_text, accent_color):
        """Format the pole details with dark mode colors and modern styling."""
        self.text_widget.config(state="normal")
        
        # Configure text tags for dark mode styling
        self.text_widget.tag_configure("header", 
                                     font=("Segoe UI", 14, "bold"), 
                                     foreground=accent_color)
        self.text_widget.tag_configure("section", 
                                     font=("Segoe UI", 12, "bold"), 
                                     foreground="#ffffff")
        self.text_widget.tag_configure("label", 
                                     font=("JetBrains Mono", 10, "bold"), 
                                     foreground="#b9bbbe")
        self.text_widget.tag_configure("value", 
                                     font=("JetBrains Mono", 10), 
                                     foreground="#dcddde")
        self.text_widget.tag_configure("tier", 
                                     font=("Segoe UI", 11, "bold"), 
                                     foreground=accent_color)
        self.text_widget.tag_configure("highlight",
                                     font=("JetBrains Mono", 10, "bold"),
                                     foreground="#43b581")
        
        lines = details_text.split('\n')
        
        for line in lines:
            if not line.strip():
                self.text_widget.insert("end", "\n")
                continue
                
            if line.startswith("🔍 Match Tier:"):
                self.text_widget.insert("end", line + "\n\n", "header")
            elif line.startswith("📊 SPIDA Data:") or line.startswith("📋 Katapult Data:"):
                self.text_widget.insert("end", line + "\n", "section")
            elif line.startswith("Match Distance:"):
                self.text_widget.insert("end", "\n" + line + "\n", "tier")
            elif ":" in line and line.startswith("   "):
                # Split label and value with enhanced formatting
                parts = line.split(":", 1)
                if len(parts) == 2:
                    label = parts[0] + ":"
                    value = parts[1]
                    
                    # Add padding for better alignment
                    self.text_widget.insert("end", "  " + label.ljust(18), "label")
                    
                    # Highlight important values
                    if any(keyword in value.lower() for keyword in ["yes", "true", "match"]):
                        self.text_widget.insert("end", value + "\n", "highlight")
                    else:
                        self.text_widget.insert("end", value + "\n", "value")
                else:
                    self.text_widget.insert("end", line + "\n", "value")
            else:
                self.text_widget.insert("end", line + "\n", "value")
        
        # Add some spacing at the end
        self.text_widget.insert("end", "\n")
    
    def center_on_parent(self, parent):
        """Center the dialog on the parent window with modern positioning."""
        self.update_idletasks()
        
        # Get parent window geometry
        parent.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        # Calculate center position
        dialog_width = self.winfo_width()
        dialog_height = self.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Ensure dialog stays on screen
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = max(50, min(x, screen_width - dialog_width - 50))
        y = max(50, min(y, screen_height - dialog_height - 50))
        
        self.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")


class CompareApp(ttk.Window):
    """Main application window for the SPIDA ↔ Katapult comparison tool."""
    
    def __init__(self):
        print("🏗️ Initializing CompareApp...")
        super().__init__(
            title="QuiC - SPIDA ↔ Katapult Comparer",
            themename="darkly",  # Modern dark bootstrap theme
            size=(1500, 900),
            resizable=(True, True)  # (width_resizable, height_resizable)
        )
        
        print("🎨 Setting up window...")
        # Set app icon and styling
        self.iconify()
        self.deiconify()
        
        # Center window on screen
        self.center_window()

        print("🔘 Initializing circle icons...")
        # Initialize circle icons for map markers
        init_circle_icons()

        print("💾 Initializing data storage...")
        # Initialize data storage
        self.spida_path: Path | None = None
        self.kat_path: Path | None = None
        self.df: pd.DataFrame | None = None
        self.spida_data: dict | None = None  # Store original SPIDA data for editing

        print("🖼️ Setting up icon...")
        # Define icon path
        self.icon_path = ROOT_DIR / 'logo.png'
        self.iconphoto(False, tk.PhotoImage(file=self.icon_path))

        print("🎮 Creating widgets...")
        self.create_widgets()
        print("✅ CompareApp initialization complete")

    # ------------------------------------------------------------------
    # basic window helpers
    # ------------------------------------------------------------------
    def center_window(self):
        """Center the application window on the screen."""
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ------------------------------------------------------------------
    # widget construction
    # ------------------------------------------------------------------
    def create_widgets(self):
        """Create and layout all GUI widgets with modern dark styling."""
        # Configure dark mode colors
        bg_dark = "#2b3035"
        bg_darker = "#1e2124" 
        accent_primary = "#7289da"
        accent_success = "#43b581"
        accent_warning = "#faa61a"
        accent_danger = "#f04747"
        
        # ---------- Modern Header/Toolbar ----------
        header_frame = ttk.Frame(self, style="Dark.TFrame")
        header_frame.pack(fill="x", padx=0, pady=0)
        
        # App title bar
        title_frame = ttk.Frame(header_frame, padding=(20, 15))
        title_frame.pack(fill="x")

        # Title and subtitle only
        title_label = ttk.Label(
            title_frame, 
            text="QuiC", 
            font=("Segoe UI", 20, "bold"),
            foreground="#ffffff",
            background=bg_dark
        )
        title_label.pack(side=LEFT)

        subtitle_label = ttk.Label(
            title_frame, 
            text="SPIDA ↔ Katapult Comparison Tool", 
            font=("Segoe UI", 10),
            foreground="#b9bbbe",
            background=bg_dark
        )
        subtitle_label.pack(side=LEFT, padx=(10, 0))
        
        # Modern toolbar with cards
        toolbar_frame = ttk.Frame(self, padding=(20, 10, 20, 20))
        toolbar_frame.pack(fill="x")
        
        # File operations card
        file_card = ttk.Labelframe(toolbar_frame, text="📁 Data Sources", padding=15)
        file_card.pack(side=LEFT, fill="y", padx=(0, 20))
        
        ttk.Button(
            file_card, 
            text="📊 Load SPIDA JSON", 
            command=self.load_spida, 
            style="info.TButton",
            width=22
        ).pack(pady=(0, 8))
        
        ttk.Button(
            file_card, 
            text="⚡ Load Katapult JSON", 
            command=self.load_katapult, 
            style="info.TButton",
            width=22
        ).pack()
        
        # Analysis card
        analysis_card = ttk.Labelframe(toolbar_frame, text="🔍 Analysis", padding=15)
        analysis_card.pack(side=LEFT, fill="y", padx=(0, 20))
        
        self.compare_btn = ttk.Button(
            analysis_card, 
            text="🚀 Run Comparison", 
            command=self.run_compare, 
            state=DISABLED, 
            style="success.TButton",
            width=22
        )
        self.compare_btn.pack()
        
        # Export card  
        export_card = ttk.Labelframe(toolbar_frame, text="📤 Export", padding=15)
        export_card.pack(side=LEFT, fill="y", padx=(0, 20))
        
        self.export_btn = ttk.Button(
            export_card, 
            text="📈 Export Excel", 
            command=self.export_xlsx, 
            state=DISABLED,
            width=15
        )
        self.export_btn.pack(pady=(0, 8))
        
        self.save_btn = ttk.Button(
            export_card, 
            text="💾 Save SPIDA JSON", 
            command=self.save_new_json, 
            state=DISABLED, 
            style="warning.TButton",
            width=15
        )
        self.save_btn.pack()
        
        # Status card
        status_card = ttk.Labelframe(toolbar_frame, text="📊 Status", padding=15)
        status_card.pack(side=LEFT, fill="both", expand=True, padx=(0, 0))
        
        self.status_label = ttk.Label(
            status_card, 
            text="🎯 Ready to load data files...", 
            font=("Segoe UI", 10),
            foreground="#b9bbbe"
        )
        self.status_label.pack(anchor="w")
        
        # Progress indicator
        self.progress = ttk.Progressbar(
            status_card, 
            mode='indeterminate',
            style="info.Horizontal.TProgressbar"
        )
        self.progress.pack(fill="x", pady=(10, 0))

        # ---------- Main content area with modern panes ----------
        main_container = ttk.Frame(self, padding=(20, 0, 20, 20))
        main_container.pack(fill=BOTH, expand=YES)
        
        main_paned = ttk.PanedWindow(main_container, orient=VERTICAL)
        main_paned.pack(fill=BOTH, expand=YES)

        # ---------- Enhanced Table pane ----------
        table_frame = ttk.Labelframe(main_paned, text="📋 Pole Comparison Data", padding=15)
        
        # Table container with modern styling
        tree_container = ttk.Frame(table_frame, style="Card.TFrame")
        tree_container.pack(fill=BOTH, expand=YES)

        self.tree = EditableTree(
            tree_container,
            editable_cols={"SPIDA Pole Spec", "SPIDA Existing %", "SPIDA Final %", "Com Drop? (SPIDA)"},
            show="headings",
            height=18,
        )
        
        # Enhanced scrollbars
        v_scrollbar = ttk.Scrollbar(tree_container, orient=VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # ---------- Enhanced Map pane ----------
        map_frame = ttk.Labelframe(main_paned, text="🗺️ Interactive Pole Map", padding=15)
        
        # Map controls
        map_controls = ttk.Frame(map_frame)
        map_controls.pack(fill="x", pady=(0, 10))
        
        # Legend with modern styling
        legend_text = ("🟢 SCID Match  🔵 Pole # Match  🟡 Coord <1m  "
                      "🟠 Coord+Spec  🔴 Unmatched  🟣 Katapult Only")
        
        ttk.Label(
            map_controls, 
            text=legend_text, 
            font=("Segoe UI", 9),
            foreground="#b9bbbe"
        ).pack(side=LEFT)
        
        # Map view options
        ttk.Button(
            map_controls, 
            text="🎯 Fit All", 
            command=self.fit_map_to_markers,
            style="secondary.TButton"
        ).pack(side=RIGHT, padx=(5, 0))
        
        # Map widget container
        map_container = ttk.Frame(map_frame, relief="sunken", borderwidth=1)
        map_container.pack(fill=BOTH, expand=YES)
        
        try:
            self.map_widget = tkm.TkinterMapView(
                map_container, 
                width=800, 
                height=350, 
                corner_radius=8
            )
            self.map_widget.pack(fill=BOTH, expand=YES, padx=2, pady=2)
        except Exception as e:
            error_label = ttk.Label(
                map_container, 
                text=f"🗺️ Map widget unavailable: {e}", 
                foreground="#f04747",
                font=("Segoe UI", 10)
            )
            error_label.pack(expand=YES)
            self.map_widget = None

        # Add panes to main container
        main_paned.add(table_frame, weight=3)
        main_paned.add(map_frame, weight=2)
        
    def fit_map_to_markers(self):
        """Helper method to fit map view to all markers."""
        if hasattr(self, 'map_widget') and self.map_widget:
            # This would be called from update_map() when we have marker data
            pass

    # ------------------------------------------------------------------
    # file loading helpers
    # ------------------------------------------------------------------
    def load_spida(self):
        filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select SPIDA JSON file", filetypes=filetypes)
        if filename:
            try:
                self.progress.start(10)
                self.status_label.config(text="📊 Loading SPIDA data...")
                self.update()
                
                self.spida_path = Path(filename)
                with open(self.spida_path, "r", encoding="utf-8") as f:
                    self.spida_data = json.load(f)
                
                self.progress.stop()
                self.status_label.config(text=f"✅ SPIDA loaded: {self.spida_path.name}")
                self.check_ready_to_compare()
            except Exception as e:
                self.progress.stop()
                messagebox.showerror("Error", f"Failed to load SPIDA file:\n{e}")
                self.spida_path = None
                self.spida_data = None
                self.status_label.config(text="❌ Failed to load SPIDA file")

    def load_katapult(self):
        filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select Katapult JSON file", filetypes=filetypes)
        if filename:
            try:
                self.progress.start(10)
                self.status_label.config(text="⚡ Loading Katapult data...")
                self.update()
                
                self.kat_path = Path(filename)
                with open(self.kat_path, "r", encoding="utf-8") as f:
                    json.load(f)
                
                self.progress.stop()
                self.status_label.config(text=f"✅ Katapult loaded: {self.kat_path.name}")
                self.check_ready_to_compare()
            except Exception as e:
                self.progress.stop()
                messagebox.showerror("Error", f"Failed to load Katapult file:\n{e}")
                self.kat_path = None
                self.status_label.config(text="❌ Failed to load Katapult file")

    def check_ready_to_compare(self):
        if self.spida_path and self.kat_path:
            self.compare_btn.config(state=NORMAL)
            self.status_label.config(text="🚀 Ready to compare - click Run Comparison button")
        else:
            self.compare_btn.config(state=DISABLED)

    # ------------------------------------------------------------------
    # main logic
    # ------------------------------------------------------------------
    def run_compare(self):
        if not self.spida_path or not self.kat_path:
            messagebox.showwarning("Warning", "Please load both SPIDA and Katapult files first.")
            return
        try:
            self.progress.start(10)
            self.status_label.config(text="🔍 Analyzing and comparing datasets...")
            self.update()
            
            self.df = compare(self.spida_path, self.kat_path)
            
            # ---- rename / reorder columns per README ----
            rename_map = {
                "SCID": "SPIDA SCID #",
                "Katapult SCID #": "Katapult SCID #",
                "SPIDA Pole #": "SPIDA Pole #",
                "Katapult Pole #": "Katapult Pole #",
                "SPIDA Spec": "SPIDA Pole Spec",
                "Katapult Spec": "Katapult Pole Spec",
                "SPIDA Existing %": "SPIDA Existing %",
                "Katapult Existing %": "Katapult Existing %",
                "SPIDA Final %": "SPIDA Final %",
                "Katapult Final %": "Katapult Final %",
                "SPIDA Charter Drop": "Com Drop? (SPIDA)",
                "Com Drop?": "Com Drop? (Kat)",  # Fix for service drop column
            }
            self.df = self.df.rename(columns=rename_map)
            wanted = [
                "SPIDA SCID #",
                "Katapult SCID #",
                "SPIDA Pole #",
                "Katapult Pole #",
                "SPIDA Pole Spec",
                "Katapult Pole Spec",
                "SPIDA Existing %",
                "Katapult Existing %",
                "SPIDA Final %",
                "Katapult Final %",
                "Com Drop? (SPIDA)",
                "Com Drop? (Kat)",
            ]
            self.df = self.df.reindex(columns=wanted + [c for c in self.df.columns if c not in wanted])
            
            # Recalculate match indicators after column renaming with data cleaning
            def clean_value(val):
                """Clean and normalize values for comparison."""
                if pd.isna(val) or val is None:
                    return None
                val_str = str(val).strip()
                return val_str if val_str else None
            
            def normalize_charter_drop(val):
                """Normalize charter drop values for comparison (True/False <-> Yes/No)."""
                if pd.isna(val) or val is None:
                    return None
                val_str = str(val).strip().lower()
                if val_str in ['true', 'yes']:
                    return 'yes'
                elif val_str in ['false', 'no']:
                    return 'no'
                return val_str
            
            def normalize_spec(val):
                """Normalize pole specifications for comparison by removing formatting differences."""
                if pd.isna(val) or val is None:
                    return None
                val_str = str(val).strip()
                # Remove prime symbols and normalize spacing
                normalized = val_str.replace("′", "").replace("'", "")
                # Normalize multiple spaces to single spaces
                normalized = " ".join(normalized.split())
                return normalized.lower() if normalized else None
            
            # Recalculate match indicators with cleaned data
            self.df["Spec Match"] = self.df.apply(
                lambda row: normalize_spec(row.get("SPIDA Pole Spec")) == normalize_spec(row.get("Katapult Pole Spec")), 
                axis=1
            )
            self.df["Existing % Match"] = self.df.apply(
                lambda row: clean_value(row.get("SPIDA Existing %")) == clean_value(row.get("Katapult Existing %")), 
                axis=1
            )
            self.df["Final % Match"] = self.df.apply(
                lambda row: clean_value(row.get("SPIDA Final %")) == clean_value(row.get("Katapult Final %")), 
                axis=1
            )
            self.df["Charter Drop Match"] = self.df.apply(
                lambda row: normalize_charter_drop(row.get("Com Drop? (SPIDA)")) == normalize_charter_drop(row.get("Com Drop? (Kat)")), 
                axis=1
            )
            
            # track original editable cols
            for col in ["SPIDA Pole Spec", "SPIDA Existing %", "SPIDA Final %", "Com Drop? (SPIDA)"]:
                if col in self.df.columns:
                    self.df[f"__orig_{col}"] = self.df[col].copy()
            
            # Update UI with results
            self.progress.stop()
            self.status_label.config(text="🎨 Updating interface...")
            self.update()
            
            # refresh UI
            self.populate_tree()
            self.update_map()
            self.export_btn.config(state=NORMAL)
            self.save_btn.config(state=NORMAL)
            
            # Final status with statistics
            total_poles = len(self.df)
            matched_poles = len(self.df[self.df.get('Match Tier', 'unmatched') != 'unmatched'])
            match_rate = (matched_poles / total_poles * 100) if total_poles > 0 else 0
            
            self.status_label.config(text=f"✅ Analysis complete: {total_poles} poles, {match_rate:.1f}% matched")
            
            # Remove the original Charter Drop Match column entirely (no longer needed)
            if "Charter Drop Match" in self.df.columns:
                self.df = self.df.drop(columns=["Charter Drop Match"])
                
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("Comparison Error", f"Error during comparison:\n{e}\n\n{traceback.format_exc()}")
            self.status_label.config(text="❌ Comparison failed")

    # ------------------------------------------------------------------
    # tree handling
    # ------------------------------------------------------------------
    def populate_tree(self):
        if self.df is None:
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Filter out internal/analysis columns from display
        display_cols = [c for c in self.df.columns if not c.startswith("__")]
        
        # Hide coordinate columns and match analysis columns from table view
        hidden_cols = [
            "SPIDA Coord", "Katapult Coord",  # Coordinate data
            "Match Tier", "Match Distance (m)",  # Match analysis
            "Spec Match", "Existing % Match", "Final % Match", "Charter Drop Match",  # Comparison flags
            "SCID Status", "Pole # Status",  # Status columns
            "Matched by Coord"  # Legacy flag
        ]
        
        visible_cols = [c for c in display_cols if c not in hidden_cols]
        
        self.tree["columns"] = visible_cols
        self.tree["show"] = "headings"
        
        for col in visible_cols:
            self.tree.heading(col, text=col)
            if "%" in col:
                width = 100
            elif "Coord" in col:
                width = 150
            elif col == "SPIDA SCID #":
                width = 80
            elif "Pole #" in col:  # Give pole number columns more space
                width = 120
            else:
                width = 150
            self.tree.column(col, width=width, anchor=CENTER)
        
        # Define column pairs and their corresponding match indicators
        mismatch_indicators = {
            "SPIDA Pole Spec": "Spec Match",
            "Katapult Pole Spec": "Spec Match",
            "SPIDA Existing %": "Existing % Match", 
            "Katapult Existing %": "Existing % Match",
            "SPIDA Final %": "Final % Match",
            "Katapult Final %": "Final % Match",
            "Com Drop? (SPIDA)": "Charter Drop Match",
            "Com Drop? (Kat)": "Charter Drop Match"
        }
        
        for idx, row in self.df.iterrows():
            values = []
            
            for col in visible_cols:
                raw_value = row[col] if pd.notna(row[col]) and row[col] is not None else ""
                display_value = str(raw_value)
                
                # Convert True/False to Yes/No for Com Drop (SPIDA) column
                if col == "Com Drop? (SPIDA)":
                    if str(raw_value).lower() == "true":
                        display_value = "Yes"
                    elif str(raw_value).lower() == "false":
                        display_value = "No"
                    elif raw_value == "":
                        display_value = ""
                
                # Check if this specific column should be highlighted for mismatches
                if col in mismatch_indicators:
                    match_col = mismatch_indicators[col]
                    if match_col in self.df.columns:
                        is_match = row.get(match_col, True)
                        if is_match is False or (isinstance(is_match, str) and is_match.lower() == "false"):
                            display_value = f"❌ {display_value}" if display_value else "❌ [Empty]"
                
                values.append(display_value)
            
            # Insert the row with normal styling
            self.tree.insert("", END, values=values)

    # ------------------------------------------------------------------
    # map handling with rich visual grammar
    # ------------------------------------------------------------------
    
    def _mk_circle(self, lat, lon, tooltip_text: str, tier="unmatched"):
        """
        Place a small circle marker with an optional click-tooltip.
        Uses pre-created PhotoImage icons for clean, properly sized markers.
        """
        if not self.map_widget:
            return None
            
        # Get the appropriate circle icon for this tier
        icon = TIER_CIRCLE_ICONS.get(tier, TIER_CIRCLE_ICONS.get("unmatched"))
        
        marker = self.map_widget.set_marker(
            lat, lon,
            icon=icon,                     # Use the circular PhotoImage
            text=None                      # No text overlay
        )
        
        # Add tooltip on click
        if tooltip_text:
            def show_tooltip(marker_obj):  # Accept the marker argument passed by tkintermapview
                PoleDetailDialog(self, "Pole Details", tooltip_text, tier)
            marker.command = show_tooltip
            
        return marker

    def update_map(self):
        """Update map with color-coded markers, connecting lines, and enhanced legend."""
        if not self.map_widget or self.df is None:
            return
        
        try:
            # Clear existing markers and paths
            self.map_widget.delete_all_marker()
            self.map_widget.delete_all_path()
            
            marker_lats: list[float] = []
            marker_lons: list[float] = []
            edges = []  # Store matched pairs for drawing lines
            
            stats = {
                "scid": 0,
                "pole_num": 0,
                "coord_direct": 0,
                "coord_spec_verified": 0,
                "unmatched_spida": 0,
                "katapult_only": 0
            }
            
            # Color palette for connecting lines (since we can't extract from PhotoImage)
            tier_line_colors = {
                "scid": "#00c853",
                "pole_num": "#2979ff", 
                "coord_direct": "#ffb300",
                "coord_spec_verified": "#ff9800",
                "katapult_only": "#d500f9",
                "unmatched": "#d50000"
            }
            
            for _, row in self.df.iterrows():
                tier = row.get("Match Tier", "unmatched")
                spida_coord = row.get("SPIDA Coord")
                kat_coord = row.get("Katapult Coord")
                
                # Build comprehensive tooltip info from both datasets
                spida_scid = row.get("SPIDA SCID #") or "—"
                spida_spec = row.get("SPIDA Pole Spec") or "—"
                spida_pole = row.get("SPIDA Pole #") or "—"
                spida_existing = row.get("SPIDA Existing %") or "—"
                spida_final = row.get("SPIDA Final %") or "—"
                spida_charter = row.get("Com Drop? (SPIDA)") or "—"
                
                kat_scid = row.get("Katapult SCID #") or "—"
                kat_spec = row.get("Katapult Pole Spec") or "—"
                kat_pole = row.get("Katapult Pole #") or "—"
                kat_existing = row.get("Katapult Existing %") or "—"
                kat_final = row.get("Katapult Final %") or "—"
                kat_charter = row.get("Com Drop? (Kat)") or "—"
                
                # Add match distance info for coordinate matches
                match_info = ""
                if tier in ["coord_direct", "coord_spec_verified"]:
                    distance = row.get("Match Distance (m)")
                    if distance:
                        match_info = f"\n\nMatch Distance: {distance}m"
                
                # Create comprehensive tooltip showing both datasets
                def create_comprehensive_tooltip():
                    tooltip_parts = []
                    
                    # Header with match tier
                    tooltip_parts.append(f"🔍 Match Tier: {tier.replace('_', ' ').title()}")
                    
                    # SPIDA data section
                    if spida_scid != "—" or spida_spec != "—" or spida_pole != "—":
                        tooltip_parts.append("\n📊 SPIDA Data:")
                        tooltip_parts.append(f"   SCID: {spida_scid}")
                        tooltip_parts.append(f"   Pole #: {spida_pole}")
                        tooltip_parts.append(f"   Spec: {spida_spec}")
                        tooltip_parts.append(f"   Existing %: {spida_existing}")
                        tooltip_parts.append(f"   Final %: {spida_final}")
                        tooltip_parts.append(f"   Charter Drop: {spida_charter}")
                    
                    # Katapult data section  
                    if kat_scid != "—" or kat_spec != "—" or kat_pole != "—":
                        tooltip_parts.append("\n📋 Katapult Data:")
                        tooltip_parts.append(f"   SCID: {kat_scid}")
                        tooltip_parts.append(f"   Pole #: {kat_pole}")
                        tooltip_parts.append(f"   Spec: {kat_spec}")
                        tooltip_parts.append(f"   Existing %: {kat_existing}")
                        tooltip_parts.append(f"   Final %: {kat_final}")
                        tooltip_parts.append(f"   Charter Drop: {kat_charter}")
                    
                    return "\n".join(tooltip_parts) + match_info
                
                comprehensive_tooltip = create_comprehensive_tooltip()
                
                # Place SPIDA marker if coordinates exist
                if spida_coord:
                    self._mk_circle(spida_coord[0], spida_coord[1], comprehensive_tooltip, tier)
                    marker_lats.append(spida_coord[0])
                    marker_lons.append(spida_coord[1])
                
                # Place Katapult marker if coordinates exist (only if different from SPIDA)
                if kat_coord and (not spida_coord or kat_coord != spida_coord):
                    kat_tier = tier if tier != "unmatched" else "katapult_only"
                    self._mk_circle(kat_coord[0], kat_coord[1], comprehensive_tooltip, kat_tier)
                    marker_lats.append(kat_coord[0])
                    marker_lons.append(kat_coord[1])
                
                # Collect matched pairs for drawing connecting lines
                if spida_coord and kat_coord and tier in ["scid", "pole_num", "coord_direct", "coord_spec_verified"]:
                    line_color = tier_line_colors.get(tier, "#gray")
                    edges.append((spida_coord, kat_coord, line_color, tier))
                
                # Update statistics
                if tier in stats:
                    stats[tier] += 1
                elif tier == "katapult_only":
                    stats["katapult_only"] += 1
                else:
                    stats["unmatched_spida"] += 1
            
            # Draw connecting lines between matched pairs
            for spida_coord, kat_coord, color, tier in edges:
                # Use different line styles for different tiers
                width = 3 if tier == "scid" else 2
                self.map_widget.set_path(
                    [spida_coord, kat_coord], 
                    width=width, 
                    color=color
                )
            
            # Auto-zoom to fit all markers
            if marker_lats and marker_lons:
                centroid_lat = sum(marker_lats) / len(marker_lats)
                centroid_lon = sum(marker_lons) / len(marker_lons)
                lat_span = max(marker_lats) - min(marker_lats)
                lon_span = max(marker_lons) - min(marker_lons)
                max_span = max(lat_span, lon_span)
                
                # Add padding around the markers
                desired_coverage = 0.8  # Use 80% of viewport
                if max_span == 0:
                    zoom = 18
                else:
                    import math
                    zoom = int(max(3, min(18, math.log2(360.0 / (max_span / desired_coverage)))))
                
                self.map_widget.set_position(centroid_lat, centroid_lon)
                self.map_widget.set_zoom(zoom)
            
            total_poles = len(self.df)
            print(f"📍 Map updated: {total_poles} poles")
            print(f"   🎯 SCID matches: {stats['scid']}")
            print(f"   🏷️  Pole # matches: {stats['pole_num']}")
            print(f"   📍 Coord direct: {stats['coord_direct']}")
            print(f"   🔍 Coord+spec: {stats['coord_spec_verified']}")
            print(f"   ❌ Unmatched SPIDA: {stats['unmatched_spida']}")
            print(f"   💜 Katapult only: {stats['katapult_only']}")
            print(f"   🔗 Connecting lines: {len(edges)}")
            
        except Exception as e:
            print(f"Error updating map: {e}")
            import traceback
            traceback.print_exc()

    # ------------------------------------------------------------------
    # export / save helpers
    # ------------------------------------------------------------------
    def export_xlsx(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data to export. Please run comparison first.")
            return
        filename = filedialog.asksaveasfilename(
            title="Save Excel file", 
            defaultextension=".xlsx", 
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                self.progress.start(10)
                self.status_label.config(text="📈 Exporting to Excel...")
                self.update()
                
                export_cols = [c for c in self.df.columns if not c.startswith("__") and "Coord" not in c]
                self.df[export_cols].to_excel(filename, index=False)
                
                self.progress.stop()
                self.status_label.config(text=f"✅ Excel exported: {Path(filename).name}")
                messagebox.showinfo("Export Complete", f"📊 Data exported successfully to:\n{filename}")
            except Exception as e:
                self.progress.stop()
                self.status_label.config(text="❌ Excel export failed")
                messagebox.showerror("Export Error", f"Failed to export Excel file:\n{e}")

    def save_new_json(self):
        if self.df is None or self.spida_data is None:
            messagebox.showwarning("Warning", "No data to save. Please run comparison first.")
            return
        filename = filedialog.asksaveasfilename(
            title="Save updated SPIDA JSON", 
            defaultextension=".json", 
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        if not filename:
            return
        try:
            self.progress.start(10)
            self.status_label.config(text="💾 Saving SPIDA JSON...")
            self.update()
            
            updated_spida = json.loads(json.dumps(self.spida_data))
            current_data = [self.tree.item(item, "values") for item in self.tree.get_children()]
            if current_data:
                visible_cols = [c for c in self.df.columns if not c.startswith("__") and "Coord" not in c]
                for i, values in enumerate(current_data):
                    if i < len(self.df):
                        for j, col in enumerate(visible_cols):
                            if j < len(values):
                                self.df.iat[i, self.df.columns.get_loc(col)] = values[j]
            changes_made = 0
            editable_cols = ["SPIDA Pole Spec", "SPIDA Existing %", "SPIDA Final %", "Com Drop? (SPIDA)"]
            for _, row in self.df.iterrows():
                scid = row["SPIDA SCID #"]
                for col in editable_cols:
                    orig = f"__orig_{col}"
                    if orig in self.df.columns and str(row[col]) != str(row[orig]):
                        apply_edit(updated_spida, scid, col, str(row[col]))
                        changes_made += 1
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(updated_spida, f, indent=2, ensure_ascii=False)
            
            self.progress.stop()
            self.status_label.config(text=f"✅ SPIDA JSON saved: {Path(filename).name}")
            messagebox.showinfo(
                "Save Complete", 
                f"✅ Updated SPIDA JSON saved successfully!\n\n"
                f"📄 File: {filename}\n"
                f"🔧 Changes applied: {changes_made}"
            )
        except Exception as e:
            self.progress.stop()
            self.status_label.config(text="❌ SPIDA save failed")
            messagebox.showerror("Save Error", f"Failed to save JSON file:\n{e}\n\n{traceback.format_exc()}")


# ----------------------------------------------------------------------
# entry point
# ----------------------------------------------------------------------

def main():
    print("🚀 Starting QuiC application...")
    try:
        print("📱 Creating GUI application...")
        app = CompareApp()
        print("🎯 Starting mainloop...")
        app.mainloop()
        print("✅ Application closed normally")
    except Exception as e:
        print(f"❌ Application error: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main() 