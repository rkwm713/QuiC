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

# Ensure project root is on sys.path so sibling modules (compare.py, spida_writer.py) resolve
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


class CompareApp(ttk.Window):
    """Main application window for the SPIDA ↔ Katapult comparison tool."""
    
    def __init__(self):
        super().__init__(
            title="QuiC",
            themename="flatly",  # Modern bootstrap theme
            size=(1400, 800)
        )
        # Center window on screen
        self.center_window()

        # Initialize data storage
        self.spida_path: Path | None = None
        self.kat_path: Path | None = None
        self.df: pd.DataFrame | None = None
        self.spida_data: dict | None = None  # Store original SPIDA data for editing

        self.create_widgets()

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
        """Create and layout all GUI widgets."""
        # ---------- Top toolbar ----------
        toolbar = ttk.Frame(self, padding=(10, 8))
        toolbar.pack(fill="x", side="top")

        ttk.Button(toolbar, text="📁 Open SPIDA", command=self.load_spida, style="info.TButton").pack(side=LEFT, padx=(0, 5))
        ttk.Button(toolbar, text="📁 Open Katapult", command=self.load_katapult, style="info.TButton").pack(side=LEFT, padx=(0, 15))

        self.compare_btn = ttk.Button(toolbar, text="🔍 Compare", command=self.run_compare, state=DISABLED, style="success.TButton")
        self.compare_btn.pack(side=LEFT, padx=(0, 15))

        self.export_btn = ttk.Button(toolbar, text="📊 Export Excel", command=self.export_xlsx, state=DISABLED)
        self.export_btn.pack(side=LEFT, padx=(0, 5))

        self.save_btn = ttk.Button(toolbar, text="💾 Save New SPIDA JSON", command=self.save_new_json, state=DISABLED, style="warning.TButton")
        self.save_btn.pack(side=LEFT, padx=(0, 15))

        self.status_label = ttk.Label(toolbar, text="Select both SPIDA and Katapult files to begin comparison...", font=("Segoe UI", 9))
        self.status_label.pack(side=LEFT, padx=(10, 0))

        # ---------- Main panes ----------
        main_paned = ttk.PanedWindow(self, orient=VERTICAL)
        main_paned.pack(fill=BOTH, expand=YES, padx=10, pady=(0, 10))

        # ---------- Table pane ----------
        table_frame = ttk.Labelframe(main_paned, text="Pole Comparison Data", padding=10)
        tree_container = ttk.Frame(table_frame)
        tree_container.pack(fill=BOTH, expand=YES)

        self.tree = EditableTree(
            tree_container,
            editable_cols={"SPIDA Pole Spec", "SPIDA Existing %", "SPIDA Final %", "Com Drop? (SPIDA)"},
            show="headings",
            height=15,
        )
        v_scrollbar = ttk.Scrollbar(tree_container, orient=VERTICAL, command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(tree_container, orient=HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # ---------- Map pane ----------
        map_frame = ttk.Labelframe(main_paned, text="Pole Locations", padding=10)
        try:
            self.map_widget = tkm.TkinterMapView(map_frame, width=800, height=300, corner_radius=0)
            self.map_widget.pack(fill=BOTH, expand=YES)
            legend_frame = ttk.Frame(map_frame)
            legend_frame.pack(fill="x", pady=(5, 0))
            ttk.Label(legend_frame, text="🟢 SPIDA Only    🔵 Katapult Only    🟦 Overlapping (< 1m)", font=("Segoe UI", 9, "italic")).pack(side=RIGHT)
        except Exception as e:
            ttk.Label(map_frame, text=f"Map widget unavailable: {e}", foreground="red").pack(expand=YES)
            self.map_widget = None

        main_paned.add(table_frame, weight=3)
        main_paned.add(map_frame, weight=2)

    # ------------------------------------------------------------------
    # file loading helpers
    # ------------------------------------------------------------------
    def load_spida(self):
        filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select SPIDA JSON file", filetypes=filetypes)
        if filename:
            try:
                self.spida_path = Path(filename)
                with open(self.spida_path, "r", encoding="utf-8") as f:
                    self.spida_data = json.load(f)
                self.status_label.config(text=f"SPIDA loaded: {self.spida_path.name}")
                self.check_ready_to_compare()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load SPIDA file:\n{e}")
                self.spida_path = None
                self.spida_data = None

    def load_katapult(self):
        filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
        filename = filedialog.askopenfilename(title="Select Katapult JSON file", filetypes=filetypes)
        if filename:
            try:
                self.kat_path = Path(filename)
                with open(self.kat_path, "r", encoding="utf-8") as f:
                    json.load(f)
                self.status_label.config(text=f"Katapult loaded: {self.kat_path.name}")
                self.check_ready_to_compare()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Katapult file:\n{e}")
                self.kat_path = None

    def check_ready_to_compare(self):
        if self.spida_path and self.kat_path:
            self.compare_btn.config(state=NORMAL)
            self.status_label.config(text="Ready to compare - click Compare button")
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
            self.status_label.config(text="Comparing files...")
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
                "Katapult Charter Drop": "Com Drop? (Kat)",
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
            # track original editable cols
            for col in ["SPIDA Pole Spec", "SPIDA Existing %", "SPIDA Final %", "Com Drop? (SPIDA)"]:
                if col in self.df.columns:
                    self.df[f"__orig_{col}"] = self.df[col].copy()
            # refresh UI
            self.populate_tree()
            self.update_map()
            self.export_btn.config(state=NORMAL)
            self.save_btn.config(state=NORMAL)
            self.status_label.config(text=f"Comparison complete - {len(self.df)} poles found")
            # Remove the original Charter Drop Match column entirely (no longer needed)
            if "Charter Drop Match" in self.df.columns:
                self.df = self.df.drop(columns=["Charter Drop Match"])
        except Exception as e:
            messagebox.showerror("Comparison Error", f"Error during comparison:\n{e}\n\n{traceback.format_exc()}")
            self.status_label.config(text="Comparison failed")

    # ------------------------------------------------------------------
    # tree handling
    # ------------------------------------------------------------------
    def populate_tree(self):
        if self.df is None:
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        display_cols = [c for c in self.df.columns if not c.startswith("__")]
        visible_cols = [c for c in display_cols if c not in ["SPIDA Coord", "Katapult Coord"]]
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
        for _, row in self.df.iterrows():
            self.tree.insert("", END, values=[str(row[c]) if pd.notna(row[c]) and row[c] is not None else "" for c in visible_cols])

    # ------------------------------------------------------------------
    # map handling (same as before)
    # ------------------------------------------------------------------
    def update_map(self):
        if not self.map_widget or self.df is None:
            return
        try:
            self.map_widget.delete_all_marker()
            marker_lats: list[float] = []
            marker_lons: list[float] = []
            marker_count = 0
            for _, row in self.df.iterrows():
                scid = row["SPIDA SCID #"]
                spida_coord = row["SPIDA Coord"]
                kat_coord = row["Katapult Coord"]
                if spida_coord and kat_coord:
                    dist = haversine_m(spida_coord, kat_coord)
                    if dist < 1.0:
                        self.map_widget.set_marker(spida_coord[0], spida_coord[1], text="●", marker_color_circle="darkturquoise", marker_color_outside="teal", text_color="darkturquoise")
                        marker_count += 1
                    else:
                        # Show both markers when far apart, but use blue for both instead of green
                        self.map_widget.set_marker(spida_coord[0], spida_coord[1], text="●", marker_color_circle="lightblue", marker_color_outside="blue", text_color="lightblue")
                        self.map_widget.set_marker(kat_coord[0], kat_coord[1], text="●", marker_color_circle="darkblue", marker_color_outside="darkblue", text_color="darkblue")
                        marker_count += 2
                elif spida_coord:
                    # Show SPIDA-only markers but use blue instead of green
                    self.map_widget.set_marker(spida_coord[0], spida_coord[1], text="●", marker_color_circle="lightblue", marker_color_outside="blue", text_color="lightblue")
                    marker_count += 1
                elif kat_coord:
                    # Show Katapult-only markers in dark blue to distinguish from SPIDA
                    self.map_widget.set_marker(kat_coord[0], kat_coord[1], text="●", marker_color_circle="darkblue", marker_color_outside="darkblue", text_color="darkblue")
                    marker_count += 1
                for c in (spida_coord, kat_coord):
                    if c:
                        marker_lats.append(c[0])
                        marker_lons.append(c[1])
            if marker_lats and marker_lons:
                centroid_lat = sum(marker_lats) / len(marker_lats)
                centroid_lon = sum(marker_lons) / len(marker_lons)
                lat_span = max(marker_lats) - min(marker_lats)
                lon_span = max(marker_lons) - min(marker_lons)
                max_span = max(lat_span, lon_span)
                desired_tiles = 0.75
                if max_span == 0:
                    zoom = 18
                else:
                    import math
                    zoom = int(max(3, min(18, math.log2(360.0 / (max_span * desired_tiles)))))
                self.map_widget.set_position(centroid_lat, centroid_lon)
                self.map_widget.set_zoom(zoom)
            print(f"Map updated with {marker_count} markers")
        except Exception as e:
            print(f"Error updating map: {e}")

    # ------------------------------------------------------------------
    # export / save helpers
    # ------------------------------------------------------------------
    def export_xlsx(self):
        if self.df is None:
            messagebox.showwarning("Warning", "No data to export. Please run comparison first.")
            return
        filename = filedialog.asksaveasfilename(title="Save Excel file", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
        if filename:
            try:
                export_cols = [c for c in self.df.columns if not c.startswith("__") and "Coord" not in c]
                self.df[export_cols].to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export Excel file:\n{e}")

    def save_new_json(self):
        if self.df is None or self.spida_data is None:
            messagebox.showwarning("Warning", "No data to save. Please run comparison first.")
            return
        filename = filedialog.asksaveasfilename(title="Save updated SPIDA JSON", defaultextension=".json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if not filename:
            return
        try:
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
            messagebox.showinfo("Success", f"Updated SPIDA JSON saved to {filename}\n{changes_made} changes applied")
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save JSON file:\n{e}\n\n{traceback.format_exc()}")


# ----------------------------------------------------------------------
# entry point
# ----------------------------------------------------------------------

def main():
    try:
        CompareApp().mainloop()
    except Exception as e:
        print(f"Application error: {e}")
        traceback.print_exc()


if __name__ == "__main__":
    main() 