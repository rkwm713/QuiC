from __future__ import annotations

"""
editable_tree.py – Custom Treeview widget with inline editing capability.
Double-click a cell to edit it in place.
"""

import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class EditableTree(ttk.Treeview):
    """Double-click a cell → inline Entry → <Return> saves."""
    
    def __init__(self, *args, editable_cols=None, **kwargs):
        super().__init__(*args, **kwargs)
        self.editable_cols = editable_cols or set()
        self.bind("<Double-1>", self._edit_cell)
        self.current_entry = None

    def _edit_cell(self, evt):
        """Handle double-click event to start editing a cell."""
        region = self.identify_region(evt.x, evt.y)
        if region != "cell":
            return
            
        row_id = self.identify_row(evt.y)
        col_id = self.identify_column(evt.x)  # format '#3'
        
        if not row_id or not col_id:
            return
            
        try:
            col_idx = int(col_id[1:]) - 1
        except (ValueError, IndexError):
            return
            
        if col_idx < 0 or col_idx >= len(self["columns"]):
            return
            
        heading = self["columns"][col_idx]
        if heading not in self.editable_cols:
            return

        # Destroy any existing entry widget
        if self.current_entry:
            self.current_entry.destroy()
            self.current_entry = None

        # Get cell position and current value
        try:
            x, y, w, h = self.bbox(row_id, col_id)
        except tk.TclError:
            return
            
        values = self.item(row_id, "values")
        if col_idx >= len(values):
            return
            
        old_value = values[col_idx]

        # Create entry widget for editing
        entry = ttk.Entry(self)
        entry.insert(0, str(old_value))
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        entry.select_range(0, tk.END)
        
        self.current_entry = entry

        def save_edit(event=None):
            """Save the edited value back to the tree."""
            if not entry.winfo_exists():
                return
                
            new_val = entry.get()
            vals = list(self.item(row_id, "values"))
            vals[col_idx] = new_val
            self.item(row_id, values=vals)
            
            entry.destroy()
            self.current_entry = None
            self.focus_set()

        def cancel_edit(event=None):
            """Cancel editing without saving."""
            if entry.winfo_exists():
                entry.destroy()
            self.current_entry = None
            self.focus_set()

        # Bind events
        entry.bind("<Return>", save_edit)
        entry.bind("<Escape>", cancel_edit)
        entry.bind("<FocusOut>", cancel_edit)
        
        # Handle Tab to move to next editable cell
        entry.bind("<Tab>", lambda e: self._move_to_next_editable(row_id, col_idx, save_edit))

    def _move_to_next_editable(self, current_row, current_col, save_callback):
        """Move to the next editable cell when Tab is pressed."""
        save_callback()
        
        # Find next editable column in current row
        for col_idx in range(current_col + 1, len(self["columns"])):
            if self["columns"][col_idx] in self.editable_cols:
                # Simulate double-click on next cell
                col_id = f"#{col_idx + 1}"
                try:
                    x, y, w, h = self.bbox(current_row, col_id)
                    # Create a fake event
                    fake_event = type('Event', (), {})()
                    fake_event.x = x + w // 2
                    fake_event.y = y + h // 2
                    self._edit_cell(fake_event)
                    return
                except tk.TclError:
                    pass
        
        # If no more editable columns in current row, move to next row
        children = self.get_children()
        try:
            current_idx = children.index(current_row)
            if current_idx + 1 < len(children):
                next_row = children[current_idx + 1]
                # Find first editable column in next row
                for col_idx, col_name in enumerate(self["columns"]):
                    if col_name in self.editable_cols:
                        col_id = f"#{col_idx + 1}"
                        try:
                            x, y, w, h = self.bbox(next_row, col_id)
                            fake_event = type('Event', (), {})()
                            fake_event.x = x + w // 2
                            fake_event.y = y + h // 2
                            self._edit_cell(fake_event)
                            return
                        except tk.TclError:
                            pass
        except ValueError:
            pass

    def destroy(self):
        """Clean up when widget is destroyed."""
        if self.current_entry:
            self.current_entry.destroy()
            self.current_entry = None
        super().destroy()