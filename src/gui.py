"""
GUI module for Invoice Management System.
Uses CustomTkinter for modern UI.
"""

import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from typing import List, Dict, Any, Optional
from datetime import datetime

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.database import Database
from src.excel_parser import ExcelParser, DataValidator


# Set appearance mode and default color theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class InvoiceTable(ctk.CTkFrame):
    """Custom table widget for displaying invoice data."""

    # Display column configuration (subset of columns for readability)
    DISPLAY_COLUMNS = [
        ("id", "ID", 50),
        ("lab_name", "Lab", 60),
        ("tajan_bidding_tracking_number", "Tracking #", 150),
        ("amazon_test_request", "Test Request", 120),
        ("asin", "ASIN", 100),
        ("test_service_type", "Test Type", 150),
        ("invoice_number", "Invoice #", 100),
        ("amount_usd", "Amount (USD)", 100),
        ("invoice_date", "Invoice Date", 100),
        ("report_number", "Report #", 120),
    ]

    def __init__(self, parent, on_select_callback=None, show_validation_error=False, **kwargs):
        super().__init__(parent, **kwargs)

        self.on_select_callback = on_select_callback
        self.show_validation_error = show_validation_error
        self.data = []
        self.selected_id = None

        self._create_table()

    def _create_table(self):
        """Create the treeview table."""
        # Create frame for treeview and scrollbars
        self.tree_frame = ctk.CTkFrame(self)
        self.tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Determine columns
        columns = list(self.DISPLAY_COLUMNS)
        if self.show_validation_error:
            columns.append(("validation_error", "Validation Error", 200))

        column_ids = [col[0] for col in columns]

        # Create treeview with style
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                       background="#2b2b2b",
                       foreground="white",
                       fieldbackground="#2b2b2b",
                       rowheight=25)
        style.configure("Treeview.Heading",
                       background="#1f538d",
                       foreground="white",
                       relief="flat")
        style.map("Treeview",
                 background=[("selected", "#1f538d")])

        self.tree = ttk.Treeview(
            self.tree_frame,
            columns=column_ids,
            show="headings",
            selectmode="extended"  # Enable multi-select
        )

        # Configure columns
        for col_id, col_name, col_width in columns:
            self.tree.heading(col_id, text=col_name, anchor="w")
            self.tree.column(col_id, width=col_width, minwidth=50, anchor="w")

        # Scrollbars
        vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)

        # Bind selection event
        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click)

    def _on_select(self, event):
        """Handle row selection."""
        selection = self.tree.selection()
        if selection:
            item = self.tree.item(selection[0])
            values = item['values']
            if values:
                self.selected_id = values[0]  # ID is first column
                if self.on_select_callback:
                    # Find the full record
                    record = next((r for r in self.data if r.get('id') == self.selected_id), None)
                    self.on_select_callback(record)

    def _on_double_click(self, event):
        """Handle double-click for editing."""
        pass  # Will be implemented for inline editing

    def load_data(self, data: List[Dict[str, Any]]):
        """Load data into the table."""
        self.data = data
        self.tree.delete(*self.tree.get_children())

        columns = list(self.DISPLAY_COLUMNS)
        if self.show_validation_error:
            columns.append(("validation_error", "Validation Error", 200))

        for record in data:
            values = []
            for col_id, _, _ in columns:
                value = record.get(col_id, '')
                if col_id == 'amount_usd' and value:
                    try:
                        value = f"${float(value):,.2f}"
                    except (ValueError, TypeError):
                        pass
                values.append(value)
            self.tree.insert('', 'end', values=values)

    def get_selected_id(self) -> Optional[int]:
        """Get the ID of the selected row."""
        return self.selected_id

    def clear_selection(self):
        """Clear the current selection."""
        self.tree.selection_remove(self.tree.selection())
        self.selected_id = None

    def get_selected_records(self) -> List[Dict[str, Any]]:
        """Get all selected records (for multi-select)."""
        selected_records = []
        for item_id in self.tree.selection():
            item = self.tree.item(item_id)
            values = item['values']
            if values:
                record_id = values[0]  # ID is first column
                record = next((r for r in self.data if r.get('id') == record_id), None)
                if record:
                    selected_records.append(record)
        return selected_records

    def get_selected_ids(self) -> List[int]:
        """Get all selected record IDs (for multi-select)."""
        selected_ids = []
        for item_id in self.tree.selection():
            item = self.tree.item(item_id)
            values = item['values']
            if values:
                selected_ids.append(values[0])  # ID is first column
        return selected_ids


class EditDialog(ctk.CTkToplevel):
    """Dialog for editing invoice records."""

    FIELD_CONFIG = [
        ("tajan_bidding_tracking_number", "Tracking Number"),
        ("amazon_test_request", "Amazon Test Request"),
        ("amazon_tracker_number", "Amazon Tracker Number"),
        ("asin", "ASIN"),
        ("development_center", "Development Center"),
        ("category", "Category"),
        ("product_brand", "Product Brand"),
        ("product_description", "Product Description"),
        ("testing_sla", "Testing SLA"),
        ("test_service_type", "Test/Service Type"),
        ("quotation_order_number", "Quotation/Order Number"),
        ("request_date", "Request Date"),
        ("test_start_date", "Test Start Date"),
        ("report_delivered_date", "Report Delivered Date"),
        ("report_number", "Report Number"),
        ("test_inspection_location", "Test Location"),
        ("product_line", "Product Line"),
        ("amazon_quality_manager", "Quality Manager"),
        ("amazon_sourcing_manager", "Sourcing Manager"),
        ("invoice_number", "Invoice Number"),
        ("lab_contact", "Lab Contact"),
        ("comment", "Comment"),
        ("amount_usd", "Amount (USD)"),
        ("lab_name", "Lab Name"),
        ("invoice_date", "Invoice Date"),
    ]

    def __init__(self, parent, record: Dict[str, Any] = None, title: str = "Edit Invoice"):
        super().__init__(parent)

        self.record = record or {}
        self.result = None
        self.entries = {}

        self.title(title)
        self.geometry("600x700")
        self.resizable(True, True)

        # Make dialog modal
        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        # Center the dialog
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """Create dialog widgets."""
        # Scrollable frame for fields
        scroll_frame = ctk.CTkScrollableFrame(self, width=560, height=600)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Create fields
        for i, (field_id, field_label) in enumerate(self.FIELD_CONFIG):
            frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            frame.pack(fill="x", pady=2)

            label = ctk.CTkLabel(frame, text=field_label, width=150, anchor="w")
            label.pack(side="left", padx=5)

            entry = ctk.CTkEntry(frame, width=380)
            entry.pack(side="left", padx=5)

            # Pre-fill with existing value
            value = self.record.get(field_id, '')
            if value:
                entry.insert(0, str(value))

            self.entries[field_id] = entry

        # Buttons frame
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        save_btn = ctk.CTkButton(btn_frame, text="Save", command=self._on_save, width=100)
        save_btn.pack(side="right", padx=5)

        cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", command=self._on_cancel, width=100,
                                   fg_color="gray", hover_color="darkgray")
        cancel_btn.pack(side="right", padx=5)

    def _on_save(self):
        """Save the record."""
        self.result = {}
        for field_id, entry in self.entries.items():
            value = entry.get().strip()
            if field_id == 'amount_usd':
                try:
                    # Remove currency symbols and commas
                    value = value.replace('$', '').replace(',', '')
                    value = float(value) if value else 0.0
                except ValueError:
                    value = 0.0
            self.result[field_id] = value

        # Keep the ID if editing
        if 'id' in self.record:
            self.result['id'] = self.record['id']

        self.destroy()

    def _on_cancel(self):
        """Cancel the dialog."""
        self.result = None
        self.destroy()

    def get_result(self) -> Optional[Dict[str, Any]]:
        """Get the result after dialog closes."""
        self.wait_window()
        return self.result


class BulkEditDialog(ctk.CTkToplevel):
    """Dialog for bulk editing multiple invoice records."""

    FIELD_CONFIG = [
        ("lab_name", "Lab Name"),
        ("development_center", "Development Center"),
        ("category", "Category"),
        ("test_service_type", "Test/Service Type"),
        ("testing_sla", "Testing SLA"),
        ("test_inspection_location", "Test Location"),
        ("product_line", "Product Line"),
        ("amazon_quality_manager", "Quality Manager"),
        ("amazon_sourcing_manager", "Sourcing Manager"),
        ("invoice_date", "Invoice Date"),
        ("comment", "Comment"),
    ]

    def __init__(self, parent, record_count: int, title: str = "Bulk Edit"):
        super().__init__(parent)

        self.record_count = record_count
        self.result = None
        self.checkboxes = {}
        self.entries = {}

        self.title(title)
        self.geometry("650x550")
        self.resizable(True, True)

        # Make dialog modal
        self.transient(parent)
        self.grab_set()

        self._create_widgets()

        # Center the dialog
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_width()) // 2
        y = (self.winfo_screenheight() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        """Create dialog widgets."""
        # Info label
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.pack(fill="x", padx=10, pady=10)

        info_label = ctk.CTkLabel(
            info_frame,
            text=f"Editing {self.record_count} selected records.\n"
                 "Check the fields you want to update and enter new values.",
            font=ctk.CTkFont(size=12),
            justify="left"
        )
        info_label.pack(anchor="w")

        # Scrollable frame for fields
        scroll_frame = ctk.CTkScrollableFrame(self, width=610, height=400)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Create fields with checkboxes
        for field_id, field_label in self.FIELD_CONFIG:
            frame = ctk.CTkFrame(scroll_frame, fg_color="transparent")
            frame.pack(fill="x", pady=3)

            # Checkbox to enable/disable field
            checkbox_var = ctk.BooleanVar(value=False)
            checkbox = ctk.CTkCheckBox(
                frame,
                text="",
                variable=checkbox_var,
                width=24,
                command=lambda fid=field_id: self._toggle_field(fid)
            )
            checkbox.pack(side="left", padx=5)
            self.checkboxes[field_id] = checkbox_var

            # Label
            label = ctk.CTkLabel(frame, text=field_label, width=150, anchor="w")
            label.pack(side="left", padx=5)

            # Entry (initially disabled)
            entry = ctk.CTkEntry(frame, width=350, state="disabled")
            entry.pack(side="left", padx=5)
            self.entries[field_id] = entry

        # Buttons frame
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=10)

        apply_btn = ctk.CTkButton(
            btn_frame,
            text="Apply to All Selected",
            command=self._on_apply,
            width=150,
            fg_color="green",
            hover_color="darkgreen"
        )
        apply_btn.pack(side="right", padx=5)

        cancel_btn = ctk.CTkButton(
            btn_frame,
            text="Cancel",
            command=self._on_cancel,
            width=100,
            fg_color="gray",
            hover_color="darkgray"
        )
        cancel_btn.pack(side="right", padx=5)

        # Select all / Deselect all buttons
        select_all_btn = ctk.CTkButton(
            btn_frame,
            text="Select All Fields",
            command=self._select_all,
            width=120
        )
        select_all_btn.pack(side="left", padx=5)

        deselect_all_btn = ctk.CTkButton(
            btn_frame,
            text="Deselect All",
            command=self._deselect_all,
            width=100,
            fg_color="gray",
            hover_color="darkgray"
        )
        deselect_all_btn.pack(side="left", padx=5)

    def _toggle_field(self, field_id: str):
        """Toggle field entry state based on checkbox."""
        if self.checkboxes[field_id].get():
            self.entries[field_id].configure(state="normal")
        else:
            self.entries[field_id].configure(state="disabled")

    def _select_all(self):
        """Select all field checkboxes."""
        for field_id in self.checkboxes:
            self.checkboxes[field_id].set(True)
            self.entries[field_id].configure(state="normal")

    def _deselect_all(self):
        """Deselect all field checkboxes."""
        for field_id in self.checkboxes:
            self.checkboxes[field_id].set(False)
            self.entries[field_id].configure(state="disabled")

    def _on_apply(self):
        """Apply the bulk edit."""
        self.result = {}
        for field_id, checkbox_var in self.checkboxes.items():
            if checkbox_var.get():
                value = self.entries[field_id].get().strip()
                self.result[field_id] = value

        if not self.result:
            from tkinter import messagebox
            messagebox.showwarning("Warning", "Please select at least one field to update")
            return

        self.destroy()

    def _on_cancel(self):
        """Cancel the dialog."""
        self.result = None
        self.destroy()

    def get_result(self) -> Optional[Dict[str, Any]]:
        """Get the result after dialog closes."""
        self.wait_window()
        return self.result


class SettingsFrame(ctk.CTkFrame):
    """Settings panel for managing lab names and test types."""

    def __init__(self, parent, database: Database, **kwargs):
        super().__init__(parent, **kwargs)

        self.database = database
        self._create_widgets()

    def _create_widgets(self):
        """Create settings widgets."""
        # Title
        title = ctk.CTkLabel(self, text="Settings", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(pady=10)

        # Create notebook-like tabs using frames
        self.tab_frame = ctk.CTkFrame(self)
        self.tab_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Tab buttons
        tab_btn_frame = ctk.CTkFrame(self.tab_frame, fg_color="transparent")
        tab_btn_frame.pack(fill="x", pady=5)

        self.lab_btn = ctk.CTkButton(tab_btn_frame, text="Lab Names",
                                     command=lambda: self._show_tab("labs"))
        self.lab_btn.pack(side="left", padx=5)

        self.test_btn = ctk.CTkButton(tab_btn_frame, text="Test Types",
                                      command=lambda: self._show_tab("tests"))
        self.test_btn.pack(side="left", padx=5)

        # Content frames
        self.content_frame = ctk.CTkFrame(self.tab_frame)
        self.content_frame.pack(fill="both", expand=True, pady=10)

        self._create_labs_panel()
        self._create_tests_panel()

        # Show labs by default
        self._show_tab("labs")

    def _create_labs_panel(self):
        """Create lab names management panel."""
        self.labs_frame = ctk.CTkFrame(self.content_frame)

        # List of lab names
        list_frame = ctk.CTkFrame(self.labs_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.labs_listbox = tk.Listbox(list_frame, height=15, font=("Arial", 12),
                                       bg="#2b2b2b", fg="white", selectbackground="#1f538d")
        self.labs_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.labs_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.labs_listbox.config(yscrollcommand=scrollbar.set)

        # Buttons
        btn_frame = ctk.CTkFrame(self.labs_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)

        self.lab_entry = ctk.CTkEntry(btn_frame, placeholder_text="New lab name", width=200)
        self.lab_entry.pack(side="left", padx=5)

        add_btn = ctk.CTkButton(btn_frame, text="Add", command=self._add_lab, width=80)
        add_btn.pack(side="left", padx=5)

        delete_btn = ctk.CTkButton(btn_frame, text="Delete", command=self._delete_lab, width=80,
                                   fg_color="red", hover_color="darkred")
        delete_btn.pack(side="left", padx=5)

        self._refresh_labs()

    def _create_tests_panel(self):
        """Create test types management panel."""
        self.tests_frame = ctk.CTkFrame(self.content_frame)

        # List of test types
        list_frame = ctk.CTkFrame(self.tests_frame)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.tests_listbox = tk.Listbox(list_frame, height=15, font=("Arial", 12),
                                        bg="#2b2b2b", fg="white", selectbackground="#1f538d")
        self.tests_listbox.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tests_listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.tests_listbox.config(yscrollcommand=scrollbar.set)

        # Buttons
        btn_frame = ctk.CTkFrame(self.tests_frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=5)

        self.test_entry = ctk.CTkEntry(btn_frame, placeholder_text="New test type", width=200)
        self.test_entry.pack(side="left", padx=5)

        add_btn = ctk.CTkButton(btn_frame, text="Add", command=self._add_test, width=80)
        add_btn.pack(side="left", padx=5)

        delete_btn = ctk.CTkButton(btn_frame, text="Delete", command=self._delete_test, width=80,
                                   fg_color="red", hover_color="darkred")
        delete_btn.pack(side="left", padx=5)

        self._refresh_tests()

    def _show_tab(self, tab: str):
        """Show the specified tab."""
        self.labs_frame.pack_forget()
        self.tests_frame.pack_forget()

        if tab == "labs":
            self.labs_frame.pack(fill="both", expand=True)
            self.lab_btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
            self.test_btn.configure(fg_color="gray")
        else:
            self.tests_frame.pack(fill="both", expand=True)
            self.test_btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
            self.lab_btn.configure(fg_color="gray")

    def _refresh_labs(self):
        """Refresh the lab names list."""
        self.labs_listbox.delete(0, tk.END)
        for lab in self.database.get_lab_names():
            self.labs_listbox.insert(tk.END, lab)

    def _refresh_tests(self):
        """Refresh the test types list."""
        self.tests_listbox.delete(0, tk.END)
        for test in self.database.get_test_types():
            self.tests_listbox.insert(tk.END, test)

    def _add_lab(self):
        """Add a new lab name."""
        name = self.lab_entry.get().strip()
        if not name:
            messagebox.showwarning("Warning", "Please enter a lab name")
            return

        if self.database.add_lab_name(name):
            self._refresh_labs()
            self.lab_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Lab name already exists")

    def _delete_lab(self):
        """Delete selected lab name."""
        selection = self.labs_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a lab to delete")
            return

        name = self.labs_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Delete lab '{name}'?"):
            self.database.delete_lab_name(name)
            self._refresh_labs()

    def _add_test(self):
        """Add a new test type."""
        name = self.test_entry.get().strip()
        if not name:
            messagebox.showwarning("Warning", "Please enter a test type")
            return

        if self.database.add_test_type(name):
            self._refresh_tests()
            self.test_entry.delete(0, tk.END)
        else:
            messagebox.showerror("Error", "Test type already exists")

    def _delete_test(self):
        """Delete selected test type."""
        selection = self.tests_listbox.curselection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a test type to delete")
            return

        name = self.tests_listbox.get(selection[0])
        if messagebox.askyesno("Confirm", f"Delete test type '{name}'?"):
            self.database.delete_test_type(name)
            self._refresh_tests()


class SummaryFrame(ctk.CTkFrame):
    """Summary panel for invoice analytics and statistics."""

    MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

    def __init__(self, parent, database: Database, **kwargs):
        super().__init__(parent, **kwargs)

        self.database = database
        self.current_year = datetime.now().year
        self.current_view = "cost"  # cost, volume, unit_cost

        self._create_widgets()

    def _create_widgets(self):
        """Create summary widgets."""
        # Title and controls
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=5)

        title = ctk.CTkLabel(header_frame, text="Invoice Summary",
                            font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(side="left", padx=10)

        # Year selector
        year_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        year_frame.pack(side="right", padx=10)

        ctk.CTkLabel(year_frame, text="Year:").pack(side="left", padx=5)
        years = [str(y) for y in range(2020, datetime.now().year + 2)]
        self.year_combo = ctk.CTkComboBox(year_frame, values=years, width=100,
                                          command=self._on_year_change)
        self.year_combo.set(str(self.current_year))
        self.year_combo.pack(side="left", padx=5)

        # Refresh button
        refresh_btn = ctk.CTkButton(year_frame, text="Refresh",
                                    command=self.refresh_data, width=80)
        refresh_btn.pack(side="left", padx=10)

        # View selector tabs
        view_frame = ctk.CTkFrame(self)
        view_frame.pack(fill="x", padx=10, pady=5)

        self.view_buttons = {}

        cost_btn = ctk.CTkButton(view_frame, text="Total Cost (USD)",
                                 command=lambda: self._switch_view("cost"))
        cost_btn.pack(side="left", padx=5)
        self.view_buttons["cost"] = cost_btn

        volume_btn = ctk.CTkButton(view_frame, text="Number of Tests",
                                   command=lambda: self._switch_view("volume"))
        volume_btn.pack(side="left", padx=5)
        self.view_buttons["volume"] = volume_btn

        unit_btn = ctk.CTkButton(view_frame, text="Unit Cost (USD)",
                                 command=lambda: self._switch_view("unit_cost"))
        unit_btn.pack(side="left", padx=5)
        self.view_buttons["unit_cost"] = unit_btn

        lab_btn = ctk.CTkButton(view_frame, text="By Lab",
                                command=lambda: self._switch_view("by_lab"))
        lab_btn.pack(side="left", padx=5)
        self.view_buttons["by_lab"] = lab_btn

        # Summary statistics cards
        self._create_stats_cards()

        # Data table frame
        self.table_frame = ctk.CTkFrame(self)
        self.table_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self._create_summary_table()
        self._switch_view("cost")

    def _create_stats_cards(self):
        """Create summary statistics cards."""
        cards_frame = ctk.CTkFrame(self, fg_color="transparent")
        cards_frame.pack(fill="x", padx=10, pady=10)

        # Total Cost Card
        self.total_cost_card = self._create_card(cards_frame, "Total Cost", "$0.00", "#1f538d")
        self.total_cost_card.pack(side="left", padx=10, expand=True, fill="x")

        # Total Tests Card
        self.total_tests_card = self._create_card(cards_frame, "Total Tests", "0", "#2d7d46")
        self.total_tests_card.pack(side="left", padx=10, expand=True, fill="x")

        # Avg Cost Card
        self.avg_cost_card = self._create_card(cards_frame, "Avg Cost/Test", "$0.00", "#6B4C9A")
        self.avg_cost_card.pack(side="left", padx=10, expand=True, fill="x")

        # Active Labs Card
        self.labs_card = self._create_card(cards_frame, "Active Labs", "0", "#b8860b")
        self.labs_card.pack(side="left", padx=10, expand=True, fill="x")

    def _create_card(self, parent, title: str, value: str, color: str) -> ctk.CTkFrame:
        """Create a statistics card."""
        card = ctk.CTkFrame(parent, fg_color=color, corner_radius=10)

        title_label = ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=12))
        title_label.pack(pady=(10, 2))

        value_label = ctk.CTkLabel(card, text=value, font=ctk.CTkFont(size=24, weight="bold"))
        value_label.pack(pady=(2, 10))

        # Store reference to value label for updates
        card.value_label = value_label

        return card

    def _create_summary_table(self):
        """Create the summary data table."""
        # Create treeview with style
        style = ttk.Style()
        style.configure("Summary.Treeview",
                       background="#2b2b2b",
                       foreground="white",
                       fieldbackground="#2b2b2b",
                       rowheight=25)
        style.configure("Summary.Treeview.Heading",
                       background="#1f538d",
                       foreground="white",
                       relief="flat")

        # Columns: Source + Test Type + 12 months + Total
        columns = ["source", "test_type"] + self.MONTHS + ["Total"]

        self.summary_tree = ttk.Treeview(
            self.table_frame,
            columns=columns,
            show="headings",
            style="Summary.Treeview"
        )

        # Configure columns
        self.summary_tree.heading("source", text="Source", anchor="w")
        self.summary_tree.column("source", width=80, minwidth=60, anchor="w")

        self.summary_tree.heading("test_type", text="Test/Service Type", anchor="w")
        self.summary_tree.column("test_type", width=200, minwidth=150, anchor="w")

        for month in self.MONTHS:
            self.summary_tree.heading(month, text=month, anchor="e")
            self.summary_tree.column(month, width=80, minwidth=60, anchor="e")

        self.summary_tree.heading("Total", text="Total", anchor="e")
        self.summary_tree.column("Total", width=100, minwidth=80, anchor="e")

        # Scrollbars
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.summary_tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.summary_tree.xview)
        self.summary_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout
        self.summary_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

    def _switch_view(self, view: str):
        """Switch between different summary views."""
        self.current_view = view

        # Update button colors
        for v, btn in self.view_buttons.items():
            if v == view:
                btn.configure(fg_color=("#3B8ED0", "#1F6AA5"))
            else:
                btn.configure(fg_color="gray")

        self.refresh_data()

    def _on_year_change(self, value):
        """Handle year selection change."""
        self.current_year = int(value)
        self.refresh_data()

    def refresh_data(self):
        """Refresh summary data."""
        # Get both valid and abnormal invoices
        valid_invoices = self.database.get_all_invoices()
        abnormal_invoices = self.database.get_all_abnormal_invoices()

        # Add source indicator to each record
        for inv in valid_invoices:
            inv['_source'] = 'Valid'
        for inv in abnormal_invoices:
            inv['_source'] = 'Abnormal'

        # Filter by year - valid invoices
        year_valid = []
        for inv in valid_invoices:
            inv_date = inv.get('invoice_date', '') or inv.get('request_date', '')
            if inv_date and str(self.current_year) in str(inv_date)[:4]:
                year_valid.append(inv)

        # Filter by year - abnormal invoices
        year_abnormal = []
        for inv in abnormal_invoices:
            inv_date = inv.get('invoice_date', '') or inv.get('request_date', '')
            if inv_date and str(self.current_year) in str(inv_date)[:4]:
                year_abnormal.append(inv)

        # Combine for statistics (all invoices)
        all_year_invoices = year_valid + year_abnormal

        # Calculate summary statistics
        self._update_stats_cards(all_year_invoices, len(year_valid), len(year_abnormal))

        # Update table based on current view
        if self.current_view == "cost":
            self._show_cost_summary(year_valid, year_abnormal)
        elif self.current_view == "volume":
            self._show_volume_summary(year_valid, year_abnormal)
        elif self.current_view == "unit_cost":
            self._show_unit_cost_summary(year_valid, year_abnormal)
        elif self.current_view == "by_lab":
            self._show_lab_summary(year_valid, year_abnormal)

    def _update_stats_cards(self, invoices: List[Dict[str, Any]], valid_count: int = 0, abnormal_count: int = 0):
        """Update the statistics cards."""
        total_cost = sum(float(inv.get('amount_usd', 0) or 0) for inv in invoices)
        total_tests = len(invoices)
        avg_cost = total_cost / total_tests if total_tests > 0 else 0
        active_labs = len(set(inv.get('lab_name', '') for inv in invoices if inv.get('lab_name')))

        self.total_cost_card.value_label.configure(text=f"${total_cost:,.2f}")
        self.total_tests_card.value_label.configure(text=f"{total_tests:,}\n(V:{valid_count} A:{abnormal_count})")
        self.avg_cost_card.value_label.configure(text=f"${avg_cost:,.2f}")
        self.labs_card.value_label.configure(text=str(active_labs))

    def _get_month_from_date(self, date_str: str) -> Optional[int]:
        """Extract month (1-12) from date string."""
        if not date_str:
            return None
        try:
            # Try YYYY-MM-DD format
            if '-' in str(date_str):
                parts = str(date_str).split('-')
                if len(parts) >= 2:
                    return int(parts[1])
            # Try other formats
            for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y']:
                try:
                    dt = datetime.strptime(str(date_str)[:10], fmt)
                    return dt.month
                except ValueError:
                    continue
        except (ValueError, IndexError):
            pass
        return None

    def _show_cost_summary(self, valid_invoices: List[Dict[str, Any]], abnormal_invoices: List[Dict[str, Any]]):
        """Show total cost by test type and month."""
        self.summary_tree.delete(*self.summary_tree.get_children())

        def process_invoices(invoices, source_label):
            """Process invoices and return grouped data."""
            data = {}
            for inv in invoices:
                test_type = inv.get('test_service_type', 'Unknown') or 'Unknown'
                amount = float(inv.get('amount_usd', 0) or 0)
                date_str = inv.get('invoice_date', '') or inv.get('request_date', '')
                month = self._get_month_from_date(date_str)

                if test_type not in data:
                    data[test_type] = {m: 0 for m in range(1, 13)}

                if month:
                    data[test_type][month] += amount
            return data

        # Process valid and abnormal separately
        valid_data = process_invoices(valid_invoices, "Valid")
        abnormal_data = process_invoices(abnormal_invoices, "Abnormal")

        grand_totals = {m: 0 for m in range(1, 13)}

        # Add valid invoice rows first
        for test_type in sorted(valid_data.keys()):
            row_data = ["Valid", test_type]
            row_total = 0
            for m in range(1, 13):
                val = valid_data[test_type][m]
                row_data.append(f"${val:,.0f}" if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(f"${row_total:,.0f}")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add abnormal invoice rows
        for test_type in sorted(abnormal_data.keys()):
            row_data = ["Abnormal", test_type]
            row_total = 0
            for m in range(1, 13):
                val = abnormal_data[test_type][m]
                row_data.append(f"${val:,.0f}" if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(f"${row_total:,.0f}")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add grand total row
        total_row = ["", "TOTAL"]
        grand_total = 0
        for m in range(1, 13):
            total_row.append(f"${grand_totals[m]:,.0f}" if grand_totals[m] > 0 else "-")
            grand_total += grand_totals[m]
        total_row.append(f"${grand_total:,.0f}")
        self.summary_tree.insert('', 'end', values=total_row, tags=('total',))

    def _show_volume_summary(self, valid_invoices: List[Dict[str, Any]], abnormal_invoices: List[Dict[str, Any]]):
        """Show test volume by test type and month."""
        self.summary_tree.delete(*self.summary_tree.get_children())

        def process_invoices(invoices):
            """Process invoices and return grouped data."""
            data = {}
            for inv in invoices:
                test_type = inv.get('test_service_type', 'Unknown') or 'Unknown'
                date_str = inv.get('invoice_date', '') or inv.get('request_date', '')
                month = self._get_month_from_date(date_str)

                if test_type not in data:
                    data[test_type] = {m: 0 for m in range(1, 13)}

                if month:
                    data[test_type][month] += 1
            return data

        # Process valid and abnormal separately
        valid_data = process_invoices(valid_invoices)
        abnormal_data = process_invoices(abnormal_invoices)

        grand_totals = {m: 0 for m in range(1, 13)}

        # Add valid invoice rows first
        for test_type in sorted(valid_data.keys()):
            row_data = ["Valid", test_type]
            row_total = 0
            for m in range(1, 13):
                val = valid_data[test_type][m]
                row_data.append(str(val) if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(str(row_total))
            self.summary_tree.insert('', 'end', values=row_data)

        # Add abnormal invoice rows
        for test_type in sorted(abnormal_data.keys()):
            row_data = ["Abnormal", test_type]
            row_total = 0
            for m in range(1, 13):
                val = abnormal_data[test_type][m]
                row_data.append(str(val) if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(str(row_total))
            self.summary_tree.insert('', 'end', values=row_data)

        # Add grand total row
        total_row = ["", "TOTAL"]
        grand_total = 0
        for m in range(1, 13):
            total_row.append(str(grand_totals[m]) if grand_totals[m] > 0 else "-")
            grand_total += grand_totals[m]
        total_row.append(str(grand_total))
        self.summary_tree.insert('', 'end', values=total_row, tags=('total',))

    def _show_unit_cost_summary(self, valid_invoices: List[Dict[str, Any]], abnormal_invoices: List[Dict[str, Any]]):
        """Show unit cost (avg cost per test) by test type and month."""
        self.summary_tree.delete(*self.summary_tree.get_children())

        def process_invoices(invoices):
            """Process invoices and return cost and count data."""
            cost_data = {}
            count_data = {}
            for inv in invoices:
                test_type = inv.get('test_service_type', 'Unknown') or 'Unknown'
                amount = float(inv.get('amount_usd', 0) or 0)
                date_str = inv.get('invoice_date', '') or inv.get('request_date', '')
                month = self._get_month_from_date(date_str)

                if test_type not in cost_data:
                    cost_data[test_type] = {m: 0 for m in range(1, 13)}
                    count_data[test_type] = {m: 0 for m in range(1, 13)}

                if month:
                    cost_data[test_type][month] += amount
                    count_data[test_type][month] += 1
            return cost_data, count_data

        # Process valid and abnormal separately
        valid_cost, valid_count = process_invoices(valid_invoices)
        abnormal_cost, abnormal_count = process_invoices(abnormal_invoices)

        grand_costs = {m: 0 for m in range(1, 13)}
        grand_counts = {m: 0 for m in range(1, 13)}

        # Add valid invoice rows first
        for test_type in sorted(valid_cost.keys()):
            row_data = ["Valid", test_type]
            row_total_cost = 0
            row_total_count = 0
            for m in range(1, 13):
                cost = valid_cost[test_type][m]
                count = valid_count[test_type][m]
                if count > 0:
                    unit_cost = cost / count
                    row_data.append(f"${unit_cost:,.0f}")
                else:
                    row_data.append("-")
                row_total_cost += cost
                row_total_count += count
                grand_costs[m] += cost
                grand_counts[m] += count

            if row_total_count > 0:
                row_data.append(f"${row_total_cost / row_total_count:,.0f}")
            else:
                row_data.append("-")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add abnormal invoice rows
        for test_type in sorted(abnormal_cost.keys()):
            row_data = ["Abnormal", test_type]
            row_total_cost = 0
            row_total_count = 0
            for m in range(1, 13):
                cost = abnormal_cost[test_type][m]
                count = abnormal_count[test_type][m]
                if count > 0:
                    unit_cost = cost / count
                    row_data.append(f"${unit_cost:,.0f}")
                else:
                    row_data.append("-")
                row_total_cost += cost
                row_total_count += count
                grand_costs[m] += cost
                grand_counts[m] += count

            if row_total_count > 0:
                row_data.append(f"${row_total_cost / row_total_count:,.0f}")
            else:
                row_data.append("-")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add grand total row (overall averages)
        total_row = ["", "TOTAL"]
        grand_total_cost = 0
        grand_total_count = 0
        for m in range(1, 13):
            if grand_counts[m] > 0:
                total_row.append(f"${grand_costs[m] / grand_counts[m]:,.0f}")
            else:
                total_row.append("-")
            grand_total_cost += grand_costs[m]
            grand_total_count += grand_counts[m]

        if grand_total_count > 0:
            total_row.append(f"${grand_total_cost / grand_total_count:,.0f}")
        else:
            total_row.append("-")
        self.summary_tree.insert('', 'end', values=total_row, tags=('total',))

    def _show_lab_summary(self, valid_invoices: List[Dict[str, Any]], abnormal_invoices: List[Dict[str, Any]]):
        """Show summary by lab."""
        self.summary_tree.delete(*self.summary_tree.get_children())

        def process_invoices(invoices):
            """Process invoices and return grouped data."""
            data = {}
            for inv in invoices:
                lab = inv.get('lab_name', 'Unknown') or 'Unknown'
                amount = float(inv.get('amount_usd', 0) or 0)
                date_str = inv.get('invoice_date', '') or inv.get('request_date', '')
                month = self._get_month_from_date(date_str)

                if lab not in data:
                    data[lab] = {m: 0 for m in range(1, 13)}

                if month:
                    data[lab][month] += amount
            return data

        # Process valid and abnormal separately
        valid_data = process_invoices(valid_invoices)
        abnormal_data = process_invoices(abnormal_invoices)

        grand_totals = {m: 0 for m in range(1, 13)}

        # Add valid invoice rows first
        for lab in sorted(valid_data.keys()):
            row_data = ["Valid", lab]
            row_total = 0
            for m in range(1, 13):
                val = valid_data[lab][m]
                row_data.append(f"${val:,.0f}" if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(f"${row_total:,.0f}")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add abnormal invoice rows
        for lab in sorted(abnormal_data.keys()):
            row_data = ["Abnormal", lab]
            row_total = 0
            for m in range(1, 13):
                val = abnormal_data[lab][m]
                row_data.append(f"${val:,.0f}" if val > 0 else "-")
                row_total += val
                grand_totals[m] += val
            row_data.append(f"${row_total:,.0f}")
            self.summary_tree.insert('', 'end', values=row_data)

        # Add grand total row
        total_row = ["", "TOTAL"]
        grand_total = 0
        for m in range(1, 13):
            total_row.append(f"${grand_totals[m]:,.0f}" if grand_totals[m] > 0 else "-")
            grand_total += grand_totals[m]
        total_row.append(f"${grand_total:,.0f}")
        self.summary_tree.insert('', 'end', values=total_row, tags=('total',))


class FilterFrame(ctk.CTkFrame):
    """Filter panel for searching invoice data."""

    def __init__(self, parent, database: Database, on_filter_callback=None, **kwargs):
        super().__init__(parent, **kwargs)

        self.database = database
        self.on_filter_callback = on_filter_callback
        self._create_widgets()

    def _create_widgets(self):
        """Create filter widgets."""
        # Search text
        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=5, pady=2)

        ctk.CTkLabel(search_frame, text="Search:", width=80).pack(side="left")
        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="Search...", width=200)
        self.search_entry.pack(side="left", padx=5)

        # Lab filter
        lab_frame = ctk.CTkFrame(self, fg_color="transparent")
        lab_frame.pack(fill="x", padx=5, pady=2)

        ctk.CTkLabel(lab_frame, text="Lab:", width=80).pack(side="left")
        lab_names = ["All"] + self.database.get_lab_names()
        self.lab_combo = ctk.CTkComboBox(lab_frame, values=lab_names, width=200)
        self.lab_combo.set("All")
        self.lab_combo.pack(side="left", padx=5)

        # Date range
        date_frame = ctk.CTkFrame(self, fg_color="transparent")
        date_frame.pack(fill="x", padx=5, pady=2)

        ctk.CTkLabel(date_frame, text="Date From:", width=80).pack(side="left")
        self.date_from = ctk.CTkEntry(date_frame, placeholder_text="YYYY-MM-DD", width=120)
        self.date_from.pack(side="left", padx=5)

        ctk.CTkLabel(date_frame, text="To:", width=30).pack(side="left")
        self.date_to = ctk.CTkEntry(date_frame, placeholder_text="YYYY-MM-DD", width=120)
        self.date_to.pack(side="left", padx=5)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=5, pady=5)

        filter_btn = ctk.CTkButton(btn_frame, text="Apply Filter", command=self._apply_filter, width=100)
        filter_btn.pack(side="left", padx=5)

        clear_btn = ctk.CTkButton(btn_frame, text="Clear", command=self._clear_filter, width=80,
                                  fg_color="gray", hover_color="darkgray")
        clear_btn.pack(side="left", padx=5)

    def _apply_filter(self):
        """Apply the filter."""
        if self.on_filter_callback:
            filters = {
                'search_text': self.search_entry.get().strip(),
                'lab_name': self.lab_combo.get() if self.lab_combo.get() != "All" else None,
                'date_from': self.date_from.get().strip() or None,
                'date_to': self.date_to.get().strip() or None,
            }
            self.on_filter_callback(filters)

    def _clear_filter(self):
        """Clear all filters."""
        self.search_entry.delete(0, tk.END)
        self.lab_combo.set("All")
        self.date_from.delete(0, tk.END)
        self.date_to.delete(0, tk.END)

        if self.on_filter_callback:
            self.on_filter_callback({})

    def refresh_lab_list(self):
        """Refresh the lab names in the combo box."""
        lab_names = ["All"] + self.database.get_lab_names()
        self.lab_combo.configure(values=lab_names)


class MainApplication(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title("Invoice Management System")
        self.geometry("1400x800")
        self.minsize(1200, 600)

        # Initialize database
        self.database = Database()

        # Current state
        self.current_tab = "invoices"
        self.selected_record = None

        self._create_widgets()
        self._refresh_data()

    def _create_widgets(self):
        """Create main application widgets."""
        # Top toolbar
        self._create_toolbar()

        # Main content area with tabs
        self._create_main_content()

        # Status bar
        self._create_status_bar()

    def _create_toolbar(self):
        """Create the top toolbar."""
        toolbar = ctk.CTkFrame(self, height=50)
        toolbar.pack(fill="x", padx=10, pady=5)
        toolbar.pack_propagate(False)

        # Upload button
        upload_btn = ctk.CTkButton(toolbar, text="📁 Upload Invoice", command=self._upload_invoice, width=140)
        upload_btn.pack(side="left", padx=5, pady=5)

        # Add row button
        add_btn = ctk.CTkButton(toolbar, text="➕ Add Row", command=self._add_row, width=100)
        add_btn.pack(side="left", padx=5, pady=5)

        # Edit button
        edit_btn = ctk.CTkButton(toolbar, text="✏️ Edit", command=self._edit_row, width=80)
        edit_btn.pack(side="left", padx=5, pady=5)

        # Delete button
        delete_btn = ctk.CTkButton(toolbar, text="🗑️ Delete", command=self._delete_row, width=80,
                                   fg_color="red", hover_color="darkred")
        delete_btn.pack(side="left", padx=5, pady=5)

        # Separator
        separator = ctk.CTkLabel(toolbar, text="|", width=10)
        separator.pack(side="left", padx=2, pady=5)

        # Bulk Edit button
        bulk_edit_btn = ctk.CTkButton(toolbar, text="📝 Bulk Edit", command=self._bulk_edit, width=100,
                                      fg_color="#6B4C9A", hover_color="#4A3570")
        bulk_edit_btn.pack(side="left", padx=5, pady=5)

        # Bulk Delete button
        bulk_delete_btn = ctk.CTkButton(toolbar, text="🗑️ Bulk Delete", command=self._bulk_delete, width=110,
                                        fg_color="#8B0000", hover_color="#5C0000")
        bulk_delete_btn.pack(side="left", padx=5, pady=5)

        # Refresh button
        refresh_btn = ctk.CTkButton(toolbar, text="🔄 Refresh", command=self._refresh_data, width=100)
        refresh_btn.pack(side="left", padx=5, pady=5)

        # Export button
        export_btn = ctk.CTkButton(toolbar, text="📊 Export", command=self._export_data, width=100,
                                   fg_color="green", hover_color="darkgreen")
        export_btn.pack(side="right", padx=5, pady=5)

    def _create_main_content(self):
        """Create the main content area with tabs."""
        # Tab buttons
        tab_frame = ctk.CTkFrame(self)
        tab_frame.pack(fill="x", padx=10, pady=5)

        self.tab_buttons = {}

        invoices_btn = ctk.CTkButton(tab_frame, text="Valid Invoices",
                                     command=lambda: self._switch_tab("invoices"))
        invoices_btn.pack(side="left", padx=5)
        self.tab_buttons["invoices"] = invoices_btn

        abnormal_btn = ctk.CTkButton(tab_frame, text="Abnormal Items",
                                     command=lambda: self._switch_tab("abnormal"))
        abnormal_btn.pack(side="left", padx=5)
        self.tab_buttons["abnormal"] = abnormal_btn

        settings_btn = ctk.CTkButton(tab_frame, text="Settings",
                                     command=lambda: self._switch_tab("settings"))
        settings_btn.pack(side="left", padx=5)
        self.tab_buttons["settings"] = settings_btn

        summary_btn = ctk.CTkButton(tab_frame, text="Summary",
                                    command=lambda: self._switch_tab("summary"),
                                    fg_color="#2d7d46", hover_color="#1d5d30")
        summary_btn.pack(side="left", padx=5)
        self.tab_buttons["summary"] = summary_btn

        # Content container
        self.content_container = ctk.CTkFrame(self)
        self.content_container.pack(fill="both", expand=True, padx=10, pady=5)

        # Create content frames
        self._create_invoices_tab()
        self._create_abnormal_tab()
        self._create_settings_tab()
        self._create_summary_tab()

        # Show invoices tab by default
        self._switch_tab("invoices")

    def _create_invoices_tab(self):
        """Create the valid invoices tab."""
        self.invoices_frame = ctk.CTkFrame(self.content_container)

        # Filter panel
        self.filter_frame = FilterFrame(
            self.invoices_frame,
            self.database,
            on_filter_callback=self._on_filter
        )
        self.filter_frame.pack(fill="x", padx=5, pady=5)

        # Table
        self.invoices_table = InvoiceTable(
            self.invoices_frame,
            on_select_callback=self._on_select_invoice
        )
        self.invoices_table.pack(fill="both", expand=True, padx=5, pady=5)

    def _create_abnormal_tab(self):
        """Create the abnormal items tab."""
        self.abnormal_frame = ctk.CTkFrame(self.content_container)

        # Info label
        info_label = ctk.CTkLabel(
            self.abnormal_frame,
            text="Items that failed validation during upload are shown here.",
            font=ctk.CTkFont(size=12)
        )
        info_label.pack(pady=5)

        # Table
        self.abnormal_table = InvoiceTable(
            self.abnormal_frame,
            on_select_callback=self._on_select_abnormal,
            show_validation_error=True
        )
        self.abnormal_table.pack(fill="both", expand=True, padx=5, pady=5)

    def _create_settings_tab(self):
        """Create the settings tab."""
        self.settings_frame = SettingsFrame(self.content_container, self.database)

    def _create_summary_tab(self):
        """Create the summary/analytics tab."""
        self.summary_frame = SummaryFrame(self.content_container, self.database)

    def _create_status_bar(self):
        """Create the status bar."""
        self.status_bar = ctk.CTkLabel(self, text="Ready", anchor="w")
        self.status_bar.pack(fill="x", padx=10, pady=5)

    def _switch_tab(self, tab: str):
        """Switch to the specified tab."""
        # Hide all frames
        self.invoices_frame.pack_forget()
        self.abnormal_frame.pack_forget()
        self.settings_frame.pack_forget()
        self.summary_frame.pack_forget()

        # Reset button colors
        for btn in self.tab_buttons.values():
            btn.configure(fg_color="gray")

        # Show selected frame and highlight button
        self.current_tab = tab
        if tab == "invoices":
            self.invoices_frame.pack(fill="both", expand=True)
            self.tab_buttons["invoices"].configure(fg_color=("#3B8ED0", "#1F6AA5"))
        elif tab == "abnormal":
            self.abnormal_frame.pack(fill="both", expand=True)
            self.tab_buttons["abnormal"].configure(fg_color=("#3B8ED0", "#1F6AA5"))
        elif tab == "settings":
            self.settings_frame.pack(fill="both", expand=True)
            self.tab_buttons["settings"].configure(fg_color=("#3B8ED0", "#1F6AA5"))
        elif tab == "summary":
            self.summary_frame.pack(fill="both", expand=True)
            self.summary_frame.refresh_data()  # Refresh data when switching to summary
            self.tab_buttons["summary"].configure(fg_color=("#2d7d46", "#1d5d30"))

    def _upload_invoice(self):
        """Handle invoice file upload."""
        file_path = filedialog.askopenfilename(
            title="Select Invoice Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        # Ask for invoice date
        invoice_date = self._ask_invoice_date()

        # Parse the file
        lab_names = self.database.get_lab_names()
        parser = ExcelParser(lab_names)
        records, detected_lab, error = parser.parse_file(file_path, invoice_date)

        if error:
            messagebox.showerror("Error", error)
            return

        if not detected_lab:
            messagebox.showwarning(
                "Lab Not Detected",
                f"Could not detect lab name from filename.\nKnown labs: {', '.join(lab_names)}\n\n"
                "Records will be imported without lab name."
            )

        if not records:
            messagebox.showinfo("Info", "No valid records found in the file.")
            return

        # Validate records
        test_types = self.database.get_test_types()
        validator = DataValidator(test_types)
        valid_records, invalid_records = validator.validate_records(records)

        # Insert valid records
        for record in valid_records:
            self.database.insert_invoice(record)

        # Insert invalid records
        for record, error_msg in invalid_records:
            self.database.insert_abnormal_invoice(record, error_msg)

        # Show summary
        messagebox.showinfo(
            "Import Complete",
            f"Successfully imported {len(valid_records)} valid records.\n"
            f"{len(invalid_records)} records moved to Abnormal Items."
        )

        self._refresh_data()
        self._update_status(f"Imported {len(records)} records from {os.path.basename(file_path)}")

    def _ask_invoice_date(self) -> str:
        """Ask user for invoice date."""
        dialog = ctk.CTkInputDialog(
            text="Enter invoice date (YYYY-MM-DD):\nLeave empty to skip.",
            title="Invoice Date"
        )
        result = dialog.get_input()
        return result if result else ""

    def _add_row(self):
        """Add a new invoice row."""
        dialog = EditDialog(self, title="Add New Invoice")
        result = dialog.get_result()

        if result:
            if self.current_tab == "invoices":
                self.database.insert_invoice(result)
            else:
                self.database.insert_abnormal_invoice(result, "Manually Added")

            self._refresh_data()
            self._update_status("New record added")

    def _edit_row(self):
        """Edit the selected row."""
        if not self.selected_record:
            messagebox.showwarning("Warning", "Please select a record to edit")
            return

        dialog = EditDialog(self, record=self.selected_record, title="Edit Invoice")
        result = dialog.get_result()

        if result:
            record_id = result.get('id')
            if self.current_tab == "invoices":
                self.database.update_invoice(record_id, result)
                self._refresh_data()
                self._update_status(f"Record {record_id} updated")
            else:
                # Re-validate abnormal item after edit
                test_types = self.database.get_test_types()
                validator = DataValidator(test_types)
                is_valid, errors = validator.validate_record(result)

                if is_valid:
                    # Item now passes validation - move to valid invoices
                    # Remove the 'id' field for insertion
                    insert_data = {k: v for k, v in result.items() if k != 'id'}
                    new_id = self.database.insert_invoice(insert_data)
                    self.database.delete_abnormal_invoice(record_id)
                    self._refresh_data()
                    self._update_status(f"Record moved to Valid Invoices (new ID: {new_id})")
                    messagebox.showinfo(
                        "Validation Passed",
                        f"Record passed validation and has been moved to Valid Invoices.\n"
                        f"New Invoice ID: {new_id}"
                    )
                else:
                    # Still invalid - update with new validation error
                    result['validation_error'] = "; ".join(errors)
                    self.database.update_abnormal_invoice(record_id, result)
                    self._refresh_data()
                    self._update_status(f"Record {record_id} updated (still abnormal)")

    def _delete_row(self):
        """Delete the selected row."""
        if not self.selected_record:
            messagebox.showwarning("Warning", "Please select a record to delete")
            return

        record_id = self.selected_record.get('id')
        if messagebox.askyesno("Confirm", f"Delete record {record_id}?"):
            if self.current_tab == "invoices":
                self.database.delete_invoice(record_id)
            else:
                self.database.delete_abnormal_invoice(record_id)

            self.selected_record = None
            self._refresh_data()
            self._update_status(f"Record {record_id} deleted")

    def _bulk_edit(self):
        """Bulk edit multiple selected rows."""
        # Get the current table
        if self.current_tab == "invoices":
            table = self.invoices_table
        elif self.current_tab == "abnormal":
            table = self.abnormal_table
        else:
            messagebox.showwarning("Warning", "Bulk edit is not available in Settings")
            return

        selected_ids = table.get_selected_ids()
        selected_records = table.get_selected_records()

        if len(selected_ids) < 2:
            messagebox.showwarning("Warning", "Please select at least 2 records for bulk edit.\n"
                                              "Use Ctrl+Click or Shift+Click to select multiple rows.")
            return

        # Show bulk edit dialog
        dialog = BulkEditDialog(self, record_count=len(selected_ids))
        result = dialog.get_result()

        if result:
            if self.current_tab == "invoices":
                # For valid invoices, just update
                success_count = 0
                for record_id in selected_ids:
                    if self.database.update_invoice(record_id, result):
                        success_count += 1
                self._refresh_data()
                self._update_status(f"Bulk edit: {success_count} records updated")
                messagebox.showinfo("Success", f"Successfully updated {success_count} records.")
            else:
                # For abnormal items, update and re-validate each record
                test_types = self.database.get_test_types()
                validator = DataValidator(test_types)

                updated_count = 0
                moved_count = 0

                for record in selected_records:
                    record_id = record.get('id')

                    # Merge the bulk edit changes into the record
                    updated_record = dict(record)
                    updated_record.update(result)

                    # Re-validate
                    is_valid, errors = validator.validate_record(updated_record)

                    if is_valid:
                        # Move to valid invoices
                        insert_data = {k: v for k, v in updated_record.items() if k not in ('id', 'validation_error', '_source')}
                        self.database.insert_invoice(insert_data)
                        self.database.delete_abnormal_invoice(record_id)
                        moved_count += 1
                    else:
                        # Update in abnormal with new validation error
                        updated_record['validation_error'] = "; ".join(errors)
                        self.database.update_abnormal_invoice(record_id, updated_record)
                        updated_count += 1

                self._refresh_data()
                self._update_status(f"Bulk edit: {updated_count} updated, {moved_count} moved to Valid")
                messagebox.showinfo(
                    "Bulk Edit Complete",
                    f"Results:\n"
                    f"- {moved_count} records passed validation and moved to Valid Invoices\n"
                    f"- {updated_count} records updated (still abnormal)"
                )

    def _bulk_delete(self):
        """Bulk delete multiple selected rows."""
        # Get the current table
        if self.current_tab == "invoices":
            table = self.invoices_table
        elif self.current_tab == "abnormal":
            table = self.abnormal_table
        else:
            messagebox.showwarning("Warning", "Bulk delete is not available in Settings")
            return

        selected_ids = table.get_selected_ids()

        if len(selected_ids) < 2:
            messagebox.showwarning("Warning", "Please select at least 2 records for bulk delete.\n"
                                              "Use Ctrl+Click or Shift+Click to select multiple rows.")
            return

        # Confirm deletion
        if not messagebox.askyesno("Confirm Bulk Delete",
                                   f"Are you sure you want to delete {len(selected_ids)} selected records?\n\n"
                                   "This action cannot be undone."):
            return

        # Delete all selected records
        success_count = 0
        for record_id in selected_ids:
            if self.current_tab == "invoices":
                if self.database.delete_invoice(record_id):
                    success_count += 1
            else:
                if self.database.delete_abnormal_invoice(record_id):
                    success_count += 1

        self.selected_record = None
        self._refresh_data()
        self._update_status(f"Bulk delete: {success_count} records deleted")
        messagebox.showinfo("Success", f"Successfully deleted {success_count} records.")

    def _refresh_data(self):
        """Refresh all data tables."""
        invoices = self.database.get_all_invoices()
        self.invoices_table.load_data(invoices)

        abnormal = self.database.get_all_abnormal_invoices()
        self.abnormal_table.load_data(abnormal)

        # Update filter lab list
        self.filter_frame.refresh_lab_list()

    def _on_filter(self, filters: Dict[str, Any]):
        """Handle filter application."""
        if any(filters.values()):
            results = self.database.search_invoices(filters)
            self.invoices_table.load_data(results)
            self._update_status(f"Filter applied: {len(results)} records found")
        else:
            self._refresh_data()
            self._update_status("Filter cleared")

    def _on_select_invoice(self, record: Optional[Dict[str, Any]]):
        """Handle invoice selection."""
        self.selected_record = record

    def _on_select_abnormal(self, record: Optional[Dict[str, Any]]):
        """Handle abnormal invoice selection."""
        self.selected_record = record

    def _export_data(self):
        """Export data to Excel."""
        file_path = filedialog.asksaveasfilename(
            title="Save Export",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        try:
            from openpyxl import Workbook

            wb = Workbook()

            # Valid invoices sheet
            ws1 = wb.active
            ws1.title = "Valid Invoices"
            invoices = self.database.get_all_invoices()
            if invoices:
                headers = list(invoices[0].keys())
                ws1.append(headers)
                for record in invoices:
                    ws1.append(list(record.values()))

            # Abnormal invoices sheet
            ws2 = wb.create_sheet("Abnormal Items")
            abnormal = self.database.get_all_abnormal_invoices()
            if abnormal:
                headers = list(abnormal[0].keys())
                ws2.append(headers)
                for record in abnormal:
                    ws2.append(list(record.values()))

            wb.save(file_path)
            messagebox.showinfo("Success", f"Data exported to {file_path}")
            self._update_status(f"Data exported to {os.path.basename(file_path)}")

        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def _update_status(self, message: str):
        """Update the status bar."""
        self.status_bar.configure(text=message)

    def on_closing(self):
        """Handle window close."""
        self.database.close()
        self.destroy()


def main():
    """Main entry point."""
    app = MainApplication()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()


if __name__ == "__main__":
    main()
