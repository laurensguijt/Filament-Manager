import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import os
import shutil

from Filament_Manager.ui_components import ColorPreviewCanvas, configure_treeview_style
from Filament_Manager.data_operations import (
    init_excel, 
    read_excel_data, 
    write_excel_data, 
    read_print_log
)
from Filament_Manager.report_generator import generate_inventory_report
from Filament_Manager.dialogs.filament_edit_dialog import FilamentEditDialog
from Filament_Manager.dialogs.add_filament_dialog import AddFilamentDialog
from Filament_Manager.dialogs.usage_dialog import FilamentUsageDialog
from Filament_Manager.dialogs.settings_dialog import SettingsDialog
from Filament_Manager.dialogs.filter_dialog import FilterRecommendationsDialog
from Filament_Manager.dialogs.print_history_edit_dialog import PrintHistoryEditDialog
from Filament_Manager.models import FilamentData


class FilamentManagerApp(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Set up the main window
        self.title("Filament Manager")
        self.geometry("1800x1000")
        
        # Set theme based on setting

        
        # Initialize tooltip
        self.tooltip = None
        
        # Initialize ttk style
        configure_treeview_style()
        
        # Store color images
        self.color_images = {}
        
        # Create main container
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Create main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=3)

        # Create sidebar for input
        self.create_sidebar()

        # Create right frame for display
        self.create_display_frame()

        # Initialize data
        self.refresh_data()

        # Create settings button
        self.create_settings_button()

        # Load print history from Excel
        self.load_print_history()
        
        # Create report button
        self.create_report_button()

    def create_settings_button(self):
        # Create settings button frame
        settings_frame = ctk.CTkFrame(self)
        settings_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="w")
        
        # Settings button
        settings_button = ctk.CTkButton(
            settings_frame,
            text="⚙️ Settings",
            command=self.open_settings,
            width=120,
            height=32,
            font=ctk.CTkFont(size=14)
        )
        settings_button.pack(side="left", padx=5)

    def open_settings(self):
        """Open the settings dialog"""
        dialog = SettingsDialog(self)
        self.wait_window(dialog)

    def change_appearance_mode(self, new_appearance_mode):
        """Change the appearance mode and update all Treeview styles"""
        ctk.set_appearance_mode(new_appearance_mode.lower())
        
        # Reconfigure treeview style
        configure_treeview_style()
        
        # Force update of all treeviews
        if hasattr(self, 'tree'):
            self.tree.configure(style="Treeview")
            for item in self.tree.get_children():
                self.tree.item(item, tags=())
        
        if hasattr(self, 'print_history'):
            self.print_history.configure(style="Treeview")
            for item in self.print_history.get_children():
                self.print_history.item(item, tags=())
        
        # Refresh data to ensure proper colors
        self.refresh_data()

    def create_sidebar(self):
        sidebar = ctk.CTkFrame(self.main_frame, width=400)
        sidebar.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        sidebar.grid_propagate(False)

        # Title with modern styling
        title_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        title_frame.pack(fill="x", pady=20)
        title_frame.grid_columnconfigure(0, weight=1)
        
        title_label = ctk.CTkLabel(
            title_frame, 
            text="Filament Manager",
            font=("Roboto", 24, "bold")
        )
        title_label.pack(anchor="center")

        # Create buttons frame
        buttons_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=20, pady=20)

        # Add New Filament button
        add_button = ctk.CTkButton(
            buttons_frame,
            text="Add New Filament",
            command=self.open_add_filament_dialog,
            height=50,
            font=("Roboto", 16),
            fg_color="#2196F3",  # Blue color
            hover_color="#1976D2",  # Darker blue
            corner_radius=8
        )
        add_button.pack(fill="x", pady=10)

        # Register Usage button
        usage_button = ctk.CTkButton(
            buttons_frame,
            text="Register Filament Usage",
            command=self.open_usage_dialog,
            height=50,
            font=("Roboto", 16),
            fg_color="#4CAF50",  # Green color
            hover_color="#388E3C",  # Darker green
            corner_radius=8
        )
        usage_button.pack(fill="x", pady=10)

        # Filter Recommendations button
        filter_button = ctk.CTkButton(
            buttons_frame,
            text="Filter Recommendations",
            command=self.open_filter_dialog,
            height=50,
            font=("Roboto", 16),
            fg_color="#9C27B0",  # Purple color
            hover_color="#7B1FA2",  # Darker purple
            corner_radius=8
        )
        filter_button.pack(fill="x", pady=10)

    def open_add_filament_dialog(self):
        """Open a dialog to add a new filament"""
        dialog = AddFilamentDialog(self)
        self.wait_window(dialog)

    def open_usage_dialog(self):
        """Open a dialog to register filament usage"""
        dialog = FilamentUsageDialog(self)
        self.wait_window(dialog)

    def open_filter_dialog(self):
        """Open the filter recommendations dialog"""
        dialog = FilterRecommendationsDialog(self)
        self.wait_window(dialog)

    def create_display_frame(self):
        """Create the main display frame with the filament overview"""
        display_frame = ctk.CTkFrame(self.main_frame)
        display_frame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")
        display_frame.grid_columnconfigure(0, weight=1)
        display_frame.grid_rowconfigure(1, weight=1)

        # Title with search bar
        header_frame = ctk.CTkFrame(display_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        header_frame.grid_columnconfigure(0, weight=1)
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="Filament Overview",
            font=("Roboto", 24, "bold")
        )
        title_label.grid(row=0, column=0, sticky="w")

        # Create scrollable frame for filament overview
        scroll_frame = ctk.CTkScrollableFrame(display_frame, width=800, height=400)
        scroll_frame.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")

        # Define columns
        headers = ["Material", "Description", "Code", "Supplier", "Date", "Weight (g)", "Empty Spool (g)", ""]  # Extra column for edit button
        
        # Add headers
        for i, header in enumerate(headers):
            label = ctk.CTkLabel(
                scroll_frame,
                text=header,
                font=ctk.CTkFont(weight="bold"),
                height=35  # Set fixed height for the header
            )
            label.grid(row=0, column=i, padx=10, pady=(10, 10), sticky="nsew")  # Changed sticky to nsew for full cell coverage
            scroll_frame.grid_columnconfigure(i, weight=1)

        # Store reference to scroll_frame for updating data
        self.scroll_frame = scroll_frame
        
        # Print History Section
        history_label = ctk.CTkLabel(
            display_frame,
            text="Print History",
            font=("Roboto", 16, "bold")
        )
        history_label.grid(row=2, column=0, padx=20, pady=(20, 5), sticky="w")

        history_frame = ctk.CTkFrame(display_frame)
        history_frame.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        history_frame.grid_columnconfigure(0, weight=1)
        history_frame.grid_rowconfigure(0, weight=1)

        # Create scrollable frame for print history with matching style
        if ctk.get_appearance_mode() == "Dark":
            frame_color = "#2b2b2b"
        else:
            frame_color = "#dbdbdb"
            
        history_scroll_frame = ctk.CTkScrollableFrame(history_frame, width=800, height=200, fg_color=frame_color)
        history_scroll_frame.grid(row=0, column=0, sticky="nsew")

        # Define history columns
        history_headers = ["Date & Time", "Print Name", "Material", "Used", "Remaining", ""]  # Added empty column for edit button
        
        # Add history headers
        for i, header in enumerate(history_headers):
            label = ctk.CTkLabel(
                history_scroll_frame,
                text=header,
                font=ctk.CTkFont(weight="bold"),
                height=35  # Set fixed height for the header
            )
            label.grid(row=0, column=i, padx=10, pady=(10, 10), sticky="nsew")  # Changed sticky to nsew for full cell coverage
            history_scroll_frame.grid_columnconfigure(i, weight=1)

        # Store reference to history_scroll_frame for updating data
        self.history_scroll_frame = history_scroll_frame

        # Configure row weights for display_frame to make print history smaller
        display_frame.grid_rowconfigure(1, weight=3)  # Filament overview gets more space
        display_frame.grid_rowconfigure(3, weight=1)  # Print history gets less space

    def refresh_data(self):
        # Clear existing items
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        
        # Add headers
        headers = ["Material", "Description", "Code", "Supplier", "Date", "Weight (g)", "Empty Spool (g)", ""]
        for col_idx, header in enumerate(headers):
            label = ctk.CTkLabel(
                self.scroll_frame,
                text=header,
                font=ctk.CTkFont(weight="bold")
            )
            label.grid(row=0, column=col_idx, padx=10, pady=(5, 10), sticky="w")
        
        data = read_excel_data()
        
        for row_idx, filament in enumerate(data, start=1):
            try:
                # Create color preview frame
                color_frame = ctk.CTkFrame(self.scroll_frame)
                color_frame.grid(row=row_idx, column=0, padx=10, pady=3, sticky="w")
                
                # Add color preview to frame
                color_preview = ColorPreviewCanvas(color_frame, width=20, height=20)
                color_preview.set_color(filament.hex_color)
                color_preview.pack(side="left", padx=(0, 5))
                
                # Add material and variant label next to color preview
                material_variant_label = ctk.CTkLabel(color_frame, text=f"{filament.material} {filament.variant}")
                material_variant_label.pack(side="left")
                
                # Add description
                description_label = ctk.CTkLabel(self.scroll_frame, text=filament.description or "")
                description_label.grid(row=row_idx, column=1, padx=10, pady=3, sticky="w")
                
                # Format the date
                date_str = filament.date_opened.strftime("%Y-%m-%d") if isinstance(filament.date_opened, datetime) else str(filament.date_opened).split(' ')[0]
                
                # Create the display values
                display_values = [
                    filament.code,  # Code
                    filament.supplier,  # Supplier
                    date_str,  # Date
                    f"{filament.weight:,.0f}g",  # Weight
                    f"{filament.empty_spool_weight:,.0f}g"  # Empty spool weight
                ]
                
                # Add labels for each value
                for col_idx, value in enumerate(display_values, start=2):
                    label = ctk.CTkLabel(self.scroll_frame, text=value)
                    label.grid(row=row_idx, column=col_idx, padx=10, pady=3, sticky="w")
                
                # Add edit button with tooltip
                edit_button = ctk.CTkButton(
                    self.scroll_frame,
                    text="⚙️",
                    width=30,
                    height=24,
                    command=lambda f=filament: self.edit_filament([
                        f.code,            # code
                        f.material,        # material
                        f.variant,         # variant
                        f.supplier,        # supplier
                        f.date_opened,     # date_opened
                        f.weight,          # weight
                        f.empty_spool_weight,  # empty_spool_weight
                        f.hex_color,       # hex_color
                        f.description      # description
                    ]),
                    fg_color="transparent",
                    hover_color=("gray80", "gray20"),
                    font=ctk.CTkFont(size=16)
                )
                edit_button.grid(row=row_idx, column=len(headers)-1, padx=10, pady=3)
                
                # Create tooltip
                tooltip_text = "Configure Filament"
                edit_button.bind("<Enter>", lambda event, btn=edit_button, tip=tooltip_text: 
                    self.show_tooltip(event, btn, tip))
                edit_button.bind("<Leave>", self.hide_tooltip)
                
            except Exception as e:
                print(f"Error refreshing data row {row_idx}: {str(e)}")
                continue

    def load_print_history(self):
        """Load print history from Excel and display it in the scrollable frame"""
        # Clear existing items
        for widget in self.history_scroll_frame.winfo_children():
            widget.destroy()

        # Add headers
        headers = ["Date & Time", "Print Name", "Material", "Used", "Remaining", ""]  # Added empty column for edit button
        for col_idx, header in enumerate(headers):
            label = ctk.CTkLabel(
                self.history_scroll_frame,
                text=header,
                font=ctk.CTkFont(weight="bold")
            )
            label.grid(row=0, column=col_idx, padx=10, pady=(5, 10), sticky="w")

        # Get print log entries and filament data
        log_entries = read_print_log()
        filament_data = read_excel_data()
        
        # Create a dictionary of filament codes to hex colors
        filament_colors = {filament.code: filament.hex_color for filament in filament_data}
        
        # Add entries in reverse order (newest first)
        for row_idx, entry in enumerate(reversed(log_entries), start=1):
            try:
                # Get the hex color for this filament
                hex_color = filament_colors.get(entry.filament_code, "#000000")
                
                # Format the material string
                material_str = f"{entry.material} {entry.variant}"
                
                # Format the weights
                used_str = f"{entry.used_weight:,.0f}g"
                remaining_str = f"{entry.remaining_weight:,.0f}g"
                
                # Create the display values for timestamp and print name
                display_values = [
                    entry.timestamp,
                    entry.print_name
                ]
                
                # Add labels for timestamp and print name
                for col_idx, value in enumerate(display_values):
                    label = ctk.CTkLabel(self.history_scroll_frame, text=value)
                    label.grid(row=row_idx, column=col_idx, padx=10, pady=3, sticky="w")
                
                # Create material frame with color circle and text
                material_frame = ctk.CTkFrame(self.history_scroll_frame)
                material_frame.grid(row=row_idx, column=2, padx=10, pady=3, sticky="w")
                
                # Add color preview to material frame
                color_preview = ColorPreviewCanvas(material_frame, width=20, height=20)
                color_preview.set_color(hex_color)
                color_preview.pack(side="left", padx=(0, 5))
                
                # Add material text next to color preview
                material_label = ctk.CTkLabel(material_frame, text=material_str)
                material_label.pack(side="left")
                
                # Add remaining values (Used and Remaining)
                remaining_values = [
                    used_str,
                    remaining_str
                ]
                
                # Add labels for remaining values
                for col_idx, value in enumerate(remaining_values, start=3):
                    label = ctk.CTkLabel(self.history_scroll_frame, text=value)
                    label.grid(row=row_idx, column=col_idx, padx=10, pady=3, sticky="w")
                
                # Add edit button with tooltip
                edit_button = ctk.CTkButton(
                    self.history_scroll_frame,
                    text="⚙️",
                    width=30,
                    height=24,
                    command=lambda d={"timestamp": entry.timestamp, 
                                     "print_name": entry.print_name,
                                     "material": material_str,
                                     "used_weight": used_str,
                                     "remaining_weight": remaining_str,
                                     "filament_code": entry.filament_code}: self._on_print_history_double_click(d),
                    fg_color="transparent",
                    hover_color=("gray80", "gray20"),
                    font=ctk.CTkFont(size=16)
                )
                edit_button.grid(row=row_idx, column=len(headers)-1, padx=10, pady=3)
                
                # Create tooltip
                tooltip_text = "Edit Print Usage"
                edit_button.bind("<Enter>", lambda event, btn=edit_button, tip=tooltip_text: 
                    self.show_tooltip(event, btn, tip))
                edit_button.bind("<Leave>", self.hide_tooltip)
                
            except Exception as e:
                print(f"Error loading print history row {row_idx}: {str(e)}")
                continue

    def edit_filament(self, row_data):
        """Edit a filament entry"""
        try:
            # Create FilamentData object with the correct order of fields
            filament_data = FilamentData(
                code=row_data[0],
                material=row_data[1],
                variant=row_data[2],
                supplier=row_data[3],
                date_opened=row_data[4],
                weight=float(row_data[5]),
                hex_color=row_data[7],  # hex_color is at index 7
                empty_spool_weight=float(row_data[6]),  # empty_spool_weight is at index 6
                description=row_data[8] if len(row_data) > 8 else ""  # description is at index 8 if present
            )
            
            dialog = FilamentEditDialog(self, filament_data)
            self.wait_window(dialog)
            
            if dialog.result:
                if dialog.result["action"] == "delete":
                    self._handle_delete(dialog.result["code"])
                elif dialog.result["action"] == "save":
                    self._handle_save(dialog.result)
                    
        except Exception as e:
            messagebox.showerror("Error", f"Failed to edit filament: {str(e)}")
            return

    def _handle_delete(self, code):
        """Handle deletion of a filament entry"""
        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete filament {code}?"):
            data = read_excel_data()
            data = [row for row in data if row.code != code]
            write_excel_data(data)
            
            # Refresh both the filament list and print history
            self.refresh_data()
            self.load_print_history()
            
            messagebox.showinfo("Deleted", "Filament successfully deleted!")

    def _handle_save(self, result):
        """Handle saving of a filament entry"""
        data = read_excel_data()
        for filament in data:
            if filament.code == result['code']:
                try:
                    # Update filament with new values
                    filament.material = result['color']
                    filament.variant = result['variant']
                    filament.supplier = result['supplier']
                    filament.date_opened = result['date']
                    filament.weight = float(result['weight'])
                    filament.hex_color = result['hex_color']
                    filament.empty_spool_weight = float(result['empty_spool_weight'])
                    filament.description = result['description']
                except Exception as e:
                    self.show_error("Update Error", f"Failed to update filament: {str(e)}")
                    return
                break
        
        write_excel_data(data)
        self.refresh_data()
        self.show_info("Updated", "Filament successfully updated!")

    def _on_print_history_double_click(self, print_data):
        """Handle click on print history edit button"""
        try:
            dialog = PrintHistoryEditDialog(self, print_data)
            self.wait_window(dialog)
                    
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while opening the edit dialog: {str(e)}")

    def create_backup(self):
        """Create a backup of the Excel files"""
        try:
            # Get the current date and time for the backup filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"filament_manager_backup_{timestamp}.xlsx"
            
            # Ask user for backup location
            backup_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=default_filename,
                title="Save Backup As"
            )
            
            if backup_path:
                # Get the directory of the current script
                current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                
                # Copy the Excel file to the backup location
                shutil.copy2(os.path.join(current_dir, "filament_data.xlsx"), backup_path)
                
                messagebox.showinfo(
                    "Backup Created",
                    f"Backup successfully created at:\n{backup_path}"
                )
        except Exception as e:
            messagebox.showerror(
                "Backup Error",
                f"Failed to create backup:\n{str(e)}"
            )

    def restore_backup(self):
        """Restore from a backup Excel file"""
        try:
            # Ask user for backup file
            backup_path = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx")],
                title="Select Backup File to Restore"
            )
            
            if backup_path:
                # Confirm restoration
                if messagebox.askyesno(
                    "Confirm Restore",
                    "This will replace your current data with the backup data.\n" +
                    "Are you sure you want to continue?"
                ):
                    # Get the directory of the current script
                    current_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                    target_path = os.path.join(current_dir, "filament_data.xlsx")
                    
                    # Create a backup of the current file first
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    auto_backup_path = os.path.join(
                        current_dir,
                        f"filament_data_auto_backup_{timestamp}.xlsx"
                    )
                    shutil.copy2(target_path, auto_backup_path)
                    
                    # Copy the backup file to the current location
                    shutil.copy2(backup_path, target_path)
                    
                    messagebox.showinfo(
                        "Restore Complete",
                        "Backup has been successfully restored.\n" +
                        f"A backup of your previous data was created at:\n{auto_backup_path}"
                    )
                    
                    # Refresh the display
                    self.refresh_data()
        except Exception as e:
            messagebox.showerror(
                "Restore Error",
                f"Failed to restore backup:\n{str(e)}"
            )

    def show_tooltip(self, event, widget, text):
        """Show tooltip near the widget"""
        if self.tooltip:
            self.hide_tooltip()
            
        x, y, _, _ = widget.bbox("insert")
        x += widget.winfo_rootx() + 25
        y += widget.winfo_rooty() + 20

        # Creates a toplevel window
        self.tooltip = tk.Toplevel()
        # Leaves only the label and removes the app window
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        
        # Creates tooltip label
        label = tk.Label(
            self.tooltip,
            text=text,
            justify='left',
            background='#2b2b2b' if ctk.get_appearance_mode() == "Dark" else '#ffffff',
            foreground='white' if ctk.get_appearance_mode() == "Dark" else 'black',
            relief='solid',
            borderwidth=1,
            font=("Roboto", "10", "normal")
        )
        label.pack()

    def hide_tooltip(self, event=None):
        """Hide the tooltip"""
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def show_error(self, title, message):
        """Show error message dialog"""
        messagebox.showerror(title, message)

    def show_info(self, title, message):
        """Show info message dialog"""
        messagebox.showinfo(title, message)

    def create_report_button(self):
        """Create the generate report button"""
        report_button = ctk.CTkButton(
            self.main_frame,
            text="Generate Inventory Report",
            command=generate_inventory_report,
            height=32,
            font=ctk.CTkFont(size=14),
            fg_color="#2196F3",  # Blue color
            hover_color="#1976D2"  # Darker blue
        )
        report_button.grid(row=1, column=1, padx=20, pady=(0, 10), sticky="e") 