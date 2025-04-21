import customtkinter as ctk
from tkinter import messagebox, filedialog
import os
import shutil
import webbrowser
from datetime import datetime

from Filament_Manager.report_generator import generate_inventory_report


class SettingsDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Settings")
        self.geometry("550x750")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Store parent for reference (to use parent methods)
        self.parent = parent
        
        # Create the settings form
        self.create_form()

    def create_form(self):
        # Main content frame
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title
        title_label = ctk.CTkLabel(
            content_frame, 
            text="Settings",
            font=("Roboto", 24, "bold")
        )
        title_label.pack(pady=(0, 20))
        
        # Appearance settings
        appearance_frame = ctk.CTkFrame(content_frame)
        appearance_frame.pack(fill="x", pady=10)
        
        appearance_label = ctk.CTkLabel(
            appearance_frame,
            text="Appearance",
            font=("Roboto", 16, "bold")
        )
        appearance_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Theme selection
        theme_frame = ctk.CTkFrame(appearance_frame, fg_color="transparent")
        theme_frame.pack(fill="x", padx=20, pady=5)
        
        theme_label = ctk.CTkLabel(theme_frame, text="Theme:")
        theme_label.pack(side="left", padx=(0, 10))
        
        # Get current appearance mode
        current_appearance = ctk.get_appearance_mode()
        
        self.appearance_var = ctk.StringVar(value=current_appearance)
        appearance_combobox = ctk.CTkComboBox(
            theme_frame,
            values=["System", "Light", "Dark"],
            variable=self.appearance_var,
            width=120,
            command=self.change_appearance
        )
        appearance_combobox.pack(side="left")
        
        # Reports section
        reports_frame = ctk.CTkFrame(content_frame)
        reports_frame.pack(fill="x", pady=10)
        
        reports_label = ctk.CTkLabel(
            reports_frame,
            text="Reports",
            font=("Roboto", 16, "bold")
        )
        reports_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Generate inventory report button
        inventory_button = ctk.CTkButton(
            reports_frame,
            text="Generate Inventory Report",
            command=generate_inventory_report,
            width=200,
            fg_color="#2196F3",
            hover_color="#1976D2"
        )
        inventory_button.pack(pady=10, padx=20)
        
        # Backup and restore section
        backup_frame = ctk.CTkFrame(content_frame)
        backup_frame.pack(fill="x", pady=10)
        
        backup_label = ctk.CTkLabel(
            backup_frame,
            text="Backup and Restore",
            font=("Roboto", 16, "bold")
        )
        backup_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        # Backup button
        backup_button = ctk.CTkButton(
            backup_frame,
            text="Create Backup",
            command=self.create_backup,
            width=200,
            fg_color="#4CAF50",
            hover_color="#388E3C"
        )
        backup_button.pack(pady=10, padx=20)
        
        # Restore button
        restore_button = ctk.CTkButton(
            backup_frame,
            text="Restore from Backup",
            command=self.restore_backup,
            width=200,
            fg_color="#FFA000",
            hover_color="#FF8F00"
        )
        restore_button.pack(pady=10, padx=20)
        
        # About section
        about_frame = ctk.CTkFrame(content_frame)
        about_frame.pack(fill="x", pady=10)
        
        about_label = ctk.CTkLabel(
            about_frame,
            text="About Filament Manager",
            font=("Roboto", 16, "bold")
        )
        about_label.pack(anchor="w", padx=10, pady=(10, 5))
        
        version_label = ctk.CTkLabel(
            about_frame,
            text="Version: 1.0.0",
            font=("Roboto", 12)
        )
        version_label.pack(anchor="w", padx=20, pady=2)
        
        author_label = ctk.CTkLabel(
            about_frame,
            text="Created by: Laurens Guijt",
            font=("Roboto", 12)
        )
        author_label.pack(anchor="w", padx=20, pady=2)
        
        # GitHub link
        github_frame = ctk.CTkFrame(about_frame, fg_color="transparent")
        github_frame.pack(fill="x", padx=20, pady=5, anchor="w")
        
        github_label = ctk.CTkLabel(
            github_frame,
            text="GitHub Repository:",
            font=("Roboto", 12)
        )
        github_label.pack(side="left", pady=2)
        
        github_link = ctk.CTkButton(
            github_frame,
            text="View on GitHub",
            command=lambda: webbrowser.open("https://github.com/laurensguijt/Filament-Manager"),
            fg_color="transparent",
            text_color=("#1a73e8", "#61afef"),
            hover=True,
            font=("Roboto", 12, "underline"),
            width=250,
            height=20
        )
        github_link.pack(side="left", padx=5, pady=2)
        
        # Close button at the bottom
        close_button = ctk.CTkButton(
            self,
            text="Close",
            command=self.destroy,
            width=120,
            fg_color="gray",
            hover_color="darkgray"
        )
        close_button.pack(pady=20)

    def change_appearance(self, new_appearance_mode):
        """Change the appearance mode (Light/Dark)"""
        ctk.set_appearance_mode(new_appearance_mode.lower())
        self.parent.change_appearance_mode(new_appearance_mode)

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
                current_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
                
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
                    current_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
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
                    self.parent.refresh_data()
        except Exception as e:
            messagebox.showerror(
                "Restore Error",
                f"Failed to restore backup:\n{str(e)}"
            ) 