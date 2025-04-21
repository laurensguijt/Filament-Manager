import customtkinter as ctk
from tkinter import colorchooser, messagebox
from tkcalendar import DateEntry
import subprocess
import os
import tempfile

from Filament_Manager.ui_components import ColorPreviewCanvas
from Filament_Manager.report_generator import generate_filament_label
from Filament_Manager.models import FilamentData


class FilamentEditDialog(ctk.CTkToplevel):
    def __init__(self, parent, filament_data: FilamentData):
        super().__init__(parent)
        
        self.title("Edit Filament")
        self.geometry("450x700")  # Increased from 400x650
        
        # Store the original data
        self.filament_data = filament_data
        self.result = None
        
        # Create the form
        self.create_form()
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
    def create_form(self):
        # Main content frame
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Color preview and picker
        color_frame = ctk.CTkFrame(content_frame)
        color_frame.pack(fill="x", pady=10)
        
        self.color_preview = ColorPreviewCanvas(color_frame)
        self.color_preview.pack(side="left", padx=10)
        self.color_preview.set_color(self.filament_data.hex_color)
        
        pick_color_btn = ctk.CTkButton(
            color_frame,
            text="Choose Color",
            command=self.pick_color,
            width=120,
            fg_color="#4CAF50",  # Match the green color
            hover_color="#388E3C"
        )
        pick_color_btn.pack(side="left", padx=10)

        # Material type selection
        material_frame = ctk.CTkFrame(content_frame)
        material_frame.pack(fill="x", pady=10)
        
        material_label = ctk.CTkLabel(material_frame, text="Material Type", font=("Roboto", 14))
        material_label.pack()
        
        self.material_combobox = ctk.CTkComboBox(
            material_frame,
            values=[
                "PLA",
                "PETG",
                "ABS",
                "TPU",
                "Nylon",
                "ASA",
                "PC (Polycarbonate)",
                "PVA",
                "HIPS",
                "PP (Polypropylene)",
                "Wood Fill",
                "Metal Fill",
                "Carbon Fiber"
            ],
            width=200,
            state="readonly"
        )
        self.material_combobox.pack(pady=5)
        self.material_combobox.set(self.filament_data.material)

        # Variant selection
        variant_frame = ctk.CTkFrame(content_frame)
        variant_frame.pack(fill="x", pady=10)
        
        variant_label = ctk.CTkLabel(variant_frame, text="Variant", font=("Roboto", 14))
        variant_label.pack()
        
        self.variant_combobox = ctk.CTkComboBox(
            variant_frame,
            values=[
                "Mat",
                "Glossy",
                "Silk",
                "Carbon Filled",
                "Glass Filled",
                "Metal Filled",
                "Glow in the Dark",
                "Transparent",
                "UV Resistant",
                "High Flow",
                "High Impact",
                "Standard"
            ],
            width=200,
            state="readonly"
        )
        self.variant_combobox.pack(pady=5)
        self.variant_combobox.set(self.filament_data.variant)
        
        # Supplier
        self.supplier_entry = ctk.CTkEntry(content_frame, placeholder_text="Supplier")
        self.supplier_entry.pack(fill="x", pady=10)
        self.supplier_entry.insert(0, self.filament_data.supplier)
        
        # Description
        self.description_entry = ctk.CTkEntry(content_frame, placeholder_text="Description (optional)")
        self.description_entry.pack(fill="x", pady=10)
        self.description_entry.insert(0, self.filament_data.description or "")
        
        # Weight
        weight_frame = ctk.CTkFrame(content_frame)
        weight_frame.pack(fill="x", pady=10)
        
        # Total weight
        weight_label = ctk.CTkLabel(weight_frame, text="Total Weight (g)", font=("Roboto", 14))
        weight_label.pack()
        
        self.weight_entry = ctk.CTkEntry(weight_frame, placeholder_text="Weight (grams)")
        self.weight_entry.pack(fill="x", pady=5)
        self.weight_entry.insert(0, str(self.filament_data.weight))
        
        # Empty spool weight
        empty_spool_label = ctk.CTkLabel(weight_frame, text="Empty Spool Weight (g)", font=("Roboto", 14))
        empty_spool_label.pack(pady=(10,0))
        
        self.empty_spool_weight_entry = ctk.CTkEntry(weight_frame, placeholder_text="Empty Spool Weight (grams)")
        self.empty_spool_weight_entry.pack(fill="x", pady=5)
        self.empty_spool_weight_entry.insert(0, str(self.filament_data.empty_spool_weight))
        
        # Date
        date_frame = ctk.CTkFrame(content_frame)
        date_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(date_frame, text="Date opened:").pack(side="left", padx=5)
        self.date_entry = DateEntry(
            date_frame,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        self.date_entry.pack(side="left", padx=5)
        self.date_entry.set_date(self.filament_data.date_opened)

        # Save button above the bottom buttons
        save_button = ctk.CTkButton(
            self,
            text="Save",
            command=self.save,
            width=120,
            fg_color="#4CAF50",  # Green color
            hover_color="#388E3C"  # Darker green
        )
        save_button.pack(pady=(20, 10))
        
        # Button frame at the bottom
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=(0, 20), side="bottom")
        
        # Label button
        label_button = ctk.CTkButton(
            button_frame,
            text="Generate Label",
            command=self.generate_label,
            fg_color="#2196F3",  # Blue color
            hover_color="#1976D2",  # Darker blue
            width=120
        )
        label_button.pack(side="left", padx=5)
        
        # Delete button
        delete_button = ctk.CTkButton(
            button_frame,
            text="Delete",
            command=self.delete,
            fg_color="red",
            hover_color="darkred",
            width=120
        )
        delete_button.pack(side="left", padx=5)
        
        # Cancel button
        cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self.cancel,
            fg_color="gray",
            hover_color="darkgray",
            width=120
        )
        cancel_button.pack(side="right", padx=5)

    def save(self):
        try:
            weight = float(self.weight_entry.get())
            empty_spool_weight = float(self.empty_spool_weight_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Weight and Empty Spool Weight must be numbers")
            return
            
        self.result = {
            "action": "save",
            'code': self.filament_data.code,
            'color': self.material_combobox.get(),
            'variant': self.variant_combobox.get(),
            'supplier': self.supplier_entry.get(),
            'date': self.date_entry.get_date(),
            'weight': weight,
            'hex_color': self.color_preview.color,
            'empty_spool_weight': empty_spool_weight,
            'description': self.description_entry.get()
        }
        self.destroy()

    def delete(self):
        if messagebox.askyesno("Delete", "Are you sure you want to delete this filament?"):
            self.result = {"action": "delete", "code": self.filament_data.code}
            self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()

    def pick_color(self):
        color = colorchooser.askcolor(title="Choose Filament Color")
        if color[1]:
            self.color_preview.set_color(color[1])

    def generate_label(self):
        """Generate a PDF label for the filament"""
        # Update filament_data with current values so the label uses the current information
        try:
            weight = float(self.weight_entry.get())
            empty_spool_weight = float(self.empty_spool_weight_entry.get())
            
            updated_filament_data = FilamentData(
                code=self.filament_data.code,
                material=self.material_combobox.get(),
                variant=self.variant_combobox.get(),
                supplier=self.supplier_entry.get(),
                date_opened=self.date_entry.get_date(),
                weight=weight,
                hex_color=self.color_preview.color,
                empty_spool_weight=empty_spool_weight,
                description=self.description_entry.get()
            )
            
            generate_filament_label(updated_filament_data)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate label:\n{str(e)}")

    def show_error(self, title, message):
        messagebox.showerror(title, message)

    def show_info(self, title, message):
        messagebox.showinfo(title, message) 