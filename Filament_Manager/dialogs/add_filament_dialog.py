import customtkinter as ctk
from tkinter import colorchooser, messagebox
from tkcalendar import DateEntry

from Filament_Manager.ui_components import ColorPreviewCanvas
from Filament_Manager.data_operations import get_next_code, read_excel_data, write_excel_data
from Filament_Manager.models import FilamentData
from Filament_Manager.report_generator import generate_filament_label


class AddFilamentDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Add New Filament")
        self.geometry("450x700")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Default color
        self.current_color_hex = "#000000"
        
        # Create the form
        self.create_form()

    def create_form(self):
        # Main content frame
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Color preview and picker
        color_frame = ctk.CTkFrame(content_frame)
        color_frame.pack(fill="x", pady=10)
        
        self.color_preview = ColorPreviewCanvas(color_frame)
        self.color_preview.pack(side="left", padx=10)
        
        pick_color_btn = ctk.CTkButton(
            color_frame,
            text="Choose Color",
            command=self.pick_color,
            width=120
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
        self.material_combobox.set("Select material")

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
        self.variant_combobox.set("Select variant")

        # Input fields
        fields_frame = ctk.CTkFrame(content_frame)
        fields_frame.pack(fill="x", pady=10)

        self.supplier_entry = ctk.CTkEntry(
            fields_frame,
            placeholder_text="Supplier",
            width=200,
            height=35
        )
        self.supplier_entry.pack(pady=10)

        self.weight_entry = ctk.CTkEntry(
            fields_frame,
            placeholder_text="Weight (grams)",
            width=200,
            height=35
        )
        self.weight_entry.pack(pady=10)

        self.empty_spool_weight_entry = ctk.CTkEntry(
            fields_frame,
            placeholder_text="Empty Spool Weight (grams)",
            width=200,
            height=35
        )
        self.empty_spool_weight_entry.pack(pady=10)

        # Description field
        self.description_entry = ctk.CTkEntry(
            fields_frame,
            placeholder_text="Description (optional)",
            width=200,
            height=35
        )
        self.description_entry.pack(pady=10)

        # Date picker
        date_frame = ctk.CTkFrame(fields_frame)
        date_frame.pack(pady=10)
        
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
        
        # Label options frame
        label_options_frame = ctk.CTkFrame(button_frame)
        label_options_frame.pack(side="left", padx=5, pady=5)
        
        # QR code checkbox
        self.include_qr_var = ctk.BooleanVar(value=True)
        qr_checkbox = ctk.CTkCheckBox(
            label_options_frame,
            text="Include QR Code",
            variable=self.include_qr_var
        )
        qr_checkbox.pack(side="left", padx=5)
        
        # Barcode checkbox
        self.include_barcode_var = ctk.BooleanVar(value=True)
        barcode_checkbox = ctk.CTkCheckBox(
            label_options_frame,
            text="Include Barcode",
            variable=self.include_barcode_var
        )
        barcode_checkbox.pack(side="left", padx=5)
        
        # Label button
        label_button = ctk.CTkButton(
            button_frame,
            text="Generate Label",
            command=lambda: self.generate_label(new_filament),
            fg_color="#2196F3",  # Blue color
            hover_color="#1976D2",  # Darker blue
            width=120
        )
        label_button.pack(side="left", padx=5)

    def pick_color(self):
        color = colorchooser.askcolor(title="Choose Filament Color")
        if color[1]:
            self.color_preview.set_color(color[1])
            self.current_color_hex = color[1]

    def cancel(self):
        self.destroy()

    def save(self):
        material = self.material_combobox.get()
        variant = self.variant_combobox.get()
        supplier = self.supplier_entry.get()
        date_opened = self.date_entry.get_date()
        weight = self.weight_entry.get()
        empty_spool_weight = self.empty_spool_weight_entry.get()
        description = self.description_entry.get()
        hex_color = getattr(self, 'current_color_hex', '#000000')

        if material == "Select material" or variant == "Select variant" or not (supplier and date_opened and weight and empty_spool_weight):
            messagebox.showerror("Empty Field", "Please fill in all required fields.")
            return

        try:
            weight_gram = float(weight)
            empty_spool_weight = float(empty_spool_weight)
        except ValueError:
            messagebox.showerror("Error", "Please make sure the weight and empty spool weight are numbers.")
            return

        try:
            # Read existing data
            data = read_excel_data()
            
            # Generate next available code
            code = get_next_code(data)
            
            # Create new FilamentData object
            new_filament = FilamentData(
                code=code,
                material=material,
                variant=variant,
                supplier=supplier,
                date_opened=date_opened,
                weight=weight_gram,
                hex_color=hex_color,
                empty_spool_weight=empty_spool_weight,
                description=description
            )
            
            # Add to data and write back to Excel
            data.append(new_filament)
            write_excel_data(data)

            # Refresh the main window's display
            self.master.refresh_data()
            
            # Show success message
            messagebox.showinfo("Success", f"New filament {code} added successfully!")
            
            # Close the dialog
            self.destroy()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save filament: {str(e)}")
            return

    def generate_label(self, filament_data):
        """Generate a label for a newly created filament"""
        try:
            generate_filament_label(
                filament_data,
                include_qr=self.include_qr_var.get(),
                include_barcode=self.include_barcode_var.get()
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate label:\n{str(e)}") 