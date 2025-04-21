import customtkinter as ctk
from tkinter import ttk, messagebox
from datetime import datetime
from PIL import Image, ImageTk, ImageDraw

from Filament_Manager.ui_components import configure_treeview_style
from Filament_Manager.data_operations import read_excel_data, write_excel_data, add_print_log_entry


class FilamentUsageDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Register Filament Usage")
        self.geometry("600x700")  # Increased from 500x600
        
        # Store color images
        self.color_images = {}
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Create the form
        self.create_form()

    def create_form(self):
        # Main content frame
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Filament selection frame
        filament_frame = ctk.CTkFrame(content_frame)
        filament_frame.pack(fill="x", pady=10)
        
        ctk.CTkLabel(filament_frame, text="Select Filament:", font=("Roboto", 14, "bold")).pack(pady=(10,0))
        
        # Create Treeview for filament selection instead of Combobox
        tree_frame = ctk.CTkFrame(filament_frame)
        tree_frame.pack(fill="x", pady=10)
        
        # Configure style for dark/light mode compatibility
        configure_treeview_style()
        
        # Create and configure the Treeview
        columns = ("material", "description", "variant", "supplier", "weight")
        self.filament_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", height=6, selectmode="browse", style="Treeview")
        
        # Configure tree column (for color circles)
        self.filament_tree.column("#0", width=30, anchor="center", stretch=False)
        
        # Configure columns
        self.filament_tree.heading("material", text="Material")
        self.filament_tree.heading("description", text="Description")
        self.filament_tree.heading("variant", text="Variant")
        self.filament_tree.heading("supplier", text="Supplier")
        self.filament_tree.heading("weight", text="Available")
        
        self.filament_tree.column("material", width=100)
        self.filament_tree.column("description", width=150)
        self.filament_tree.column("variant", width=100)
        self.filament_tree.column("supplier", width=100)
        self.filament_tree.column("weight", width=100, anchor="center")
        
        # Create and configure scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.filament_tree.yview)
        self.filament_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack the Treeview and scrollbar
        self.filament_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Filament info frame
        self.filament_info_frame = ctk.CTkFrame(filament_frame)
        self.filament_info_frame.pack(fill="x", pady=10)
        
        self.filament_info_label = ctk.CTkLabel(
            self.filament_info_frame,
            text="Select a filament to see details",
            font=("Roboto", 12)
        )
        self.filament_info_label.pack(pady=5)

        # Print name input
        self.print_name_entry = ctk.CTkEntry(
            content_frame,
            placeholder_text="Print name",
            width=300,
            height=35
        )
        self.print_name_entry.pack(pady=10)

        # Required weight input
        self.print_weight_entry = ctk.CTkEntry(
            content_frame,
            placeholder_text="Required weight (grams)",
            width=300,
            height=35
        )
        self.print_weight_entry.pack(pady=10)

        # Button frame at the bottom
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=20, side="bottom")
        
        # Cancel button
        cancel_button = ctk.CTkButton(
            button_frame,
            text="Cancel",
            command=self.cancel,
            fg_color="gray",
            hover_color="darkgray",
            width=120
        )
        cancel_button.pack(side="left", padx=5)
        
        # Register button
        register_button = ctk.CTkButton(
            button_frame,
            text="Register",
            command=self.register,
            fg_color="#4CAF50",  # Green color
            hover_color="#388E3C",  # Darker green
            width=120
        )
        register_button.pack(side="right", padx=5)

        # Bind selection event
        self.filament_tree.bind("<<TreeviewSelect>>", self.on_filament_selected)
        
        # Initial update of the filament list
        self.update_filament_list()

    def create_circle_image(self, hex_color, size=16):
        """Create a circle image for the color preview"""
        if hex_color in self.color_images:
            return self.color_images[hex_color]
            
        # Create a higher resolution image for better anti-aliasing
        scale = 4  # Scale factor for anti-aliasing
        img = Image.new("RGBA", (size * scale, size * scale), (255, 255, 255, 0))
        draw = ImageDraw.Draw(img)
        
        # Draw a larger circle with anti-aliasing
        padding = scale  # Add padding for smoother edges
        draw.ellipse((padding, padding, size * scale - padding, size * scale - padding), 
                     fill=hex_color)
        
        # Resize the image down to the desired size with anti-aliasing
        img = img.resize((size, size), Image.Resampling.LANCZOS)
        
        # Convert to PhotoImage and store
        photo_image = ImageTk.PhotoImage(img)
        self.color_images[hex_color] = photo_image
        return photo_image

    def update_filament_list(self):
        """Update the list of available filaments in the treeview"""
        # Clear existing items
        for item in self.filament_tree.get_children():
            self.filament_tree.delete(item)
        
        data = read_excel_data()
        
        for filament in data:
            try:
                # Create color image
                color_image = self.create_circle_image(filament.hex_color)
                
                # Format weight
                weight_str = f"{filament.weight:,.0f}g"
                
                # Insert item with all information
                self.filament_tree.insert("", "end",
                                        text="",
                                        image=color_image,
                                        values=(filament.material,  # Material
                                               filament.description or "",  # Description
                                               filament.variant,   # Variant
                                               filament.supplier,   # Supplier
                                               weight_str),  # Weight
                                        tags=(filament.code, filament.hex_color))  # Store code and color as tags
                
            except Exception as e:
                print(f"Error processing filament {filament.code}: {str(e)}")
                continue

    def on_filament_selected(self, event):
        """Update the filament info when a selection is made"""
        selection = self.filament_tree.selection()
        if not selection:
            return
            
        item = selection[0]
        tags = self.filament_tree.item(item)["tags"]
        if not tags:
            return
            
        code = tags[0]  # First tag is the filament code
        values = self.filament_tree.item(item)["values"]
        
        if values:
            material, description, variant, supplier, weight = values
            info_text = f"Selected: {material} {variant}"
            if description:
                info_text += f"\nDescription: {description}"
            info_text += f"\nSupplier: {supplier}\nAvailable: {weight}"
            
            self.filament_info_label.configure(text=info_text)

    def cancel(self):
        self.destroy()

    def register(self):
        selection = self.filament_tree.selection()
        if not selection:
            messagebox.showerror("No Selection", "Please select a filament spool first.")
            return

        print_name = self.print_name_entry.get()
        if not print_name:
            messagebox.showerror("No Name", "Please enter a name for the print.")
            return

        try:
            required_weight = float(self.print_weight_entry.get())
            if required_weight <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Weight", "Please enter a valid weight (greater than 0).")
            return

        item = selection[0]
        code = self.filament_tree.item(item)["tags"][0]  # Get code from tags
        
        # Read data
        data = read_excel_data()
        selected_filament = None
        
        # Find the filament and check weight
        for filament in data:
            if filament.code == code:
                selected_filament = filament
                break

        if not selected_filament:
            messagebox.showerror("Error", "Selected filament not found.")
            return

        if selected_filament.weight < required_weight:
            messagebox.showerror("Insufficient Filament", 
                              f"Not enough filament available.\nAvailable: {selected_filament.weight:,.0f}g\nRequired: {required_weight:,.0f}g")
            return

        # Update the filament weight
        for filament in data:
            if filament.code == code:
                new_weight = filament.weight - required_weight
                filament.weight = new_weight
                break

        # Write back to Excel
        write_excel_data(data)

        # Get current timestamp
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")

        # Add to print log in Excel
        add_print_log_entry(
            current_time,
            print_name,
            code,
            selected_filament.material,
            selected_filament.variant,
            required_weight,
            new_weight
        )

        # Refresh the main window's displays
        self.master.refresh_data()  # Refresh filament display
        self.master.load_print_history()  # Reload print history

        messagebox.showinfo("Print Registered", 
                          f"Print '{print_name}' registered!\n{required_weight:,.0f}g filament used.")
        self.destroy() 