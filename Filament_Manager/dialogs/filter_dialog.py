import customtkinter as ctk
from tkinter import ttk, messagebox
import random

from Filament_Manager.ui_components import configure_treeview_style, create_circle_image
from Filament_Manager.data_operations import read_excel_data


class FilterRecommendationsDialog(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        
        self.title("Filter Recommendations")
        self.geometry("900x700")
        
        # Make dialog modal
        self.transient(parent)
        self.grab_set()
        
        # Store color images
        self.color_images = {}
        
        # Load filament data
        self.filament_data = read_excel_data()
        
        # Create the form
        self.create_form()

    def create_form(self):
        # Main content frame with grid
        content_frame = ctk.CTkFrame(self)
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)
        content_frame.grid_columnconfigure(0, weight=1)
        
        # Header
        header_label = ctk.CTkLabel(
            content_frame,
            text="Filter Recommendations",
            font=("Roboto", 24, "bold")
        )
        header_label.grid(row=0, column=0, pady=(0, 20), sticky="w")
        
        # Create search criteria frame
        criteria_frame = ctk.CTkFrame(content_frame)
        criteria_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        criteria_frame.grid_columnconfigure(1, weight=1)
        
        # Material type
        ctk.CTkLabel(criteria_frame, text="Material type:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        
        # Get unique materials
        materials = sorted(set(f.material for f in self.filament_data))
        
        # Add default option
        self.material_var = ctk.StringVar(value="All Materials")
        material_options = ["All Materials"] + materials
        
        material_dropdown = ctk.CTkComboBox(
            criteria_frame,
            values=material_options,
            variable=self.material_var,
            width=200
        )
        material_dropdown.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Minimum weight
        ctk.CTkLabel(criteria_frame, text="Minimum weight (g):").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        
        self.weight_var = ctk.StringVar(value="0")
        weight_entry = ctk.CTkEntry(
            criteria_frame,
            textvariable=self.weight_var,
            width=100
        )
        weight_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        
        # Supplier
        ctk.CTkLabel(criteria_frame, text="Supplier:").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        
        # Get unique suppliers
        suppliers = self.get_unique_suppliers()
        
        self.supplier_var = ctk.StringVar(value="All Suppliers")
        supplier_options = ["All Suppliers"] + [s for s in suppliers if s != "Select supplier"]
        
        supplier_dropdown = ctk.CTkComboBox(
            criteria_frame,
            values=supplier_options,
            variable=self.supplier_var,
            width=200
        )
        supplier_dropdown.grid(row=2, column=1, padx=10, pady=10, sticky="w")
        
        # Purpose
        ctk.CTkLabel(criteria_frame, text="Purpose:").grid(row=3, column=0, padx=10, pady=10, sticky="w")
        
        purposes = ["General", "Strength", "Visual", "Flexibility", "Heat Resistance"]
        self.purpose_var = ctk.StringVar(value=purposes[0])
        
        purpose_dropdown = ctk.CTkComboBox(
            criteria_frame,
            values=purposes,
            variable=self.purpose_var,
            width=200
        )
        purpose_dropdown.grid(row=3, column=1, padx=10, pady=10, sticky="w")
        
        # Results frame (Treeview)
        results_frame = ctk.CTkFrame(content_frame)
        results_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(0, weight=1)
        content_frame.grid_rowconfigure(2, weight=1)
        
        # Configure treeview style
        style = configure_treeview_style()
        
        # Create treeview for displaying results
        columns = ("material", "variant", "supplier", "weight", "suitability")
        self.tree = ttk.Treeview(results_frame, columns=columns, show="headings", selectmode="browse")
        
        # Configure columns
        self.tree.heading("material", text="Material")
        self.tree.heading("variant", text="Variant")
        self.tree.heading("supplier", text="Supplier")
        self.tree.heading("weight", text="Available")
        self.tree.heading("suitability", text="Suitability")
        
        self.tree.column("material", width=150)
        self.tree.column("variant", width=150)
        self.tree.column("supplier", width=150)
        self.tree.column("weight", width=100, anchor="center")
        self.tree.column("suitability", width=150, anchor="center")
        
        # Scrollbars
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        # Pack tree and scrollbar
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        
        # Double click event
        self.tree.bind("<Double-1>", self.on_recommendation_select)
        
        # Button frame at the bottom
        button_frame = ctk.CTkFrame(content_frame)
        button_frame.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        button_frame.grid_columnconfigure(0, weight=1)
        
        # Search button
        search_button = ctk.CTkButton(
            button_frame,
            text="Search Recommendations",
            command=self.search_recommendations,
            height=40,
            font=("Roboto", 14),
            fg_color="#4CAF50",  # Green color
            hover_color="#388E3C"  # Darker green
        )
        search_button.pack(pady=10)
        
        # Close button
        close_button = ctk.CTkButton(
            button_frame,
            text="Close",
            command=self.destroy,
            height=40,
            font=("Roboto", 14)
        )
        close_button.pack(pady=5)

    def get_unique_suppliers(self):
        """Get list of unique suppliers from the database"""
        try:
            data = read_excel_data()
            suppliers = sorted(set(filament.supplier for filament in data))
            return ["Select supplier"] + list(suppliers)  # Add default option
        except Exception as e:
            print(f"Error getting suppliers: {str(e)}")
            return ["Select supplier"]

    def search_recommendations(self):
        """Search for filament recommendations based on criteria"""
        try:
            # Get and validate input values
            material_filter = self.material_var.get()
            supplier_filter = self.supplier_var.get()
            purpose_filter = self.purpose_var.get()
            
            try:
                weight_filter = float(self.weight_var.get())
                if weight_filter < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Error", "Please enter a valid weight value.")
                return
                
            # Clear existing items in the treeview
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            # Filter the filaments
            filtered_results = []
            
            for filament in self.filament_data:
                # Apply material filter
                if material_filter != "All Materials" and filament.material != material_filter:
                    continue
                    
                # Apply supplier filter
                if supplier_filter != "All Suppliers" and filament.supplier != supplier_filter:
                    continue
                    
                # Apply weight filter
                if filament.weight < weight_filter:
                    continue
                    
                # Calculate suitability score based on purpose
                suitability = self.calculate_suitability(filament, purpose_filter)
                
                # Add to filtered results
                filtered_results.append((filament, suitability))
                
            # Sort by suitability score (highest first)
            filtered_results.sort(key=lambda x: x[1], reverse=True)
            
            # Add to treeview
            for filament, suitability in filtered_results:
                # Format weight
                weight_str = f"{filament.weight:,.0f}g"
                
                # Format suitability string
                if suitability >= 80:
                    suitability_str = "★★★★★ Excellent"
                elif suitability >= 60:
                    suitability_str = "★★★★☆ Very Good"
                elif suitability >= 40:
                    suitability_str = "★★★☆☆ Good"
                elif suitability >= 20:
                    suitability_str = "★★☆☆☆ Fair"
                else:
                    suitability_str = "★☆☆☆☆ Poor"
                
                # Insert into treeview
                self.tree.insert("", "end", values=(
                    filament.material,
                    filament.variant,
                    filament.supplier,
                    weight_str,
                    suitability_str
                ), tags=(filament.code,))
                
            # Show no results message if needed
            if not filtered_results:
                messagebox.showinfo("No Results", "No filaments match your search criteria.")
                
        except Exception as e:
            messagebox.showerror("Search Error", f"An error occurred: {str(e)}")

    def calculate_suitability(self, filament, purpose):
        """Calculate a suitability score (0-100) for a filament based on purpose"""
        base_score = random.randint(50, 90)  # Base randomized score
        
        # Apply material-based adjustments for each purpose
        if purpose == "Strength":
            if filament.material in ["PETG", "ABS", "Nylon"]:
                base_score += 20
            elif filament.material in ["PLA"]:
                base_score -= 10
                
        elif purpose == "Visual":
            if "Silk" in filament.variant or "Glossy" in filament.variant:
                base_score += 20
            if filament.material == "PLA":
                base_score += 10
                
        elif purpose == "Flexibility":
            if filament.material == "TPU":
                base_score += 30
            elif filament.material in ["PLA", "PETG"]:
                base_score -= 15
                
        elif purpose == "Heat Resistance":
            if filament.material in ["ABS", "ASA", "PC (Polycarbonate)"]:
                base_score += 25
            elif filament.material == "PLA":
                base_score -= 20
        
        # Ensure the score is within 0-100
        return max(0, min(100, base_score))

    def on_recommendation_select(self, event):
        """Handle double-click on a recommendation"""
        selection = self.tree.selection()
        if not selection:
            return
            
        item = selection[0]
        filament_code = self.tree.item(item, "tags")[0]
        
        # Find the filament data
        selected_filament = None
        for filament in self.filament_data:
            if filament.code == filament_code:
                selected_filament = filament
                break
        
        if selected_filament:
            # Show detailed information
            material_variant = f"{selected_filament.material} {selected_filament.variant}"
            weight = f"{selected_filament.weight:,.0f}g"
            date_str = selected_filament.date_opened.strftime("%Y-%m-%d")
            
            message = (
                f"Filament: {material_variant}\n"
                f"Code: {selected_filament.code}\n"
                f"Supplier: {selected_filament.supplier}\n"
                f"Available weight: {weight}\n"
                f"Date opened: {date_str}\n"
            )
            
            if selected_filament.description:
                message += f"Description: {selected_filament.description}\n"
                
            messagebox.showinfo("Filament Details", message) 