import os
import customtkinter as ctk
from Filament_Manager.data_operations import init_excel
from Filament_Manager.app import FilamentManagerApp


def main():
    # Set appearance mode and default color theme
    ctk.set_appearance_mode("system")
    ctk.set_default_color_theme("blue")
    
    # Ensure Excel file exists
    init_excel()
    
    # Create and run the application
    app = FilamentManagerApp()
    app.mainloop()


if __name__ == "__main__":
    main() 