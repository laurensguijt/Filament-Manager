import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw
from tkinter import ttk, colorchooser, messagebox


class ColorPreviewCanvas(ctk.CTkCanvas):
    """A custom canvas widget for previewing filament colors"""
    
    def __init__(self, parent, width=20, height=20, **kwargs):
        super().__init__(parent, width=width, height=height, **kwargs)
        self.width = width
        self.height = height
        self.color = "#000000"
        
        # Get the correct background color based on appearance mode
        if ctk.get_appearance_mode() == "Dark":
            bg_color = "#2b2b2b"
        else:
            bg_color = "#ffffff"
            
        self.configure(highlightthickness=0, bg=bg_color)
        self.bind("<Configure>", self._on_resize)
        
    def set_color(self, hex_color):
        self.color = hex_color
        self.delete("all")
        # Create a rounded rectangle with the specified color
        self.create_oval(2, 2, self.width-2, self.height-2, fill=hex_color, outline=hex_color)
        
    def _on_resize(self, event):
        self.width = event.width
        self.height = event.height
        self.delete("all")
        # Redraw the color preview when resized
        self.create_oval(2, 2, self.width-2, self.height-2, fill=self.color, outline=self.color)


def create_circle_image(hex_color, size=16):
    """Return a tkinter-compatible image with a filled circle of the specified color."""
    # Create a higher resolution image for better anti-aliasing
    scale = 4  # Scale factor for anti-aliasing
    img = Image.new("RGBA", (size * scale, size * scale), (255, 255, 255, 0))  # transparante achtergrond
    draw = ImageDraw.Draw(img)
    
    # Draw a larger circle with anti-aliasing
    padding = scale  # Add padding for smoother edges
    draw.ellipse((padding, padding, size * scale - padding, size * scale - padding), 
                 fill=hex_color)
    
    # Resize the image down to the desired size with anti-aliasing
    img = img.resize((size, size), Image.Resampling.LANCZOS)
    
    return ImageTk.PhotoImage(img)


def configure_treeview_style():
    """Configure the ttk.Treeview style based on the current appearance mode"""
    style = ttk.Style()
    
    # Set theme to 'clam' for better customization
    style.theme_use('clam')
    
    # Get colors based on current mode
    if ctk.get_appearance_mode() == "Dark":
        bg_color = "#2b2b2b"
        text_color = "white"
        selected_color = "#1f538d"
        heading_bg = "#2b2b2b"
    else:
        bg_color = "#ffffff"
        text_color = "black"
        selected_color = "#93c5fd"
        heading_bg = "#f0f0f0"

    # Configure the Treeview style
    style.configure("Treeview",
                   background=bg_color,
                   foreground=text_color,
                   fieldbackground=bg_color,
                   borderwidth=0,
                   font=("Roboto", 11))
    
    style.configure("Treeview.Heading",
                   background=heading_bg,
                   foreground=text_color,
                   borderwidth=1,
                   relief="flat",
                   font=("Roboto", 11, "bold"))
    
    style.map("Treeview",
             background=[("selected", selected_color)],
             foreground=[("selected", "white")])
    
    return style


def show_error(title, message):
    """Show error message dialog"""
    messagebox.showerror(title, message)


def show_info(title, message):
    """Show info message dialog"""
    messagebox.showinfo(title, message) 