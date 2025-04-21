import os
import tempfile
import subprocess
from datetime import datetime
import qrcode
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from tkinter import messagebox

from Filament_Manager.data_operations import read_excel_data


def draw_color_circle(canvas_obj, x, y, diameter, hex_color):
    """Draw a filled circle with the specified color and a black border"""
    try:
        # Ensure valid hex color
        if not hex_color:
            hex_color = '#000000'
        
        if not hex_color.startswith('#'):
            hex_color = f"#{hex_color}"
            
        # Remove '#' for processing
        color_value = hex_color[1:]
        
        # Convert short form to long form
        if len(color_value) == 3:
            color_value = ''.join(c + c for c in color_value)
            
        # Validate hex color
        if len(color_value) != 6 or not all(c in '0123456789ABCDEFabcdef' for c in color_value):
            print(f"Invalid hex color: {hex_color}, defaulting to black")
            color_value = '000000'
            
        # Save canvas state
        canvas_obj.saveState()
        
        # Draw filled circle
        canvas_obj.setFillColor(colors.HexColor(f"#{color_value}"))
        canvas_obj.setStrokeColor(colors.black)
        canvas_obj.setLineWidth(0.5)
        
        # Calculate circle path
        path = canvas_obj.beginPath()
        path.circle(x + diameter/2, y + diameter/2, diameter/2)
        canvas_obj.drawPath(path, fill=1, stroke=1)
        
        # Restore canvas state
        canvas_obj.restoreState()
        
    except Exception as e:
        print(f"Error drawing color circle: {str(e)}")
        # If any error occurs, draw a black circle
        canvas_obj.saveState()
        canvas_obj.setFillColor(colors.black)
        canvas_obj.setStrokeColor(colors.black)
        path = canvas_obj.beginPath()
        path.circle(x + diameter/2, y + diameter/2, diameter/2)
        canvas_obj.drawPath(path, fill=1, stroke=1)
        canvas_obj.restoreState()


def generate_filament_label(filament_data):
    """Generate a PDF label for the filament"""
    try:
        # Create a temporary file
        temp_dir = tempfile.gettempdir()
        
        # Get the filament code
        filament_code = filament_data.code
        # Remove any non-alphanumeric characters
        filament_code = ''.join(c for c in filament_code if c.isalnum())
        
        label_path = os.path.join(temp_dir, f"filament_label_{filament_code}.pdf")
        
        # Generate QR code - smaller size for better positioning
        qr = qrcode.QRCode(version=1, box_size=10, border=0)
        qr.add_data(filament_code)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        qr_path = os.path.join(temp_dir, "temp_qr.png")
        qr_img.save(qr_path)
        
        # Generate barcode
        barcode_path = os.path.join(temp_dir, "temp_barcode")
        code128 = Code128(filament_code, writer=ImageWriter())
        code128.save(barcode_path)
        
        # Create the PDF with a larger size to match the image
        c = canvas.Canvas(label_path, pagesize=(100*mm, 60*mm))
        
        # Draw background with rounded corners
        c.setFillColor(colors.white)
        c.roundRect(2*mm, 2*mm, 96*mm, 56*mm, 5*mm, fill=1, stroke=1)
        
        # Draw color preview box with border (smaller and in top right)
        c.setFillColor(filament_data.hex_color)
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.5)
        c.roundRect(75*mm, 40*mm, 15*mm, 15*mm, 2*mm, fill=1, stroke=1)
        
        # Add QR code in top left
        qr_size = 20*mm
        qr_x = 8*mm
        qr_y = 35*mm
        c.drawImage(qr_path, qr_x, qr_y, qr_size, qr_size)
        
        # Add barcode under QR code with increased size
        barcode_width = 35*mm  # Increased width from 30mm to 35mm
        barcode_height = 30*mm  # Increased height from 25mm to 30mm
        barcode_x = qr_x - 7.5*mm  # Adjusted x position to center the wider barcode
        barcode_y = 5*mm  # Lowered position to prevent overlap with QR code
        c.drawImage(barcode_path + ".png", barcode_x, barcode_y, barcode_width, barcode_height, preserveAspectRatio=True)
        
        # Add text with better formatting and larger font
        c.setFillColor(colors.black)
        
        # Filament code - now just the code without any black square
        c.setFont("Helvetica-Bold", 14)
        c.drawString(35*mm, 47*mm, f"Code: {filament_code}")
        
        # Material and variant - using actual font size from image
        material_variant = f"{filament_data.material} {filament_data.variant}"
        c.setFont("Helvetica", 14)
        c.drawString(35*mm, 40*mm, material_variant)
        
        # Supplier
        c.drawString(35*mm, 33*mm, f"Supplier: {filament_data.supplier}")
        
        # Date opened
        date_str = filament_data.date_opened.strftime("%Y-%m-%d") if isinstance(filament_data.date_opened, datetime) else str(filament_data.date_opened).split(' ')[0]
        c.drawString(35*mm, 26*mm, f"Opened: {date_str}")
        
        # Empty spool weight
        empty_spool_weight_formatted = "{:,.0f}".format(filament_data.empty_spool_weight)
        c.drawString(35*mm, 19*mm, f"Empty Spool: {empty_spool_weight_formatted}g")
        
        # Add description if present
        if filament_data.description:
            c.drawString(35*mm, 12*mm, f"Note: {filament_data.description}")
        
        # Add decorative elements - thinner line, moved down
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.25)
        c.line(8*mm, 4*mm, 92*mm, 4*mm)  # Line moved down further
        
        # Save the PDF
        c.save()
        
        # Clean up temporary files
        os.remove(qr_path)
        os.remove(barcode_path + ".png")
        
        # Open the PDF with the default viewer
        if os.name == 'nt':  # Windows
            os.startfile(label_path)
        else:  # macOS and Linux
            subprocess.call(('open', label_path))
        
        return True
            
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate label:\n{str(e)}")
        return False


def generate_inventory_report():
    """Generate a PDF report of the filament inventory"""
    try:
        # Get current timestamp
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        
        # Create a temporary file
        temp_dir = tempfile.gettempdir()
        pdf_file = os.path.join(temp_dir, f"filament_inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        
        # Create the PDF with A4 size
        c = canvas.Canvas(pdf_file, pagesize=A4)
        
        # Calculate page width and margins
        page_width = A4[0]
        page_height = A4[1]
        margin = 40
        content_width = page_width - 2 * margin
        
        # Set font
        c.setFont("Helvetica-Bold", 24)
        
        # Add title
        c.drawString(margin, page_height - margin, "Filament Inventory Overview")
        
        # Add generation timestamp
        c.setFont("Helvetica", 12)
        c.drawString(margin, page_height - margin - 30, f"Generated on: {timestamp}")
        
        # Add summary section
        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin, page_height - margin - 70, "Summary")
        
        # Get data for summary
        data = read_excel_data()
        total_spools = len(data)
        total_weight = sum(filament.weight for filament in data)
        
        # Add summary details
        c.setFont("Helvetica", 12)
        c.drawString(margin + 20, page_height - margin - 90, f"Total number of spools:")
        c.drawString(margin + 200, page_height - margin - 90, str(total_spools))
        c.drawString(margin + 20, page_height - margin - 110, f"Total available weight:")
        c.drawString(margin + 200, page_height - margin - 110, f"{total_weight:,.0f}g")
        
        # Add filament overview section
        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin, page_height - margin - 150, "Filament Overview")
        
        # Create table headers and calculate column widths
        headers = ["Code", "Material", "Variant", "Supplier", "Opened", "Available", "Color"]
        col_widths = [60, 80, 80, 80, 70, 70, 50]  # Adjusted widths
        x_positions = [margin]
        for width in col_widths[:-1]:
            x_positions.append(x_positions[-1] + width)
        
        # Function to center text in column
        def center_text(text, x_pos, width):
            text_width = c.stringWidth(str(text), c._fontname, c._fontsize)
            return x_pos + (width - text_width) / 2
        
        # Set starting y position for table
        y = page_height - margin - 180
        
        # Draw header background
        c.setFillColor(colors.HexColor('#1e3d7b'))  # Dark blue color
        c.rect(margin - 5, y - 15, content_width + 10, 20, fill=True)
        
        # Draw headers
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 10)
        for i, header in enumerate(headers):
            x = center_text(header, x_positions[i], col_widths[i])
            c.drawString(x, y-7, header)
        
        # Reset fill color for data
        c.setFillColor(colors.black)
        
        # Add data rows
        c.setFont("Helvetica", 10)
        y -= 30
        
        row_height = 25  # Define consistent row height
        circle_diameter = 15  # Define circle diameter
        
        for filament in data:
            # Check if we need a new page
            if y < margin + 50:
                c.showPage()
                c.setFont("Helvetica", 10)
                y = page_height - margin - 30
                
                # Redraw headers on new page
                c.setFillColor(colors.HexColor('#1e3d7b'))
                c.rect(margin - 5, y - 15, content_width + 10, 20, fill=True)
                c.setFillColor(colors.white)
                c.setFont("Helvetica-Bold", 10)
                for i, header in enumerate(headers):
                    x = center_text(header, x_positions[i], col_widths[i])
                    c.drawString(x, y, header)
                c.setFont("Helvetica", 10)
                y -= 30
            
            # Format date
            date_str = filament.date_opened.strftime("%Y-%m-%d") if isinstance(filament.date_opened, datetime) else str(filament.date_opened).split(' ')[0]
            
            # Draw row data
            row_data = [
                filament.code,
                filament.material,
                filament.variant,
                filament.supplier,
                date_str,
                f"{filament.weight:,.0f}g"
            ]
            
            # Center each value in its column
            c.setFillColor(colors.black)
            for i, value in enumerate(row_data):
                x = center_text(value, x_positions[i], col_widths[i])
                c.drawString(x, y, str(value))
            
            # Calculate color circle position
            color_x = x_positions[-1] + (col_widths[-1] - circle_diameter) / 2
            color_y = y - circle_diameter/2 + 5  # Adjusted to center vertically with text
            
            # Draw color circle with border
            draw_color_circle(c, color_x, color_y, circle_diameter, filament.hex_color)
            
            # Draw light gray line under each row
            c.setStrokeColor(colors.lightgrey)
            c.line(margin - 5, y - 12, margin + content_width + 5, y - 12)
            
            y -= row_height
        
        # Add footer
        c.setFont("Helvetica-Oblique", 8)
        c.drawString(margin, 30, "This report was automatically generated by Filament Manager Pro")
        
        # Save the PDF
        c.save()
        
        # Open the PDF with the default viewer
        if os.name == 'nt':  # Windows
            os.startfile(pdf_file)
        else:  # macOS and Linux
            subprocess.call(('open', pdf_file))
            
        return True
        
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate PDF report:\n{str(e)}")
        return False 