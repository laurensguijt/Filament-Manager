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


def generate_filament_label(filament_data, include_qr=True, include_barcode=True):
    """Generate a PDF label for the filament"""
    try:
        # Create a temporary file
        temp_dir = tempfile.gettempdir()
        
        # Get the filament code
        filament_code = filament_data.code
        # Remove any non-alphanumeric characters
        filament_code = ''.join(c for c in filament_code if c.isalnum())
        
        label_path = os.path.join(temp_dir, f"filament_label_{filament_code}.pdf")
        
        # Generate QR code if requested
        qr_path = None
        if include_qr:
            qr = qrcode.QRCode(version=1, box_size=10, border=0)
            qr.add_data(filament_code)
            qr.make(fit=True)
            qr_img = qr.make_image(fill_color="black", back_color="white")
            qr_path = os.path.join(temp_dir, "temp_qr.png")
            qr_img.save(qr_path)
        
        # Generate barcode if requested
        barcode_path = None
        if include_barcode:
            barcode_path = os.path.join(temp_dir, "temp_barcode")
            code128 = Code128(filament_code, writer=ImageWriter())
            code128.save(barcode_path)
        
        # Create the PDF with a larger size to match the image
        c = canvas.Canvas(label_path, pagesize=(100*mm, 60*mm))
        
        # Define margins
        left_margin = 6*mm
        right_margin = 94*mm
        top_margin = 54*mm
        bottom_margin = 6*mm
        
        # Draw background with rounded corners
        c.setFillColor(colors.white)
        c.roundRect(2*mm, 2*mm, 96*mm, 56*mm, 5*mm, fill=1, stroke=1)
        
        # Draw color preview box with border (smaller and in top right)
        c.setFillColor(filament_data.hex_color)
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.5)
        c.roundRect(75*mm, 40*mm, 15*mm, 15*mm, 2*mm, fill=1, stroke=1)

        # Define standard element sizes
        qr_size = 20*mm
        barcode_width = 25*mm  # Reduced to 25mm width
        barcode_height = qr_size  # Same height as QR code
        
        # Calculate text positions and widths based on which codes are included
        if include_qr and include_barcode:
            # Both codes included
            text_x = 35*mm
            max_text_width = right_margin - text_x - 5*mm
            code_y = 47*mm
            material_y = 40*mm
            supplier_y = 33*mm
            date_y = 26*mm
            weight_y = 19*mm
            desc_y = 12*mm
        elif include_qr:
            # Only QR code
            text_x = 35*mm
            max_text_width = right_margin - text_x - 5*mm
            code_y = 47*mm
            material_y = 40*mm
            supplier_y = 33*mm
            date_y = 26*mm
            weight_y = 19*mm
            desc_y = 12*mm
        elif include_barcode:
            # Only barcode
            text_x = 35*mm  # Keep text position same as with QR code
            max_text_width = right_margin - text_x - 5*mm
            code_y = 47*mm
            material_y = 40*mm
            supplier_y = 33*mm
            date_y = 26*mm
            weight_y = 19*mm
            desc_y = 12*mm
        else:
            # No codes included
            text_x = left_margin
            max_text_width = right_margin - text_x - 5*mm
            code_y = 50*mm
            material_y = 43*mm
            supplier_y = 36*mm
            date_y = 29*mm
            weight_y = 22*mm
            desc_y = 15*mm
        
        # Add QR code in top left if requested
        if include_qr and qr_path:
            qr_x = left_margin + 2*mm
            qr_y = 35*mm
            c.drawImage(qr_path, qr_x, qr_y, qr_size, qr_size)
        
        # Add barcode with consistent positioning
        if include_barcode and barcode_path:
            barcode_x = left_margin + 2*mm
            if include_qr:
                barcode_y = 8*mm  # Position below QR code
            else:
                barcode_y = 35*mm  # Same vertical position as where QR would be
            c.drawImage(barcode_path + ".png", barcode_x, barcode_y, barcode_width, barcode_height)
        
        # Add text with better formatting and larger font
        c.setFillColor(colors.black)
        
        # Filament code - now just the code without any black square
        c.setFont("Helvetica-Bold", 14)
        c.drawString(text_x, code_y, f"Code: {filament_code}")
        
        # Material and variant - using actual font size from image
        material_variant = f"{filament_data.material} {filament_data.variant}"
        c.setFont("Helvetica", 14)
        c.drawString(text_x, material_y, material_variant)
        
        # Supplier
        c.drawString(text_x, supplier_y, f"Supplier: {filament_data.supplier}")
        
        # Date opened
        date_str = filament_data.date_opened.strftime("%Y-%m-%d") if isinstance(filament_data.date_opened, datetime) else str(filament_data.date_opened).split(' ')[0]
        c.drawString(text_x, date_y, f"Opened: {date_str}")
        
        # Empty spool weight
        empty_spool_weight_formatted = "{:,.0f}".format(filament_data.empty_spool_weight)
        c.drawString(text_x, weight_y, f"Empty Spool: {empty_spool_weight_formatted}g")
        
        # Add description if present with text wrapping
        if filament_data.description:
            c.setFont("Helvetica", 12)  # Slightly smaller font for description
            # Split description into words
            words = filament_data.description.split()
            lines = []
            current_line = []
            
            for word in words:
                # Test if adding the word would exceed the width
                test_line = ' '.join(current_line + [word])
                if c.stringWidth(f"Note: {test_line}" if not lines else test_line, "Helvetica", 12) <= max_text_width:
                    current_line.append(word)
                else:
                    if current_line:
                        lines.append(' '.join(current_line))
                    current_line = [word]
            
            if current_line:
                lines.append(' '.join(current_line))
            
            # Draw each line of the description
            for i, line in enumerate(lines):
                # Check if there's enough vertical space
                current_y = desc_y - (i * 4*mm)
                if current_y > bottom_margin:  # Only draw if above bottom margin
                    c.drawString(text_x, current_y, f"Note: {line}" if i == 0 else line)
        
        # Add decorative elements - thinner line, moved down
        c.setStrokeColor(colors.black)
        c.setLineWidth(0.25)
        c.line(left_margin, bottom_margin - 2*mm, right_margin, bottom_margin - 2*mm)
        
        # Save the PDF
        c.save()
        
        # Clean up temporary files
        if qr_path and os.path.exists(qr_path):
            os.remove(qr_path)
        if barcode_path and os.path.exists(barcode_path + ".png"):
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