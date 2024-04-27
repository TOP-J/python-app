from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib import colors
from app import table_frame




def pdfgeneration():
    

    for elem in table_frame.get_children():
       print(elem)

    doc = canvas.Canvas("table.pdf", pagesize=letter)

    # Create the table data as a nested list
    table_data = [
        ["Title", "CVL", "CP", "HW", "CA Mark", "Exam Mark", "Grade", "GP"]
        # Add your table row data here
    ]
    for elem in table_frame.get_children():
       li = []
       for each in elem.winfo_children():
          text = each.cget("text")
          if text == None or text == "":
             li.append(each.get())
          else:
             li.append(text)
       table_data.append(li)

    # Define the table style
    table_style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightblue),  # Header row background color
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),  # Header row text color
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),  # Center align all cells
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),  # Header row font
        ("FONTSIZE", (0, 0), (-1, 0), 12),  # Header row font size
        ("BOTTOMPADDING", (0, 0), (-1, 0), 12),  # Header row bottom padding
        ("BACKGROUND", (0, 1), (-1, -1), colors.beige),  # Data row background color
        ("GRID", (0, 0), (-1, -1), 1, colors.black),  # Table grid color and thickness
    ])

    # Create the table and apply the style
    table = Table(table_data)
    table.setStyle(table_style)

    # Draw the table on the PDF canvas
    table.wrapOn(doc, 0, 0)
    table.drawOn(doc, 50, 50)

    # Save and close the PDF document
    doc.save()