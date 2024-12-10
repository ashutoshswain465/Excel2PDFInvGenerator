import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

# Glob the Excel files in the invoices directory
filepaths = glob.glob("invoices/*.xlsx")

# Loop through each file path in the list
for filepath in filepaths:
    # Read the first sheet of each Excel file as a dataframe
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Create a PDF object
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # Add a new page to the PDF
    pdf.add_page()

    # Extract the filename without the extension for use in the PDF
    filename = Path(filepath).stem
    # Split filename into invoice number and date
    invoice_nr, date = filename.split("-")

    # Add invoice number to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    # Add date to the PDF
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    # Process column headers for display
    columns = list(df.columns)
    # Replace underscores in column names and capitalize them
    columns = [item.replace("_", " ").title() for item in columns]
    # Set font for column headers
    pdf.set_font(family="Times", size=10)
    # Set gray text color for aesthetics
    pdf.set_text_color(80, 80, 80)
    # Create table headers in the PDF
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Loop through each row in the dataframe
    for index, row in df.iterrows():
        # Set font for row content
        pdf.set_font(family="Times", size=10)
        # Set text color
        pdf.set_text_color(80, 80, 80)
        # Add cells with data from the dataframe
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Calculate and display the total price
    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    # Display cells for layout consistency
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Display the total price in bold
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=f"The total price is {total_sum}", ln=1)

    # Add a footer with the company logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=25, h=8, txt=f"PythonHow")
    # Insert company logo
    pdf.image("pythonhow.png", w=10)

    # Save the PDF to a file in the specified directory
    pdf.output(f"PDFs/{filename}.pdf")
