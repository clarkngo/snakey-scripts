from datetime import datetime, timedelta
from docx import Document

def generate_date_range(start_date, weeks):
    date_format = "%m/%d"
    start_date = datetime.strptime(start_date, date_format)

    # Create a new Word document
    doc = Document()

    # Add a table to the document
    table = doc.add_table(rows=1, cols=1)
    table.style = 'Table Grid'
    table.autofit = False

    for week in range(weeks):
        end_date = start_date + timedelta(days=6)
        date_range = f"{start_date.strftime(date_format)} - {end_date.strftime(date_format)}"

        # Insert a new row in the table
        row = table.add_row().cells
        row[0].text = str(week + 1) + "\n" + date_range

        start_date = end_date + timedelta(days=1)

    # Save the document
    doc.save('output.docx')

# User input
start_date_input = input("Enter the starting date (MM/DD): ")
weeks_to_generate = 10

# Generate and save the Word document
generate_date_range(start_date_input, weeks_to_generate)
