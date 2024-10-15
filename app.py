from flask import Flask, request, send_file
import os
import json
from docx import Document  # Ensure correct import from python-docx

app = Flask(__name__)

# Endpoint to handle document updating
@app.route('/update-document', methods=['POST'])
def update_document():
    # Ensure the uploads directory exists
    if not os.path.exists('uploads'):
        os.makedirs('uploads')

    # Get the JSON data from the form-data (it comes as a string, so we need to parse it)
    json_data = request.form['json']
    data = json.loads(json_data)  # Parse the JSON string

    # Get the uploaded document from the client
    uploaded_file = request.files['document']
    document_path = os.path.join('uploads', uploaded_file.filename)

    # Save the uploaded file locally
    uploaded_file.save(document_path)

    # Open and modify the Word document using python-docx
    doc = Document(document_path)

    entries = data.get('entries', [])

    # Assuming the document has a table
    for table in doc.tables:
        for entry in entries:
            found_row = False
            # Iterate over the rows and check for existing 'No.'
            for current_row_index in range(1, len(table.rows)):  # Skip header row
                row = table.rows[current_row_index]
                if row.cells[0].text.strip() == entry['No.']:  # Check 'No.' match
                    found_row = True
                    break

            if found_row:
                # 'No.' already exists, skip to next entry
                continue

            # If no match found, find the first empty row or add a new one
            for current_row_index in range(1, len(table.rows)):
                row = table.rows[current_row_index]
                if not row.cells[0].text.strip():  # Find empty row
                    break
            else:
                # If no empty row is found, add a new row
                last_row = table.rows[-1]
                row = table.add_row()
                for i, cell in enumerate(last_row.cells):
                    row.cells[i].text = cell.text  # Copy structure from the last row

            # Update the row with data from the JSON
            row.cells[0].text = entry['No.']  # No.
            row.cells[1].text = entry['Drawing Number']  # Drawing Number
            row.cells[2].text = entry['Drawing Title']  # Drawing Title
            row.cells[3].text = entry['Revision Number']  # Revision Number
            row.cells[4].text = entry['Date of Issue']  # Date of Issue
            row.cells[5].text = entry['Prepared By']  # Prepared By
            row.cells[6].text = entry['Approved By']  # Approved By
            row.cells[7].text = entry['Client Approval Status']  # Client Approval Status
            row.cells[8].text = entry['File Location/Reference']  # File Location/Reference
            row.cells[9].text = entry['Remarks']  # Remarks

    # Save the updated document
    updated_path = os.path.join('uploads', 'updated_' + uploaded_file.filename)
    doc.save(updated_path)

    # Send back the updated document
    return send_file(updated_path, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True)
