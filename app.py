import os
import openpyxl
from flask import Flask, render_template, request, redirect

app = Flask(__name__)
app.secret_key = os.urandom(24)

def save_to_excel(data, sheet_name):
    # Load the existing workbook or create a new one
    try:
        wb = openpyxl.load_workbook('feedback_data.xlsx')
    except FileNotFoundError:
        wb = openpyxl.Workbook()

    # Select the active sheet (create a new one if it doesn't exist)
    if sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
    else:
        sheet = wb.create_sheet(title=sheet_name)

    # Insert the date in the first row
    date = data['Date']

    # Check if the sheet is empty or if the title row is missing
    if sheet.max_row == 0 or sheet.cell(row=1, column=1).value != 'Date':
        sheet.insert_rows(1)
        sheet.cell(row=1, column=1, value='Date')
        for idx, key in enumerate(data.keys(), start=2):
            sheet.cell(row=1, column=idx, value=key)

    # Find the next empty row
    next_row = sheet.max_row + 1

    # Insert date value
    sheet.cell(row=next_row, column=1, value=date)

    # Insert data values
    for idx, value in enumerate(data.values(), start=2):
        sheet.cell(row=next_row, column=idx, value=value)

    # Save the workbook
    wb.save('feedback_data.xlsx')


@app.route('/')
def index():
    return render_template('curriculum.html')

@app.route('/library-feedback', methods=['GET', 'POST'])
def library_feedback():
    if request.method == 'POST':
        # Process form data and store it
        return redirect('/success')  
    return render_template('library.html') 

@app.route('/submit-curriculum-feedback', methods=['POST'])
def submit_curriculum_feedback():
    curriculum_data = request.form
    save_to_excel(curriculum_data, 'Curriculum Feedback')
    return redirect('/library-feedback')

@app.route('/ambience-feedback', methods=['GET', 'POST'])
def ambience_feedback():
    if request.method == 'POST':
        return redirect('/success')  
    return render_template('ambience.html') 

@app.route('/submit-library-feedback', methods=['POST'])
def submit_library_feedback():
    library_data = request.form
    save_to_excel(library_data, 'Library Feedback')
    return redirect('/ambience-feedback')

@app.route('/submit-ambience-feedback', methods=['POST'])
def submit_ambience_feedback():
    ambience_data = request.form
    save_to_excel(ambience_data, 'Ambience Feedback')
    return redirect('/success')

@app.route('/success')
def success():
    return render_template('success.html')

if __name__ == '__main__':
    app.run(debug=True)
