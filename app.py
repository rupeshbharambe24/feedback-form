import os
from flask import Flask, render_template, request, redirect, url_for, session
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Set a secret key for session management

# Route to serve the HTML form
@app.route('/')
def index():
    if 'feedback_submitted' in session:
        return redirect(url_for('feedback_submitted'))
    return render_template('index.html')

# Route to handle form submission and store data in Excel sheet
@app.route('/submit-feedback', methods=['POST'])
def submit_feedback():
    if 'feedback_submitted' in session:
        return redirect(url_for('feedback_submitted'))
    
    name = request.form['name']
    roll_number = request.form['rollNumber']
    branch = request.form['branch']
    answers = [request.form[f'q{i}'] for i in range(1, 10)]

    # Load workbook and select active sheet
    try:
        wb = load_workbook('feedback.xlsx')
        ws = wb.active
    except FileNotFoundError:
        # If file doesn't exist, create a new workbook
        wb = Workbook()
        ws = wb.active
        ws.append(['Name', 'Roll Number', 'Branch', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9'])

    # Append data to the worksheet
    ws.append([name, roll_number, branch] + answers)

    # Save the workbook
    wb.save('feedback.xlsx')

    # Mark feedback as submitted in session to prevent resubmission from the same device
    session['feedback_submitted'] = True

    return redirect(url_for('feedback_submitted'))

# Route to display success message after feedback submission
@app.route('/feedback-submitted')
def feedback_submitted():
    return render_template('success.html')

if __name__ == '__main__':
    app.run(debug=True)
