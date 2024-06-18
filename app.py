from flask import Flask, render_template, request
from openpyxl import load_workbook

app = Flask(__name__)

# Dictionary to map categories to their respective cell values
category_cells = {
    "father": {
        "junk food": "B2",
        "veg": "C2",
        "non-veg": "D2",
        "medical": "E2",
        "T": "F2",
        "worker": "G2",
        "M": "H2",
        "A": "I2",
        # Add more categories for father here
    },
    "mother": {
        "junk food": "B3",
        "veg": "C3",
        "non-veg": "D3",
        "medical": "E3",
        "T": "F3",
        "worker": "G3",
        "M": "H3",
        "A": "I3",
        # Add more categories for mother here
    },
    "manju": {
        "junk food": "B4",
        "veg": "C4",
        "non-veg": "D4",
        "medical": "E4",
        "T": "F4",
        "worker": "G4",
        "M": "H4",
        "A": "I4",
        # Add more categories for manju here
    },
    "nancy": {
        "junk food": "B5",
        "veg": "C5",
        "non-veg": "D5",
        "medical": "E5",
        "T": "F5",
        "worker": "G5",
        "M": "H5",
        "A": "I5",
        # Add more categories for nancy here
    },
    "deepak": {
        "junk food": "B6",
        "veg": "C6",
        "non-veg": "D6",
        "medical": "E6",
        "T": "F6",
        "worker": "G6",
        "M": "H6",
        "A": "I6",
        # Add more categories for deepak here
    },
    "mayur": {
        "junk food": "B7",
        "veg": "C7",
        "non-veg": "D7",
        "medical": "E7",
        "T": "F7",
        "worker": "G7",
        "M": "H7",
        "A": "I7",
        # Add more categories for house here
    },
    "house": {
        "junk food": "B8",
        "veg": "C8",
        "non-veg": "D8",
        "medical": "E8",
        "T": "F8",
        "worker": "G8",
        "M": "H8",
        "A": "I8",
    }

    # Add more names here
}

# Store the calculator result
calc_result = None

@app.route('/')
def index():
    return render_template('index.html', result=calc_result)

@app.route('/update', methods=['POST'])
def update_data():
    name = request.form['name']
    category = request.form['category']
    amount = int(request.form['amount'])

    book = load_workbook("Book1.xlsx")
    sheet = book.active

    if name in category_cells and category in category_cells[name]:
        cell = category_cells[name][category]
        sheet[cell].value = int(sheet[cell].value) + amount
        book.save("Book1.xlsx")
        return render_template('index.html', success=True)
    else:
        return "INVALID OUTPUT"

@app.route('/calculate', methods=['POST'])
def calculate():
    global calc_result
    num1 = float(request.form['num1'])
    num2 = float(request.form['num2'])
    operator = request.form['operator']

    if operator == '+':
        calc_result = num1 + num2
    elif operator == '-':
        calc_result = num1 - num2
    elif operator == '*':
        calc_result = num1 * num2
    elif operator == '/':
        if num2 == 0:
            calc_result = 'Error: Division by zero!'
        else:
            calc_result = num1 / num2
    else:
        calc_result = 'Error: Invalid operator!'

    return render_template('index.html', result=calc_result)

if __name__ == '__main__':
    app.run(debug=True)
