from flask import Flask, render_template
from openpyxl import load_workbook

app = Flask(__name__)


@app.route('/')
def homepage():
    excel = load_workbook('my_report.xlsx')
    sheet = excel['Sheet1']
    column = sheet['A']
    return render_template('index.html', goods=column)


if __name__ == "__main__":
    app.run(debug=True)
