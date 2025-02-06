from flask import Flask, jsonify, request, send_from_directory
import openpyxl
import os

app = Flask(__name__)

# The path to the Excel file
EXCEL_FILE = 'data.xlsx'

# Initialize the Excel file if it doesn't exist
if not os.path.exists(EXCEL_FILE):
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'Data'
    sheet.append(['Date', 'Place', 'Time', 'Kms'])
    wb.save(EXCEL_FILE)

@app.route('/')
def index():
    # Serve the whatsapp.html file from the static folder
    return send_from_directory('static', 'whatsapp.html')

@app.route('/save-to-excel', methods=['POST'])
def save_to_excel():
    data = request.get_json()
    date = data['date']
    place = data['place']
    time = data['time']
    kms = data['kms']

    wb = openpyxl.load_workbook(EXCEL_FILE)
    sheet = wb.active
    sheet.append([date, place, time, kms])
    wb.save(EXCEL_FILE)

    return jsonify({"message": "Data saved successfully!"})

if __name__ == '__main__':
    app.run(debug=True)
