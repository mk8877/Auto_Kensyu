from flask import Flask
from views import spreadsheet_view, update_spreadsheet
app = Flask(__name__)

@app.route('/')
def home():
    return spreadsheet_view()

@app.route('/update', methods=['POST'])
def update():
    return update_spreadsheet()

if __name__ == '__main__':
    app.run(debug=True)
