from flask import Flask
from views import spreadsheet_view
app = Flask(__name__)

@app.route('/')
def home():
    return spreadsheet_view()

if __name__ == '__main__':
    app.run(debug=True)
