import decimal

from flask import Flask, jsonify, render_template
import openpyxl

from concrexit.legacy import ConcrexitLegacyAPIService

app = Flask(__name__)

USERNAME = "USERNAME"
PASSWORD = "PASSWORD"
CLIENT_ID = "dHhHjKsdRj7L7jeUPaenPZWXpdYGgUunm14HD09F"
CLIENT_SECRET = "OjRnPXBIN5sueeIWChHrV5WD9XAmIJUtHx1uH4korcOWmPXP89zLxNSmBZcjXDhqsIdD1x2KCd0es806E6nkq2kCSRxfy7Q17BsWmBnaOE2kb48qrfV8ztH1SHa2SOqz"
SHIFT_ID = 49

FILENAME = "auction.xlsx"
SHEET = "Sheet1"
CELL = "G2"


def get_total_sales_revenue():
    try:
        concrexit = ConcrexitLegacyAPIService(base_url="https://thalia.nu", username=USERNAME,password=PASSWORD, client_id=CLIENT_ID, client_secret=CLIENT_SECRET)
        response = concrexit.get(f"/api/v2/admin/sales/shifts/{SHIFT_ID}/")
    except Exception:
        return 0

    if response.status_code == 200:
        shift_data = response.json()
        return decimal.Decimal(shift_data["total_revenue"])
    else:
        return 0


def get_auction_revenue():
    try:
        workbook = openpyxl.load_workbook(FILENAME, data_only=True)
        worksheet = workbook[SHEET]
        cell = worksheet[CELL]
        return cell.value
    except Exception:
        return 0


@app.route("/total")
def display_value():
    value = get_total_sales_revenue() / 2 + get_auction_revenue()
    return jsonify({'value': value})


@app.route("/")
def index():
    return render_template('index.html')


if __name__ == '__main__':
    app.run()
