from flask import Flask, render_template, request, send_from_directory, flash, redirect, url_for
from flask_cors import CORS
import io
from logic.autofill_api import autofill
import pandas as pd

app = Flask(__name__)
app.secret_key = "very secret key"
CORS(app)


@app.route("/autofill", methods=["POST"])
def autofill_route():
    data = request.json
    cookies = data["cookies"]
    current_url = data["current_url"]
    semesters = int(data["semesters"])
    autofill(cookies, current_url, semesters)
    return {"message": "Autofill complete. Now you can download the contract."}, 200


@app.route("/download_contract", methods=["GET"])
def download_contract():
    return send_from_directory(path="output_contract.docx", directory="../logic/data", as_attachment=True)


@app.route("/download_receipt", methods=["GET"])
def download_receipt():
    return send_from_directory(path="output_receipt.docx", directory="../logic/data", as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
