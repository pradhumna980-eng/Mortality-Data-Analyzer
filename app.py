from flask import Flask, render_template, request
from actuarial import MortalityDataset, ActuarialAnalyzer, get_age_band

app = Flask(__name__)
import os
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "patient_mortality_dataset.xlsx")

@app.route("/")
def home():
    return render_template("index.html")

@app.route("/analyze", methods=["POST"])
def analyze():
    try:
        age = int(request.form["age"])
    except:
        return "Invalid age input."
    disease = request.form["disease"]

    data = MortalityDataset(EXCEL_FILE)
    if not data.records:
        return "No data loaded."

    analyzer = ActuarialAnalyzer(data)
    age_band = get_age_band(age)
    probs, n = analyzer.conditional_death_prob_3yr(age_band, disease)
    if n == 0:
        return "No matching records found."

    result_files = analyzer.export_results(age_band, disease)

    return render_template(
        "results.html",
        age_band=age_band,
        disease=disease,
        n=n,
        p1=probs[1],
        p2=probs[2],
        p3=probs[3],
        outfile=result_files["outfile"],
        cond_prob_file=result_files["conditional_probability"],
        seasonal_file=result_files["seasonal_mortality"],
        life_table_file=result_files["life_table"]
    )

if __name__ == "__main__":
    app.run(debug=True)
