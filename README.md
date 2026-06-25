# Mortality Data Analyzer

A Python actuarial tool that analyzes patient mortality data using life table 
methods and conditional probability models.

## Features
- **Life Table Construction** — computes lx, qx, dx by age band
- **3-Year Conditional Death Probability** — by age group and disease
- **Seasonal Mortality Analysis** — mortality rates by infection month
- **Excel Export** — results across 3 structured sheets
- **Flask Web Interface** — browser-based input and chart visualization

## Background
Built as an extension of actuarial coursework in the Bachelor of Mathematical 
Sciences program at TU IOST (School of Mathematical Sciences, Balkhu). 
Applies concepts from Actuarial Statistics I & II including life table 
construction and mortality modeling.

## Tech Stack
Python 3 · Flask · openpyxl · matplotlib

## Setup
```bash
pip install flask openpyxl matplotlib
python app.py        # Web interface at http://localhost:5000
python pythonproject.py  # CLI version
```

## Dataset
`patient_mortality_dataset.xlsx` — synthetic dataset for demonstration.  
Columns: PatientID, Age, Disease, InfectionMonth, Death (0/1), YearsUntilDeath

## Author
Pradhumna Shrestha — BMS Student, TU IOST  
Actuarial Science Specialization | Kathmandu, Nepal