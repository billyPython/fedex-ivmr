from flask import Flask

app = Flask(__name__)

@app.route('/report_gen')
def generate_ivmr_report():
    return 'IVMR.csv report will be generated here and returned in xls/xlsx'