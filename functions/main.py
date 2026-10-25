from flask import Flask, request, send_file
from flask_cors import CORS
import xlsxwriter
import io

app = Flask(__name__)
CORS(app)

@app.route('/export', methods=['POST'])
def export_excel():
    req_data = request.get_json() or {}
    rule_stats = req_data.get('ruleStats', [])
    trend_stats = req_data.get('trendStats', [])
    student_stats = req_data.get('studentStats', [])

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # Sheet 1: 違規項目統計
    ws_rules = workbook.add_worksheet('違規項目統計')
    ws_rules.write_row('A1', ['違規項目', '次數'])
    for i, item in enumerate(rule_stats, 2):
        ws_rules.write_row(f'A{i}', [item.get('name', '未分類'), item.get('count', 0)])
    
    if rule_stats:
        chart_rules = workbook.add_chart({'type': 'column'})
        chart_rules.add_series({
            'name': '違規次數',
            'categories': f'=違規項目統計!$A$2:$A${len(rule_stats)+1}',
            'values': f'=違規項目統計!$B$2:$B${len(rule_stats)+1}',
            'fill': {'color': '#4F46E5'}
        })
        ws_rules.insert_chart('D2', chart_rules)

    workbook.close()
    output.seek(0)

    return send_file(
        output,
        download_name="analytics_report.xlsx",
        as_attachment=True,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    app.run(port=8080)