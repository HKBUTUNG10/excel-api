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
    class_stats = req_data.get('classStats', [])
    student_stats = req_data.get('studentStats', [])

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

    # 📊 Sheet 1: 違規趨勢分析 (折線圖)
    ws_trend = workbook.add_worksheet('違規趨勢')
    ws_trend.write_row('A1', ['日期', '宗數'])
    for i, item in enumerate(trend_stats, 2):
        ws_trend.write_row(f'A{i}', [item.get('date', ''), item.get('count', 0)])

    if trend_stats:
        chart_trend = workbook.add_chart({'type': 'line'})
        chart_trend.add_series({
            'name': '違規宗數',
            'categories': f'=違規趨勢!$A$2:$A${len(trend_stats)+1}',
            'values': f'=違規趨勢!$B$2:$B${len(trend_stats)+1}',
            'line': {'color': '#10B981', 'width': 2.25}
        })
        chart_trend.set_title({'name': '每日違規趨勢'})
        ws_trend.insert_chart('D2', chart_trend)

    # 📊 Sheet 2: 各班違規次數 (柱狀圖)
    ws_class = workbook.add_worksheet('各班違規統計')
    ws_class.write_row('A1', ['班別', '違規次數'])
    for i, item in enumerate(class_stats, 2):
        ws_class.write_row(f'A{i}', [item.get('name', ''), item.get('count', 0)])

    if class_stats:
        chart_class = workbook.add_chart({'type': 'column'})
        chart_class.add_series({
            'name': '違規次數',
            'categories': f'=各班違規統計!$A$2:$A${len(class_stats)+1}',
            'values': f'=各班違規統計!$B$2:$B${len(class_stats)+1}',
            'fill': {'color': '#3B82F6'}
        })
        chart_class.set_title({'name': '各班違規次數 (最高至最低)'})
        ws_class.insert_chart('D2', chart_class)

    # 📊 Sheet 3: 欠帶事項分佈 (條形圖)
    ws_rules = workbook.add_worksheet('欠帶事項分佈')
    ws_rules.write_row('A1', ['違規項目', '次數'])
    for i, item in enumerate(rule_stats, 2):
        ws_rules.write_row(f'A{i}', [item.get('name', '未分類'), item.get('count', 0)])

    if rule_stats:
        chart_rules = workbook.add_chart({'type': 'bar'})
        chart_rules.add_series({
            'name': '違規次數',
            'categories': f'=欠帶事項分佈!$A$2:$A${len(rule_stats)+1}',
            'values': f'=欠帶事項分佈!$B$2:$B${len(rule_stats)+1}',
            'fill': {'color': '#6366F1'}
        })
        chart_rules.set_title({'name': '欠帶事項分佈'})
        ws_rules.insert_chart('D2', chart_rules)

    # 📊 Sheet 4: 高頻違規學生 (前10名)
    ws_students = workbook.add_worksheet('違規最高學生')
    ws_students.write_row('A1', ['排名', '班別', '姓名', '學號', '違規次數'])
    for i, item in enumerate(student_stats, 2):
        ws_students.write_row(f'A{i}', [
            i - 1,
            item.get('class', ''),
            item.get('name', ''),
            item.get('id', ''),
            item.get('count', 0)
        ])

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