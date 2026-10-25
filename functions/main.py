from firebase_functions import https_fn
from firebase_admin import initialize_app
import xlsxwriter
import io

initialize_app()

@https_fn.on_request()
def export_analytics_excel(req: https_fn.Request) -> https_fn.Response:
    # 處理跨域請求 (CORS)
    if req.method == 'OPTIONS':
        headers = {
            'Access-Control-Allow-Origin': '*',
            'Access-Control-Allow-Methods': 'POST',
            'Access-Control-Allow-Headers': 'Content-Type'
        }
        return https_fn.Response('', status=204, headers=headers)

    # 取得前端傳來的統計數據
    req_data = req.get_json() or {}
    rule_stats = req_data.get('ruleStats', [])
    trend_stats = req_data.get('trendStats', [])
    student_stats = req_data.get('studentStats', [])

    # 建立記憶體中的 Excel 檔案
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)

    # ==========================================
    # Sheet 1: 違規項目統計 (含柱狀圖)
    # ==========================================
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
        chart_rules.set_title({'name': '欠帶事項分佈'})
        ws_rules.insert_chart('D2', chart_rules)

    # ==========================================
    # Sheet 2: 日期趨勢分析 (含折線圖)
    # ==========================================
    ws_trend = workbook.add_worksheet('日期趨勢分析')
    ws_trend.write_row('A1', ['日期', '人次'])
    
    for i, item in enumerate(trend_stats, 2):
        ws_trend.write_row(f'A{i}', [item.get('date', '未知'), item.get('count', 0)])

    if trend_stats:
        chart_trend = workbook.add_chart({'type': 'line'})
        chart_trend.add_series({
            'name': '當日違規人次',
            'categories': f'=日期趨勢分析!$A$2:$A${len(trend_stats)+1}',
            'values': f'=日期趨勢分析!$B$2:$B${len(trend_stats)+1}',
            'line': {'color': '#059669', 'width': 2.25}
        })
        chart_trend.set_title({'name': '每日違規人次趨勢'})
        ws_trend.insert_chart('D2', chart_trend)

    # ==========================================
    # Sheet 3: 班級與學生排行
    # ==========================================
    ws_students = workbook.add_worksheet('班級與學生排行')
    ws_students.write_row('A1', ['排名', '班別', '姓名', '學號', '違規總次數'])
    
    for i, item in enumerate(student_stats, 2):
        ws_students.write_row(f'A{i}', [i-1, item.get('class', ''), item.get('name', ''), item.get('id', ''), item.get('count', 0)])

    # 封裝並回傳二進位檔案
    workbook.close()
    output.seek(0)

    headers = {
        'Access-Control-Allow-Origin': '*',
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': 'attachment; filename="analytics_report.xlsx"'
    }
    return https_fn.Response(output.read(), headers=headers)