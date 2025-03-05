import os
import pandas as pd
from flask import Flask, render_template_string, request

app = Flask(__name__)

# อนุญาตให้อัปโหลดเฉพาะไฟล์ Excel เท่านั้น
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_filename(filename, expected_prefix):
    # แทนที่ช่องว่างเป็น "_" เพื่อรองรับไฟล์ที่มีช่องว่าง
    formatted_filename = filename.replace(" ", "_")
    return formatted_filename.startswith(expected_prefix) and formatted_filename.endswith(('.xlsx', '.xls'))

# HTML Template
TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Kittipoom - Test</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-gray-100 flex items-center justify-center min-h-screen">
    <div class="bg-white shadow-lg rounded-lg p-8 w-full max-w-2xl">
        <h2 class="text-2xl font-bold text-center mb-6">Upload Daily Report and New Employee Files</h2>
        <form action="/" method="post" enctype="multipart/form-data" class="space-y-4">
            <div>
                <label class="block text-sm font-medium text-gray-700">Daily Reports (Multiple):</label>
                <input type="file" name="daily_report" multiple required class="mt-1 p-2 w-full border rounded-md">
            </div>
            
            <div>
                <label class="block text-sm font-medium text-gray-700">New Employee:</label>
                <input type="file" name="new_employee" required class="mt-1 p-2 w-full border rounded-md">
            </div>
            
            <button type="submit" class="w-full bg-blue-500 text-white py-2 rounded-md hover:bg-blue-600">Upload and Process</button>
        </form>
        
        {% if data %}
        <h2 class="text-xl font-semibold text-center mt-8">New Employees Under Each Team Member</h2>
        <div class="overflow-x-auto mt-4">
            <table class="min-w-full bg-white border border-gray-200 rounded-lg">
                <thead>
                    <tr class="bg-gray-200 text-gray-600 uppercase text-sm leading-normal">
                        <th class="py-3 px-6 text-left">Employee Name</th>
                        <th class="py-3 px-6 text-left">Join Date</th>
                        <th class="py-3 px-6 text-left">Role</th>
                        <th class="py-3 px-6 text-left">Team Member</th>
                    </tr>
                </thead>
                <tbody class="text-gray-600 text-sm font-light">
                    {% for row in data %}
                    <tr class="border-b border-gray-200 hover:bg-gray-100">
                        <td class="py-3 px-6 text-left">{{ row["Employee Name"] }}</td>
                        <td class="py-3 px-6 text-left">{{ row["Join Date"] }}</td>
                        <td class="py-3 px-6 text-left">{{ row["Role"] }}</td>
                        <td class="py-3 px-6 text-left">{{ row["Team Member"] }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        {% endif %}
    </div>
</body>
</html>
"""

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        daily_report_files = request.files.getlist('daily_report')
        new_employee_file = request.files['new_employee']
        
        # ตรวจสอบประเภทไฟล์ก่อนอ่าน
        if not allowed_file(new_employee_file.filename):
            return "Error: Only Excel files (.xlsx, .xls) are allowed!", 400
        
        for file in daily_report_files:
            if not allowed_file(file.filename):
                return "Error: Only Excel files (.xlsx, .xls) are allowed!", 400
        
        # ตรวจสอบชื่อไฟล์ก่อนอ่าน
        if not validate_filename(new_employee_file.filename, "New_Employee") and not validate_filename(new_employee_file.filename, "New Employee"):
            return "Error: New Employee file must be named 'New Employee_YYYYMM.xlsx'", 400

        for file in daily_report_files:
            if not validate_filename(file.filename, "Daily_report") and not validate_filename(file.filename, "Daily Report"):
                return "Error: Daily Report files must be named 'Daily report_YYYYMMDD_Name_Surname.xlsx'", 400
        
        # อ่านข้อมูลจากไฟล์ที่อัปโหลด
        df_new_employee = pd.read_excel(new_employee_file)
        
        all_reports = []
        for file in daily_report_files:
            df = pd.read_excel(file)
            filename_parts = file.filename.replace("Daily_report_", "").replace(".xlsx", "").split("_")
            team_member = f"{filename_parts[2]} {filename_parts[3]}"
            df["Team Member"] = team_member
            all_reports.append(df)
        
        # รวมข้อมูลจากไฟล์ Daily Reports ทั้งหมด
        df_daily = pd.concat(all_reports, ignore_index=True)
        
        # รวมข้อมูล
        merged_data = pd.merge(df_new_employee, df_daily, on="Role", how="left")
        final_data = merged_data[["Employee Name", "Join Date", "Role", "Team Member"]].dropna()
        
        return render_template_string(TEMPLATE, data=final_data.to_dict(orient='records'))
    
    return render_template_string(TEMPLATE, data=None)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))