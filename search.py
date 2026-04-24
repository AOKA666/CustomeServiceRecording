import requests
import json
import openpyxl
import sys

auth = open('C:/Users/18441/Desktop/客服记录台/Cookie.txt').read()

headers = {
    'sec-ch-ua-platform': '"Windows"',
    'x_gduserid': '97196',
    'Referer': 'https://next.gaodun.com/',
    'sec-ch-ua': '"Google Chrome";v="147", "Not.A/Brand";v="8", "Chromium";v="147"',
    'Authentication': auth,
    'sec-ch-ua-mobile': '?0',
    'X-Requested-Extend': '{"systemName":"SYSTEM-OCRM"}',
    'x_gdsid': 'no-dbU1pRhAxQ2SdP_Bvp_hACCokwMDtmtLFXI_6',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36',
    'Accept': 'application/json',
    'Content-Type': 'application/json; charset=UTF-8',
}

json_data = {
    'pageNum': 1,
    'pageSize': 10,
    'clueNo': '',
    'customerName': '',
    'customerNo': '',
    'mobile': sys.argv[1],
    'otherContact': '',
    'overseaMobile': '',
    'searchColumn': 'mobile',
}

response = requests.post('https://apigateway.gaodun.com/solon/api/v1/clues/global-quick-search', headers=headers, json=json_data)
data = json.loads(response.text)

# 老师
owner_name = data["result"]["list"][0]["ownerName"]
real_name = owner_name.split("-")[-1]

# 项目
intent_project = data["result"]["list"][0]["intentProjectName"].strip()

# 如果 real_name 为空 或者 等于“新海”，走Excel查找
if not real_name or real_name == "新海":
    try:
        # 打开同文件夹下的 紧急联系人.xlsx
        wb = openpyxl.load_workbook("紧急联系人.xlsx")
        ws = wb.active

        # 遍历A列，查找 intent_project
        emergency_name = None
        for row in range(1, ws.max_row + 1):
            a_value = ws.cell(row=row, column=1).value
            if a_value and str(a_value).strip() == intent_project:
                emergency_name = ws.cell(row=row, column=2).value
                break

        # 赋值并输出
        if emergency_name:
            real_name = str(emergency_name).strip()
            print(f"线索老师不存在!")
        else:
            print(f"未找到【{intent_project}】对应的紧急联系人")
            
    except Exception as e:
        print(f"读取Excel失败：{str(e)}")

print(real_name)