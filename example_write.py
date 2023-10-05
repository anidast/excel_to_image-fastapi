import requests
import json
import pandas as pd

data1 = pd.DataFrame()
data2 = pd.DataFrame()

data = dict(
    df1 = data1.to_dict(),
    df2 = data2.to_dict()
)

api_url = "http://127.0.0.1:8000/write2excel/"
form_data = {
    'outputname': 'report.xlsx',
    'data': str(data)
    }
file = {'template_file': ('template.xlsx', open('template.xlsx', 'rb'))}

response = requests.post(api_url, files=file, data=form_data) 
result = json.loads(response.content)

file_output = requests.get(result).content
with open('report.xlsx', 'wb') as handler:
    handler.write(file_output)
