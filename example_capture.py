import requests
import json

api_url = "http://127.0.0.1:8000/excel2img/"
form_data = {
    'outputnames': 'dashboard.png',
    'sheets': 'Dashboard',
    'cells': 'A1:R29',
    }
file = {'file': ('example.xlsx', open('example.xlsx', 'rb'))}

response = requests.post(api_url, files=file, data=form_data) 
result = json.loads(response.content)

image_output = requests.get(result).content
with open('image_name.png', 'wb') as handler:
    handler.write(image_output)