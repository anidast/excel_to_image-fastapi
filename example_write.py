import requests
import json
import pandas as pd

data1 = pd.DataFrame({
    'Name': ['Alice', 'Bob', 'Charlie'],
    'Age': [25, 30, 35],
    'City': ['New York', 'Los Angeles', 'Chicago']
})

data2 = pd.DataFrame({
    'Product': ['Apple', 'Banana', 'Cherry'],
    'Price': [1.2, 0.8, 2.5],
    'Quantity': [10, 20, 15]
})

mytext = """\
# Title

Text **bold** and *italic*

* A first bullet
* A second bullet

# Another Title

This paragraph has a line break.
Another line.
"""

data = dict(
    df1 = data1.to_dict(orient='records'),
    df2 = data2.to_dict(orient='records'),
    text = mytext
)

api_url = "http://127.0.0.1:8000/write2excel/"
form_data = {
    'outputname': 'report.xlsx',
    'data': str(json.dumps(data))
    }
file = {'template_file': ('template.xlsx', open('template.xlsx', 'rb'))}

response = requests.post(api_url, files=file, data=form_data) 

if response.status_code == 200:
    try:
        result = response.json()

        if 'error' in result:
            print(f"Error from API: {result['error']}")
        else:
            excel_output = requests.get(result["excel_url"]).content
            pdf_output = requests.get(result["pdf_url"]).content
            with open('report.xlsx', 'wb') as handler:
                handler.write(excel_output)
            with open('report.pdf', 'wb') as handler:
                handler.write(pdf_output)
                
    except json.JSONDecodeError as e:
        print(f"Error decoding JSON response: {e}")
        print("Response content:", response.text)
else:
    print(f"Failed to get a response from the API. Status code: {response.status_code}")
    print("Response content:", response.text)