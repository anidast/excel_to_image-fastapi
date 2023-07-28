import uvicorn
from typing import Annotated
from fastapi import FastAPI, Request, File, Form, UploadFile
from fastapi.staticfiles import StaticFiles
import excel2img
import os

appdesc = """
This API helps you to capture excel file into image file.

Example usage in python script:
```
import requests
import json

api_url = "http://127.0.0.1:8000/excel2img/"
form_data = {
    'outputnames': 'dashboard.png',
    'sheets': 'Dashboard',
    'cells': 'A1:R29',
    }
file = {'file': ('example.xlsx', open('D:\\Documents\\example.xlsx', 'rb'))}

response = requests.post(api_url, files=file, data=form_data) 
result = json.loads(response.content)
```

`result`(response) only contains the image url, so you still need to get the image file. You can get image file with this:

```
image_output = requests.get(result).content
with open('image_name.png', 'wb') as handler:
    handler.write(image_output)
```

Note: To consume api you need the `requests` library in your environment, if you don't have it try `pip install requests`.
"""

outputnamedesc = """
Images filename(with format) as outputs. Image format allowed:JPEG/PNG. 
It can contain a single filename or multiple filenames. 
For multiple filenames use separator \", \"(comma+space) for each filename.

Example: `filename.png` or `filename1.png, filename2.png, filename3.png`
"""
sheetdesc = """
The worksheet title in the excel file to be captured. It can contain a single sheet or multiple sheet. 
For multiple sheet use separator \", \"(comma+space) for each sheet

Example: `Sheet1` or `Sheet1, Sheet2, Sheet3`
"""
celldesc = """
Range of cells to be captured. It can contain a single range or multiple ranges. 
For multiple ranges use separator \", \"(comma+space) for each cell

Example: `A1:E5` or `A1:E5, G1:T25, B3:M40`

__outputnames, sheets, and cells must be the same length!__
"""

respdesc = """
Return image url string (single) or list of image url (multiple). 

Example: 

_"http://127.0.0.1:8000/output/filename.png"_ (single)

_["http://127.0.0.1:8000/output/filename1.png", "http://127.0.0.1:8000/output/filename2.png"]_ (multiple)
"""

app = FastAPI(
    title="excel2img API", 
    description=appdesc
    )

app.mount("/output", StaticFiles(directory="output"), name="output")

@app.post(
    "/excel2img/", 
    responses={
        200: {
            "description": respdesc,
        }
    }
)
def capture(
    file: Annotated[UploadFile, File(description="Excel file")],
    outputnames: Annotated[str, Form(description=outputnamedesc)],
    sheets: Annotated[str, Form(description=sheetdesc)],
    cells: Annotated[str, Form(description=celldesc)],
    request: Request
):
    with open(os.path.join('input/', file.filename), "wb+") as file_object:
        file_object.write(file.file.read())

    result = []

    for outputname, sheet, cell in zip(outputnames.split(', '), sheets.split(', '), cells.split(', ')):
        excel2img.export_img(os.path.join('input/', file.filename), 'output/' + outputname, sheet, cell)
        result.append(str(request.base_url) + 'output/' + outputname)
    
    if (len(result) == 1):
        return result[0]
    else:
        return result
