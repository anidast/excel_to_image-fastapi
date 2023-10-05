from typing import Annotated
from fastapi import FastAPI, Request, File, Form, UploadFile
from fastapi.staticfiles import StaticFiles
import excel2img
import os
import xlwings as xw
import pandas as pd
import json

appdesc = """
This API helps you work with excel without having to install it on your machine.
"""

tags_metadata = [
    {
        "name": "excel2img",
        "description": "This endpoint helps you to capture excel file into image file.",
    },
    {
        "name": "write2excel",
        "description": "This endpoint helps you to write data into excel file.",
    },
]

outputnamesdesc = """
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
    title="Excel API", 
    description=appdesc,
    openapi_tags=tags_metadata
    )

app.mount("/output", StaticFiles(directory="output"), name="output")

@app.post(
    "/excel2img/", 
    tags=["excel2img"],
    responses={
        200: {
            "description": respdesc,
        }
    }
)
def capture(
    file: Annotated[UploadFile, File(description="Excel file")],
    outputnames: Annotated[str, Form(description=outputnamesdesc)],
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



outputnamedesc = """
Excel filename(with format) as outputs. The file format must be the same as the template file format.

Example: `Report.xlsx` 
"""

respdesc2 = """
Return excel file url in string. 

Example: 

_"http://127.0.0.1:8000/output/report.xlsx"_ 
"""

@app.post(
    "/write2excel/", 
    tags=["write2excel"],
    responses={
        200: {
            "description": respdesc2,
        }
    }
)
def write(
    template_file: Annotated[UploadFile, File(description="Excel file template.")],
    outputname: Annotated[str, Form(description=outputnamedesc)],
    data: Annotated[str, Form(description="Dictionary data to be filled in the template.")],
    request: Request
):
    with open(os.path.join('input/', template_file.filename), "wb+") as file_object:
        file_object.write(template_file.file.read())

    data_dict = json.loads(data)
    dataframe = {}

    for key in data_dict:
        dataframe[key] = pd.DataFrame(data_dict[key])

    with xw.App(visible=False) as app:
        book = app.render_template(os.path.join('input/', template_file.filename),
                                os.path.join('output/', outputname),
                                **dataframe)
        return str(request.base_url) + 'output/' + outputname