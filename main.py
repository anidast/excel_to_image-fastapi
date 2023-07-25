import uvicorn
from typing import Annotated
from fastapi import FastAPI, Request, File, Form, UploadFile
from fastapi.staticfiles import StaticFiles
import excel2img
import os

app = FastAPI(
    servers=[
        {"url": "http://127.0.0.1:8000"},
    ],
    title="excel2img API", 
    description="This API helps you to capture excel file into image file."
    )

app.mount("/output", StaticFiles(directory="output"), name="output")

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

_outputnames, sheets, and cells must be the same length!_
"""

@app.post(
    "/excel2img/", 
    responses={
        200: {
            "description": "Return (list of) image url. Example: _http://127.0.0.1:8000/output/filename.png_",
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
    _, file_extension = os.path.splitext(file.filename)
    # file.save(os.path.join('files/', file.filename))
    # file.save('file_excel' + file_extension)
    with open('file_excel' + file_extension, "wb+") as file_object:
        file_object.write(file.file.read())

    result = []

    for outputname, sheet, cell in zip(outputnames.split(', '), sheets.split(', '), cells.split(', ')):
        excel2img.export_img('file_excel' + file_extension, 'output/' + outputname, sheet, cell)
        result.append(str(request.base_url) + 'output/' + outputname)
    
    return '\n'.join(result)
