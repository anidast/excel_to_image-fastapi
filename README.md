# Excel to Image - FastAPI

## 1. Create Virtual Environment

```powershell
python -m venv env_name
```

## 2. Activate Virtual Environment

```powershell
& path_to/env_name/Scripts/Activate.ps1
```

## 3. Install Library

```powershell
pip install -r requirements.txt
```

## 4. Create `input` and `output` directory


## 5. Run Server

Run the live server:

```powershell
uvicorn main:app --reload
```

The command uvicorn main:app refers to:

main: the file main.py (the Python "module").

app: the object created inside of main.py with the line app = FastAPI().

--reload: make the server restart after code changes. Only use for development.
The --reload option consumes much more resources, is more unstable, etc.
It helps a lot during development, but you shouldn't use it in production.


Or you can run the live server with customable host and/or port:
```
uvicorn main:app --host 0.0.0.0 --port 80
```

In the output, there's a line with something like:

INFO:     Uvicorn running on http://127.0.0.1:8000 (Press CTRL+C to quit)

That line shows the URL where your app is being served, in your local machine.

## Check Documentation 

Swagger API documentation can be found at {apiurl}/docs. 

Examples of how to use API with code can be seen in the files example_capture.py for capture image and example_write.py for write to excel.