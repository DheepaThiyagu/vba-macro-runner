from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
import os
import uuid
import pythoncom
import win32com.client as win32
import shutil

app = FastAPI(title="VBA Macro Runner")

EXCEL_TEMP_DIR = r"C:\excel-runner\work"
os.makedirs(EXCEL_TEMP_DIR, exist_ok=True)

def run_macro(file_path: str, macro_name: str):
    """
    Run a VBA macro in the given Excel file and return the C2 value.
    Ensures COM is initialized in this thread.
    """
    # Normalize Windows path
    file_path = os.path.normpath(file_path)

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")

        wb = excel.Workbooks.Open(file_path)

        # Run macro
        excel.Application.Run(f"'{wb.Name}'!{macro_name}")

        # Read debug value
        sheet = wb.Sheets(1)
        c2_value = sheet.Range("C2").Value

        # Save & close
        wb.SaveAs(file_path)
        wb.Close(SaveChanges=True)

        return c2_value

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()


@app.post("/run-macro-upload/")
async def run_macro_upload(
        file: UploadFile = File(...),
        macro_name: str = Form(...)
):
    """
    Upload an XLSM file, run the macro, and return the processed file.
    """
    # Generate temp file path
    temp_filename = f"temp_{uuid.uuid4().hex}.xlsm"
    temp_file_path = os.path.join(EXCEL_TEMP_DIR, temp_filename)

    # Save uploaded file to Windows temp folder
    with open(temp_file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)

    try:
        # Run macro
        c2_value = run_macro(temp_file_path, macro_name)

        # Return processed file and debug value
        return FileResponse(
            temp_file_path,
            media_type="application/vnd.ms-excel.sheet.macroEnabled.12",
            filename=f"processed_{file.filename}"
        )
    finally:
        file.file.close()
        # Optionally, delete temp_file_path after download if needed
        # os.remove(temp_file_path)
