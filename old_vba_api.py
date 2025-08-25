from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import os
import pythoncom
import win32com.client as win32

app = FastAPI(title="VBA Macro Runner")

# Pydantic model for request body
class MacroRequest(BaseModel):
    file_path: str
    macro_name: str

def run_macro(file_path: str, macro_name: str):
    """
    Run a VBA macro in the given Excel file and return the C2 value.
    Ensures COM is initialized in this thread.
    Normalizes file paths for Windows Excel COM.
    """
    # Normalize path for Windows
    file_path = os.path.normpath(file_path)

    # Initialize COM for this thread
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

        return {"status": "success", "c2_value": c2_value}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        if excel:
            excel.Quit()
        pythoncom.CoUninitialize()


# POST endpoint: expects JSON body
@app.post("/run-macro/")
def run_excel_macro(req: MacroRequest):
    return run_macro(req.file_path, req.macro_name)
