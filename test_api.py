import requests

url = "http://127.0.0.1:8000/run-macro/"
data = {
    "file_path": r"C:\Users\ASCII\IdeaProjects\vba\sample_button.xlsm",
    "macro_name": "HighlightMaxMinAndDifference"
}

resp = requests.post(url, json=data)
print(resp.json())
