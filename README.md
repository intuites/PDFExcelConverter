# PDF ↔ Excel converter

Install the Python dependencies and start the API/web server from this folder:

```powershell
python -m pip install -r requirements.txt
python server.py
```

Open <http://127.0.0.1:8000>. Do not open `web/index.html` directly or use a static file server: conversion requests are handled by `server.py` at `POST /convert`.

Run the checks with:

```powershell
python test_convert.py
python test_server.py
```
