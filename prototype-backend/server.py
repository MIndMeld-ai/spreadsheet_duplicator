from fastapi import FastAPI, File, UploadFile, Form, HTTPException
from fastapi.responses import StreamingResponse
import uvicorn
import tempfile
import zipfile
import io
import json
from openpyxl import load_workbook

app = FastAPI()

@app.post('/generate')
async def generate(template: UploadFile = File(...), mapping: str = Form(...)):
    try:
        mapping_obj = json.loads(mapping)
    except Exception as e:
        raise HTTPException(status_code=400, detail='Invalid mapping JSON')
    # For prototype: simply echo back the template or return a zip with copies per mapping row
    try:
        content = await template.read()
        wb = load_workbook(filename=io.BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=400, detail='Failed to read template')

    rows = mapping_obj.get('rows', [])
    if not rows:
        # return the original template
        return StreamingResponse(io.BytesIO(content), media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, mode='w') as zf:
        for i, row in enumerate(rows):
            # For prototype just store the original file under a new name
            name = f'output_{i+1}.xlsx'
            zf.writestr(name, content)
    mem.seek(0)
    return StreamingResponse(mem, media_type='application/zip')

if __name__ == '__main__':
    uvicorn.run(app, host='0.0.0.0', port=8000, reload=True)
