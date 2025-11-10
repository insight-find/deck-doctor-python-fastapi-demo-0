from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse
from pptx import Presentation
import io

app = FastAPI(title="PPTX Edit Demo")


def replace_text_in_presentation(pres: Presentation, token: str, replacement: str):
    """Simple pass: replace token in any text frame."""
    for slide in pres.slides:
        for shape in slide.shapes:
            if not hasattr(shape, "text"):
                continue
            if token in shape.text:
                shape.text = shape.text.replace(token, replacement)


@app.post("/modify-pptx")
async def modify_pptx(file: UploadFile = File(...), name: str = Form(...)):
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Upload a .pptx file.")

    # Read file bytes
    contents = await file.read()
    # Load Presentation from bytes buffer
    bio = io.BytesIO(contents)
    pres = Presentation(bio)

    # Perform modification (replace token with name)
    replace_text_in_presentation(pres, "{{NAME}}", name)

    # Save to bytes buffer
    out = io.BytesIO()
    pres.save(out)
    out.seek(0)

    headers = {
        "Content-Disposition": f'attachment; filename="modified-{file.filename}"'
    }
    return StreamingResponse(out,
                             media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                             headers=headers)
