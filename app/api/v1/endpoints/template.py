"""Word template management endpoints."""

from fastapi import APIRouter, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.services.template_service import (
    DOCX_MIME,
    get_template_info,
    get_template_path,
    update_template_bytes,
)

router = APIRouter(prefix="/api/template", tags=["template"])


@router.get("")
async def template_info():
    """Get current template metadata."""
    try:
        return {"status": "ok", "template": get_template_info()}
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))


@router.get("/download")
async def download_template():
    """Download current template file."""
    try:
        path = get_template_path()
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))

    return FileResponse(
        path=path,
        media_type=DOCX_MIME,
        filename=path.name,
    )


@router.post("")
async def upload_template(file: UploadFile = File(...)):
    """Upload and activate a new Word template (.docx) without restarting the service."""
    data = await file.read()
    try:
        info = update_template_bytes(
            data,
            filename=file.filename,
            content_type=file.content_type,
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except FileNotFoundError as e:
        raise HTTPException(status_code=404, detail=str(e))
    except OSError as e:
        raise HTTPException(status_code=500, detail=f"Failed to save template: {e}")

    return {"status": "ok", "template": info}
