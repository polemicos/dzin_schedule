import os
import re
import ezodf
import tempfile
import logging
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse, FileResponse
from ics import Calendar, Event
import aiofiles
import isodate 
app = FastAPI(title="Schedule to ICS API")

# Logging setup
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def iso_to_hms(iso_duration):
    # Parse the ISO 8601 duration string
    duration = isodate.parse_duration(iso_duration)
    # Convert the duration to total seconds
    total_seconds = int(duration.total_seconds())
    # Format as HH:MM:SS
    return f"{total_seconds // 3600:02}:{(total_seconds % 3600) // 60:02}:{total_seconds % 60:02}"


def is_valid_time(time_str: str) -> bool:
    """Validate if a time string matches HH:MM or HH:MM:SS format and has valid hours."""
    pattern = r'^([0-1]?[0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$'
    return bool(re.match(pattern, time_str))


@app.post("/upload-schedule")
async def upload_schedule(
    file: UploadFile = File(...),
    month: int = Form(...),
    year: int = Form(...)
):
    logger.info("Request received for /upload-schedule")

    filename = file.filename.lower()
    if not filename.endswith(".ods"):
        return JSONResponse(
            status_code=400,
            content={"error": "Unsupported file type. Only .ods files are accepted."}
        )

    # Save the uploaded file to disk
    try:
        with tempfile.NamedTemporaryFile(suffix=".ods", delete=False) as tmp:
            tmp_path = tmp.name
        async with aiofiles.open(tmp_path, "wb") as out_file:
            while chunk := await file.read(1024 * 1024):
                await out_file.write(chunk)
        logger.info(f"File saved to {tmp_path}")
    except Exception as e:
        logger.error(f"Failed to save uploaded ODS file: {str(e)}")
        return JSONResponse(status_code=500, content={"error": str(e)})

    # Open ODS directly with ezodf
    try:
        doc = ezodf.opendoc(tmp_path)
        sheet = next((s for s in doc.sheets if s.name.lower() == "plan"), None)
        if not sheet:
            return JSONResponse(
                status_code=400,
                content={"error": "Sheet 'Plan' not found in ODS file"}
            )
    except Exception as e:
        logger.error(f"Failed to open ODS: {e}")
        return JSONResponse(status_code=500, content={"error": str(e)})
    finally:
        try:
            if os.path.exists(tmp_path):
                os.unlink(tmp_path)
        except Exception as cleanup_err:
            logger.warning(f"Failed to remove temp file: {cleanup_err}")

    # -------------------- Process schedule --------------------
    cal = Calendar()
    n_rows = sheet.nrows()
    n_cols = sheet.ncols()
    logger.info(f"Processing sheet: {n_rows} rows x {n_cols} cols")
    for row_idx in range(n_rows):
        for col_idx in range(n_cols):
            val = sheet[row_idx, col_idx].value
            if isinstance(val, str) and val.strip().lower() == "diana":
                
                # Check that next row has 'dzień'
                if row_idx + 1 < n_rows:
                    next_val = sheet[row_idx + 1, col_idx].value
                    if isinstance(next_val, str) and next_val.strip().lower() == "dzień":
                        c = col_idx + 1
                        while c < n_cols:
                            day_val = sheet[row_idx + 1, c].value
                            if not isinstance(day_val, float):
                                c += 1
                                continue

                            day = int(day_val)
                            start_raw = sheet[row_idx + 2, c].value
                            end_raw = sheet[row_idx + 3, c].value

                            try:
                                start_dt = f"{year}-{month:02d}-{day:02d} {iso_to_hms(start_raw)}"
                                end_dt = f"{year}-{month:02d}-{day:02d} {iso_to_hms(end_raw)}"
                                if start_dt == end_dt:
                                    c += 1
                                    continue
                                cal.events.add(Event(name="Работа", begin=start_dt, end=end_dt))
                                logger.info(f"Added event for day {day}: {start_dt} → {end_dt}")
                            except Exception as e:
                                    logger.error(f"Error adding event for day {day}: {e}")

                            c += 1

    if not cal.events:
        return JSONResponse(status_code=400, content={"error": "No valid events found in schedule."})

    # -------------------- Write and return ICS --------------------
    ics_content = cal.serialize()
    try:
        with tempfile.NamedTemporaryFile(mode="w", suffix=".ics", delete=False, encoding="utf-8") as tmp_ics:
            ics_path = tmp_ics.name
            tmp_ics.write(ics_content)
            tmp_ics.flush()
        logger.info(f"ICS file created at {ics_path}")
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to write ICS file: {e}"})

    return FileResponse(
        path=ics_path,
        filename="work_schedule.ics",
        media_type="text/calendar",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "close",
            "Content-Disposition": "attachment; filename=work_schedule.ics"
        },
    )

