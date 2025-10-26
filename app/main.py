from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import JSONResponse, FileResponse
import pandas as pd
from ics import Calendar, Event
from datetime import datetime
import io
import pyexcel
import os
import logging

app = FastAPI(title="Schedule to ICS API")

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

@app.post("/upload-schedule")
async def upload_schedule(
    file: UploadFile = File(...),
    month: int = Form(...),
    year: int = Form(...)
):
    logger.info("Request received for /upload-schedule")
    # === Read the uploaded file into a BytesIO buffer ===
    logger.info("Starting file read")
    file.file.seek(0)
    file_bytes = file.file.read()
    logger.info(f"File read complete, size: {len(file_bytes)} bytes")

    # Determine file type and read with appropriate engine
    filename = file.filename.lower()
    
    try:
        if filename.endswith(".ods"):
            logger.info("Processing ODS file")
            book_data = pyexcel.get_book_dict(file_content=file_bytes, file_type="ods")
            if "Plan" not in book_data:
                logger.error("Sheet 'Plan' not found in ODS file")
                return JSONResponse(status_code=400, content={"error": "Sheet 'Plan' not found in ODS file."})
            sheet_data = book_data["Plan"]
            df = pd.DataFrame(sheet_data).astype(str).fillna("")
        elif filename.endswith((".xlsx", ".xls")):
            logger.info("Processing Excel file")
            buffer = io.BytesIO(file_bytes)
            df = pd.read_excel(buffer, engine="openpyxl", header=None, sheet_name="Plan", dtype=str)
            df = df.fillna("")
        else:
            logger.error("Unsupported file type")
            return JSONResponse(status_code=400, content={"error": "Unsupported file type. Use .ods or .xlsx"})
        logger.info(f"DataFrame created: {df.shape[0]} rows, {df.shape[1]} columns")
    except Exception as e:
        logger.error(f"Failed to read file: {e}")
        return JSONResponse(status_code=500, content={"error": f"Failed to read file: {e}"})

    # === Generate ICS calendar ===
    cal = Calendar()
    rows, cols = df.shape

    for row in range(rows):
        for col in range(cols):
            cell_value = df.iat[row, col].strip()
            if cell_value.lower() == "diana":
                if row + 1 < rows and df.iat[row + 1, col].strip().lower() == "dzień":
                    c = col + 1
                    while c < cols:
                        val = df.iat[row + 1, c].strip()
                        if not val:
                            c += 1
                            continue
                        try:
                            day = int(val)
                        except ValueError:
                            c += 1
                            continue
                        if 1 <= day <= 31:
                            start_raw = df.iat[row + 2, c].strip() if row + 2 < rows else ""
                            end_raw = df.iat[row + 3, c].strip() if row + 3 < rows else ""
                            if start_raw and end_raw:
                                try:
                                    start_dt = datetime.strptime(f"{year}-{month:02d}-{day:02d} {start_raw}", "%Y-%m-%d %H:%M:%S")
                                    if end_raw == "24:00":
                                        end_dt_base = datetime(year, month, day) + pd.Timedelta(days=1)
                                        end_dt = end_dt_base.replace(hour=0, minute=0)
                                    else:
                                        end_dt = datetime.strptime(f"{year}-{month:02d}-{day:02d} {end_raw}", "%Y-%m-%d %H:%M:%S")
                                    if end_dt <= start_dt:
                                        end_dt += pd.Timedelta(days=1)
                                    event = Event(
                                        name="Работа",
                                        begin=start_dt,
                                        end=end_dt
                                    )
                                    cal.events.add(event)
                                except Exception as e:
                                    print(f"⛔ Error processing day {day} (Times: '{start_raw}' -> '{end_raw}'): {e}")
                        c += 1

    # === Return ICS file ===
    logger.info("Generating ICS content")
    ics_content = cal.serialize()
    # Save ICS content to a persistent file on the system with a unique name
    output_path = f"work_schedule_{year}{month:02d}_{int(datetime.now().timestamp())}.ics"
    try:
        logger.info(f"Writing ICS file to {output_path}")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(ics_content)
        logger.info(f"ICS file successfully written to {output_path}")
    except Exception as e:
        logger.error(f"Failed to write ICS file to {output_path}: {e}")
        return JSONResponse(status_code=500, content={"error": f"Failed to write ICS file: {e}"})

    # Ensure file exists before sending
    if not os.path.exists(output_path):
        logger.error(f"ICS file not found at {output_path}")
        return JSONResponse(status_code=500, content={"error": "Failed to create ICS file on server."})

    logger.info(f"Sending ICS file: {output_path}")
    return FileResponse(
        path=output_path,
        filename="work_schedule.ics",
        media_type="text/calendar",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "close",
            "Content-Disposition": f"attachment; filename=work_schedule.ics"
        }
    )