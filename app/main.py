from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import StreamingResponse, JSONResponse
import pandas as pd
from ics import Calendar, Event
from datetime import datetime
import io
import pyexcel  # We'll use this for robust ODS reading

app = FastAPI(title="Schedule to ICS API")

@app.post("/upload-schedule")
async def upload_schedule(
    file: UploadFile = File(...),
    month: int = Form(...),
    year: int = Form(...)
):
    # === Read the uploaded file into a BytesIO buffer ===
    file.file.seek(0)
    file_bytes = file.file.read()

    # Determine file type and read with appropriate engine
    filename = file.filename.lower()
    
    try:
        if filename.endswith(".ods"):
            # WORKAROUND: pd.read_excel(engine="odf") fails on duration formats (e.g., "154:00")
            # It tries to parse them as time, ignoring dtype=str.
            # We read the raw data with pyexcel first, then load into pandas.
            
            # 1. Read raw data using pyexcel
            book_data = pyexcel.get_book_dict(file_content=file_bytes, file_type="ods")
            
            if "Plan" not in book_data:
                return JSONResponse(status_code=400, content={"error": "Sheet 'Plan' not found in ODS file."})
                
            sheet_data = book_data["Plan"]
            
            # 2. Load list-of-lists into pandas, forcing string type and filling empty cells
            df = pd.DataFrame(sheet_data).astype(str).fillna("")

        elif filename.endswith((".xlsx", ".xls")):
            # The openpyxl engine respects dtype=str much better
            buffer = io.BytesIO(file_bytes)
            df = pd.read_excel(buffer, engine="openpyxl", header=None, sheet_name="Plan", dtype=str)
            df = df.fillna("") # Ensure NaNs become empty strings
        else:
            return JSONResponse(status_code=400, content={"error": "Unsupported file type. Use .ods or .xlsx"})
            
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": f"Failed to read file: {e}"})

    # === Generate ICS calendar ===
    cal = Calendar()
    rows, cols = df.shape

    for row in range(rows):
        for col in range(cols):
            # All data is now guaranteed to be a string, so we can safely .strip()
            cell_value = df.iat[row, col].strip()

            # 1️⃣ Look for "Diana"
            if cell_value.lower() == "diana":
                # Check if the next row has "Dzień"
                if row + 1 < rows and df.iat[row + 1, col].strip().lower() == "dzień":
                    # 2️⃣ Go right and find numbers 1–31
                    c = col + 1
                    while c < cols:
                        val = df.iat[row + 1, c].strip()
                        if not val: # Skip empty cells
                            c += 1
                            continue

                        try:
                            day = int(val)
                        except ValueError:
                            c += 1
                            continue

                        if 1 <= day <= 31:
                            # 3️⃣ Start time (row below)
                            start_raw = df.iat[row + 2, c].strip() if row + 2 < rows else ""
                            # 4️⃣ End time (row below start)
                            end_raw = df.iat[row + 3, c].strip() if row + 3 < rows else ""
                            if start_raw and end_raw:
                                try:
                                    # 5️⃣ Create event
                                    if start_raw == end_raw:
                                        continue
                                    # Base start datetime
                                    start_dt = datetime.strptime(f"{year}-{month:02d}-{day:02d} {start_raw}", "%Y-%m-%d %H:%M:%S")
                                    
                                    # Base end datetime
                                    if end_raw == "24:00":
                                        # "24:00" means 00:00 of the *next* day.
                                        # Use Timedelta to safely handle month/year rollovers.
                                        end_dt_base = datetime(year, month, day) + pd.Timedelta(days=1)
                                        end_dt = end_dt_base.replace(hour=0, minute=0)
                                    else:
                                        # Standard time on the same day
                                        end_dt = datetime.strptime(f"{year}-{month:02d}-{day:02d} {end_raw}", "%Y-%m-%d %H:%M:%S")

                                    # Handle other overnight shifts (e.g., 22:00 - 06:00)
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

    # === Return .ics file ===
    ics_buffer = io.StringIO(cal.serialize())
    response = StreamingResponse(
        iter([ics_buffer.getvalue()]),
        media_type="text/calendar"
    )
    response.headers["Content-Disposition"] = "attachment; filename=work_schedule.ics"
    return response