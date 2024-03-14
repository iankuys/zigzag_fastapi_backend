from typing import List
import pyodbc
from collections import defaultdict
import win32com.client
import win32com
from datetime import datetime
from pathlib import Path

import uvicorn
from pydantic import BaseModel
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

app = FastAPI(openapi_url=None)
patient_dict = defaultdict()
queried = False
PATH_TO_THIS_FOLDER = Path(__file__).resolve().parent
PATH_TO_SESSIONS_FOLDER = PATH_TO_THIS_FOLDER / "sessions"
PPT_FILE = Path(PATH_TO_THIS_FOLDER, "stdbatt_v2022.pptm").resolve()
BAS_FILE = Path(PATH_TO_THIS_FOLDER, "modMain.bas").resolve()

origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

SQLMasterData = (
    "Provider=SQLOLEDB;"
    "Server=Spinal;"
    "Database=IBACohortReports;"
    "Integrated Security=SSPI;"
    "DataTypeCompatibility=80;"
    "MARS Connection=True;"
)

cnxn = pyodbc.connect(
    "DRIVER={SQL Server};"
    "Provider=SQLOLEDB;"
    "Server=Spinal;"
    "Database=IBACohortReports;"
    "Integrated Security=SSPI;"
    "DataTypeCompatibility=80;"
    "MARS Connection=True;"
)


def timestamp_now(compact=False, only_ymd=False) -> str:
    """Credit: Brandon
        Returns a string of the current date+time in the form of
        YYYY-MM-DD hh:mm:ss
    If `compact` == True, then returns in the form of
        YYYYMMDD_hhmmss
    If `only_ymd` == True, then only the first "year/month/day" portion is returned:
        YYYY-MM-DD or YYYYMMDD
    """
    timestamp = datetime.now()
    if compact:
        if only_ymd:
            return timestamp.strftime("%Y%m%d")
        return timestamp.strftime("%Y%m%d_%H%M%S")
    if only_ymd:
        return timestamp.strftime("%Y-%m-%d")
    return timestamp.strftime("%Y-%m-%d %H:%M:%S")

class Patients(BaseModel):
    patients: List[int]

class Visits(BaseModel):
    visits: List[int]

class RequestedZigzag(BaseModel):
    p_id: int
    visits: List[int]

@app.get("/get_patients")
async def get_patient():
    global queried
    global patient_dict
    cursor = cnxn.cursor()
    
    print(f'Will Query DB: {not queried}')
    if not queried:
        cursor.execute(
            "SELECT [PatientID], [VisitNumber] FROM [IBACohort].[Npsych].[vwScores_StdBatt_v2022]"
        )
        rows = cursor.fetchall()
        for row in rows:
            if row.PatientID not in patient_dict:
                patient_dict[row.PatientID] = []
            patient_dict[row.PatientID].append(row.VisitNumber)
        queried = True
        print('Queried Complete!')

    return Patients(patients=list(patient_dict.keys()))


@app.get("/get_visits")
async def get_visits(p_id: int):
    return Visits(visits=patient_dict[int(p_id)])


@app.post("/get_zigzag")
async def get_zigzag(request: RequestedZigzag):
    try:
        global path_to_pptx

        p_id = request.p_id
        visits = list(request.visits)
        print(f'Requested data: {p_id}, {request.visits}')

        cur_time = timestamp_now(compact=True)
        session_id = f"{cur_time}-{'-'.join(map(str, list(request.visits)))}"
        this_session_folder = PATH_TO_SESSIONS_FOLDER / session_id
        
        # make new folder for session
        this_session_folder.mkdir()

        # copy operation
        destination = this_session_folder / f"{session_id}.pptm"
        destination.write_bytes(PPT_FILE.read_bytes())

        ppt = win32com.client.Dispatch("PowerPoint.Application")

        wb = ppt.Presentations.Open(destination)
        # Original location: \\marcfs\Database\Reports\ZigZag\stdbatt_v2022.pptm
        
        ppt.VBE.ActiveVBProject.VBComponents.Import(BAS_FILE)
        # Original location: \\marcfs\Database\Reports\ZigZag\StdBatt_v2022_Vue.JS\modMain.bas

        ppt.Run("SetDBMaster", SQLMasterData)

        # if visits are more than one we run the macro for multiple zigzags
        if len(visits) > 1:
            color = 1
            for visit in visits:
                ppt.Run("SetSubject", int(p_id), 1, int(visit), color)
                color += 1
        else:
            ppt.Run("SetSubject", int(p_id), 1, int(visits[0]), 1)

        print("Completed Zig Zag")

        # Save the PowerPoint file to a temporary location
        wb.SaveAs(destination)
        ppt.Quit()

        print("Sending Powerpoint File.")

        headers = {
            "Access-Control-Expose-Headers": "*"
        }
        
        return FileResponse(path=destination, headers=headers, filename=f'{session_id}.pptm')

    except Exception as e:
        print("Error Loading Zig Zag", str(e))
        return ({"detail": "Not Found", 
                 "error": str(e)})

if __name__ == "__main__":
    uvicorn.run(app, port=8000, host="0.0.0.0")