import pyodbc
from flask_cors import CORS
from collections import defaultdict
import win32com.client
import win32com
from datetime import datetime
from pathlib import Path
import atexit
import threading
from threading import Event

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from utils.background_task import BackgroundTasks
from models.pydantic_models import Patients, Visits, RequestedZigzag
from utils.utils import get_connection_str, timestamp_now, print_log

patient_dict = defaultdict()
queried = False
stop_event = Event()

PATH_TO_THIS_FOLDER = Path(__file__).resolve().parent
PATH_TO_SESSIONS_FOLDER = PATH_TO_THIS_FOLDER / "sessions"
PPT_FILE = Path(PATH_TO_THIS_FOLDER, "stdbatt_v2022.pptm").resolve()
BAS_FILE = Path(PATH_TO_THIS_FOLDER, "modMain.bas").resolve()
CONNECTION_STR = Path(PATH_TO_THIS_FOLDER, "connection_string.txt").resolve()

if PATH_TO_SESSIONS_FOLDER.is_dir() == False:
    PATH_TO_SESSIONS_FOLDER.mkdir()

BACKGROUND_TASK = BackgroundTasks(path=PATH_TO_SESSIONS_FOLDER, stop_event=stop_event)

print_log("Starting background thread")
BACKGROUND_TASK.start()

SQLMasterData = (get_connection_str(CONNECTION_STR, 1))
cnxn = pyodbc.connect(get_connection_str(CONNECTION_STR, 0))

app = FastAPI(openapi_url=None)

origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/get_patients")
def get_patient():
    try:
        global queried
        global patient_dict
        cursor = cnxn.cursor()
        
        print_log(f'Will Query DB: {not queried}')
        if not queried:
            cursor.execute(
                "SELECT [PatientID], [VisitNumber], YEAR(ExamDate) as [Year] FROM [IBACohort].[Npsych].[vwScores_StdBatt_v2022]"
            )
            rows = cursor.fetchall()
            for row in rows:
                if row.PatientID not in patient_dict:
                    patient_dict[row.PatientID] = []
                patient_dict[row.PatientID].append((row.VisitNumber, row.Year))
            queried = True
            print_log('Queried Complete!')

        return Patients(patients=list(patient_dict.keys()))
    
    except Exception as e:
        print_log(f"Error Loading Zig Zag, {str(e)}")
        return ({"detail": "Not Found", 
                 "error": str(e)})

@app.get("/get_visits")
def get_visits(p_id: int):
    try:
        return Visits(visits=[visit[0] for visit in patient_dict[int(p_id)]], years=[visit[1] for visit in patient_dict[int(p_id)]])
    except Exception as e:
        print_log(f"Unable to find patient ID {str(e)}")
        return Visits(visits=[], years=[])

@app.post("/get_zigzag")
async def get_zigzag(request: RequestedZigzag):
    try:
        p_id = request.p_id
        visits = list(request.visits)
        print_log(f'Requested data: {p_id}, {request.visits}')

        cur_time = timestamp_now(compact=True)
        session_id = f"{cur_time}-{p_id}-{'-'.join(map(str, list(request.visits)))}"
        this_session_folder = PATH_TO_SESSIONS_FOLDER / session_id
        
        # make new folder for session
        this_session_folder.mkdir()

        # copy operation
        destination = this_session_folder / f"{session_id}.pptm"
        destination.write_bytes(PPT_FILE.read_bytes())

        ppt = win32com.client.Dispatch("PowerPoint.Application")

        ppt_file = ppt.Presentations.Open(destination)
        # Original location: \\marcfs\Database\Reports\ZigZag\stdbatt_v2022.pptm
        
        ppt.VBE.ActiveVBProject.VBComponents.Import(BAS_FILE)
        # Original location: \\marcfs\Database\Reports\ZigZag\StdBatt_v2022_Vue.JS\modMain.bas

        ppt.Run("SetDBMaster", SQLMasterData)

        # if visits are more than one we run the macro for multiple zigzags
        if len(visits) > 1:
            color = len(visits)
            for visit in visits:
                ppt.Run("SetSubject", int(p_id), 1, int(visit), color)
                color -= 1
        else:
            ppt.Run("SetSubject", int(p_id), 1, int(visits[0]), 1)

        print_log("Completed Zig Zag")

        # Save the PowerPoint file to a temporary location
        # copy operation
        pdf_destination = this_session_folder / f"{session_id}.pdf"

        ppt_file.SaveAs(pdf_destination, 32)
        ppt_file.Close()

        print_log("Sending PDF File.")

        headers = {
            "Access-Control-Expose-Headers": "*",
            "Content-Type": "application/pdf"
        }
        
        return FileResponse(path=pdf_destination, headers=headers, filename=f'{session_id}.pdf')
        
    except Exception as e:
        print_log(f"Error Loading Zig Zag, {str(e)}")
        return ({"detail": "Not Found", 
                 "error": str(e)})

@app.post("/get_ppt")
async def get_ppt(request: RequestedZigzag):
    try:
        p_id = request.p_id
        visits = list(request.visits)
        print_log(f'Requested data: {p_id}, {request.visits}')

        cur_time = timestamp_now(compact=True)
        session_id = f"{cur_time}-{p_id}-{'-'.join(map(str, list(request.visits)))}"
        this_session_folder = PATH_TO_SESSIONS_FOLDER / session_id
        
        # make new folder for session
        this_session_folder.mkdir()

        # copy operation
        destination = this_session_folder / f"{session_id}.pptm"
        destination.write_bytes(PPT_FILE.read_bytes())

        ppt = win32com.client.Dispatch("PowerPoint.Application")

        ppt_file = ppt.Presentations.Open(destination)
        # Original location: \\marcfs\Database\Reports\ZigZag\stdbatt_v2022.pptm
        
        ppt.VBE.ActiveVBProject.VBComponents.Import(BAS_FILE)
        # Original location: \\marcfs\Database\Reports\ZigZag\StdBatt_v2022_Vue.JS\modMain.bas

        ppt.Run("SetDBMaster", SQLMasterData)

        # if visits are more than one we run the macro for multiple zigzags
        if len(visits) > 1:
            color = len(visits)
            for visit in visits:
                ppt.Run("SetSubject", int(p_id), 1, int(visit), color)
                color -= 1
        else:
            ppt.Run("SetSubject", int(p_id), 1, int(visits[0]), 1)

        print_log("Completed Zig Zag")

        # Save the PowerPoint file to a temporary location
        ppt_file.SaveAs(destination)
        ppt_file.Close()

        print_log("Sending Powerpoint File.")

        headers = {
            "Access-Control-Expose-Headers": "*"
        }
        
        return FileResponse(path=destination, headers=headers, filename=f'{session_id}.pptm')
    
    except Exception as e:
        print_log(f"Error Loading Zig Zag, {str(e)}")
        return ({"detail": "Not Found", 
                 "error": str(e)})

@atexit.register
def terminate_background():
    print_log("Terminating background thread")
    BACKGROUND_TASK.join()

if __name__ == "__main__":
    uvicorn.run(app, host="0.0.0.0", port=8000)
    stop_event.set()
