<<<<<<< HEAD
<<<<<<< HEAD
import flask
from flask import Blueprint, Flask, jsonify, send_file, request
=======
>>>>>>> 08cce2e (moved classes to different files for better code readability, added backgroundtask class for deleting old generated zig zag folders)
import pyodbc
from flask_cors import CORS
from collections import defaultdict
import win32com.client
import win32com
<<<<<<< HEAD

from typing import List
from pydantic import BaseModel
from pathlib import Path
import waitress

=======
from datetime import datetime
from pathlib import Path
import atexit
import threading
from threading import Event

import uvicorn
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from background_task import BackgroundTasks
from models.pydantic_models import Patients, Visits, RequestedZigzag
from helper import get_connection_str, timestamp_now

patient_dict = defaultdict()
queried = False
<<<<<<< HEAD
>>>>>>> 08cce2e (moved classes to different files for better code readability, added backgroundtask class for deleting old generated zig zag folders)
=======
stop_event = Event()
>>>>>>> ddc0d41 (moved classes into a different file to improve the structure of project)
PATH_TO_THIS_FOLDER = Path(__file__).resolve().parent
PPT_FILE = Path(PATH_TO_THIS_FOLDER, "stdbatt_v2022.pptm").resolve()
BAS_FILE = Path(PATH_TO_THIS_FOLDER, "modMain.bas").resolve()
<<<<<<< HEAD
OUT_PPT = Path(PATH_TO_THIS_FOLDER, "output.pptm").resolve()
OUT_PDF = Path(PATH_TO_THIS_FOLDER, "output.pdf").resolve()
HOST = "127.0.0.1"
PORT = "4997"
URL_PREFIX = "/zigzag_backend"
=======
CONNECTION_STR = Path(PATH_TO_THIS_FOLDER, "connection_string.txt").resolve()
>>>>>>> ddc0d41 (moved classes into a different file to improve the structure of project)

<<<<<<< HEAD
bp = Blueprint(
    "main_blueprint", __name__, static_folder="static", template_folder="templates"
)
=======
if PATH_TO_SESSIONS_FOLDER.is_dir() == False:
    PATH_TO_SESSIONS_FOLDER.mkdir()

BACKGROUND_TASK = BackgroundTasks(path=PATH_TO_SESSIONS_FOLDER, stop_event=stop_event)

print("Starting background thread")
BACKGROUND_TASK.start()

<<<<<<< HEAD
origins = ["*"]
>>>>>>> 08cce2e (moved classes to different files for better code readability, added backgroundtask class for deleting old generated zig zag folders)

SQLMasterData = (
    "Provider=SQLOLEDB;Server=Spinal;Database=IBACohortReports;Integrated Security=SSPI;DataTypeCompatibility=80;MARS Connection=True;"
)
cnxn = pyodbc.connect(
    "DRIVER={SQL Server};Provider=SQLOLEDB;Server=Spinal;Database=IBACohortReports;Integrated Security=SSPI;DataTypeCompatibility=80;MARS Connection=True;"
)

patient_dict = defaultdict()
queried = False

<<<<<<< HEAD
class Patient(BaseModel):
    id: int
=======
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

app = FastAPI(openapi_url=None)

<<<<<<< HEAD
class Visits(BaseModel):
>>>>>>> 415cdb1 (added new get_zigzag endpoint which returns pdf file, get ppt endpoint returns ppt file)
    visits: List[int]


class PatientData(BaseModel):
    datas: List[Patient]

@bp.route("/get_patients", methods=["GET"])
def get_patient():
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

    response = flask.jsonify(patients=list(patient_dict.keys()))
    response.headers.add("Access-Control-Allow-Origin", "*")
    return response


@bp.route("/get_visits", methods=["GET"])
def get_visits():
    args = request.args
    patient_id = args.get("patient_id")
    response = jsonify(visits=patient_dict[int(patient_id)])
    response.headers.add("Access-Control-Allow-Origin", "*")
    print("Visits request completed ")
    return response


@bp.route("/get_zigzag", methods=["POST"])
def get_zigzag():
=======
=======
SQLMasterData = (get_connection_str(CONNECTION_STR, 1))
cnxn = pyodbc.connect(get_connection_str(CONNECTION_STR, 0))

app = FastAPI(openapi_url=None)

origins = ["*"]

>>>>>>> ddc0d41 (moved classes into a different file to improve the structure of project)
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
        
        print(f'Will Query DB: {not queried}')
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
            print('Queried Complete!')

        return Patients(patients=list(patient_dict.keys()))
    
    except Exception as e:
        print("Error Loading Patients", str(e))
        return ({"detail": "Not Found", 
                 "error": str(e)})

@app.get("/get_visits")
def get_visits(p_id: int):
    try:
        return Visits(visits=[visit[0] for visit in patient_dict[int(p_id)]], years=[visit[1] for visit in patient_dict[int(p_id)]])
    except Exception as e:
        print("Unable to find patient ID", str(e))
        return Visits(visits=[], years=[])

@app.post("/get_zigzag")
def get_zigzag(request: RequestedZigzag):
>>>>>>> 08cce2e (moved classes to different files for better code readability, added backgroundtask class for deleting old generated zig zag folders)
    try:
        data = request.json  # Assuming the data is in JSON format
        print(f'Requested data: {data}')

        global path_to_pptx

<<<<<<< HEAD
        visits = data.get("visits")
        p_id = data.get("patient_id")
=======
        p_id = request.p_id
        visits = list(request.visits)
        print(f'Requested data: {p_id}, {request.visits}')

        cur_time = timestamp_now(compact=True)
        session_id = f"{cur_time}-{p_id}-{'-'.join(map(str, list(request.visits)))}"
        this_session_folder = PATH_TO_SESSIONS_FOLDER / session_id
        
        # make new folder for session
        this_session_folder.mkdir()

        # copy operation
        destination = this_session_folder / f"{session_id}.pptm"
        destination.write_bytes(PPT_FILE.read_bytes())
>>>>>>> 04c55a9 (added p_id to sessionid)

        ppt = win32com.client.Dispatch("PowerPoint.Application")

<<<<<<< HEAD
        # Process the data as needed
        response_data = {"message": "Data received successfully"}
        wb = ppt.Presentations.Open(PPT_FILE)
        # Original location: \\marcfs\Database\Reports\ZigZag\AutoZigZagChartSQL C2.ppt

=======
        ppt_file = ppt.Presentations.Open(destination)
        # Original location: \\marcfs\Database\Reports\ZigZag\stdbatt_v2022.pptm
        
>>>>>>> ddc0d41 (moved classes into a different file to improve the structure of project)
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

        print("Completed Zig Zag")

        # Save the PowerPoint file to a temporary location
<<<<<<< HEAD
        wb.SaveAs(OUT_PDF, 32)
=======
        # copy operation
        destination = this_session_folder / f"{session_id}.pdf"

        ppt_file.SaveAs(destination, 32)
        ppt_file.Close()

        print("Sending PDF File.")

        headers = {
            "Access-Control-Expose-Headers": "*",
            "Content-Type": "application/pdf"
        }
        
        return FileResponse(path=destination, headers=headers, filename=f'{session_id}.pdf')
        
    except Exception as e:
        print("Error Loading Zig Zag", str(e))
        return ({"detail": "Not Found", 
                 "error": str(e)})

@app.post("/get_ppt")
def get_ppt(request: RequestedZigzag):
    try:
        global path_to_pptx

        p_id = request.p_id
        visits = list(request.visits)
        print(f'Requested data: {p_id}, {request.visits}')

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

        print("Completed Zig Zag")

        # Save the PowerPoint file to a tempor  ary location
<<<<<<< HEAD
        wb.SaveAs(destination)
>>>>>>> 415cdb1 (added new get_zigzag endpoint which returns pdf file, get ppt endpoint returns ppt file)
        ppt.Quit()
=======
        ppt_file.SaveAs(destination)
        ppt_file.Close()
>>>>>>> ddc0d41 (moved classes into a different file to improve the structure of project)

        print("Sending Powerpoint File.")
        
<<<<<<< HEAD
        return send_file(OUT_PDF, as_attachment=True, download_name='output.pdf')

=======
        return FileResponse(path=destination, headers=headers, filename=f'{session_id}.pptm')
    
>>>>>>> 415cdb1 (added new get_zigzag endpoint which returns pdf file, get ppt endpoint returns ppt file)
    except Exception as e:
        print("Error Loading Zig Zag", str(e))
        return jsonify({"error": str(e)}), 500
    

@bp.route("get_ppt", methods=["POST"])
def get_ppt():
    try:
        data = request.json  # Assuming the data is in JSON format
        print(f'Requested data: {data}')

        global path_to_pptx

        visits = data.get("visits")
        p_id = data.get("patient_id")

        ppt = win32com.client.Dispatch("PowerPoint.Application")

        # Process the data as needed
        response_data = {"message": "Data received successfully"}
        wb = ppt.Presentations.Open(PPT_FILE)
        # Original location: \\marcfs\Database\Reports\ZigZag\AutoZigZagChartSQL C2.ppt

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
        wb.SaveAs(OUT_PPT)
        ppt.Quit()

        print("Sending Powerpoint File.")
        
        return send_file(OUT_PPT, as_attachment=True, download_name='output.pptm')

    except Exception as e:
        print("Error Loading Zig Zag", str(e))
        return jsonify({"error": str(e)}), 500

flask_app = Flask(__name__)
CORS(flask_app)  # Enable CORS for all routes
flask_app.config["APPLICATION_ROOT"] = URL_PREFIX
flask_app.register_blueprint(bp, url_prefix=URL_PREFIX)

@atexit.register
def terminate_background():
    print("Terminating background thread")
    BACKGROUND_TASK.join()

if __name__ == "__main__":
<<<<<<< HEAD
<<<<<<< HEAD
    print(f"Attempting to serve on http://{HOST}:{PORT}{URL_PREFIX}")

    listenStr = f"{HOST}:{PORT}"
    waitress.serve(flask_app, listen=listenStr, log_untrusted_proxy_headers=True)
=======
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

if PATH_TO_SESSIONS_FOLDER.is_dir() == False:
    PATH_TO_SESSIONS_FOLDER.mkdir()

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
>>>>>>> refs/rewritten/main-3
=======
    uvicorn.run(app, port=8000)
=======
    uvicorn.run(app, host="0.0.0.0", port=8000)
>>>>>>> 04c55a9 (added p_id to sessionid)
    stop_event.set()

    
>>>>>>> 08cce2e (moved classes to different files for better code readability, added backgroundtask class for deleting old generated zig zag folders)
