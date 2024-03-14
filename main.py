import flask
from flask import Blueprint, Flask, jsonify, send_file, request
import pyodbc
from flask_cors import CORS
from collections import defaultdict
import win32com.client
import win32com

from typing import List
from pydantic import BaseModel
from pathlib import Path
import waitress

PATH_TO_THIS_FOLDER = Path(__file__).resolve().parent
PPT_FILE = Path(PATH_TO_THIS_FOLDER, "stdbatt_v2022.pptm").resolve()
BAS_FILE = Path(PATH_TO_THIS_FOLDER, "modMain.bas").resolve()
OUT_FILE = Path(PATH_TO_THIS_FOLDER, "output.pptm").resolve()
HOST = "127.0.0.1"
PORT = "4997"
URL_PREFIX = "/zigzag_backend"

bp = Blueprint(
    "main_blueprint", __name__, static_folder="static", template_folder="templates"
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

patient_dict = defaultdict()
queried = False

class Patient(BaseModel):
    id: int
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
        wb.SaveAs(OUT_FILE)
        ppt.Quit()

        print("Sending Powerpoint File.")
        
        return send_file(OUT_FILE, as_attachment=True, download_name='output.pptm')

    except Exception as e:
        print("Error Loading Zig Zag", str(e))
        return jsonify({"error": str(e)}), 500

flask_app = Flask(__name__)
CORS(flask_app)  # Enable CORS for all routes
flask_app.config["APPLICATION_ROOT"] = URL_PREFIX
flask_app.register_blueprint(bp, url_prefix=URL_PREFIX)

if __name__ == "__main__":
    print(f"Attempting to serve on http://{HOST}:{PORT}{URL_PREFIX}")

    listenStr = f"{HOST}:{PORT}"
    waitress.serve(flask_app, listen=listenStr, log_untrusted_proxy_headers=True)