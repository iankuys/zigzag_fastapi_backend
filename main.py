import flask
from flask import Flask, jsonify, render_template, request
import pyodbc
from flask_cors import CORS
from collections import defaultdict 
import win32com.client
import win32com

from typing import List
from pydantic import BaseModel

SQLMasterData = (
    "Provider=SQLOLEDB;"
    "Server=Spinal;"
    "Database=IBACohortReports;"
    "Integrated Security=SSPI;"
    "DataTypeCompatibility=80;"
    "MARS Connection=True;"
)
cnxn = pyodbc.connect("DRIVER={SQL Server};"
                    "Provider=SQLOLEDB;"
                    "Server=Spinal;"
                    "Database=IBACohortReports;"
                    "Integrated Security=SSPI;"
                    "DataTypeCompatibility=80;"
                    "MARS Connection=True;")

flask_app = Flask(__name__)
CORS(flask_app)  # Enable CORS for all routes
patient_dict = defaultdict()
queried = False

class Patient(BaseModel):
    id: int 
    visits: List[int]

class PatientData(BaseModel):
    datas: List[Patient]

@flask_app.route("/")
def root():
    return render_template("index.html")

@flask_app.route("/get_patients", methods=["GET"])
def get_patient():
    global queried
    global patient_dict
    cursor = cnxn.cursor()
    cursor.execute('SELECT [PatientID], [VisitNumber] FROM [IBACohort].[Npsych].[vwScores_StdBatt_v2022]')

    rows = cursor.fetchall()

    print(queried)
    if not queried:
        for row in rows:
            if row.PatientID not in patient_dict:
                patient_dict[row.PatientID] = []
            patient_dict[row.PatientID].append(row.VisitNumber)
        queried = True    

    response = flask.jsonify(patients=list(patient_dict.keys()))
    response.headers.add('Access-Control-Allow-Origin', '*')
    return response

@flask_app.route("/get_visits", methods=["GET"])
def get_visits():
    args = request.args
    patient_id = args.get('patient_id')
    response = jsonify(visits=patient_dict[int(patient_id)])
    response.headers.add('Access-Control-Allow-Origin', '*')
    print(response)
    return response

@flask_app.route("/get_zigzag", methods=["POST"])
def get_zigzag():

    try:
        data = request.json  # Assuming the data is in JSON format

        global path_to_pptx

        visits = data.get("visits")
        p_id = data.get("patient_id")
        
        ppt = win32com.client.Dispatch('PowerPoint.Application')
        ppt.Visible = True

        # Process the data as needed
        response_data = {'message': 'Data received successfully'}
        ppt.Presentations.Open(r'Z:\Reports\ZigZag\AutoZigZagChartSQL C2.ppt')
        ppt.VBE.ActiveVBProject.VBComponents.Import(r'C:\Users\ysku\Desktop\modMain.bas')
        ppt.Run("SetDBMaster", SQLMasterData)
    
        # if visits are more than one we run the macro for multiple zigzags
        if len(visits) > 1:    
            color = 1
            for visit in visits:
                ppt.Run("SetSubject", int(p_id), 1, int(visit), color)
                color += 1
        else:
            ppt.Run("SetSubject", int(p_id), 1, int(visits[0]), 1)
        
        return jsonify(response_data)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    
if __name__ == "__main__":
    flask_app.run()

