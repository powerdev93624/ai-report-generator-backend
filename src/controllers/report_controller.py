from flask import request, Response, json, Blueprint
from src.models.user_model import User
from src import bcrypt, db
from datetime import datetime, timedelta, timezone
from src.middlewares import authentication_required
import jwt
import os
import uuid
from src.services.functions import *
import pythoncom
import pandas as pd




# user controller blueprint to be registered with api blueprint
report = Blueprint("report", __name__)
original_doc_path = "src/files/sample/sample.docx"
output_doc_path = "src/files/report/report.docx"
@report.route('/generate', methods = ["POST"])
def generate():
    pythoncom.CoInitializeEx(0)
    scores = pd.read_csv("src/files/score/score.csv")
    result_path = "src/files/result/result.csv"
    result_df = pd.read_csv(result_path)
    doc = Document(original_doc_path)
    paragraphs = doc.paragraphs
    try:
        data = request.json
        full_name = data["firstName"]+" "+ data["lastName"]
        with open('src/files/transcript/transcript.txt', 'r') as file:
            transcript = file.read()
        replace_full_name(paragraphs, full_name)
        print(1)
        replace_assessor_and_organization(paragraphs, data["assessor"], data["organization"])
        print(2)
        replace_background(paragraphs, get_background(full_name, transcript))
        print(3)
        replace_strategic_partner(paragraphs, get_strategic_partner(full_name, transcript))
        print(4)
        replace_innovate(paragraphs, get_innovate(full_name, transcript))
        print(5)
        replace_connect(paragraphs, get_connect(full_name, transcript))
        print(6)
        replace_simplify(paragraphs, get_simplify(full_name, transcript))
        print(7)
        replace_talent_enabler(paragraphs, get_talent_enabler(full_name, transcript))
        print(8)
        replace_coach(paragraphs, get_coach(full_name, transcript))
        print(9)
        replace_empower(paragraphs, get_empower(full_name, transcript))
        print(10)
        replace_elevate(paragraphs, get_elevate(full_name, transcript))
        print(11)
        replace_agile_executor(paragraphs, get_agile_executor(full_name, transcript))
        print(12)
        replace_deliver(paragraphs, get_deliver(full_name, transcript))
        print(13)
        replace_adapt(paragraphs, get_adapt(full_name, transcript))
        print(14)
        replace_pioneer(paragraphs, get_pioneer(full_name, transcript))
        print(15)
        replace_resilient_steward(paragraphs, get_resilient_steward(full_name, transcript))
        print(16)
        replace_serve(paragraphs, get_serve(full_name, transcript))
        print(17)
        replace_inspire(paragraphs, get_inspire(full_name, transcript))
        print(18)
        replace_sustain(paragraphs, get_sustain(full_name, transcript))
        print(19)
        replace_recommendation(paragraphs,full_name, transcript, scores)
        print(20)
        replace_next_steps(paragraphs, get_next_steps(full_name, transcript))
        print(21)
        doc.save(output_doc_path)
        fill_score(scores, output_doc_path)
        
        current_date = datetime.now().strftime('%Y-%m-%d')
        new_row = {
            'id': len(result_df), "Full Name": full_name, "Date": current_date
        }
        for col in result_df.columns:
            if col == "id":
                new_row["id"] = len(result_df)
            elif col == "Full Name":
                new_row["Full Name"] = full_name
            elif col == "Date":
                new_row["Date"] = current_date
            else:
                new_row[col] = str(scores.loc[scores['id'] == col, 'score'].values[0])
        result_df.loc[len(result_df)] = new_row
        result_df.to_csv(result_path, index=None)
        return Response(
            response=json.dumps({
                'status': True,
                "message": "Generate Successful",
                }),
            status=201,
            mimetype='application/json'
        )
        
    except Exception as e:
        return Response(
            response=json.dumps({
                'status': False,
                "message": "Error Occured",
                "error": str(e)}),
            status=500,
            mimetype='application/json'
        )
    
        

        
