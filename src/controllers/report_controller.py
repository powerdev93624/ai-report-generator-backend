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




# user controller blueprint to be registered with api blueprint
report = Blueprint("report", __name__)
original_doc_path = "src/files/sample/sample.docx"
output_doc_path = "src/files/result/result.docx"
@report.route('/generate', methods = ["POST"])
def generate():
    pythoncom.CoInitializeEx(0)
    scores = pd.read_csv("src/files/score/score.csv")
    doc = Document(original_doc_path)
    paragraphs = doc.paragraphs
    try: 
        data = request.json
        print(data)
        print(data["firstName"])
        with open('src/files/transcript/transcript.txt', 'r') as file:
            transcript = file.read()
        replace_full_name(paragraphs, data["firstName"]+" "+ data["lastName"])
        # print(1)
        replace_assessor_and_organization(paragraphs, data["assessor"], data["organization"])
        # print(2)
        replace_background(paragraphs, get_background(transcript))
        # print(3)
        replace_strategic_partner(paragraphs, get_strategic_partner(transcript))
        # print(4)
        replace_innovate(paragraphs, get_innovate(transcript))
        # print(5)
        replace_connect(paragraphs, get_connect(transcript))
        # print(6)
        replace_simplify(paragraphs, get_simplify(transcript))
        # print(7)
        replace_talent_enabler(paragraphs, get_talent_enabler(transcript))
        # print(8)
        replace_coach(paragraphs, get_coach(transcript))
        # print(9)
        replace_empower(paragraphs, get_empower(transcript))
        # print(10)
        replace_elevate(paragraphs, get_elevate(transcript))
        # print(11)
        replace_agile_executor(paragraphs, get_agile_executor(transcript))
        # print(12)
        replace_deliver(paragraphs, get_deliver(transcript))
        # print(13)
        replace_adapt(paragraphs, get_adapt(transcript))
        # print(14)
        replace_pioneer(paragraphs, get_pioneer(transcript))
        # print(15)
        replace_resilient_steward(paragraphs, get_resilient_steward(transcript))
        # print(16)
        replace_serve(paragraphs, get_serve(transcript))
        # print(17)
        replace_inspire(paragraphs, get_inspire(transcript))
        # print(18)
        replace_sustain(paragraphs, get_sustain(transcript))
        # print(19)
        replace_recommendation(paragraphs, transcript, scores)
        # print(20)
        replace_next_steps(paragraphs, get_next_steps(transcript))
        print(21)
        doc.save(output_doc_path)
        fill_score(scores, output_doc_path)
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
    
        

        
