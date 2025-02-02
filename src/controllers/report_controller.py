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
        
@report.route('/generate', methods = ["POST"])
def generate():
    pythoncom.CoInitializeEx(0)
    try: 
        data = request.json
        print(data)
        print(data["firstName"])
        with open('src/files/transcript/transcript.txt', 'r') as file:
            transcript = file.read()
        replace_full_name(data["firstName"]+" "+ data["lastName"])
        print(1)
        replace_assessor_and_organization(data["assessor"], data["organization"])
        print(2)
        replace_background(get_background(transcript))
        print(3)
        replace_strategic_partner(get_strategic_partner(transcript))
        print(4)
        replace_innovate(get_innovate(transcript))
        print(5)
        replace_connect(get_connect(transcript))
        print(6)
        replace_simplify(get_simplify(transcript))
        print(7)
        replace_talent_enabler(get_talent_enabler(transcript))
        print(8)
        replace_coach(get_coach(transcript))
        print(9)
        replace_empower(get_empower(transcript))
        print(10)
        replace_elevate(get_elevate(transcript))
        print(11)
        replace_agile_executor(get_agile_executor(transcript))
        print(12)
        replace_deliver(get_deliver(transcript))
        print(13)
        replace_adapt(get_adapt(transcript))
        print(14)
        replace_pioneer(get_pioneer(transcript))
        print(15)
        replace_resilient_steward(get_resilient_steward(transcript))
        print(16)
        replace_serve(get_serve(transcript))
        print(17)
        replace_inspire(get_inspire(transcript))
        print(18)
        replace_sustain(get_sustain(transcript))
        print(19)
        replace_recommendation(transcript)
        print(20)
        replace_next_steps(get_next_steps(transcript))
        print(21)
        fill_score()
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
    
        

        
