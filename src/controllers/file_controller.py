from flask import request, Response, json, Blueprint, send_file
from src.models.user_model import User
from src import bcrypt, db
from datetime import datetime, timedelta, timezone
from src.middlewares import authentication_required
from werkzeug.utils import secure_filename
# import jwt
import os
import uuid


file = Blueprint("file", __name__)

# login api/auth/signin
@file.route('/transcript', methods = ["POST"])
def upload_transcript():
    try: 
        file = request.files['file']
        file.save(f"src/files/transcript/transcript.txt")
        return Response(
            response=json.dumps({
                    'status': True,
                    "message": "Upload Successful"
                }),
            status=200,
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
        
@file.route('/score', methods = ["POST"])
def upload_score():
    try: 
        file = request.files['file']
        file.save(f"src/files/score/score.csv")
        return Response(
            response=json.dumps({
                    'status': True,
                    "message": "Upload Successful"
                }),
            status=200,
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
        
@file.route('/download')
def download_file():
    file_path = 'files/result/result.docx'
    return send_file(file_path, as_attachment=True)

@file.route("/", methods=["GET"])
def delete_files():
    current_dir = os.getcwd()
    transcript_path = os.path.join(current_dir, 'src/files/transcript/transcript.txt')
    score_path = os.path.join(current_dir, 'src/files/score/score.csv')
    result_path = os.path.join(current_dir, 'src/files/result/result.docx')
    if os.path.exists(transcript_path):
        os.remove(transcript_path)
    if os.path.exists(score_path):
        os.remove(score_path)
    if os.path.exists(result_path):
        os.remove(result_path)
    return Response(
        response=json.dumps({
                'status': True,
                "message": "Delete Successful"
            }),
        status=200,
        mimetype='application/json'
    )