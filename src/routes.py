from flask import Blueprint
from src.controllers.auth_controller import auth
from src.controllers.file_controller import file
from src.controllers.report_controller import report

# main blueprint to be registered with application
api = Blueprint('api', __name__)

# register user with api blueprint
api.register_blueprint(auth, url_prefix="/auth")
api.register_blueprint(file, url_prefix="/file")
api.register_blueprint(report, url_prefix="/report")