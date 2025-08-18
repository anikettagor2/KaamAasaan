from handwritten_api import app as handwritten_app
from report_api import app as report_app
from fastapi import FastAPI

app = FastAPI()

# Mount both apps
app.mount("/handwritten", handwritten_app)
app.mount("/report", report_app)
