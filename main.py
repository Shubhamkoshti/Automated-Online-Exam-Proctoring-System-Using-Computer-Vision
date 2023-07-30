import cv2
import numpy as np
import time
import pyautogui
import tkinter as tk
from tkcalendar import DateEntry
from tkinter import ttk
import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
import requests
from flask import Flask, render_template, request

# Libraries for model
import os
import numpy as np
import pandas as pd
import tensorflow as tf
from tensorflow import keras
from tensorflow.keras import layers
import cv2
import matplotlib.pyplot as plt

app = Flask(__name__)

# Load the saved model
model = tf.keras.models.load_model('E:\\saved_model\\intership_240723_A.h5')

# OAuth scope for accessing Google Drive and Google Sheets
scope = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets']

# Google Drive link to the service account credentials JSON file
credentials_url = 'https://drive.google.com/uc?export=download&id=1jHmL1BhS6_zFqE1Y9FuMeJMwlWn8yB40'
# Replace 'YOUR_FILE_ID' with the actual file ID from the Google Drive link

# Create the service account credentials
response = requests.get(credentials_url)
response.raise_for_status()
credentials_content = response.json()

# Authorize the client using the credentials
credentials = service_account.Credentials.from_service_account_info(credentials_content, scopes=scope)

sheet_client = gspread.authorize(credentials)


@app.route('/create_folder', methods=['POST'])
def create_folder():
    roll_number = request.form['roll_number']
    date = request.form['date']
    exam_name = request.form['exam_name']
    email = request.form['email']

    if not roll_number or not date or not exam_name or not email:
        return "Please enter all the fields."

    # Update the existing spreadsheet with the input data
    spreadsheet_key = '1M_3qPiLJdE0n6zbx8LzkKLAR_LwJ3GJXRtWYgcVVTNg'  # Replace with the actual spreadsheet key
    spreadsheet = sheet_client.open_by_key(spreadsheet_key)
    sheet = spreadsheet.sheet1

    # Find the row index with the matching roll number and email
    roll_numbers = sheet.col_values(1)  # Assuming Roll Number is in the first column (column index 1)
    email_ids = sheet.col_values(4)  # Assuming Email ID is in the fourth column (column index 4)
    try:
        row_index = roll_numbers.index(roll_number) + 1  # Adding 1 to match the row index in the spreadsheet
        if email != email_ids[row_index - 1]:
            error_message = "Roll number is not associated with the provided email ID."
            # Continue with the operation and inform the user about the incorrect entry
    except ValueError:
        # Row with the matching roll number and email not found, create a new row
        row_index = len(roll_numbers) + 1
        sheet.append_row([roll_number, '', '', '', ''])  # Append an empty row for the new entry

        # Get the existing folder ID from the spreadsheet
        folder_id = sheet.cell(row_index, 5).value  # Assuming Folder ID is in the fifth column (column index 5)

        # Delete the existing folder if it exists
        if folder_id:
            drive_service = build('drive', 'v3', credentials=credentials)
            drive_service.files().delete(fileId=folder_id).execute()

    except ValueError:
        # Row with the matching roll number and email not found, create a new row
        row_index = len(roll_numbers) + 1
        sheet.append_row([roll_number, '', '', '', ''])  # Append an empty row for the new entry
        folder_id = None  # Set folder_id to None initially

    # Create a new folder
    folder_name = 'Roll Number ' + roll_number
    folder_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': ['1sXrdYQnkOu-2RdsjOfGUHk27K97YHHaV']  # Replace with the actual parent folder ID
    }
    drive_service = build('drive', 'v3', credentials=credentials)
    folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
    folder_id = folder.get('id')

    # Update the row with the new folder ID
    sheet.update_cell(row_index, 5, folder_id)  # Assuming Folder ID is in the fifth column (column index 5)

    # Update the row with the input data
    sheet.update_cell(row_index, 2, date)  # Assuming Date is in the second column (column index 2)
    sheet.update_cell(row_index, 3, exam_name)  # Assuming Exam Name is in the third column (column index 3)
    sheet.update_cell(row_index, 4, email)  # Assuming Email ID is in the fourth column (column index 4)

    # Call the capture_screen function
    capture_screen(folder_id)

    return "Folder created successfully. Data updated in the spreadsheet."


def capture_screen(folder_id):
    # Set the desired width and height of the captured screen
    width = 2000
    height = 1200

    # Set the time interval between screen captures (in seconds)
    interval = 3

    # Set the duration in seconds (1 hour and 30 minutes)
    duration = 120

    # Calculate the number of iterations based on duration and interval
    iterations = duration // interval

    # Run the loop for the specified duration
    for i in range(iterations):
        # Get the current timestamp
        timestamp = time.strftime('%Y%m%d_%H%M%S')

        # Capture the screen using pyautogui
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # Call the model to check if the screenshot is "ok"
        result = check_screenshot(model, screenshot)

        # Save the image with the timestamp to the specified folder on Google Drive if it's not "ok"
        if not result:
            save_screenshot_to_drive(screenshot, folder_id, timestamp)

        # Wait for the specified interval
        time.sleep(interval)


def check_screenshot(model, screenshot):
    # Preprocess the image for the model
    processed_image = preprocess_image(screenshot)

    # Make a prediction using the model
    prediction = model.predict(np.array([processed_image]))

    # Assuming the model returns a single value for "ok" or "not ok"
    # Adjust the condition based on your model's output format
    if prediction > 0.5 :
        return False  # Not ok
    else:
        return True  # Ok


def preprocess_image(image):
    # Resize the image to the required input shape for the model
    image = cv2.resize(image, (224, 224))

    # Preprocess the image according to your model's requirements (e.g., normalize pixel values)
    # You may need to adapt this code based on your model's preprocessing steps
    image = image.astype("float32") / 255.0

    return image


from googleapiclient.http import MediaInMemoryUpload

def save_screenshot_to_drive(screenshot, folder_id, timestamp):
    credentials = service_account.Credentials.from_service_account_info(credentials_content, scopes=scope)
    drive_service = build('drive', 'v3', credentials=credentials)

    # Convert the screenshot to bytes
    _, img_encoded = cv2.imencode('.png', screenshot)
    media_body = MediaInMemoryUpload(img_encoded.tobytes(), mimetype='image/png')

    # Create the file metadata
    file_metadata = {
        'name': f'screenshot_{timestamp}.png',
        'parents': [folder_id]
    }

    # Upload the file to Google Drive
    drive_service.files().create(body=file_metadata, media_body=media_body).execute()


@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    app.run(host='0.0.0.0')
