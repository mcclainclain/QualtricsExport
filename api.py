#!/usr/bin/env python
# coding: utf-8

# In[1]:


from customtkinter import CTkLabel as Label
import requests
import json
from zipfile import ZipFile
from io import BytesIO, TextIOWrapper
import pandas as pd
import time
import openpyxl
import csv
import calendar
from datetime import datetime
import os


def get_survey(name, id, start_date, end_date, api_key, base_url, path, ctk):

    # POST, start survey response export

    body = {
        "format": "csv",
        "startDate": start_date.strftime("%Y-%m-%dT00:00:00Z"),
        "endDate": end_date.strftime("%Y-%m-%dT23:59:59Z"),
        "timeZone": "America/Denver"
    }

    headers = {
        "Content-Type": "application/json",
        "X-API-TOKEN": api_key
    }

    url = base_url + id + "/export-responses"

    try:
        start = requests.post(url, json=body, headers=headers)
    except Exception as e:
        label = Label(ctk, text="Error: {}".format(e))
        label.grid(row=10, padx=10, pady=5)
        return

    # PROGRESS, get survey progress, make sure is done, and get progressId so we can get fileId

    start_json = json.loads(str(start.content, encoding="utf-8"))
    print(start_json)

    # Check for an error
    if start_json["meta"]["httpStatus"] != "200 - OK":
        label = Label(ctk, text="Error: {}".format(
            start_json["meta"]["error"]["errorMessage"]))
        label.grid(row = 11, padx=10, pady=5)
        return

    progressId = start_json["result"]["progressId"]

    url = base_url + id + "/export-responses/" + progressId

    # FILE, get response file using progressId

    pct_complete = 0.0

    while pct_complete < 100.0:
        time.sleep(1)
        progress = requests.get(url, headers=headers)
        progress_json = json.loads(progress.text)
        pct_complete = progress_json["result"]["percentComplete"]

    fileId = progress_json["result"]["fileId"]

    url = base_url + id + "/export-responses/" + fileId + "/file"

    file = requests.get(url, headers=headers, stream=True)

    # DATA CLEANING, use response zip, convert to csv, and make an excel file
    with ZipFile(BytesIO(file.content)) as filezip:
        files = filezip.namelist()
        with filezip.open(files[0], "r") as csv_file:
            wrap = TextIOWrapper(csv_file, encoding="utf-8")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = name if name != "DoIT-US-Techstore-Walk-in Helpdesk" else "DoIT-US-Techstore-Walk-in Helpd"
            reader = csv.reader(wrap)
            for row in reader:
                ws.append(row)

            wb.save(os.path.join(path, name + ".xlsx"))

    label = Label(ctk, text="{} has been saved".format(name))
    label.grid(row = 13, padx=10, pady=5)


def run(api_key, file_path, tk, start_date, end_date):

    # If the file path doesn't exist, throw an error
    if not os.path.exists(file_path):
        lb = Label(tk, text="The file path does not exist")
        lb.grid(row = 12, padx=10, pady=5)
        # Clear the label after 3 seconds
        tk.after(3000, lambda: lb.grid_forget())
        raise Exception("The file path does not exist")

    # Variables
    BASE_URL = "https://iad1.qualtrics.com/API/v3/surveys/"

    SURVEYS = [
        {
            "name": "DoIT-US-Helpdesk",
            "survey_id": "SV_3mkhDJSZKQWgBP7"
        },
        {
            "name": "DoIT-US-Departmental Support",
            "survey_id": "SV_bsHlaFE46Kiw9aB"
        },
        {
            "name": "DoIT-US-Techstore-Walk-in Helpdesk",
            "survey_id": "SV_e8VlH3aVWODKL4x"
        }
    ]

    start_date = datetime.strptime(start_date, "%m/%d/%Y")
    end_date = datetime.strptime(end_date, "%m/%d/%Y")

    print(start_date, end_date)

    folder_path = os.path.join(
        file_path, datetime.strftime(start_date, "%Y%m%d") + "-" + datetime.strftime(end_date, "%Y%m%d"))

    # Add folder if not there
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)

    for i in range(len(SURVEYS)):
        get_survey(SURVEYS[i]["name"], SURVEYS[i]["survey_id"], start_date, end_date,
                   api_key, BASE_URL, folder_path, tk)
