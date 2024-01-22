from flask import Flask, render_template, request
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from waitress import serve
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
from io import BytesIO

app = Flask(__name__)

# Azure Storage Account Configuration
connection_string = "DefaultEndpointsProtocol=https;AccountName=padamstorageaccount;AccountKey=0FcumCJ7pJkkmcK0oDQK0+3DowiSWXldPXnHHr2DQSi5wjtmoD/JsiHGV8qwQFJEaOQ4e4W80cOv+AStx1y7Gw==;EndpointSuffix=core.windows.net"
container_name = "padamcontainer"

# File names in Azure Storage
config_blob_name = "Configuration.xlsx"
depression_blob_name = "Depression.xlsx"
remedy_blob_name = "Remedy.xlsx"

# Names for core and secondary parameters
core_param_names = ["Depressed Mood", "Loss of interest and enjoyment", "Reduced energy leading to increased fatigability, tiredness"]
sec_param_names = ["Reduced concentration and attention", "Apprehension and worry", "Ideas of guilt",
                   "Bleak and pessimistic views of the future", "Ideas or acts of self-harm or suicide",
                   "Disturbed sleep", "Diminished appetite", "Unworthiness", "Loss of libido and sexual desires"]

# Gmail account credentials
sender_email = "padam.iit@gmail.com"
sender_password = "ojxn sjxs nfdv dmuw"

global_sum = 0


# Style for the frames
frame_style = {'bd': 5, 'relief': 'groove'}  # No background color for ScrolledText

def read_from_azure_storage(blob_name):
    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
    container_client = blob_service_client.get_container_client(container_name)
    blob_client = container_client.get_blob_client(blob_name)
    blob_data = blob_client.download_blob()
    df = pd.read_excel(BytesIO(blob_data.readall()))
    return df

#def write_to_azure_storage(df, blob_name):
#    blob_service_client = BlobServiceClient.from_connection_string(connection_string)
#    container_client = blob_service_client.get_container_client(container_name)
#    blob_client = container_client.get_blob_client(blob_name)
#    blob_client.upload_blob(df.to_excel(index=False), overwrite=True)

from io import BytesIO

def write_to_azure_storage(df, blob_name):
    try:
        # Convert DataFrame to Excel binary data
        excel_data = BytesIO()
        df.to_excel(excel_data, index=False)
        excel_data.seek(0)

        # Upload Excel binary data to Azure Blob Storage
        blob_service_client = BlobServiceClient.from_connection_string(connection_string)
        container_client = blob_service_client.get_container_client(container_name)
        blob_client = container_client.get_blob_client(blob_name)
        
        # Upload blob
        blob_client.upload_blob(excel_data.read(), overwrite=True)

        print(f"Blob {blob_name} uploaded successfully.")
    except Exception as e:
        print(f"Error uploading blob: {e}")


config_df = read_from_azure_storage(config_blob_name)
print (config_df)
# Recipient email
recipient_email = config_df.loc[0, "Value"]
#recipient_email = "padam.itbhu@gmail.com"

def find_value_from_id(input_type, col_name):
    df = read_from_azure_storage(remedy_blob_name)
    match_row = df[df['IndexVal'] == input_type]
    if not match_row.empty:
        remedy_value = match_row[col_name].iloc[0]
        return remedy_value
    else:
        return f"No value found for the given type: {input_type}"

def send_email(result_text):
    try:
        if config_df.loc[1, "Value"].lower() == "yes": #"TODO giving "MailOff as column name is not working here fix this
            print("Email sending is turned off.")
            return
        message = MIMEText(result_text)
        message["From"] = sender_email
        message["To"] = recipient_email
        message["Subject"] = "Depression Assessment Result"
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.sendmail(sender_email, recipient_email, message.as_string())
            print("Email sent successfully")
    except Exception as e:
        print(f"Error sending email: {e}")

def assess_depression(name, core_params, secondary_params):
    print("Patient Name:", name)
    print("Core Parameters:", core_params)
    print("Secondary Parameters:", secondary_params)

    core_present = sum(1 for level in core_params if level > 2)
    secondary_present = sum(1 for level in secondary_params if level > 0)
    core_levels = max(core_params)
    core_sum = sum(core_params)
    secondary_sum = sum(secondary_params)
    global global_sum
    total_sum = core_sum + secondary_sum
    global_sum = total_sum

    result = "A"
    if  total_sum <= 9:
        result = "A"
    elif  total_sum <= 18:
        result = "B"
    elif core_present == 0 and total_sum >= 19:
        result = "B"
    elif core_present >= 3 and total_sum >= 39 :
        result = "E"
    elif core_present >= 2 and total_sum >= 29:
        result = "D"
    elif core_present >= 1 and total_sum >= 19:
        result = "C"
    else:
        result = "F"

    return result


def submit_form(patient_name, core_params, secondary_params):
    result_id = assess_depression(patient_name, core_params, secondary_params)
    remedy = find_value_from_id(result_id, "Remedy")
    result_name = find_value_from_id(result_id, "Name")
    result_complete = f"Patient {patient_name} is having {result_name}. Suggested Remedy is to {remedy}. "
    send_email(result_complete)
    # Append to Excel file in Azure Storage
    basic_param_names = ['Name', 'Depression Type', 'Remedy']
    basic_params = [patient_name, result_name, remedy]
    sum_param_names = ['Sum']
    sum_params = [global_sum]
    data_excel = dict(zip(basic_param_names + core_param_names + sec_param_names + sum_param_names,
                          basic_params + core_params + secondary_params + sum_params))
    df = pd.DataFrame(data_excel, index=[0])
    existing_data = read_from_azure_storage(depression_blob_name)
    df = pd.concat([existing_data, df], ignore_index=True)
    write_to_azure_storage(df, depression_blob_name)
    return result_complete

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        patient_name = request.form['patient_name']
        core_params = [int(request.form[f'core_param_{i}']) for i in range(3)]
        secondary_params = [int(request.form[f'sec_param_{i}']) for i in range(9)]

        result = submit_form(patient_name, core_params, secondary_params)
        return render_template('result.html', result=result)

    core_param_names = ["Depressed Mood",
                        "Loss of interest and enjoyment",
                        "Reduced energy leading to increased fatigability, tiredness"
                        ]

    sec_param_names = [
        "Reduced concentration and attention",
        "Apprehension and worry",
        "Ideas of guilt",
        "Bleak and pessimistic views of the future",
        "Ideas or acts of self-harm or suicide",
        "Disturbed sleep",
        "Diminished appetite",
        "Unworthiness",
        "Loss of libido and sexual desires"
    ]

    # Use zip in the view function before passing to the template
    core_params_zipped = zip(range(len(core_param_names)), core_param_names)
    sec_params_zipped = zip(range(len(sec_param_names)), sec_param_names)

    return render_template('index.html', core_params=core_params_zipped, sec_params=sec_params_zipped)

if __name__ == '__main__':
    serve(app, host="0.0.0.0", port=8000)
