#importing libraries
from extract_txt import read_files
from txt_processing import preprocess
from txt_to_features import txt_features, feats_reduce
from extract_entities import get_number, get_email, rm_email, rm_number, get_name, get_skills
from model import simil
import pandas as pd
import json
import os
import uuid
from flask import Flask, flash, request, redirect, url_for, render_template, send_file
import csv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import PyPDF2


#used directories for data, downloading and uploading files 
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'files/resumes/')
DOWNLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'files/outputs/')
DATA_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Data/')

# Make directory if UPLOAD_FOLDER does not exist
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)

# Make directory if DOWNLOAD_FOLDER does not exist
if not os.path.isdir(DOWNLOAD_FOLDER):
    os.mkdir(DOWNLOAD_FOLDER)
#Flask app config 
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['DATA_FOLDER'] = DATA_FOLDER
app.config['SECRET_KEY'] = 'nani?!'

# Allowed extension you can set your own
ALLOWED_EXTENSIONS = set(['txt', 'pdf', 'doc','docx'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
 
@app.route('/', methods=['GET'])
def main_page():
    return _show_page()
 
@app.route('/', methods=['POST'])
def upload_folder():
    if request.method == 'POST':
        if 'pdf_folder' not in request.files:
            flash('No files part')
            return redirect(request.url)

        files = request.files.getlist('pdf_folder')
        if len(files) == 0:
            flash('No files selected')
            return redirect(request.url)

        # Get the list of uploaded file names
        uploaded_files = [file.filename for file in files if file.filename.endswith('.pdf')]

        # Process each uploaded file
        processed_texts = []
        for file in files:
            if file.filename == '':
                flash('No selected file')
                continue
            if file.filename.endswith('.pdf'):
                # Process the PDF file
                processed_text = process_pdf(file)
                processed_texts.append(processed_text)
        
        print("Uploaded Files:", uploaded_files)

        # Pass the processed_texts and uploaded_files list to the template
        return render_template('index.html', processed_texts=processed_texts, uploaded_files=uploaded_files)

    return render_template('index.html')
 
 
@app.route('/download/<code>', methods=['GET'])
def download(code):
    files = _get_files()
    if code in files:
        path = os.path.join(UPLOAD_FOLDER, code)
        if os.path.exists(path):
            return send_file(path)
    abort(404)

def process_folder(pdf_folder_path):
    """
    Process all PDF files in the specified folder.
    """
    processed_texts = []
    for filename in os.listdir(pdf_folder_path):
        if filename.endswith('.pdf'):
            pdf_file_path = os.path.join(pdf_folder_path, filename)
            processed_text = process_pdf(pdf_file_path)
            processed_texts.append(processed_text)
    return processed_texts

def process_pdf(file_content):
    # Create a PDF file reader object
    pdf_reader = PyPDF2.PdfReader(file_content)

    # Extract text from the PDF
    text = ""
    for page_number in range(len(pdf_reader.pages)):
        text += pdf_reader.pages[page_number].extract_text()

    return text


 
def _show_page():
    files = _get_files()
    return render_template('index.html', files=files)
 
def _get_files():
    file_list = os.path.join(UPLOAD_FOLDER, 'files.json')
    if os.path.exists(file_list):
        with open(file_list) as fh:
            return json.load(fh)
    return {}




# Initialize Flask-Mail
def clean_email(email):
    # Remove unwanted characters like square brackets
    return email.strip("[]'\"")

def send_email(recipient_emails, subject, message):
    for recipient_email in recipient_emails:
        clean_email_addr = clean_email(recipient_email)
        
        # Your email credentials
        sender_email = "tanishadhoot27@outlook.com"
        sender_password = "1437201@At"  # Use your app password here

        # Create message container
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = clean_email_addr
        msg['Subject'] = subject

        # Add message body
        msg.attach(MIMEText(message, 'plain'))

        try:
            # Connect to Outlook SMTP server
            server = smtplib.SMTP('smtp.office365.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)

            # Send email
            server.sendmail(sender_email, clean_email_addr, msg.as_string())
            print(f"Email sent successfully to {clean_email_addr}")

            # Close connection
            server.quit()

        except Exception as e:
            print(f"Error sending email to {clean_email_addr}: {e}")




@app.route('/process',methods=["POST"])
def process():
    if request.method == 'POST':

        rawtext = request.form['rawtext']
        jdtxt=[rawtext]
        resumetxt=read_files(UPLOAD_FOLDER)
        p_resumetxt = preprocess(resumetxt)
        p_jdtxt = preprocess(jdtxt)

        feats = txt_features(p_resumetxt, p_jdtxt)
        feats_red = feats_reduce(feats)

        df = simil(feats_red, p_resumetxt, p_jdtxt)

        t = pd.DataFrame({'Original Resume':resumetxt})
        dt = pd.concat([df,t],axis=1)

        dt['Phone No.']=dt['Original Resume'].apply(lambda x: get_number(x))
        
        dt['E-Mail ID']=dt['Original Resume'].apply(lambda x: get_email(x))

        dt['Original']=dt['Original Resume'].apply(lambda x: rm_number(x))
        dt['Original']=dt['Original'].apply(lambda x: rm_email(x))
        dt['Candidate\'s Name']=dt['Original'].apply(lambda x: get_name(x))

        skills = pd.read_csv(DATA_FOLDER+'skill_red.csv')
        skills = skills.values.flatten().tolist()
        skill = []
        for z in skills:
            r = z.lower()
            skill.append(r)

        dt['Skills']=dt['Original'].apply(lambda x: get_skills(x,skill))
        dt = dt.drop(columns=['Original','Original Resume'])
        sorted_dt = dt.sort_values(by=['JD 1'], ascending=False)

        sorted_dt_filtered = sorted_dt[sorted_dt['JD 1'] >= 0.001]

        out_path = DOWNLOAD_FOLDER+"Candidates.csv"
        sorted_dt_filtered.to_csv(out_path,index=False)

        top_rankers = []
        with open(out_path, 'r') as file:
            reader = csv.DictReader(file)
            sorted_rankers = sorted(reader, key=lambda row: float(row['JD 1']), reverse=True)
            for row in sorted_rankers[:3]:
                top_rankers.append(row['E-Mail ID'])  # Append only the email address

        subject = "Congratulations on Your Selection!"
        message_template = """
        Dear Student ,

        Congratulations! You have been selected as one of the top candidates for the position. We are excited to move forward with the hiring process. We will be in touch with you shortly with further details.

        Best regards,
        HR from CodeHerThing
        """

        # Send emails
        send_email(top_rankers, subject, message_template)

        return send_file(out_path, as_attachment=True)

    


if __name__=="__main__":
    app.run(port=8080, debug=False)