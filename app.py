from flask import Flask, render_template, request, redirect, url_for, session, send_file, jsonify
import os
import pandas as pd
from pymongo import MongoClient
from omr_script import process_omr
import json

app = Flask(__name__)
app.secret_key = 'your_secret_key'

# ---------------------------
#  MongoDB Connection (MERGED)
# ---------------------------
MONGO_URI = "mongodb+srv://cseetech123_db_user:CSEETECH@cluster0.wkwdqhd.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
client = MongoClient(MONGO_URI)
db = client['faculty']
users_collection = db['credentials']

# --------------------------
#  File storage folders
# --------------------------
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = os.path.join(UPLOAD_FOLDER, 'output')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER


# -------------------------
#  ROUTES
# -------------------------

@app.route('/')
def welcome():
    return render_template('welcomeLogin.html')


# -------------------------
#  LOGIN (MongoDB)
# -------------------------
@app.route('/login', methods=['POST'])
def login():
    username = request.form['username']
    password = request.form['password']

    user = users_collection.find_one({
        'username': username,
        'password': password
    })

    if user:
        session['logged_in'] = True
        return redirect(url_for('upload_form'))
    else:
        error = "Invalid credentials, please try again."
        return render_template('welcomeLogin.html', error=error)


# -------------------------------------
#  UPLOAD ROUTE (Merged Functionality)
# -------------------------------------
@app.route('/upload', methods=['GET', 'POST'])
def upload_form():
    if not session.get('logged_in'):
        return redirect(url_for('welcome'))

    if request.method == 'POST':
        teacher_name = request.form['teacher_name']
        subject = request.form['subject']
        num_of_questions = int(request.form['num_of_questions'])
        answer_key = request.form['answer_key'].split(',')

        # Validate answer key length
        if len(answer_key) != num_of_questions:
            error_message = (
                f"Error: The number of answers ({len(answer_key)}) "
                f"does not match the number of questions ({num_of_questions})."
            )
            return jsonify({'error': error_message})

        uploaded_file = request.files['pdf_file']

        if uploaded_file.filename != '':
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
            uploaded_file.save(pdf_path)

            # ------------------------------------------------------
            #  Write data.json (same as Node server.js)
            # ------------------------------------------------------
            json_data = {
                "input": {
                    "pdfFilePath": pdf_path,
                    "numQuestions": num_of_questions,
                    "answerKey": answer_key
                },
                "output": {}
            }

            with open("data.json", "w") as f:
                json.dump(json_data, f, indent=4)

            # ------------------------------------------------------
            #  Run OMR Processing (Option A)
            # ------------------------------------------------------
            try:
                process_omr(
                    pdf_path,
                    images_folder_path="images_output",
                    output_folder=OUTPUT_FOLDER,
                    answer_key=answer_key,
                    num_of_questions=num_of_questions
                )

                # Save output path for compatibility with Node.js version
                results_file = os.path.join(OUTPUT_FOLDER, "omr_results.xlsx")
                json_data["output"]["resultFilePath"] = results_file

                with open("data.json", "w") as f:
                    json.dump(json_data, f, indent=4)

                return redirect(url_for('results'))

            except Exception as e:
                return jsonify({"error": f"An error occurred during processing: {e}"})

    return render_template('index.html')


# -------------------------------
#  RESULTS PAGE
# -------------------------------
@app.route('/results')
def results():
    results_file = os.path.join(app.config['OUTPUT_FOLDER'], 'omr_results.xlsx')

    if not os.path.exists(results_file):
        return "<h1>Results not available yet. Please complete the OMR process.</h1>"

    results_df = pd.read_excel(results_file)
    results_html = results_df.to_html()

    return render_template('results.html',
                           results_html=results_html,
                           download_link=url_for('download_file'))


# -------------------------------
#  DOWNLOAD FILE
# -------------------------------
@app.route('/download')
def download_file():
    results_file = os.path.join(app.config['OUTPUT_FOLDER'], 'omr_results.xlsx')
    if os.path.exists(results_file):
        return send_file(results_file, as_attachment=True)
    else:
        return "<h1>File not found.</h1>"


# -------------------------------
#  LOGOUT
# -------------------------------
@app.route('/logout')
def logout():
    session.pop('logged_in', None)
    return redirect(url_for('welcome'))


# -------------------------------
#  RUN SERVER
# -------------------------------
if __name__ == "__main__":
    app.run(debug=True)
