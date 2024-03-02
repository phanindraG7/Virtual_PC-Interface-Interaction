from flask import Flask, render_template,request,redirect,url_for
from your_code import run_volume,run_vmouse
from main import ppt_to_img,gesture_control_presentation
from threading import Thread
import os

app = Flask(__name__)

# Define the upload folder for PPT files
app.config['UPLOAD_FOLDER'] = 'uploads'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Define routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run_volume', methods=['GET', 'POST'])
def run_volume_route():
    run_volume()
    return render_template('index.html')

@app.route('/run_vmouse', methods=['GET', 'POST'])
def run_vmouse_route():
    run_vmouse()
    return render_template('index.html')

# @app.route('/run_presentation', methods=['GET', 'POST'])
# def run_presentation_route():
#     ppt_to_img()
@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/run_presentation', methods=['GET', 'POST'])
def run_presentation():
    return render_template('ppt.html')

# Route to handle PPT file upload and conversion
@app.route('/upload', methods=['GET','POST'])
def upload():
    if 'ppt_file' not in request.files:
        return redirect(url_for('index'))

    ppt_file = request.files['ppt_file']

    if ppt_file.filename == '':
        return redirect(url_for('index'))

    if ppt_file:
        ppt_path = os.path.join(app.config['UPLOAD_FOLDER'], ppt_file.filename)
        ppt_file.save(ppt_path)

        output_directory = os.path.join(app.config['UPLOAD_FOLDER'], 'output')
        ppt_to_img(ppt_path, output_directory)


        # return render_template('index.html', ppt_filename=ppt_file.filename, output_directory=output_directory)
        gesture_control_presentation(output_directory)

    return redirect(url_for('index'))


# Run the app
if __name__ == '__main__':
    app.run(debug=True)