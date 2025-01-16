import pandas as pd
import os
from flask import Flask, request, jsonify, render_template
import requests
import random
import string
import threading

# Flask App for URL Shortening
app = Flask(__name__)

# Define the file upload folder
UPLOAD_FOLDER = 'uploads/'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Character Dictionary for TH Code Decoding
dict_chars = [
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
    "A", "B", "C", "D", "E", "F", "G", "H", "J", "K",
    "M", "N", "P", "Q", "R", "S", "T", "U", "V", "W",
    "X", "Y", "Z",
    "A'", "B'", "C'", "D'", "E'", "F'", "G'", "H'", "J'", "K'",
    "M'", "N'", "P'", "Q'", "R'", "S'", "T'", "U'", "V'", "W'",
    "X'", "Y'", "Z'",
    "a'", "b'", "c'", "d'", "e'", "f'", "g'", "h'", "j'", "k'",
    "m'", "n'", "p'", "q'", "r'", "s'", "t'", "u'", "v'", "w'",
    "x'", "y'", "z'",
    "a", "b", "c", "d", "e", "f", "g", "h", "j", "k",
    "m", "n", "p", "q", "r", "s", "t", "u", "v", "w",
    "x", "y", "z"
]

def th_to_lat_lon(th_code):
    try:
        expanded_code = []
        for char in th_code:
            if char == "'":
                expanded_code[-1] += "'"
            else:
                expanded_code.append(char)

        numeric_code = []
        for char in expanded_code:
            if char in dict_chars:
                numeric_code.append(str(dict_chars.index(char)).zfill(2))
            else:
                raise ValueError(f"Invalid character '{char}' in TH code.")

        numeric_string = ''.join(numeric_code)
        lat_part = int(numeric_string[:8])
        lon_part = int(numeric_string[8:16])

        latitude = (lat_part - 9000000) / 100000.0
        longitude = (lon_part - 18000000) / 100000.0

        return latitude, longitude
    except Exception as e:
        print(f"Error decoding TH code '{th_code}': {e}")
        return None, None

def shorten_url_with_tinyurl(long_url, custom_alias):
    try:
        tinyurl_url = f"https://tinyurl.com/api-create.php?url={long_url}&alias={custom_alias}"
        response = requests.get(tinyurl_url)
        if response.status_code == 200:
            shortened_url = response.text
            return shortened_url.replace('http://', '').replace('https://', '')
        else:
            return "Error generating short URL"
    except Exception as e:
        return f"Error: {e}"

def generate_random_string(length=1):
    return ''.join(random.choices(string.ascii_lowercase, k=length))

@app.route('/shorten', methods=['POST'])
def shorten_url():
    try:
        data = request.get_json()
        long_url = data['url']
        custom_alias = data['alias'].replace("'", "")
        short_url = shorten_url_with_tinyurl(long_url, custom_alias)
        return jsonify({'shortened_url': short_url}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 400

def process_excel(file_path):
    try:
        data = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error reading the file: {e}")

    if 'Home Street' in data.columns:
        # Generate latitude and longitude
        lat_long_list = []
        for th_code in data['Home Street']:
            lat, lon = th_to_lat_lon(str(th_code))
            if lat is not None and lon is not None:
                lat_long_list.append(f"{lat}, {lon}")
            else:
                lat_long_list.append("Invalid TH Code")

        data['Lat-Long'] = lat_long_list

        # Generate Web Page URLs
        data['Web Page'] = data['Lat-Long'].apply(
            lambda x: f"https://www.google.com/maps/dir/?api=1&origin=My+Location&destination={x}" if "Invalid" not in x else "Invalid URL"
        )

        # Generate Short URLs
        def generate_shortened_url(row):
            if "Invalid" in row['Lat-Long']:
                return "Invalid Short URL"
            long_url = row['Web Page']
            custom_alias = row['Home Street'].replace("'", "")  # Remove apostrophes for custom alias
            return shorten_url_with_tinyurl(long_url, custom_alias)

        data['Short URL'] = data.apply(generate_shortened_url, axis=1)

        # Save updated Excel file
        data.to_excel(file_path, index=False)
        print(f"File processed successfully: {file_path}")
    else:
        raise ValueError("'Home Street' column not found in the uploaded Excel file.")

def save_uploaded_file(file):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    return file_path

@app.route('/')
def upload_file():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    file_path = save_uploaded_file(file)
    process_excel(file_path)
    return jsonify({'message': 'File processed successfully', 'file': file_path}), 200

# Run the Flask app
if __name__ == "__main__":
    threading.Thread(target=lambda: app.run(debug=True, use_reloader=False)).start()
