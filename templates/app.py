from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import threading
import os
import time

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def automate_process(file_path, options):
    chromedriver_path = "chromedriver.exe"
    driver = webdriver.Chrome(chromedriver_path)
    df = pd.read_excel(file_path)

    try:
        for index, row in df.iterrows():
            print(f"Processing row {index+1}")
            url = row['URL']
            next_year = datetime.now().year + 1
            value1 = row['admin']
            value2 = row['pass']

            # Define all the URLs and values here as in the original code

            driver.get(url)
            login = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "_58_login")))
            login.send_keys(value1)

            logout = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "_58_password")))
            logout.send_keys(value2)
            logout.send_keys(Keys.ENTER)

            # Continue with the rest of the automation steps based on the options
            # Check for options['sync_ttc'], options['sync_dvc'], etc.
            
            # Example for syncing common procedures (TTC)
            if options['sync_ttc']:
                driver.get(value3)
                dongbo_ttc = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "btnDongBo")))
                dongbo_ttc.click()
                time.sleep(15)

            # Add other automation steps similarly

        logout_url = urljoin(driver.current_url, "/c/portal/logout")
        driver.get(logout_url)

    finally:
        driver.quit()
        print("ĐÃ HOÀN TẤT !!!")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        return redirect(request.url)
    if file:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)
        
        options = {
            'sync_ttc': 'sync_ttc' in request.form,
            'sync_dvc': 'sync_dvc' in request.form,
            'sync_linhvuc': 'sync_linhvuc' in request.form,
            'config_holidays': 'config_holidays' in request.form
        }
        
        thread = threading.Thread(target=automate_process, args=(file_path, options))
        thread.start()
        
        return redirect(url_for('index'))

@app.route('/download')
def download_sample():
    sample_file_path = "data.xlsx"  # Path to your sample file
    return send_file(sample_file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
