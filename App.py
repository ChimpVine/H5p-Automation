import csv
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from webdriver_manager.chrome import ChromeDriverManager
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from dotenv import load_dotenv

# Load environment variables from the .env file
load_dotenv()
# Access the variables
username = os.getenv("NAME")
password = os.getenv("PASSWORD")

# Function to authenticate with Google Drive API
def authenticate_google_drive():
    scopes = ["https://www.googleapis.com/auth/drive.file"]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    drive_service = build("drive", "v3", credentials=creds)
    return drive_service

# Function to check if a folder exists in Google Drive
def get_existing_folder_id(service, folder_name, parent_folder_id="1Vzky2BKAD7ReEFSB4I_z0cxqpFX6TvZZ"):
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder'"
    if parent_folder_id:
        query += f" and '{parent_folder_id}' in parents"
    
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    return None

# Function to create a folder in Google Drive if it doesn't exist
def create_or_get_drive_folder(service, folder_name, parent_folder_id="1Vzky2BKAD7ReEFSB4I_z0cxqpFX6TvZZ"):
    folder_id = get_existing_folder_id(service, folder_name, parent_folder_id)
    if folder_id:
        print(f"Folder '{folder_name}' already exists. ID: {folder_id}")
        return folder_id
    
    folder_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
    }
    if parent_folder_id:
        folder_metadata['parents'] = [parent_folder_id]

    folder = service.files().create(body=folder_metadata, fields='id').execute()
    print(f"Folder '{folder_name}' created. ID: {folder.get('id')}")
    return folder.get('id')

# Function to upload file to Google Drive
def upload_to_drive(service, file_path, folder_id):
    file_name = os.path.basename(file_path)
    file_metadata = {'name': file_name, 'parents': [folder_id]}
    
    media = MediaFileUpload(file_path, mimetype='application/octet-stream')
    uploaded_file = service.files().create(
        body=file_metadata, media_body=media, fields='id').execute()
    
    print(f"Uploaded: {file_name}, File ID: {uploaded_file.get('id')}")
    return f"https://drive.google.com/file/d/{uploaded_file.get('id')}/view?usp=sharing"

# Function to wait for the actual download to complete
def wait_for_download_complete(download_directory, timeout=60):
    start_time = time.time()
    while True:
        files = os.listdir(download_directory)
        h5p_files = [f for f in files if f.endswith('.h5p')]
        tmp_files = [f for f in files if f.endswith('.tmp')]

        if h5p_files and not tmp_files:
            print("Download complete.")
            return os.path.join(download_directory, h5p_files[0])

        if time.time() - start_time > timeout:
            raise TimeoutException("File download timed out after waiting for 60 seconds.")

        time.sleep(2)  # Wait for 2 seconds before checking again

# Function to safely delete a file with retries
def safe_delete_file(file_path, retries=3, delay=2):
    for attempt in range(retries):
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Deleted file from local directory: {file_path}")
                return True
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")
            time.sleep(delay)  # Wait before retrying
    print(f"Failed to delete file after {retries} attempts: {file_path}")
    return False

# Function to process CSV and download/upload H5P files
def process_csv_and_download(csv_file_path, google_drive_subject_folder_id, download_directory, drive_service):
    urls = []
    with open(csv_file_path, newline='') as csvfile:
        reader = csv.reader(csvfile)
        headers = next(reader)
        for row in reader:
            if len(row) > 2 and row[2].strip():
                url = row[2].strip()
                if not url.startswith(('http://', 'https://')):
                    print(f"Skipping invalid URL: {url}")
                    continue
                urls.append(row)

    results = []
    for row in urls:
        url = row[2]
        print(f"Opening URL: {url}")
        driver.get(url)
        time.sleep(10)

        try:
            driver.switch_to.frame(driver.find_element(By.TAG_NAME, 'iframe'))  
        except NoSuchElementException:
            print("No iframe found, proceeding without switching.")

        try:
            reuse_button = WebDriverWait(driver, 60).until(
                EC.element_to_be_clickable((By.XPATH, '//button[@aria-label="Reuse this content."]'))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", reuse_button)
            print("Button found and scrolled into view.")
            driver.execute_script("arguments[0].click();", reuse_button)
            print(f"Button clicked successfully for URL: {url}")

            try:
                download_button = WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[@class="h5p-big-button h5p-download-button"]'))
                )
                download_button.click()
                print("Download button clicked successfully.")

                downloaded_file_path = wait_for_download_complete(download_directory)
                print(f"Downloaded file: {downloaded_file_path}")
                
                # Upload the file to Google Drive
                drive_link = upload_to_drive(drive_service, downloaded_file_path, google_drive_subject_folder_id)
                print(f"File uploaded to Google Drive: {drive_link}")
                results.append(row + ["Downloaded", downloaded_file_path, drive_link])

                # Safely delete the file from the local download folder after successful upload
                safe_delete_file(downloaded_file_path)

            except TimeoutException:
                print("Timeout: The download button was not found or clickable within the specified time.")
                results.append(row + ["Error Downloading, Please work on this manually", "N/A", ""])
            except NoSuchElementException:
                print("Error: The download button was not found.")
                results.append(row + ["Error Downloading, Please work on this manually", "N/A", ""])
            except ElementClickInterceptedException:
                print("Error: The download button could not be clicked, possibly due to an overlay or another element blocking it.")
                results.append(row + ["Error Downloading, Please work on this manually", "N/A", ""])
            except Exception as e:
                print(f"An unexpected error occurred while trying to click the download button - {e}")
                results.append(row + ["Error Downloading, Please work on this manually", "N/A", ""])

            time.sleep(20)

        except TimeoutException:
            print(f"Timeout: The 'Reuse this content' button was not found within the specified time for URL: {url}.")
            results.append(row + ["Error Downloading, Please work on this manually", "N/A", ""])

    output_csv_file_path = 'output_with_status.csv'
    with open(output_csv_file_path, mode='w', newline='') as outfile:
        writer = csv.writer(outfile)
        writer.writerow(headers + ["STATUS", "File Path", "Drive Link"])
        writer.writerows(results)

    print(f"Results written to {output_csv_file_path}")
    upload_to_drive(drive_service, output_csv_file_path, google_drive_subject_folder_id)

# Iterate through local output_files folder and get all CSV files
def process_local_output_files_folder(drive_service):
    base_dir = os.getcwd()
    output_files_dir = os.path.join(base_dir, "output_files")
    
    for root, dirs, files in os.walk(output_files_dir):
        for file in files:
            if file.endswith('.csv'):
                csv_file_path = os.path.join(root, file)
                
                relative_path = os.path.relpath(root, output_files_dir)
                grade, subject = relative_path.split(os.sep)

                google_drive_grade_folder_id = create_or_get_drive_folder(drive_service, grade)
                google_drive_subject_folder_id = create_or_get_drive_folder(drive_service, subject, parent_folder_id=google_drive_grade_folder_id)

                download_directory = os.path.join(base_dir, "downloads")
                os.makedirs(download_directory, exist_ok=True)

                process_csv_and_download(csv_file_path, google_drive_subject_folder_id, download_directory, drive_service)

# Configure Chrome Options for custom download directory
chrome_options = Options()
prefs = {
    "download.default_directory": os.path.join(os.getcwd(), "downloads"),
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

# Chrome Driver setup with options
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Main execution flow
if __name__ == '__main__':
    driver.get('https://np.chimpvine.com/admin')
    time.sleep(5)

    # Fill the login form
    username_input = driver.find_element(By.ID, 'login_username')
    password_input = driver.find_element(By.ID, 'login_password')
    username_input.send_keys(username)
    password_input.send_keys(password)
    login_button = driver.find_element(By.XPATH, '//button[@type="submit"]')
    login_button.click()

    time.sleep(10)

    if "admin" in driver.current_url:
        print("Logged in successfully.")
        drive_service = authenticate_google_drive()  # Authenticate once here
        process_local_output_files_folder(drive_service)
    else:
        print("Login failed. Check credentials or URL.")
    
    driver.quit()
