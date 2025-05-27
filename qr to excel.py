import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import cv2
from pyzbar.pyzbar import decode
from PIL import Image, ImageTk
import openpyxl
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials # For service account
# If using OAuth2 for installed applications (more common for desktop apps):
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import threading # To prevent UI freezing during webcam scan
import pandas as pd # For easier Excel/Sheet header checking

# --- Configuration ---
APP_NAME = "QR Object Logger"
EXCEL_DEFAULT_FILENAME = "qr_scans.xlsx"
GSHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
                 'https://www.googleapis.com/auth/drive.file'] # Drive scope for finding sheets by name
GSHEET_CREDENTIALS_FILE = 'credentials.json' # Downloaded from Google Cloud Console
GSHEET_TOKEN_FILE = 'token.json' # Will be created after first authorization

class QRScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("600x700")

        self.scanned_data = None
        self.excel_file_path = EXCEL_DEFAULT_FILENAME
        self.gspread_client = None
        self.gspread_sheet_name = tk.StringVar(value="My QR Scans") # Default Google Sheet name

        # --- UI Elements ---
        # Style
        style = ttk.Style()
        style.configure("TButton", padding=6, relief="flat", font=('Helvetica', 10))
        style.configure("TLabel", padding=6, font=('Helvetica', 10))
        style.configure("TEntry", padding=6, font=('Helvetica', 10))
        style.configure("Header.TLabel", font=('Helvetica', 12, 'bold'))

        # Frame for QR actions
        qr_frame = ttk.LabelFrame(root, text="QR Code Actions", padding=10)
        qr_frame.pack(padx=10, pady=10, fill="x")

        self.btn_scan_webcam = ttk.Button(qr_frame, text="Scan from Webcam", command=self.start_webcam_scan_thread)
        self.btn_scan_webcam.pack(side=tk.LEFT, padx=5, expand=True, fill="x")

        self.btn_scan_file = ttk.Button(qr_frame, text="Scan from Desktop File", command=self.scan_from_file)
        self.btn_scan_file.pack(side=tk.LEFT, padx=5, expand=True, fill="x")

        # Frame for Scanned Data Display
        data_frame = ttk.LabelFrame(root, text="Scanned Data", padding=10)
        data_frame.pack(padx=10, pady=10, fill="x")

        self.lbl_qr_result = ttk.Label(data_frame, text="Scan a QR code...", wraplength=550)
        self.lbl_qr_result.pack(pady=5, fill="x")

        # Frame for Excel Options
        excel_frame = ttk.LabelFrame(root, text="Excel Options", padding=10)
        excel_frame.pack(padx=10, pady=10, fill="x")

        self.lbl_excel_path = ttk.Label(excel_frame, text=f"Excel Path: {os.path.abspath(self.excel_file_path)}")
        self.lbl_excel_path.pack(pady=5, fill="x")

        self.btn_choose_excel_path = ttk.Button(excel_frame, text="Change Excel Save Location", command=self.choose_excel_path)
        self.btn_choose_excel_path.pack(pady=5, fill="x")

        self.btn_save_excel = ttk.Button(excel_frame, text="Save to Excel", command=self.save_to_excel, state=tk.DISABLED)
        self.btn_save_excel.pack(pady=5, fill="x")

        # Frame for Google Sheets Options
        gsheet_frame = ttk.LabelFrame(root, text="Google Sheets Options", padding=10)
        gsheet_frame.pack(padx=10, pady=10, fill="x")

        ttk.Label(gsheet_frame, text="Google Sheet Name:").pack(side=tk.LEFT, padx=5)
        self.entry_gsheet_name = ttk.Entry(gsheet_frame, textvariable=self.gspread_sheet_name)
        self.entry_gsheet_name.pack(side=tk.LEFT, padx=5, expand=True, fill="x")

        self.btn_auth_gsheet = ttk.Button(gsheet_frame, text="Authenticate Google Sheets", command=self.authenticate_gsheets)
        self.btn_auth_gsheet.pack(pady=5, fill="x")
        
        self.btn_save_gsheet = ttk.Button(gsheet_frame, text="Save to Google Sheets", command=self.save_to_google_sheets, state=tk.DISABLED)
        self.btn_save_gsheet.pack(pady=5, fill="x")

        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.set_status("Ready.")

        # Webcam display area (optional, can be a separate window)
        self.webcam_window = None
        self.webcam_label = None
        self.cap = None
        self.stop_webcam_event = threading.Event()

    def set_status(self, message):
        self.status_var.set(message)
        print(f"STATUS: {message}") # Also print to console for debugging

    def _process_qr_content(self, qr_content_str):
        """
        Processes the raw QR content string.
        Expected format: "ObjectID,ObjectName"
        Returns: (object_id, object_name, timestamp_str) or None if format is wrong
        """
        try:
            parts = qr_content_str.split(',', 1)
            if len(parts) == 2:
                object_id = parts[0].strip()
                object_name = parts[1].strip()
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                return object_id, object_name, timestamp
            else:
                self.set_status(f"Error: QR content format is incorrect. Expected 'ID,Name'. Got: {qr_content_str}")
                messagebox.showerror("QR Error", f"Invalid QR data format. Expected 'ObjectID,ObjectName'.\nFound: '{qr_content_str}'")
                return None
        except Exception as e:
            self.set_status(f"Error processing QR content: {e}")
            messagebox.showerror("Processing Error", f"Could not process QR content: {e}")
            return None

    def update_ui_with_scan(self, object_id, object_name, timestamp):
        self.scanned_data = {
            "Object ID": object_id,
            "Name": object_name,
            "Timestamp": timestamp
        }
        display_text = (f"Successfully Scanned!\n"
                        f"ID: {object_id}\n"
                        f"Name: {object_name}\n"
                        f"Timestamp: {timestamp}")
        self.lbl_qr_result.config(text=display_text)
        self.btn_save_excel.config(state=tk.NORMAL)
        if self.gspread_client: # Only enable if authenticated
            self.btn_save_gsheet.config(state=tk.NORMAL)
        self.set_status("QR Code successfully scanned and parsed.")
        messagebox.showinfo("Scan Success", "QR Code successfully scanned!")


    def start_webcam_scan_thread(self):
        self.stop_webcam_event.clear()
        self.set_status("Starting webcam...")
        # Disable button to prevent multiple threads
        self.btn_scan_webcam.config(state=tk.DISABLED)
        threading.Thread(target=self.scan_from_webcam, daemon=True).start()

    def scan_from_webcam(self):
        self.cap = cv2.VideoCapture(0)
        if not self.cap.isOpened():
            self.set_status("Error: Cannot open webcam.")
            messagebox.showerror("Webcam Error", "Could not open webcam. Is it connected and not in use?")
            self.btn_scan_webcam.config(state=tk.NORMAL) # Re-enable button
            return

        # Create a Toplevel window for the webcam feed
        if self.webcam_window is None or not self.webcam_window.winfo_exists():
            self.webcam_window = tk.Toplevel(self.root)
            self.webcam_window.title("Webcam Feed (Press 'Q' to close)")
            self.webcam_label = ttk.Label(self.webcam_window)
            self.webcam_label.pack()
            self.webcam_window.protocol("WM_DELETE_WINDOW", self.stop_webcam_feed) # Handle window close
        
        self.set_status("Webcam active. Looking for QR code...")

        try:
            while not self.stop_webcam_event.is_set():
                ret, frame = self.cap.read()
                if not ret:
                    self.set_status("Error: Failed to capture frame from webcam.")
                    break
                
                # Convert frame to RGB for Pillow and then Tkinter
                cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                img = Image.fromarray(cv2image)
                imgtk = ImageTk.PhotoImage(image=img)
                
                if self.webcam_label and self.webcam_label.winfo_exists():
                    self.webcam_label.imgtk = imgtk # Keep a reference
                    self.webcam_label.configure(image=imgtk)
                    self.webcam_window.update_idletasks() # Force UI update

                decoded_objects = decode(frame)
                if decoded_objects:
                    qr_content = decoded_objects[0].data.decode('utf-8')
                    self.set_status(f"QR Detected: {qr_content}")
                    
                    processed_data = self._process_qr_content(qr_content)
                    if processed_data:
                        object_id, object_name, timestamp = processed_data
                        self.root.after(0, lambda: self.update_ui_with_scan(object_id, object_name, timestamp))
                        self.stop_webcam_feed() # Stop after successful scan
                        break 
                
                # Allow Tkinter to process events and check for 'q' key press in console
                # cv2.imshow alternative is handled by Tkinter window
                if cv2.waitKey(1) & 0xFF == ord('q'): # For console-based stop if needed
                     self.stop_webcam_event.set()

        except Exception as e:
            self.set_status(f"Webcam scanning error: {e}")
            messagebox.showerror("Webcam Error", f"An error occurred during webcam scanning: {e}")
        finally:
            self.stop_webcam_feed()
            self.btn_scan_webcam.config(state=tk.NORMAL) # Re-enable button

    def stop_webcam_feed(self):
        self.stop_webcam_event.set()
        if self.cap:
            self.cap.release()
            self.cap = None
        if self.webcam_window and self.webcam_window.winfo_exists():
            self.webcam_window.destroy()
            self.webcam_window = None
        cv2.destroyAllWindows() # Ensure all OpenCV windows are closed
        self.set_status("Webcam stopped.")


    def scan_from_file(self):
        file_path = filedialog.askopenfilename(
            title="Select QR Code Image",
            filetypes=(("PNG files", "*.png"), ("JPEG files", "*.jpg;*.jpeg"), ("All files", "*.*"))
        )
        if not file_path:
            self.set_status("File selection cancelled.")
            return

        try:
            img = Image.open(file_path)
            decoded_objects = decode(img)

            if decoded_objects:
                qr_content = decoded_objects[0].data.decode('utf-8')
                self.set_status(f"QR Detected in file: {qr_content}")
                processed_data = self._process_qr_content(qr_content)
                if processed_data:
                    object_id, object_name, timestamp = processed_data
                    self.update_ui_with_scan(object_id, object_name, timestamp)
            else:
                self.set_status(f"No QR code found in {os.path.basename(file_path)}.")
                messagebox.showinfo("Scan Result", f"No QR code found in the selected image: {os.path.basename(file_path)}.")
                self.lbl_qr_result.config(text="No QR code found in image.")
                self.scanned_data = None
                self.btn_save_excel.config(state=tk.DISABLED)
                self.btn_save_gsheet.config(state=tk.DISABLED)

        except Exception as e:
            self.set_status(f"Error reading image file: {e}")
            messagebox.showerror("File Error", f"Could not read or process image: {e}")


    def choose_excel_path(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=os.path.basename(self.excel_file_path),
            title="Save Excel File As"
        )
        if file_path:
            self.excel_file_path = file_path
            self.lbl_excel_path.config(text=f"Excel Path: {self.excel_file_path}")
            self.set_status(f"Excel save path set to: {self.excel_file_path}")
        else:
            self.set_status("Excel path selection cancelled.")

    def _get_excel_headers(self):
        return ["Object ID", "Name", "Timestamp"]

    def save_to_excel(self):
        if not self.scanned_data:
            messagebox.showwarning("No Data", "No data has been scanned yet.")
            return

        headers = self._get_excel_headers()
        row_data = [self.scanned_data[h] for h in headers] # Ensure correct order

        try:
            if os.path.exists(self.excel_file_path):
                workbook = openpyxl.load_workbook(self.excel_file_path)
                sheet = workbook.active
                # Optional: Check if headers match if file already exists
                # current_headers = [cell.value for cell in sheet[1]]
                # if current_headers != headers:
                #     messagebox.askyesno("Header Mismatch", "Excel headers differ. Overwrite headers?")
            else:
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                sheet.append(headers)

            sheet.append(row_data)
            workbook.save(self.excel_file_path)
            self.set_status(f"Data saved to Excel: {self.excel_file_path}")
            messagebox.showinfo("Excel Save", f"Data successfully saved to\n{self.excel_file_path}")

        except PermissionError:
            self.set_status(f"Error: Permission denied for {self.excel_file_path}. Is the file open?")
            messagebox.showerror("Excel Error", f"Permission denied. Please close the Excel file if it's open and try again.\nPath: {self.excel_file_path}")
        except Exception as e:
            self.set_status(f"Error saving to Excel: {e}")
            messagebox.showerror("Excel Error", f"Could not save to Excel: {e}")


    def authenticate_gsheets(self):
        creds = None
        self.set_status("Authenticating Google Sheets...")
        if not os.path.exists(GSHEET_CREDENTIALS_FILE):
            messagebox.showerror("Google Auth Error", 
                                 f"{GSHEET_CREDENTIALS_FILE} not found. "
                                 "Please download it from Google Cloud Console and place it in the application directory.")
            self.set_status(f"Error: {GSHEET_CREDENTIALS_FILE} not found.")
            return

        try:
            # The file token.json stores the user's access and refresh tokens, and is
            # created automatically when the authorization flow completes for the first time.
            if os.path.exists(GSHEET_TOKEN_FILE):
                creds = Credentials.from_authorized_user_file(GSHEET_TOKEN_FILE, GSHEET_SCOPES)
            
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    try:
                        creds.refresh(Request())
                    except Exception as e: # Catch broad exception for refresh failure
                        self.set_status(f"Token refresh failed: {e}. Re-initiating auth flow.")
                        creds = None # Force re-authentication
                if not creds: # creds is None or refresh failed
                    flow = InstalledAppFlow.from_client_secrets_file(GSHEET_CREDENTIALS_FILE, GSHEET_SCOPES)
                    # Specify a port or let it choose one automatically.
                    # Forcing a specific port can be useful if localhost isn't resolving correctly
                    # or if you need to set up specific redirect URIs.
                    # Example: creds = flow.run_local_server(port=0)
                    # Forcing port 8080 for example:
                    # creds = flow.run_local_server(host='localhost', port=8080, 
                    #                             authorization_prompt_message='Please visit this URL: {url}', 
                    #                             success_message='The auth flow is complete; you may close this window.',
                    #                             open_browser=True)
                    # If run_local_server causes issues (e.g. in environments without easy browser access or with strict firewalls):
                    # You might need to use run_console() but that's less user-friendly for a GUI app.
                    # Let's try with default run_local_server first.
                    messagebox.showinfo("Google Authentication", 
                                        "A browser window will open for Google Authentication. "
                                        "Please complete the sign-in process.")
                    creds = flow.run_local_server(port=0) # port=0 means pick an available port
                
                # Save the credentials for the next run
                with open(GSHEET_TOKEN_FILE, 'w') as token_file:
                    token_file.write(creds.to_json())
            
            self.gspread_client = gspread.authorize(creds)
            self.set_status("Google Sheets authenticated successfully.")
            messagebox.showinfo("Google Auth", "Successfully authenticated with Google Sheets.")
            if self.scanned_data: # Enable save button if data is present
                self.btn_save_gsheet.config(state=tk.NORMAL)

        # Alternative: Service Account (uncomment and adjust if using service account)
        # try:
        #     creds = ServiceAccountCredentials.from_json_keyfile_name(GSHEET_CREDENTIALS_FILE, GSHEET_SCOPES)
        #     self.gspread_client = gspread.authorize(creds)
        #     self.set_status("Google Sheets authenticated successfully (Service Account).")
        #     messagebox.showinfo("Google Auth", "Successfully authenticated with Google Sheets (Service Account).")
        #     if self.scanned_data: self.btn_save_gsheet.config(state=tk.NORMAL)
        except FileNotFoundError:
            self.set_status(f"Error: {GSHEET_CREDENTIALS_FILE} not found.")
            messagebox.showerror("Google Auth Error", f"Credentials file '{GSHEET_CREDENTIALS_FILE}' not found.")
        except Exception as e:
            self.gspread_client = None # Ensure client is None on failure
            self.set_status(f"Google Sheets authentication failed: {e}")
            messagebox.showerror("Google Auth Error", f"Authentication failed: {e}")


    def save_to_google_sheets(self):
        if not self.scanned_data:
            messagebox.showwarning("No Data", "No data has been scanned yet.")
            return
        if not self.gspread_client:
            messagebox.showerror("Google Sheets Error", "Not authenticated with Google Sheets. Please authenticate first.")
            self.set_status("Error: Not authenticated with Google Sheets.")
            return

        sheet_name_to_use = self.gspread_sheet_name.get()
        if not sheet_name_to_use:
            messagebox.showerror("Google Sheets Error", "Please enter a Google Sheet name.")
            self.set_status("Error: Google Sheet name is empty.")
            return

        headers = self._get_excel_headers() # Same headers
        row_data = [self.scanned_data[h] for h in headers]

        try:
            self.set_status(f"Accessing Google Sheet: {sheet_name_to_use}...")
            # Try to open the spreadsheet by name
            try:
                spreadsheet = self.gspread_client.open(sheet_name_to_use)
            except gspread.exceptions.SpreadsheetNotFound:
                self.set_status(f"Spreadsheet '{sheet_name_to_use}' not found. Creating it...")
                spreadsheet = self.gspread_client.create(sheet_name_to_use)
                # Share it with yourself if you created it via API and want to open in browser easily
                # Or if using service account, share with user's email
                # spreadsheet.share('your-email@example.com', perm_type='user', role='writer')
                self.set_status(f"Created and opened spreadsheet: {sheet_name_to_use}")


            # Try to get the first worksheet, or create one named 'Sheet1'
            try:
                worksheet = spreadsheet.sheet1 # Default first sheet
            except gspread.exceptions.WorksheetNotFound:
                self.set_status(f"Worksheet 'Sheet1' not found in '{sheet_name_to_use}'. Creating it...")
                worksheet = spreadsheet.add_worksheet(title="Sheet1", rows="100", cols="20") # Create with a default size
                self.set_status(f"Created worksheet 'Sheet1' in '{sheet_name_to_use}'.")
            
            # Check if headers exist
            # Using pandas to read the first row for simplicity, then gspread for writing
            # This avoids gspread rate limits if sheet is very large and we only need headers
            existing_headers = []
            try:
                # Get all values, which can be slow for large sheets.
                # A more optimized way is to get just the first row.
                first_row = worksheet.row_values(1)
                if first_row:
                    existing_headers = first_row
            except Exception as e: # Handles empty sheet or other gspread errors
                self.set_status(f"Could not read headers from sheet (may be empty): {e}")

            if not existing_headers or existing_headers != headers:
                # If sheet is completely empty (no first row) or headers don't match
                if not existing_headers: # Sheet is empty
                    worksheet.insert_row(headers, 1) # Insert headers at the first row
                    self.set_status("Added headers to new Google Sheet.")
                elif existing_headers != headers and all(h == '' for h in existing_headers): # First row is empty but exists
                    worksheet.update('A1', [headers]) # Update the first row with headers
                    self.set_status("Added headers to empty first row of Google Sheet.")
                # else:
                    # Headers exist but are different. Decide on a strategy:
                    # 1. Append anyway (might mess up columns)
                    # 2. Warn user
                    # 3. Create new sheet
                    # For now, we'll just append. A more robust solution would ask the user.
                    # self.set_status("Warning: Headers in Google Sheet do not match. Appending data anyway.")

            worksheet.append_row(row_data)
            self.set_status(f"Data saved to Google Sheet: {sheet_name_to_use} (Sheet1)")
            messagebox.showinfo("Google Sheets Save", f"Data successfully saved to Google Sheet:\n'{sheet_name_to_use}' (Sheet1)")

        except gspread.exceptions.APIError as e:
            self.set_status(f"Google Sheets API Error: {e.response.json()['error']['message']}")
            messagebox.showerror("Google Sheets API Error", f"API Error: {e.response.json()['error']['message']}")
        except Exception as e:
            self.set_status(f"Error saving to Google Sheets: {e}")
            messagebox.showerror("Google Sheets Error", f"Could not save to Google Sheets: {e}")

    def on_closing(self):
        """Handle window close event."""
        if self.cap:
            self.stop_webcam_feed()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = QRScannerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing) # Handle window close
    root.mainloop()