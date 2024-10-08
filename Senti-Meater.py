import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from PIL import Image, ImageTk
import requests
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import getpass
import logging
import threading

# Configure logging
logging.basicConfig(
    filename=os.path.join(os.path.expanduser("~"), "SentimentAnalyzer.log"),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Google Sheets API setup
SCOPE = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]

# Embedded Google API credentials (Replace placeholders with your actual credentials)
GOOGLE_CREDENTIALS_JSON = '''
{
  "type": "service_account",
    "project_id": "eng-serenity-436016-b2",
    "private_key_id": "6c4e959b22dd4c9a3700fbe7e87c00454192908e",
    "private_key": "-----BEGIN PRIVATE KEY-----\\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCMtelazdsXmJi3\\nbyqrkQq9J0dWdaucvcdBod1K6D0tt26B5PrX+pc/bpnuSyfAb3eE2A3UO9eHUOtq\\nUSKcqgvORpbbKSFEyIuQoOy3t2kthbfMmOUCeUek4Ii8TsSCaB3tFeA7XJ+UODkc\\nBjEUyppTTBaN6PQbCwfqm6J8E58pfC3Won2H3OrNWPEzu9hsrQPoVGXzhgP21EOO\\nUDTrJzwroFQfNWCtGhKy+7WBq7dGAxeoDFZ/b1+RZBmCQgy1Bs2zd2i9IaWZzxSD\\ngEOeW4gMTcPB7NhdYC5RdSjANZ5FCIQ7qQOwjKcssiTV2AhOekuunWEqNokQP4vG\\nnTUzshJdAgMBAAECggEAFk2X39nk7NALViP2OPyD8xMV68k5+t2v3CIqan72LNgN\\nhDmbCFSC2GiRAzmBYwSoBp6c+TNvswCWAqOP9Ly/+KGtXq6dxIJ0cL0OgwOSKxzY\\nqG2MUGs7JRH79REXym3ItIrKSyPKyC/DNEMc+3qDofQGV9cBU4SuQWG1jFeAhmQZ\\nc0MkoK4wwwrB6NNpm1da3WE0QjmJ3MogCRrEONHQAYkJuSgSAD6whTsBqJ5FqZLv\\nEa1lZbcyxcsX/WSMJ4diute0xqZLY3oppe0AD6PSYK193IpGEbsm7VesoKP6yce6\\nWN3FITmcuLFU0NkUKkSGOBD2fkdDxiAsYkeQZM0idwKBgQDD98PqycdgQXgjsNJy\\nZn84bd6bqkgOl8PeH4MaKeUuTaJ+PiXe88HmoHYlanFVg1o/SO/b6aHA4NQp28yA\\nC3v/iXFxP+vkkcK9jcOg/K/2r29bO0LV+wXSkjNtK4ME0ZPNr/5cNXb186KOmwv8\\nkA1Xw3WtZovBJZznrNa4LUqNWwKBgQC30MAGeu7iUbrJQhkDR+TE2fvHAY2KbH3I\\nrGuz2MWY0r2JVFSs0JqXCdaM8xr+ogDri8zfQQhwDPxHBtEwFI37bx56AzUBa4os\\nj7wibtccD3f4NINtFMm+Aw37CkfhIFv82FDBnsaijEdsFWoUqBGReqhxLpmn8HPg\\n3cOMHmpUpwKBgQCTLU2S1CBNBl54T6B+EsSRWNLLDkQ30XtlIz2PNM/OyrezIHHI\\n1EFYOEMDLsIXeyMYTGr4Oqsk9LXjChS4RefGry7n4x4C+AXN3t6B1cVB+9giKIu1\\nsWVaFDtTTk6EG/JplDfwgKbraSM4/vEtqfKba0zCAjYLxXfl90T75egL6QKBgE3r\\n+ltE5dufFfWXRY80fPBOEAOuztetYi0dmpKlBC7it2JuE28nB0Gb9A3QSNNEzesM\\nWo8RvIfzmUZqx2cAb6f01RCYJ3IwqmR1kiVuo1XL4OmhKU2mkFcyaEzRcOMompY3\\nBRTvP/lMSkKxWUTkcn4fZySDwrOEpTrgB7NweVblAoGACl+RKg1hbaSaVdNtTuk1\\n2V5eABWSRHh+kG9InkOXnvdKRXFd7ShyjatQ0WjFEkGMaaxIIZFXeCE0juffdH9o\\nP9y0m+9ZdFefVXYIOnobmKWL5rBNmzq6yiTY6QFlH5T8Bejb1fvoY0G/uANfjjLE\\ncs/r4ffjx+D0dLWkXzrGANY=\\n-----END PRIVATE KEY-----\\n",
  "client_email": "sheets-api-service-account-937@eng-serenity-436016-b2.iam.gserviceaccount.com",
  "client_id": "104639935010218564276",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/sheets-api-service-account-937%40eng-serenity-436016-b2.iam.gserviceaccount.com",
  "client_x509_cert_url": "googleapis.com"
}'''

# Spreadsheet ID (Replace with your actual spreadsheet ID)
SPREADSHEET_ID = '1IHh9c1uq8Ge-nf7S8wPQRYcyWh2LkQ1O2M7iisGUyQo'

def authorize_gsheets():
    try:
        creds = ServiceAccountCredentials.from_json_keyfile_dict(json.loads(GOOGLE_CREDENTIALS_JSON), SCOPE)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(SPREADSHEET_ID).sheet1  # Access the first sheet
        logging.info("Authorized Google Sheets successfully.")
        return sheet
    except Exception as e:
        logging.error(f"Error authorizing Google Sheets: {e}")
        sys.exit(1)

def save_mood_to_gsheets(user_id, selected_mood, date, time_str, work_related=None, reason=None):
    try:
        sheet = authorize_gsheets()
        mood_data = [user_id, selected_mood, date, time_str, work_related, reason]
        sheet.append_row(mood_data)
        logging.info("Sentiment saved successfully!")
    except Exception as e:
        logging.error(f"Error saving sentiment: {e}")

def load_image_from_url(url, size=(80, 90)):
    try:
        response = requests.get(url)
        response.raise_for_status()
        image_data = response.content
        img = Image.open(BytesIO(image_data)).convert("RGBA")
        img = img.resize(size, Image.LANCZOS)  # High-quality downscaling
        return ImageTk.PhotoImage(img)
    except Exception as e:
        logging.error(f"Error loading image from {url}: {e}")
        # Return a placeholder image if loading fails
        placeholder = Image.new('RGBA', size, (190, 190, 190, 250))  # Simple gray placeholder
        return ImageTk.PhotoImage(placeholder)

class MoodTrackerApp:
    def __init__(self, master):
        self.master = master
        self.user_id = getpass.getuser()  # Get the system's logged-in user ID
        self.selected_mood = None
        self.mood_buttons = {}  # Dictionary to store buttons for mood selection
        self.text_color = "black"  # Text color for mood buttons
        self.selected_font = ("Calibri", 12, "bold")  # Default font

        # Disable window close, minimize, and movement
        self.master.protocol("WM_DELETE_WINDOW", self.disable_close)
        self.master.attributes("-topmost", True)  # Keep the window always on top
        self.master.resizable(False, False)
        self.master.overrideredirect(True)  # Removes minimize/maximize buttons

        # Set window properties
        self.master.title("Senti-Meater")
        self.master.geometry("500x500")  # Adjusted window size
        self.master.configure(bg='#FFFFFF')  # Neutral white background

        # Create widgets
        self.create_widgets()

        # Make the window modal
        self.make_modal()

        # Schedule the first popup after a short delay
        # Adjust the delay as needed (e.g., for production, set it to trigger based on your requirements)
        self.master.after(25200000, self.show_popup)  # 25200000 ms = 7 Hour

    def disable_close(self):
        messagebox.showwarning("Warning", "You cannot close this window without submitting a response.")

    def make_modal(self):
        # Make the window modal by disabling other windows and focusing on this one
        self.master.grab_set()
        self.master.focus_force()

    def create_widgets(self):
        # Header Frame for Title and Logo
        header_frame = tk.Frame(self.master, bg='#FFFFFF')
        header_frame.pack(fill=tk.X, pady=5, padx=5)

        # Title Label
        title = tk.Label(
            header_frame, 
            text="How are you feeling today?", 
            font=("Calibri", 18, 'bold'), 
            bg='#FFFFFF', 
            fg='black'
        )
        title.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        # Load and display the logo
        logo_url = 'https://annovasolutions.com/wp-content/uploads/2024/09/Annova-logo.jpg'  # Update with your logo URL
        logo_img = load_image_from_url(logo_url, size=(200, 50))  # Adjust size as needed
        self.logo_label = tk.Label(header_frame, image=logo_img, bg='#FFFFFF')
        self.logo_label.image = logo_img  # Keep a reference to avoid garbage collection
        self.logo_label.pack(side=tk.RIGHT, anchor='ne')

        # Mood Buttons Frame
        moods_frame = tk.Frame(self.master, bg='#FFFFFF')
        moods_frame.pack(pady=5)

        # Moods list with image URLs
        self.moods = [
            ('Joyful', 'Joyful', 'https://i.pinimg.com/originals/a7/aa/2d/a7aa2d632e925e570eacbfc91bb3f156.png'),
            ('Happy', 'Happy', 'https://i.pinimg.com/originals/93/76/f4/9376f4bc2cf659688e4fe9887adddc4a.png'),
            ('Neutral', 'Neutral', 'https://cdn.shopify.com/s/files/1/1061/1924/products/Slightly_Smiling_Emoji_Icon_34f238ed-d557-4161-b966-779d8f37b1ac.png?v=1485577096'),
            ('Awful', 'Awful', 'https://cdn.shopify.com/s/files/1/1061/1924/files/Very_Angry_Emoji.png?9898922749706957214'),
            ('Demotivated', 'Demotivated', 'https://cdn.shopify.com/s/files/1/1061/1924/products/Confounded_Face_Emoji_Icon_ios10_grande.png?v=1542436041'),
            ('Sad', 'Sad', 'https://stickerly.pstatic.net/sticker_pack/wDjPkDnxxu1kgeFlKMkDQ/HFS88J/16/2302d2c4-a47d-4c7c-b3c0-24c1fd0ece2e.png')
        ]

        # Create buttons for mood selection
        for idx, mood in enumerate(self.moods):
            mood_image = load_image_from_url(mood[2], size=(70, 60))  # Load image from URL
            button = tk.Button(
                moods_frame, 
                text=mood[0], 
                image=mood_image, 
                compound=tk.TOP,  # Display image on top of text
                font=self.selected_font,  # Use the selected font
                bg='#FFFFFF',  # Match the app's background
                fg=self.text_color,  # Set initial button text color
                relief="raised",  # 3D effect
                borderwidth=0,  # Remove border for cleaner look
                activebackground='#1563CA',  # Subtle color change on button press
                command=lambda m=mood[1]: self.select_mood(m)
            )
            button.image = mood_image  # Keep a reference to avoid garbage collection
            self.mood_buttons[mood[1]] = button
            # Arrange buttons in a grid (2 rows x 3 columns)
            row = idx // 6  # 3 columns
            col = idx % 6
            button.grid(row=row, column=col, padx=0, pady=5)  # Added padding for better layout

        # Response Display Label
        self.response_label = tk.Label(self.master, text="", font=("Calibri", 12), bg='#FFFFFF', fg='black')
        self.response_label.pack(pady=(10, 0))

        # Work-related Frame (Initially hidden)
        self.work_related_frame = tk.Frame(self.master, bg='#FFFFFF')

        self.work_related_var = tk.StringVar()
        self.work_related_var.set('Select')  # Default option

        self.work_related_label = tk.Label(
            self.work_related_frame, 
            text="Is it work-related?", 
            font=("Calibri", 12, 'bold'), 
            bg='#FFFFFF', 
            fg='black'
        )
        self.work_related_label.pack(anchor='w', pady=(0, 5))

        self.work_related_dropdown = ttk.Combobox(
            self.work_related_frame, 
            textvariable=self.work_related_var, 
            values=["No", "Yes"],
            state="readonly"  # Make the dropdown read-only
        )
        self.work_related_dropdown.bind("<<ComboboxSelected>>", self.on_work_related_selected)
        self.work_related_dropdown.pack(anchor='w', pady=(0, 5))

        self.reason_label = tk.Label(self.work_related_frame, text="Reason:", font=("Calibri", 12, 'bold'), bg='#FFFFFF', fg='black')
        self.reason_textbox = tk.Text(self.work_related_frame, height=5, width=40)
        self.reason_label.pack(anchor='w')
        self.reason_textbox.pack(anchor='w', pady=(0, 5))

        # Disclaimer Note
        disclaimer = tk.Label(
            self.work_related_frame, 
            text="Please Note: Your responses will remain confidential.", 
            font=("Calibri", 10,'bold', "italic"), 
            fg="red", 
            bg='#FFFFFF'
        )
        disclaimer.pack(anchor='w', pady=(5, 0))

        # Submit Button
        self.submit_button = tk.Button(
            self.master, 
            text="Submit", 
            command=self.submit_response, 
            bg='#4CAF50', 
            fg='white', 
            font=("Calibri", 12, 'bold'), 
            activebackground='#45a049'
        )
        self.submit_button.pack(pady=10)

    def show_popup(self):
        # Reset previous selections
        self.selected_mood = None
        self.work_related_var.set('Select')  # Reset dropdown
        self.reason_textbox.delete('1.0', tk.END)  # Clear text box
        self.work_related_frame.pack_forget()  # Hide the work-related frame
        self.response_label.config(text="")  # Clear the response label

        # Make the window modal
        self.make_modal()

        # Show the main app window
        self.master.deiconify()
        self.master.lift()  # Bring window to top
        self.master.attributes("-topmost", True)  # Keep the window always on top

    def select_mood(self, mood):
        self.selected_mood = mood
        self.response_label.config(text=f"Selected Response: {self.selected_mood}")  # Display selected mood
        if mood in ['Awful', 'Demotivated', 'Sad']:  # Negative moods
            self.work_related_frame.pack(pady=2)  # Show work-related frame
        else:
            self.work_related_frame.pack_forget()  # Hide it if mood is not negative

    def on_work_related_selected(self, event):
        if self.work_related_var.get() == 'No':
            self.reason_textbox.delete('1.0', tk.END)  # Clear the reason text box
            self.reason_textbox.config(state=tk.DISABLED)  # Disable the reason text box
        else:
            self.reason_textbox.config(state=tk.NORMAL)  # Enable the reason text box

    def submit_response(self):
        if not self.selected_mood:
            messagebox.showwarning("Warning", "Please select a Response before submitting.")
            return

        work_related = self.work_related_var.get()

        # Check if work-related is "Yes" and reason is provided
        if self.selected_mood in ['Awful', 'Demotivated', 'Sad']:
            if work_related == 'Select':
                messagebox.showwarning("Warning", "Please select whether the response is work-related.")
                return

            if work_related == 'Yes':
                reason = self.reason_textbox.get('1.0', tk.END).strip()
                if not reason:
                    messagebox.showwarning("Warning", "Please provide a reason for the work-related response.")
                    return
            else:
                reason = ""  # No reason required if "No" is selected
        else:
            # Non-negative moods, no need for work-related info or reason
            work_related = ""
            reason = ""

        # Save the mood response
        date = datetime.now().strftime("%Y-%m-%d")
        time_str = datetime.now().strftime("%H:%M:%S")
        save_mood_to_gsheets(self.user_id, self.selected_mood, date, time_str, work_related, reason)

        # Show confirmation message
        messagebox.showinfo("Senti-Meater", "Your Sentiment has been submitted.\n\n!! Have a Great time Ahead !! ")

        # Hide the popup window
        self.master.withdraw()

        # Schedule the next popup after 7 Hour (252,00000 ms)
        self.master.after(25200000, self.show_popup)  # 25200000 ms = 7 Hour

def run_application():
    root = tk.Tk()
    app = MoodTrackerApp(root)
    root.mainloop()

if __name__ == "__main__":
    run_application()