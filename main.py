import gspread
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from PIL import Image, ImageDraw, ImageFont
import io

# Define the Google Sheets credentials and sheet name
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials = None

try:
    credentials = service_account.Credentials.from_service_account_file('psyched-cubist-342516-bae8789627e1.json', scopes=scope)
except FileNotFoundError:
    print("Make sure to provide the correct path to your service account JSON file.")

gc = gspread.service_account(filename='psyched-cubist-342516-bae8789627e1.json')
sheet_name = 'Data form - instore'
worksheet_name = 'Website'

# Connect to the Google Sheet
worksheet = gc.open(sheet_name).worksheet(worksheet_name)

# Define the font and text to use
font_path = 'Banurasmie-RpaKE (1).otf'  # Replace with the path to your font file
font_size = 96  # Font size

# Define image properties
width = 500
height = 500
background_color = (255, 255, 255)  # White background (RGB)
text_color = (0, 0, 0)  # Black text (RGB)

# Define the font and text to use
font = ImageFont.truetype(font_path, font_size)

# Define the folder ID of the target folder in Google Drive
target_folder_id = '1oLKIYZPPWAYriAMp0FBW8QdR0_hKoAyG'

# Iterate over rows in column B starting from B2
for row_index in range(2, worksheet.row_count + 1):
    cell_value = worksheet.cell(row_index, 2).value  # Get value from column B
    if not cell_value:
        break  # Stop if no more values in column B

    # Define the coordinates for text placement
    x = (width - font.getsize(cell_value)[0]) / 2  # Center text horizontally
    y = (height - font.getsize(cell_value)[1]) / 2  # Center text vertically

    # Create an image with the specified font and text
    image = Image.new("RGB", (width, height), background_color)
    draw = ImageDraw.Draw(image)
    draw.text((x, y), cell_value, font=font, fill=text_color)

    # Save the image to a bytes buffer
    image_bytes = io.BytesIO()
    image.save(image_bytes, format='PNG')
    image_bytes.seek(0)

    # Upload the image to Google Drive in the specified folder
    drive_service = build('drive', 'v3', credentials=credentials)
    media = MediaIoBaseUpload(image_bytes, mimetype='image/png')
    file_metadata = {
        'name': f'{cell_value}.png',
        'parents': [target_folder_id]
    }

    # Create the file in the specified folder
    uploaded_file = drive_service.files().create(
        media_body=media,
        body=file_metadata
    ).execute()

    # Get the image URL
    image_url = f'https://drive.google.com/uc?export=view&id={uploaded_file["id"]}'

    # Update the Google Sheet with the image link in column H
    worksheet.update_cell(row_index, 8, image_url)  # Assuming column H is the 8th column (1-based index)
