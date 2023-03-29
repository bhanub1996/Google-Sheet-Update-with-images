import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image
from io import BytesIO

# set up credentials for Google Sheets API
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("path/to/credentials.json", scope)
client = gspread.authorize(creds)

# open the Google Sheet and select the worksheet
sheet = client.open("Sheet Name").sheet1

# create a dictionary with the data to insert
data = {
    "image": "path/to/image.png",
    "coordinates": "A1",
    "gcp_all_results": "result1, result2, result3",
    "gcp_final_label": "final_label",
    "gcp_final_label_score": "final_score"
}

# open the image file and convert to bytes
with open(data["image"], "rb") as f:
    img_bytes = BytesIO(f.read())

# open the image using the Pillow library and resize if needed
img = Image.open(img_bytes)
img.thumbnail((300, 300)) # adjust the size as needed
img_bytes = BytesIO()
img.save(img_bytes, format=img.format)
img_bytes.seek(0)

# upload the image to Google Drive and get the link
img_file = client.upload_media(img_bytes, mimetype=img.format)
img_link = img_file.get("webViewLink")

# insert the image into the worksheet
cell = sheet.cell(*sheet.find(data["coordinates"]).address)
cell.note = img_link

# insert the other data into the worksheet
for col, key in enumerate(data.keys(), start=2):
    cell = "{}{}".format(chr(ord("A") + col - 1), 1)
    sheet.update_cell(1, col, key)
    sheet.format("A1:Z1", {"backgroundColor": {"red": 0.8, "green": 0.8, "blue": 0.8}, "textFormat": {"bold": True}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP"})
    sheet.update_cell(2, col, data[key])
    sheet.format(cell, {"borders": {"top": {"style": "SOLID"}}})

# format the cell with the image
sheet.format(cell, {"verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP", "note": {"backgroundColor": {"red": 1, "green": 1, "blue": 1}, "textFormat": {"foregroundColor": {"red": 0, "green": 0, "blue": 0}}, "padding": {"top": 2, "bottom": 2, "left": 2, "right": 2}}})
